import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Geo, Bar, Page
from pyecharts.globals import ThemeType
from pyecharts.commons.utils import JsCode
import re
import os  # 用于调用系统命令

# 1. 设置文件路径
# 使用原始字符串处理 Windows 路径
file_path = r"D:\2026\03 小象BI\01 商品明细_2026-04-07.xlsx"
output_file = "xiaoxiang_bi_fixed.html"

# 2. 读取并清洗数据
try:
    # 自动识别是 Excel 还是 CSV
    if file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path)
    else:
        df = pd.read_csv(file_path)

    # A. 确保销售额是数字，把无法识别的变成0
    df['商品销售额'] = pd.to_numeric(df['商品销售额'], errors='coerce').fillna(0)

    # B. 清洗城市名称（防止括号后缀导致地图识别失败）
    # 将“广州（含佛山）”等统一清洗为“广州”
    df['城市_清洗'] = df['城市'].astype(str).apply(lambda x: re.sub(r'（.*?）', '', x))

    # C. 【关键修正：数据聚合】
    # 将同一个城市的多行数据进行求和，这样每个城市只会剩下一个总销售额数字，避免提示框数值堆叠
    df_summed = df.groupby('城市_清洗')['商品销售额'].sum().round(2).reset_index()

except Exception as e:
    print(f"数据读取或处理出错: {e}")
    exit()

# 3. 准备地图数据（排除“全国”，仅展示具体城市分布）
city_only = df_summed[df_summed['城市_清洗'] != '全国'].sort_values(by='商品销售额', ascending=False)
map_data = [list(z) for z in zip(city_only['城市_清洗'], city_only['商品销售额'])]

# 4. 准备排行榜数据（包含“全国”在内的销售额前15名）
rank_data = df_summed.sort_values(by='商品销售额', ascending=False).head(15)

# --- 绘图部分 ---

# 地图配置
geo = (
    Geo(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="500px"))
    .add_schema(maptype="china")
    .add("总销售额(元)", map_data, type_="scatter")
    .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
    .set_global_opts(
        title_opts=opts.TitleOpts(title="全国城市销售分布图（已聚合求和）"),
        visualmap_opts=opts.VisualMapOpts(max_=city_only['商品销售额'].max(), is_piecewise=False),
        tooltip_opts=opts.TooltipOpts(
            # 使用 JsCode 精确获取聚合后的数值，params.value[2] 对应 Geo 数据的数值部分
            formatter=JsCode("function(params){return params.name + ': ' + params.value[2] + ' 元';}")
        )
    )
)

# 排行榜配置
bar = (
    Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="500px"))
    .add_xaxis(rank_data['城市_清洗'].tolist()[::-1])
    .add_yaxis("总销售额", rank_data['商品销售额'].tolist()[::-1])
    .reversal_axis() # 轴转置，方便阅读城市名称
    .set_series_opts(label_opts=opts.LabelOpts(position="right"))
    .set_global_opts(
        title_opts=opts.TitleOpts(title="销售额排行榜 (Top 15)"),
        xaxis_opts=opts.AxisOpts(name="金额"),
        yaxis_opts=opts.AxisOpts(is_show=True)
    )
)

# 合并保存并渲染
page = Page(layout=Page.SimplePageLayout)
page.add(geo, bar)
page.render(output_file)

print(f"修正后的看板已生成：{output_file}")

# 5. 自动使用 Microsoft Edge 打开生成的 HTML 文件
# 使用 start msedge 命令调用浏览器
try:
    os.system(f'start msedge "{os.path.abspath(output_file)}"')
    print("已尝试调用 Microsoft Edge 打开看板。")
except Exception as e:
    print(f"自动打开浏览器失败: {e}")