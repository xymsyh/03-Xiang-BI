import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Geo, Bar, Page
from pyecharts.globals import ThemeType
from pyecharts.commons.utils import JsCode
import re

# 1. 设置文件路径
file_path = r"D:\2026\03 小象BI\01 商品明细_2026-04-07.xlsx"

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
    # 将同一个城市的多行数据进行求和，这样广州就只会剩下一个总销售额数字
    df_summed = df.groupby('城市_清洗')['商品销售额'].sum().round(2).reset_index()

except Exception as e:
    print(f"数据读取或处理出错: {e}")
    exit()

# 3. 准备地图数据（排除“全国”，仅展示城市）
city_only = df_summed[df_summed['城市_清洗'] != '全国'].sort_values(by='商品销售额', ascending=False)
map_data = [list(z) for z in zip(city_only['城市_清洗'], city_only['商品销售额'])]

# 4. 准备排行榜数据（包含“全国”在内的前15名）
rank_data = df_summed.sort_values(by='商品销售额', ascending=False).head(15)

# --- 绘图部分 ---

# 地图
geo = (
    Geo(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="500px"))
    .add_schema(maptype="china")
    .add("总销售额(元)", map_data, type_="scatter")
    .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
    .set_global_opts(
        title_opts=opts.TitleOpts(title="全国城市销售分布图（已聚合求和）"),
        visualmap_opts=opts.VisualMapOpts(max_=city_only['商品销售额'].max(), is_piecewise=False),
        tooltip_opts=opts.TooltipOpts(
            formatter=JsCode("function(params){return params.name + ': ' + params.value[2] + ' 元';}")
        )
    )
)

# 排行榜
bar = (
    Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="500px"))
    .add_xaxis(rank_data['城市_清洗'].tolist()[::-1])
    .add_yaxis("总销售额", rank_data['商品销售额'].tolist()[::-1])
    .reversal_axis()
    .set_series_opts(label_opts=opts.LabelOpts(position="right"))
    .set_global_opts(
        title_opts=opts.TitleOpts(title="销售额排行榜 (Top 15)"),
        xaxis_opts=opts.AxisOpts(name="金额"),
        yaxis_opts=opts.AxisOpts(is_show=True)
    )
)

# 合并保存
page = Page(layout=Page.SimplePageLayout)
page.add(geo, bar)
page.render("xiaoxiang_bi_fixed.html")

print("修正后的看板已生成：xiaoxiang_bi_fixed.html")
print("提示：现在广州等城市应只显示一个唯一的总数。")