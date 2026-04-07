import json
import os
import re

import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar, Geo, Map, Page
from pyecharts.commons.utils import JsCode
from pyecharts.globals import ThemeType

# ── 路径配置 ──────────────────────────────────────────────
BASE_DIR    = r"D:\2026\03 小象BI"
file_path   = os.path.join(BASE_DIR, "01 商品明细_2026-04-07.xlsx")
mapping_file = os.path.join(BASE_DIR, "01b 映射.json")
output_file  = "xiaoxiang_bi_fixed.html"

# ── 加载映射表 ────────────────────────────────────────────
with open(mapping_file, encoding="utf-8") as f:
    _mapping = json.load(f)
city_to_province   = _mapping["city_to_province"]
province_full_names = _mapping["province_full_names"]

def to_echarts_province(name):
    """将短省名转为 ECharts 地图全称（如 广东→广东省）"""
    return province_full_names.get(name, name + "省")

# ── 读取与清洗数据 ────────────────────────────────────────
try:
    df = pd.read_excel(file_path) if file_path.endswith(".xlsx") else pd.read_csv(file_path)

    df["商品销售额"] = pd.to_numeric(df["商品销售额"], errors="coerce").fillna(0)
    df["城市_清洗"]  = df["城市"].astype(str).apply(lambda x: re.sub(r"（.*?）", "", x))

    # 城市级聚合
    df_city = df.groupby("城市_清洗")["商品销售额"].sum().round(2).reset_index()

    # 省份级聚合：去掉末尾"市"防止误截，再用映射表推导省份
    df_city["城市_标准"] = df_city["城市_清洗"].apply(lambda x: re.sub(r"市$", "", x))
    df_city["省份"]      = df_city["城市_标准"].map(city_to_province)
    df_province = (
        df_city.dropna(subset=["省份"])
               .groupby("省份")["商品销售额"]
               .sum().round(2).reset_index()
               .sort_values("商品销售额", ascending=False)
    )

except Exception as e:
    print(f"数据读取或处理出错: {e}")
    exit()

# ── 准备绘图数据 ──────────────────────────────────────────
# 省份
province_map_data  = [[to_echarts_province(p), float(v)]
                      for p, v in zip(df_province["省份"], df_province["商品销售额"])]
province_rank_data = df_province.head(15)

# 城市（排除"全国"）
city_only      = df_city[df_city["城市_清洗"] != "全国"].sort_values("商品销售额", ascending=False)
city_map_data  = [list(z) for z in zip(city_only["城市_清洗"], city_only["商品销售额"])]
city_rank_data = df_city.sort_values("商品销售额", ascending=False).head(15)

# ── 图表：省份地图 ────────────────────────────────────────
chart_province_map = (
    Map(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="500px"))
    .add("省份总销售额(元)", province_map_data, "china")
    .set_global_opts(
        title_opts=opts.TitleOpts(title="省份销售额分布"),
        visualmap_opts=opts.VisualMapOpts(max_=float(df_province["商品销售额"].max()), is_piecewise=False),
        tooltip_opts=opts.TooltipOpts(formatter="{b}: {c} 元"),
    )
)

# ── 图表：省份排行榜 ──────────────────────────────────────
chart_province_bar = (
    Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="500px"))
    .add_xaxis(province_rank_data["省份"].tolist()[::-1])
    .add_yaxis("总销售额", province_rank_data["商品销售额"].tolist()[::-1])
    .reversal_axis()
    .set_series_opts(label_opts=opts.LabelOpts(position="right"))
    .set_global_opts(
        title_opts=opts.TitleOpts(title="省份销售额排行榜 (Top 15)"),
        xaxis_opts=opts.AxisOpts(name="金额"),
    )
)

# ── 图表：城市地图 ────────────────────────────────────────
chart_city_map = (
    Geo(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="500px"))
    .add_schema(maptype="china")
    .add("总销售额(元)", city_map_data, type_="scatter")
    .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
    .set_global_opts(
        title_opts=opts.TitleOpts(title="城市销售分布图"),
        visualmap_opts=opts.VisualMapOpts(max_=float(city_only["商品销售额"].max()), is_piecewise=False),
        tooltip_opts=opts.TooltipOpts(
            formatter=JsCode("function(params){return params.name + ': ' + params.value[2] + ' 元';}")
        ),
    )
)

# ── 图表：城市排行榜 ──────────────────────────────────────
chart_city_bar = (
    Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="500px"))
    .add_xaxis(city_rank_data["城市_清洗"].tolist()[::-1])
    .add_yaxis("总销售额", city_rank_data["商品销售额"].tolist()[::-1])
    .reversal_axis()
    .set_series_opts(label_opts=opts.LabelOpts(position="right"))
    .set_global_opts(
        title_opts=opts.TitleOpts(title="城市销售额排行榜 (Top 15)"),
        xaxis_opts=opts.AxisOpts(name="金额"),
    )
)

# ── 渲染输出 ──────────────────────────────────────────────
page = Page(layout=Page.SimplePageLayout)
page.add(chart_province_map, chart_province_bar, chart_city_map, chart_city_bar)
page.render(output_file)
print(f"看板已生成：{output_file}")

os.system(f'start msedge "{os.path.abspath(output_file)}"')
