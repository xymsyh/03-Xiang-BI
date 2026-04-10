import glob
import json
import os
import re

import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar, Geo, Line, Map, Page
from pyecharts.commons.utils import JsCode
from pyecharts.globals import ThemeType

# ── 路径配置 ──────────────────────────────────────────────
BASE_DIR     = r"D:\2026\03 小象BI"
INPUT_DIR    = os.path.join(BASE_DIR, "02 生成")
mapping_file = os.path.join(BASE_DIR, "01b 映射.json")

# ── 加载映射表 ────────────────────────────────────────────
with open(mapping_file, encoding="utf-8") as f:
    _mapping = json.load(f)
city_to_province    = _mapping["city_to_province"]
province_full_names = _mapping["province_full_names"]

def to_echarts_province(name):
    """将短省名转为 ECharts 地图全称（如 广东→广东省）"""
    return province_full_names.get(name, name + "省")


def format_number(value):
    """将数字格式化：>=10000用w，>=1000用k，<1000直接显示，保留1位小数"""
    value = float(value)
    if value >= 10000:
        formatted = round(value / 10000, 1)
        return f"{formatted:.1f}w"
    elif value >= 1000:
        formatted = round(value / 1000, 1)
        return f"{formatted:.1f}k"
    else:
        return str(int(value))


def process_file(file_path, output_dir):
    """处理单个 Excel 文件，为每个商品生成一个 HTML"""
    # ── 读取与清洗数据 ────────────────────────────────────────
    df = pd.read_excel(file_path) if file_path.endswith(".xlsx") else pd.read_csv(file_path)

    # 重命名列，处理乱码问题
    df.columns = ['日期', '商品ID', '商品名称', '商品69码', '城市',
                  '商品销售额', '商品销售量', '商品预订数量',
                  '供应商到大仓在途数量', '大仓库存数量', '大仓到门店在途数量',
                  '前置站点库存数量', '门店商品订单流失率', '门店商品门店订单流失率待时',
                  '门店商品缺货天数', '消库库存清零预计周期日期', '消库库存清零周期',
                  '大宗单外部销售率']

    df["商品销售额"] = pd.to_numeric(df["商品销售额"], errors="coerce").fillna(0)
    df["商品销售量"] = pd.to_numeric(df["商品销售量"], errors="coerce").fillna(0)
    # 有销售额但销售量为0时，视作销售量为1（原始数据销量为1时偶尔不计入统计）
    df.loc[(df["商品销售额"] > 0) & (df["商品销售量"] == 0), "商品销售量"] = 1
    df["总库存"]     = (
        pd.to_numeric(df["供应商到大仓在途数量"], errors="coerce").fillna(0) +
        pd.to_numeric(df["大仓库存数量"],         errors="coerce").fillna(0) +
        pd.to_numeric(df["大仓到门店在途数量"],   errors="coerce").fillna(0) +
        pd.to_numeric(df["前置站点库存数量"],     errors="coerce").fillna(0)
    )
    # 库存为快照数据，仅保留最新一天的值（非最新日期清零，避免多日累加）
    _date_str = df["日期"].astype(str).str.strip()
    _latest_date = _date_str.max()
    df.loc[_date_str != _latest_date, "总库存"] = 0

    df["城市_清洗"]  = df["城市"].astype(str).apply(lambda x: re.sub(r"（.*?）", "", x))

    # ── 统计源数据天数（A列日期去重计数）────────────────────
    num_days = df["日期"].astype(str).str.strip().nunique()
    if num_days == 0:
        num_days = 1

    # ── 按商品名称分组并按销售额排序 ────────────────────────
    product_groups = []
    for product_name, df_product in df.groupby("商品名称"):
        # 计算汇总数据用于排序（排除"全国"行，避免翻倍）
        df_no_national = df_product[df_product["城市_清洗"] != "全国"]
        total_sales = float(df_no_national["商品销售额"].sum())
        total_qty = int(df_no_national["商品销售量"].sum())
        total_stock = int(df_no_national["总库存"].sum())
        unit_price = round(total_sales / total_qty, 2) if total_qty else 0
        product_groups.append({
            "name": product_name,
            "df": df_product,
            "sales": total_sales,
            "qty": total_qty,
            "stock": total_stock,
            "price": unit_price
        })

    # 按销售额降序排序
    product_groups.sort(key=lambda x: x["sales"], reverse=True)

    # ── 插入全部商品汇总（编号 00）────────────────────────
    _df_no_national = df[df["城市_清洗"] != "全国"]
    _all_sales = float(_df_no_national["商品销售额"].sum())
    _all_qty   = int(_df_no_national["商品销售量"].sum())
    _all_stock = int(_df_no_national["总库存"].sum())
    product_groups.insert(0, {
        "name": "全部商品汇总_分城市",
        "df": df,
        "sales": _all_sales,
        "qty": _all_qty,
        "stock": _all_stock,
        "price": round(_all_sales / _all_qty, 2) if _all_qty else 0
    })

    # ── 按商品名称分组 ────────────────────────────────────
    for idx, product_info in enumerate(product_groups):
        product_name = product_info["name"]
        df_product = product_info["df"]
        total_sales = product_info["sales"]
        total_qty = product_info["qty"]
        total_stock = product_info["stock"]
        unit_price = product_info["price"]
        # 城市级聚合
        df_city = (
            df_product.groupby("城市_清洗")
              .agg(商品销售额=("商品销售额", "sum"),
                   商品销售量=("商品销售量", "sum"),
                   总库存    =("总库存",     "sum"))
              .round(2).reset_index()
        )
        df_city["单价"] = (df_city["商品销售额"] / df_city["商品销售量"].replace(0, float("nan"))).round(2)
        df_city["单日销量"] = (df_city["商品销售量"] / num_days).round(2)
        df_city["预估60天销量"] = (df_city["单日销量"] * 60).astype(int)
        df_city["周转天数"] = (df_city["总库存"] / df_city["单日销量"].replace(0, float("nan"))).round(1)

        # 省份级聚合：去掉末尾"市"防止误截，再用映射表推导省份
        df_city["城市_标准"] = df_city["城市_清洗"].apply(lambda x: re.sub(r"市$", "", x))
        df_city["省份"]      = df_city["城市_标准"].map(city_to_province)
        df_province = (
            df_city.dropna(subset=["省份"])
                   .groupby("省份")
                   .agg(商品销售额=("商品销售额", "sum"),
                        商品销售量=("商品销售量", "sum"),
                        总库存    =("总库存",     "sum"))
                   .round(2).reset_index()
                   .sort_values("商品销售额", ascending=False)
        )
        df_province["单价"] = (df_province["商品销售额"] / df_province["商品销售量"].replace(0, float("nan"))).round(2)
        df_province["单日销量"] = (df_province["商品销售量"] / num_days).round(2)
        df_province["预估60天销量"] = (df_province["单日销量"] * 60).astype(int)
        df_province["周转天数"] = (df_province["总库存"] / df_province["单日销量"].replace(0, float("nan"))).round(1)

        # ── 准备绘图数据 ──────────────────────────────────────────
        province_map_data  = [[to_echarts_province(p), float(v)]
                              for p, v in zip(df_province["省份"], df_province["商品销售额"])]
        province_rank_data = df_province

        def _row_dict(row):
            return {
                "销售额": round(float(row["商品销售额"]), 2),
                "销售量": int(row["商品销售量"]),
                "预估60天销量": int(row["预估60天销量"]),
                "总库存": int(row["总库存"]),
                "单价":   round(float(row["单价"]), 2) if pd.notna(row["单价"]) else None,
                "周转天数": round(float(row["周转天数"]), 1) if pd.notna(row["周转天数"]) else None,
            }

        def _normalize_series(values):
            """归一化到各自最大值的百分比，返回 [{value, original}, ...]"""
            clean = [v for v in values if v is not None and not (isinstance(v, float) and pd.isna(v))]
            max_val = max(abs(v) for v in clean) if clean else 1
            if max_val == 0:
                max_val = 1
            result = []
            for v in values:
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    result.append({"value": 0, "original": None})
                else:
                    result.append({"value": round(v / max_val * 100, 2),
                                   "original": round(v, 2) if isinstance(v, float) else v})
            return result

        _norm_label = opts.LabelOpts(
            position="right",
            formatter=JsCode("function(p){var o=p.data.original;"
                             "if(o==null)return'-';"
                             "return typeof o==='number'?(o%1===0?''+o:o.toFixed(2)):''+o;}")
        )

        province_tooltip_dict = {}
        for _, row in df_province.iterrows():
            d = _row_dict(row)
            province_tooltip_dict[row["省份"]]                    = d
            province_tooltip_dict[to_echarts_province(row["省份"])] = d
        city_tooltip_dict = {row["城市_清洗"]: _row_dict(row) for _, row in df_city.iterrows()}

        # 全国总计
        national_row = df_city[df_city["城市_清洗"] == "全国"]
        city_only    = df_city[df_city["城市_清洗"] != "全国"].sort_values("商品销售额", ascending=False)
        if not national_row.empty:
            national_total = float(national_row["商品销售额"].iloc[0])
            national_qty   = int(national_row["商品销售量"].iloc[0])
            national_stock = int(national_row["总库存"].iloc[0])
        else:
            national_total = float(city_only["商品销售额"].sum())
            national_qty   = int(city_only["商品销售量"].sum())
            national_stock = int(city_only["总库存"].sum())
        national_price = round(national_total / national_qty, 2) if national_qty else None

        city_map_data  = [list(z) for z in zip(city_only["城市_清洗"], city_only["商品销售额"])]
        city_rank_data = city_only.copy()

        # ── 图表：省份地图 ────────────────────────────────────────
        chart_province_map = (
            Map(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="500px"))
            .add("省份总销售额(元)", province_map_data, "china")
            .set_global_opts(
                title_opts=opts.TitleOpts(title="省份销售额分布"),
                visualmap_opts=opts.VisualMapOpts(max_=float(df_province["商品销售额"].max()) if len(df_province) > 0 else 0, is_piecewise=False),
                tooltip_opts=opts.TooltipOpts(formatter=JsCode(
                    "function(p){var d=PROVINCE_DATA[p.name]||{};"
                    "return p.name+'<br/>销售额: ¥'+p.value+' 元'"
                    "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                    "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                    "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-')"
                    "+'<br/>周转天数: '+(d.周转天数!=null?d.周转天数+' 天':'-');}"
                )),
            )
        )

        # ── 图表：省份排行榜 ──────────────────────────────────────
        _province_h = f"{max(500, len(province_rank_data) * 32 + 80)}px"
        chart_province_bar = (
            Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height=_province_h))
            .add_xaxis(province_rank_data["省份"].tolist()[::-1])
            .add_yaxis("总销售额", _normalize_series(province_rank_data["商品销售额"].tolist()[::-1]),
                       label_opts=_norm_label)
            .add_yaxis("销售量", _normalize_series(province_rank_data["商品销售量"].tolist()[::-1]),
                       label_opts=_norm_label)
            .add_yaxis("预估60天销量", _normalize_series(province_rank_data["预估60天销量"].tolist()[::-1]),
                       label_opts=_norm_label)
            .add_yaxis("总库存", _normalize_series(province_rank_data["总库存"].tolist()[::-1]),
                       label_opts=_norm_label)
            .add_yaxis("周转天数(天)", _normalize_series(province_rank_data["周转天数"].tolist()[::-1]),
                       label_opts=_norm_label)
            .reversal_axis()
            .set_global_opts(
                title_opts=opts.TitleOpts(title="省份销售额排行榜（全部）"),
                xaxis_opts=opts.AxisOpts(name="相对比例 (%)", max_=100),
                legend_opts=opts.LegendOpts(pos_top="40px", pos_left="center",
                    selected_map={"总销售额": True, "销售量": False, "预估60天销量": False, "总库存": True, "周转天数(天)": False}),
                tooltip_opts=opts.TooltipOpts(trigger="axis", formatter=JsCode(
                    "function(ps){var p=ps[0],d=PROVINCE_DATA[p.name]||{};"
                    "return p.name"
                    "+'<br/>销售额: '+(d.销售额!=null?'¥'+d.销售额+' 元':'-')"
                    "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                    "+'<br/>预估60天销量: '+(d.预估60天销量!=null?d.预估60天销量+' 件':'-')"
                    "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                    "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-')"
                    "+'<br/>周转天数: '+(d.周转天数!=null?d.周转天数+' 天':'-');}"
                )),
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
                visualmap_opts=opts.VisualMapOpts(max_=float(city_only["商品销售额"].max()) if len(city_only) > 0 else 0, is_piecewise=False),
                tooltip_opts=opts.TooltipOpts(formatter=JsCode(
                    "function(p){var d=CITY_DATA[p.name]||{};"
                    "return p.name+'<br/>销售额: ¥'+p.value[2]+' 元'"
                    "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                    "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                    "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-')"
                    "+'<br/>周转天数: '+(d.周转天数!=null?d.周转天数+' 天':'-');}"
                )),
            )
        )

        # ── 图表：城市排行榜 ──────────────────────────────────────
        city_rank_data["城市_轴标"] = city_rank_data.apply(
            lambda r: f"（{r['省份']}）{r['城市_清洗']}" if pd.notna(r.get("省份")) else r["城市_清洗"],
            axis=1,
        )

        _city_h = f"{max(500, len(city_rank_data) * 32 + 80)}px"
        chart_city_bar = (
            Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height=_city_h))
            .add_xaxis(city_rank_data["城市_轴标"].tolist()[::-1])
            .add_yaxis("总销售额", _normalize_series(city_rank_data["商品销售额"].tolist()[::-1]),
                       label_opts=_norm_label)
            .add_yaxis("销售量", _normalize_series(city_rank_data["商品销售量"].tolist()[::-1]),
                       label_opts=_norm_label)
            .add_yaxis("预估60天销量", _normalize_series(city_rank_data["预估60天销量"].tolist()[::-1]),
                       label_opts=_norm_label)
            .add_yaxis("总库存", _normalize_series(city_rank_data["总库存"].tolist()[::-1]),
                       label_opts=_norm_label)
            .add_yaxis("周转天数(天)", _normalize_series(city_rank_data["周转天数"].tolist()[::-1]),
                       label_opts=_norm_label)
            .reversal_axis()
            .set_global_opts(
                title_opts=opts.TitleOpts(title="城市销售额排行榜（全部）"),
                xaxis_opts=opts.AxisOpts(name="相对比例 (%)", max_=100),
                legend_opts=opts.LegendOpts(pos_top="40px", pos_left="center",
                    selected_map={"总销售额": True, "销售量": False, "预估60天销量": False, "总库存": True, "周转天数(天)": False}),
                tooltip_opts=opts.TooltipOpts(trigger="axis", formatter=JsCode(
                    "function(ps){"
                    "var p=ps[0];"
                    "var m=p.name.match(/^（(.+)）(.+)$/);"
                    "var city=m?m[2]:p.name,prov=m?m[1]:'';"
                    "var label=m?city+'（'+prov+'）':city;"
                    "var d=CITY_DATA[city]||{};"
                    "return label"
                    "+'<br/>销售额: '+(d.销售额!=null?'¥'+d.销售额+' 元':'-')"
                    "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                    "+'<br/>预估60天销量: '+(d.预估60天销量!=null?d.预估60天销量+' 件':'-')"
                    "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                    "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-')"
                    "+'<br/>周转天数: '+(d.周转天数!=null?d.周转天数+' 天':'-');}"
                )),
            )
        )

        # ── 生成输出文件路径 ──────────────────────────────────────
        basename = os.path.splitext(os.path.basename(file_path))[0]
        output_product_dir = os.path.join(output_dir, basename)
        os.makedirs(output_product_dir, exist_ok=True)

        safe_product_name = re.sub(r'[<>:"/\\|?*]', '', product_name)
        sales_fmt = format_number(total_sales)
        qty_fmt = format_number(total_qty)
        stock_fmt = format_number(total_stock)
        price_fmt = f"{unit_price:.2f}".rstrip('0').rstrip('.')
        daily_qty = total_qty / num_days if num_days else 0
        turnover = round(total_stock / daily_qty, 1) if daily_qty else 0
        turnover_fmt = f"{turnover:.2f}".rstrip('0').rstrip('.')
        filename = f"{idx:02d}【{sales_fmt}  {qty_fmt}  {stock_fmt}  {price_fmt}  {turnover_fmt}】{safe_product_name}.html"
        output_file = os.path.join(output_product_dir, filename)

        # ── 渲染输出 ──────────────────────────────────────────────
        page = Page(layout=Page.SimplePageLayout)
        page.add(chart_province_map, chart_province_bar, chart_city_map, chart_city_bar)
        page.render(output_file)

        # ── 注入全国总计 KPI 卡片 ─────────────────────────────────
        _price_str = f"¥ {national_price:,.2f}" if national_price else "-"
        _national_daily_qty = national_qty / num_days if num_days else 0
        _national_turnover = round(national_stock / _national_daily_qty, 1) if _national_daily_qty else None
        _turnover_str = f"{_national_turnover:.1f}".rstrip('0').rstrip('.') if _national_turnover is not None else "-"
        kpi_html = f"""
<div style="font-family:'Microsoft YaHei',sans-serif;background:#f0f4ff;padding:18px 24px 14px;border-bottom:2px solid #d0d8f0;margin-bottom:4px;">
  <div style="display:flex;justify-content:center;gap:20px;flex-wrap:wrap;">
    <div style="background:#fff;border-radius:12px;padding:16px 28px;box-shadow:0 2px 12px rgba(44,123,229,0.11);text-align:center;min-width:160px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">全国总销售额</div>
      <div style="color:#2c7be5;font-size:28px;font-weight:bold;">¥ {national_total:,.2f}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">元</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:16px 28px;box-shadow:0 2px 12px rgba(44,123,229,0.11);text-align:center;min-width:160px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">全国总销售量</div>
      <div style="color:#27ae60;font-size:28px;font-weight:bold;">{national_qty:,}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">件</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:16px 28px;box-shadow:0 2px 12px rgba(44,123,229,0.11);text-align:center;min-width:160px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">全国总库存</div>
      <div style="color:#e67e22;font-size:28px;font-weight:bold;">{national_stock:,}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">件</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:16px 28px;box-shadow:0 2px 12px rgba(44,123,229,0.11);text-align:center;min-width:160px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">全国平均单价</div>
      <div style="color:#8e44ad;font-size:28px;font-weight:bold;">{_price_str}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">元/件</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:16px 28px;box-shadow:0 2px 12px rgba(44,123,229,0.11);text-align:center;min-width:160px;cursor:help;" title="周转天数 = 总库存 ÷ 单日销量&#10;单日销量 = 总销售量 ÷ 统计天数({num_days}天)&#10;&#10;数值越小说明库存周转越快，越大说明库存积压越多">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">全国周转天数</div>
      <div style="color:#e74c3c;font-size:28px;font-weight:bold;">{_turnover_str}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">天（统计{num_days}天）</div>
    </div>
  </div>
</div>
"""
        js_data = (
            "<script>\n"
            f"var PROVINCE_DATA={json.dumps(province_tooltip_dict, ensure_ascii=False)};\n"
            f"var CITY_DATA={json.dumps(city_tooltip_dict, ensure_ascii=False)};\n"
            "</script>"
        )

        with open(output_file, "r", encoding="utf-8") as f:
            html = f.read()
        # 工具按钮脚本 —— 为排行榜柱状图添加"全部不选"和"重置排名"按钮
        toolbar_js = """<script>
(function(){
    var allDivs=document.querySelectorAll('div[id]'),charts=[];
    allDivs.forEach(function(div){
        try{var inst=echarts.getInstanceByDom(div);if(inst)charts.push({div:div,inst:inst});}catch(e){}
    });
    var names=['总销售额','销售量','预估60天销量','总库存','周转天数(天)'];
    var btnCss='padding:6px 18px;font-size:13px;border:1px solid #ccc;border-radius:6px;background:#fff;cursor:pointer;font-family:Microsoft YaHei,sans-serif;';

    function resetRank(chart){
        var opt=chart.getOption();
        var selected=opt.legend[0].selected||{};
        var cats=opt.yAxis[0].data;
        var slist=opt.series;
        var arr=cats.map(function(_,i){
            var sum=0;
            slist.forEach(function(s){
                if(selected[s.name]!==false){
                    var v=s.data[i];
                    var raw=v&&v.original!=null?v.original:(typeof v==='number'?v:0);
                    if(typeof raw==='number') sum+=raw;
                }
            });
            return {i:i,s:sum};
        });
        arr.sort(function(a,b){return a.s-b.s;});
        chart.setOption({
            yAxis:[{data:arr.map(function(x){return cats[x.i];})}],
            series:slist.map(function(s){
                return {data:arr.map(function(x){return s.data[x.i];})};
            })
        });
    }

    [1,3].forEach(function(idx){
        if(idx>=charts.length)return;
        var c=charts[idx];
        var singleMode=false;

        var wrap=document.createElement('div');
        wrap.style.cssText='margin:8px 0 4px 20px;display:flex;gap:10px;align-items:center;';

        var btn1=document.createElement('button');
        btn1.textContent='全部不选';
        btn1.style.cssText=btnCss;
        btn1.onmouseover=function(){btn1.style.background='#f0f0f0';};
        btn1.onmouseout=function(){btn1.style.background='#fff';};
        btn1.onclick=function(){
            names.forEach(function(n){c.inst.dispatchAction({type:'legendUnSelect',name:n});});
        };

        var btn2=document.createElement('button');
        btn2.textContent='重置排名';
        btn2.style.cssText=btnCss+'border-color:#2c7be5;color:#2c7be5;';
        btn2.onmouseover=function(){btn2.style.background='#e8f0fe';};
        btn2.onmouseout=function(){btn2.style.background='#fff';};
        btn2.onclick=function(){resetRank(c.inst);};

        var btn3=document.createElement('button');
        btn3.textContent='单选模式：关';
        btn3.style.cssText=btnCss+'border-color:#e67e22;color:#e67e22;';
        function updateBtn3(){
            if(singleMode){
                btn3.textContent='单选模式：开';
                btn3.style.background='#e67e22';btn3.style.color='#fff';btn3.style.borderColor='#e67e22';
            }else{
                btn3.textContent='单选模式：关';
                btn3.style.background='#fff';btn3.style.color='#e67e22';btn3.style.borderColor='#e67e22';
            }
        }
        btn3.onmouseover=function(){if(!singleMode)btn3.style.background='#fdf2e9';};
        btn3.onmouseout=function(){if(!singleMode)btn3.style.background='#fff';};
        btn3.onclick=function(){singleMode=!singleMode;updateBtn3();};

        c.inst.on('legendselectchanged',function(params){
            if(!singleMode)return;
            var sel=params.selected;
            var clicked=params.name;
            names.forEach(function(n){
                if(n===clicked){
                    if(!sel[n]) c.inst.dispatchAction({type:'legendSelect',name:n});
                }else{
                    if(sel[n]) c.inst.dispatchAction({type:'legendUnSelect',name:n});
                }
            });
        });

        var rawMode=false;
        var btn4=document.createElement('button');
        btn4.textContent='实际数值';
        btn4.style.cssText=btnCss+'border-color:#8e44ad;color:#8e44ad;';
        function updateBtn4(){
            if(rawMode){
                btn4.textContent='实际数值：开';
                btn4.style.background='#8e44ad';btn4.style.color='#fff';btn4.style.borderColor='#8e44ad';
            }else{
                btn4.textContent='实际数值';
                btn4.style.background='#fff';btn4.style.color='#8e44ad';btn4.style.borderColor='#8e44ad';
            }
        }
        btn4.onmouseover=function(){if(!rawMode)btn4.style.background='#f4ecf7';};
        btn4.onmouseout=function(){if(!rawMode)btn4.style.background='#fff';};
        btn4.onclick=function(){
            rawMode=!rawMode;updateBtn4();
            var opt=c.inst.getOption();
            if(rawMode){
                c.inst.setOption({
                    series:opt.series.map(function(s){
                        return {data:s.data.map(function(d){
                            return {value:d.original!=null?d.original:0, original:d.original};
                        })};
                    }),
                    xAxis:[{max:null, name:'实际数值'}]
                });
            }else{
                c.inst.setOption({
                    series:opt.series.map(function(s){
                        var vals=s.data.map(function(d){return d.original;}).filter(function(v){return v!=null;});
                        var mx=Math.max.apply(null,vals.map(function(v){return Math.abs(v);}));
                        if(!mx)mx=1;
                        return {data:s.data.map(function(d){
                            if(d.original==null) return {value:0,original:null};
                            return {value:Math.round(d.original/mx*10000)/100, original:d.original};
                        })};
                    }),
                    xAxis:[{max:100, name:'相对比例 (%)'}]
                });
            }
        };

        wrap.appendChild(btn1);
        wrap.appendChild(btn2);
        wrap.appendChild(btn3);
        wrap.appendChild(btn4);
        c.div.parentNode.insertBefore(wrap,c.div);
    });
})();
</script>"""

        html = html.replace("</head>", js_data + "\n</head>", 1)
        html = re.sub(r"(<body[^>]*>)", r"\1\n" + kpi_html, html, count=1)
        html = html.replace("</body>", toolbar_js + "\n</body>", 1)
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(html)

        print(f"已生成：{output_file}")

        # ── 为汇总（编号 00）生成同名表格文件 ────────────────────
        if idx == 0:
            xlsx_output = output_file.replace("_分城市.html", "_分单品.xlsx")
            product_rows = []
            for _pi, pg in enumerate(product_groups[1:], 1):
                _pg_qty = pg["qty"]
                _pg_daily = _pg_qty / num_days if num_days else 0
                _pg_est60 = int(_pg_daily * 60)
                _pg_turnover = round(pg["stock"] / _pg_daily, 1) if _pg_daily else None
                product_rows.append({
                    "商品名称": f"【{_pi:02d}】{pg['name']}",
                    "销售额": pg["sales"],
                    "销售量": _pg_qty,
                    "预估60天销量": _pg_est60,
                    "总库存": pg["stock"],
                    "单价": pg["price"],
                    "周转天数": _pg_turnover,
                })
            df_product_summary = pd.DataFrame(product_rows)
            df_province_out = df_province[["省份", "商品销售额", "商品销售量", "预估60天销量", "总库存", "单价", "周转天数"]].copy()
            df_city_out = city_only[["城市_清洗", "省份", "商品销售额", "商品销售量", "预估60天销量", "总库存", "单价", "周转天数"]].copy()
            df_city_out.rename(columns={"城市_清洗": "城市"}, inplace=True)
            with pd.ExcelWriter(xlsx_output, engine="openpyxl") as writer:
                df_product_summary.to_excel(writer, sheet_name="商品汇总", index=False)
                df_province_out.to_excel(writer, sheet_name="省份汇总", index=False)
                df_city_out.to_excel(writer, sheet_name="城市汇总", index=False)
            print(f"已生成：{xlsx_output}")

            # ── 为汇总（编号 00）生成商品排行榜 BI 看板（分单品）──────
            ranking_filename = filename.replace("_分城市.html", "_分单品.html")
            ranking_output = os.path.join(output_product_dir, ranking_filename)

            # 准备商品排行数据（跳过 product_groups[0] 即"全部商品汇总"本身）
            _rank_names = []
            _rank_sales = []
            _rank_qty = []
            _rank_est60 = []
            _rank_stock = []
            _rank_turnover = []
            _rank_tooltip = {}
            for _pi, pg in enumerate(product_groups[1:], 1):
                _label = f"【{_pi:02d}】{pg['name']}"
                _rank_names.append(_label)
                _rank_sales.append(pg["sales"])
                _rank_qty.append(pg["qty"])
                _pg_daily = pg["qty"] / num_days if num_days else 0
                _pg_est60 = int(_pg_daily * 60)
                _rank_est60.append(_pg_est60)
                _rank_stock.append(pg["stock"])
                _pg_turnover = round(pg["stock"] / _pg_daily, 1) if _pg_daily else None
                _rank_turnover.append(_pg_turnover)
                _rank_tooltip[_label] = {
                    "销售额": round(pg["sales"], 2),
                    "销售量": pg["qty"],
                    "预估60天销量": _pg_est60,
                    "总库存": pg["stock"],
                    "单价": pg["price"],
                    "周转天数": _pg_turnover,
                }

            _product_h = f"{max(600, len(_rank_names) * 32 + 80)}px"
            chart_product_bar = (
                Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height=_product_h))
                .add_xaxis(_rank_names[::-1])
                .add_yaxis("总销售额", _normalize_series(_rank_sales[::-1]),
                           label_opts=_norm_label)
                .add_yaxis("销售量", _normalize_series(_rank_qty[::-1]),
                           label_opts=_norm_label)
                .add_yaxis("预估60天销量", _normalize_series(_rank_est60[::-1]),
                           label_opts=_norm_label)
                .add_yaxis("总库存", _normalize_series(_rank_stock[::-1]),
                           label_opts=_norm_label)
                .add_yaxis("周转天数(天)", _normalize_series(_rank_turnover[::-1]),
                           label_opts=_norm_label)
                .reversal_axis()
                .set_global_opts(
                    title_opts=opts.TitleOpts(title="商品销售额排行榜（全部）"),
                    xaxis_opts=opts.AxisOpts(name="相对比例 (%)", max_=100),
                    legend_opts=opts.LegendOpts(pos_top="40px", pos_left="center",
                        selected_map={"总销售额": True, "销售量": False, "预估60天销量": False, "总库存": True, "周转天数(天)": False}),
                    tooltip_opts=opts.TooltipOpts(trigger="axis", formatter=JsCode(
                        "function(ps){var p=ps[0],d=PRODUCT_DATA[p.name]||{};"
                        "return p.name"
                        "+'<br/>销售额: '+(d.销售额!=null?'¥'+d.销售额+' 元':'-')"
                        "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                        "+'<br/>预估60天销量: '+(d.预估60天销量!=null?d.预估60天销量+' 件':'-')"
                        "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                        "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-')"
                        "+'<br/>周转天数: '+(d.周转天数!=null?d.周转天数+' 天':'-');}"
                    )),
                )
            )

            rank_page = Page(layout=Page.SimplePageLayout)
            rank_page.add(chart_product_bar)
            rank_page.render(ranking_output)

            # 注入 JS 数据、KPI 卡片和工具按钮
            rank_js_data = (
                "<script>\n"
                f"var PRODUCT_DATA={json.dumps(_rank_tooltip, ensure_ascii=False)};\n"
                "</script>"
            )

            rank_toolbar_js = """<script>
(function(){
    var allDivs=document.querySelectorAll('div[id]'),charts=[];
    allDivs.forEach(function(div){
        try{var inst=echarts.getInstanceByDom(div);if(inst)charts.push({div:div,inst:inst});}catch(e){}
    });
    var names=['总销售额','销售量','预估60天销量','总库存','周转天数(天)'];
    var btnCss='padding:6px 18px;font-size:13px;border:1px solid #ccc;border-radius:6px;background:#fff;cursor:pointer;font-family:Microsoft YaHei,sans-serif;';

    function resetRank(chart){
        var opt=chart.getOption();
        var selected=opt.legend[0].selected||{};
        var cats=opt.yAxis[0].data;
        var slist=opt.series;
        var arr=cats.map(function(_,i){
            var sum=0;
            slist.forEach(function(s){
                if(selected[s.name]!==false){
                    var v=s.data[i];
                    var raw=v&&v.original!=null?v.original:(typeof v==='number'?v:0);
                    if(typeof raw==='number') sum+=raw;
                }
            });
            return {i:i,s:sum};
        });
        arr.sort(function(a,b){return a.s-b.s;});
        chart.setOption({
            yAxis:[{data:arr.map(function(x){return cats[x.i];})}],
            series:slist.map(function(s){
                return {data:arr.map(function(x){return s.data[x.i];})};
            })
        });
    }

    charts.forEach(function(c){
        var singleMode=false;

        var wrap=document.createElement('div');
        wrap.style.cssText='margin:8px 0 4px 20px;display:flex;gap:10px;align-items:center;';

        var btn1=document.createElement('button');
        btn1.textContent='全部不选';
        btn1.style.cssText=btnCss;
        btn1.onmouseover=function(){btn1.style.background='#f0f0f0';};
        btn1.onmouseout=function(){btn1.style.background='#fff';};
        btn1.onclick=function(){
            names.forEach(function(n){c.inst.dispatchAction({type:'legendUnSelect',name:n});});
        };

        var btn2=document.createElement('button');
        btn2.textContent='重置排名';
        btn2.style.cssText=btnCss+'border-color:#2c7be5;color:#2c7be5;';
        btn2.onmouseover=function(){btn2.style.background='#e8f0fe';};
        btn2.onmouseout=function(){btn2.style.background='#fff';};
        btn2.onclick=function(){resetRank(c.inst);};

        var btn3=document.createElement('button');
        btn3.textContent='单选模式：关';
        btn3.style.cssText=btnCss+'border-color:#e67e22;color:#e67e22;';
        function updateBtn3(){
            if(singleMode){
                btn3.textContent='单选模式：开';
                btn3.style.background='#e67e22';btn3.style.color='#fff';btn3.style.borderColor='#e67e22';
            }else{
                btn3.textContent='单选模式：关';
                btn3.style.background='#fff';btn3.style.color='#e67e22';btn3.style.borderColor='#e67e22';
            }
        }
        btn3.onmouseover=function(){if(!singleMode)btn3.style.background='#fdf2e9';};
        btn3.onmouseout=function(){if(!singleMode)btn3.style.background='#fff';};
        btn3.onclick=function(){singleMode=!singleMode;updateBtn3();};

        c.inst.on('legendselectchanged',function(params){
            if(!singleMode)return;
            var sel=params.selected;
            var clicked=params.name;
            names.forEach(function(n){
                if(n===clicked){
                    if(!sel[n]) c.inst.dispatchAction({type:'legendSelect',name:n});
                }else{
                    if(sel[n]) c.inst.dispatchAction({type:'legendUnSelect',name:n});
                }
            });
        });

        var rawMode=false;
        var btn4=document.createElement('button');
        btn4.textContent='实际数值';
        btn4.style.cssText=btnCss+'border-color:#8e44ad;color:#8e44ad;';
        function updateBtn4(){
            if(rawMode){
                btn4.textContent='实际数值：开';
                btn4.style.background='#8e44ad';btn4.style.color='#fff';btn4.style.borderColor='#8e44ad';
            }else{
                btn4.textContent='实际数值';
                btn4.style.background='#fff';btn4.style.color='#8e44ad';btn4.style.borderColor='#8e44ad';
            }
        }
        btn4.onmouseover=function(){if(!rawMode)btn4.style.background='#f4ecf7';};
        btn4.onmouseout=function(){if(!rawMode)btn4.style.background='#fff';};
        btn4.onclick=function(){
            rawMode=!rawMode;updateBtn4();
            var opt=c.inst.getOption();
            if(rawMode){
                c.inst.setOption({
                    series:opt.series.map(function(s){
                        return {data:s.data.map(function(d){
                            return {value:d.original!=null?d.original:0, original:d.original};
                        })};
                    }),
                    xAxis:[{max:null, name:'实际数值'}]
                });
            }else{
                c.inst.setOption({
                    series:opt.series.map(function(s){
                        var vals=s.data.map(function(d){return d.original;}).filter(function(v){return v!=null;});
                        var mx=Math.max.apply(null,vals.map(function(v){return Math.abs(v);}));
                        if(!mx)mx=1;
                        return {data:s.data.map(function(d){
                            if(d.original==null) return {value:0,original:null};
                            return {value:Math.round(d.original/mx*10000)/100, original:d.original};
                        })};
                    }),
                    xAxis:[{max:100, name:'相对比例 (%)'}]
                });
            }
        };

        wrap.appendChild(btn1);
        wrap.appendChild(btn2);
        wrap.appendChild(btn3);
        wrap.appendChild(btn4);
        c.div.parentNode.insertBefore(wrap,c.div);
    });
})();
</script>"""

            with open(ranking_output, "r", encoding="utf-8") as f:
                rank_html = f.read()
            rank_html = rank_html.replace("</head>", rank_js_data + "\n</head>", 1)
            rank_html = re.sub(r"(<body[^>]*>)", r"\1\n" + kpi_html, rank_html, count=1)
            rank_html = rank_html.replace("</body>", rank_toolbar_js + "\n</body>", 1)
            with open(ranking_output, "w", encoding="utf-8") as f:
                f.write(rank_html)
            print(f"已生成：{ranking_output}")

            # ── 为汇总（编号 00）生成补货建议 BI 看板 ──────────────
            replenish_filename = filename.replace("_分城市.html", "_补货建议.html")
            replenish_output = os.path.join(output_product_dir, replenish_filename)

            # 准备补货建议数据
            _rep_items = []
            for _pi, pg in enumerate(product_groups[1:], 1):
                _pg_qty = pg["qty"]
                _pg_stock = pg["stock"]
                _pg_sales = pg["sales"]
                _pg_daily = _pg_qty / num_days if num_days else 0
                _pg_est60 = int(_pg_daily * 60)
                if _pg_qty > 0 and _pg_est60 > 0:
                    _pg_replenish = max(0, _pg_est60 - _pg_stock)
                    _pg_coverage = round(_pg_stock / _pg_est60 * 100, 1)
                    if _pg_coverage < 30:
                        _pg_urgency = "紧急"
                    elif _pg_coverage < 70:
                        _pg_urgency = "建议"
                    elif _pg_coverage < 100:
                        _pg_urgency = "关注"
                    else:
                        _pg_urgency = "充足"
                else:
                    _pg_replenish = None
                    _pg_coverage = None
                    _pg_urgency = "无销售历史"
                _rep_items.append({
                    "orig_idx": _pi,
                    "name": pg["name"],
                    "sales": _pg_sales,
                    "qty": _pg_qty,
                    "stock": _pg_stock,
                    "est60": _pg_est60,
                    "replenish": _pg_replenish,
                    "coverage": _pg_coverage,
                    "urgency": _pg_urgency,
                    "price": pg["price"],
                })

            # 按补货建议降序排序（None 沉底）
            _rep_items.sort(key=lambda x: (x["replenish"] if x["replenish"] is not None else -1), reverse=True)

            # 生成 label：【文件编号】商品名【顺序编号】
            for _oi, it in enumerate(_rep_items, 1):
                it["label"] = f"【{it['orig_idx']:02d}】{it['name']}【{_oi:02d}】"

            _rep_names = [it["label"] for it in _rep_items]
            _rep_est60_list = [it["est60"] for it in _rep_items]
            _rep_stock_list = [it["stock"] for it in _rep_items]
            _rep_replenish_list = [it["replenish"] if it["replenish"] is not None else 0 for it in _rep_items]

            _rep_tooltip = {}
            for it in _rep_items:
                _rep_tooltip[it["label"]] = {
                    "销售额": round(float(it["sales"]), 2),
                    "销售量": it["qty"],
                    "预估60天销量": it["est60"],
                    "总库存": it["stock"],
                    "单价": it["price"],
                    "库存覆盖率": it["coverage"],
                    "补货建议": it["replenish"],
                    "紧急度": it["urgency"],
                }

            # 补货汇总 KPI
            _rep_need_count = sum(1 for it in _rep_items if it["replenish"] is not None and it["replenish"] > 0)
            _rep_total_qty = sum(it["replenish"] for it in _rep_items if it["replenish"] is not None)
            _rep_urgent_count = sum(1 for it in _rep_items if it["urgency"] == "紧急")
            _rep_with_sales = sum(1 for it in _rep_items if it["replenish"] is not None)
            _rep_enough_count = sum(1 for it in _rep_items if it["replenish"] == 0)
            _rep_enough_rate = f"{_rep_enough_count / _rep_with_sales * 100:.1f}%" if _rep_with_sales else "-"

            _replenish_h = f"{max(600, len(_rep_names) * 32 + 80)}px"
            chart_replenish_bar = (
                Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height=_replenish_h))
                .add_xaxis(_rep_names[::-1])
                .add_yaxis("预估60天销量", _normalize_series(_rep_est60_list[::-1]),
                           label_opts=_norm_label)
                .add_yaxis("总库存", _normalize_series(_rep_stock_list[::-1]),
                           label_opts=_norm_label)
                .add_yaxis("补货建议", _normalize_series(_rep_replenish_list[::-1]),
                           label_opts=_norm_label,
                           itemstyle_opts=opts.ItemStyleOpts(color="#e74c3c"))
                .reversal_axis()
                .set_global_opts(
                    title_opts=opts.TitleOpts(title="商品补货建议（预估60天销量 − 当前库存）"),
                    xaxis_opts=opts.AxisOpts(name="相对比例 (%)", max_=100),
                    legend_opts=opts.LegendOpts(pos_top="40px", pos_left="center",
                        selected_map={"预估60天销量": True, "总库存": True, "补货建议": True}),
                    tooltip_opts=opts.TooltipOpts(trigger="axis", formatter=JsCode(
                        "function(ps){var p=ps[0],d=REPLENISH_DATA[p.name]||{};"
                        "return p.name"
                        "+'<br/>销售额: '+(d.销售额!=null?'¥'+d.销售额+' 元':'-')"
                        "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                        "+'<br/>预估60天销量: '+(d.预估60天销量!=null?d.预估60天销量+' 件':'-')"
                        "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                        "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-')"
                        "+'<br/>库存覆盖率: '+(d.库存覆盖率!=null?d.库存覆盖率+' %':'-')"
                        "+'<br/><b>补货建议: '+(d.补货建议!=null?d.补货建议+' 件':'-')+'</b>'"
                        "+'<br/>紧急度: '+(d.紧急度!=null?d.紧急度:'-');}"
                    )),
                )
            )

            replenish_page = Page(layout=Page.SimplePageLayout)
            replenish_page.add(chart_replenish_bar)
            replenish_page.render(replenish_output)

            # 注入 JS 数据、KPI 卡片和工具按钮
            replenish_js_data = (
                "<script>\n"
                f"var REPLENISH_DATA={json.dumps(_rep_tooltip, ensure_ascii=False)};\n"
                "</script>"
            )

            # 补货专用 KPI 卡片（追加在通用 kpi_html 之后）
            replenish_kpi_html = f"""
<div style="font-family:'Microsoft YaHei',sans-serif;background:#fff5f5;padding:14px 24px 12px;border-bottom:2px solid #f0d0d0;margin-bottom:4px;">
  <div style="display:flex;justify-content:center;gap:20px;flex-wrap:wrap;">
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(231,76,60,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">需补货商品数</div>
      <div style="color:#e74c3c;font-size:26px;font-weight:bold;">{_rep_need_count}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">个</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(231,76,60,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">总补货件数</div>
      <div style="color:#c0392b;font-size:26px;font-weight:bold;">{_rep_total_qty:,}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">件</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(231,76,60,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">紧急商品数</div>
      <div style="color:#d35400;font-size:26px;font-weight:bold;">{_rep_urgent_count}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">库存&lt;30%覆盖</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(231,76,60,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">库存充足率</div>
      <div style="color:#27ae60;font-size:26px;font-weight:bold;">{_rep_enough_rate}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">≥60天覆盖占比</div>
    </div>
  </div>
</div>
"""

            replenish_toolbar_js = """<script>
(function(){
    var allDivs=document.querySelectorAll('div[id]'),charts=[];
    allDivs.forEach(function(div){
        try{var inst=echarts.getInstanceByDom(div);if(inst)charts.push({div:div,inst:inst});}catch(e){}
    });
    var names=['预估60天销量','总库存','补货建议'];
    var btnCss='padding:6px 18px;font-size:13px;border:1px solid #ccc;border-radius:6px;background:#fff;cursor:pointer;font-family:Microsoft YaHei,sans-serif;';

    function resetRank(chart){
        var opt=chart.getOption();
        var selected=opt.legend[0].selected||{};
        var cats=opt.yAxis[0].data;
        var slist=opt.series;
        var arr=cats.map(function(_,i){
            var sum=0;
            slist.forEach(function(s){
                if(selected[s.name]!==false){
                    var v=s.data[i];
                    var raw=v&&v.original!=null?v.original:(typeof v==='number'?v:0);
                    if(typeof raw==='number') sum+=raw;
                }
            });
            return {i:i,s:sum};
        });
        arr.sort(function(a,b){return a.s-b.s;});
        chart.setOption({
            yAxis:[{data:arr.map(function(x){return cats[x.i];})}],
            series:slist.map(function(s){
                return {data:arr.map(function(x){return s.data[x.i];})};
            })
        });
    }

    charts.forEach(function(c){
        var singleMode=false;

        var wrap=document.createElement('div');
        wrap.style.cssText='margin:8px 0 4px 20px;display:flex;gap:10px;align-items:center;';

        var btn1=document.createElement('button');
        btn1.textContent='全部不选';
        btn1.style.cssText=btnCss;
        btn1.onmouseover=function(){btn1.style.background='#f0f0f0';};
        btn1.onmouseout=function(){btn1.style.background='#fff';};
        btn1.onclick=function(){
            names.forEach(function(n){c.inst.dispatchAction({type:'legendUnSelect',name:n});});
        };

        var btn2=document.createElement('button');
        btn2.textContent='重置排名';
        btn2.style.cssText=btnCss+'border-color:#2c7be5;color:#2c7be5;';
        btn2.onmouseover=function(){btn2.style.background='#e8f0fe';};
        btn2.onmouseout=function(){btn2.style.background='#fff';};
        btn2.onclick=function(){resetRank(c.inst);};

        var btn3=document.createElement('button');
        btn3.textContent='单选模式：关';
        btn3.style.cssText=btnCss+'border-color:#e67e22;color:#e67e22;';
        function updateBtn3(){
            if(singleMode){
                btn3.textContent='单选模式：开';
                btn3.style.background='#e67e22';btn3.style.color='#fff';btn3.style.borderColor='#e67e22';
            }else{
                btn3.textContent='单选模式：关';
                btn3.style.background='#fff';btn3.style.color='#e67e22';btn3.style.borderColor='#e67e22';
            }
        }
        btn3.onmouseover=function(){if(!singleMode)btn3.style.background='#fdf2e9';};
        btn3.onmouseout=function(){if(!singleMode)btn3.style.background='#fff';};
        btn3.onclick=function(){singleMode=!singleMode;updateBtn3();};

        c.inst.on('legendselectchanged',function(params){
            if(!singleMode)return;
            var sel=params.selected;
            var clicked=params.name;
            names.forEach(function(n){
                if(n===clicked){
                    if(!sel[n]) c.inst.dispatchAction({type:'legendSelect',name:n});
                }else{
                    if(sel[n]) c.inst.dispatchAction({type:'legendUnSelect',name:n});
                }
            });
        });

        var rawMode=false;
        var btn4=document.createElement('button');
        btn4.textContent='实际数值';
        btn4.style.cssText=btnCss+'border-color:#8e44ad;color:#8e44ad;';
        function updateBtn4(){
            if(rawMode){
                btn4.textContent='实际数值：开';
                btn4.style.background='#8e44ad';btn4.style.color='#fff';btn4.style.borderColor='#8e44ad';
            }else{
                btn4.textContent='实际数值';
                btn4.style.background='#fff';btn4.style.color='#8e44ad';btn4.style.borderColor='#8e44ad';
            }
        }
        btn4.onmouseover=function(){if(!rawMode)btn4.style.background='#f4ecf7';};
        btn4.onmouseout=function(){if(!rawMode)btn4.style.background='#fff';};
        btn4.onclick=function(){
            rawMode=!rawMode;updateBtn4();
            var opt=c.inst.getOption();
            if(rawMode){
                c.inst.setOption({
                    series:opt.series.map(function(s){
                        return {data:s.data.map(function(d){
                            return {value:d.original!=null?d.original:0, original:d.original};
                        })};
                    }),
                    xAxis:[{max:null, name:'实际数值'}]
                });
            }else{
                c.inst.setOption({
                    series:opt.series.map(function(s){
                        var vals=s.data.map(function(d){return d.original;}).filter(function(v){return v!=null;});
                        var mx=Math.max.apply(null,vals.map(function(v){return Math.abs(v);}));
                        if(!mx)mx=1;
                        return {data:s.data.map(function(d){
                            if(d.original==null) return {value:0,original:null};
                            return {value:Math.round(d.original/mx*10000)/100, original:d.original};
                        })};
                    }),
                    xAxis:[{max:100, name:'相对比例 (%)'}]
                });
            }
        };

        wrap.appendChild(btn1);
        wrap.appendChild(btn2);
        wrap.appendChild(btn3);
        wrap.appendChild(btn4);
        c.div.parentNode.insertBefore(wrap,c.div);

        // 默认开启【实际数值】
        btn4.onclick();
    });
})();
</script>"""

            with open(replenish_output, "r", encoding="utf-8") as f:
                rep_html = f.read()
            rep_html = rep_html.replace("</head>", replenish_js_data + "\n</head>", 1)
            rep_html = re.sub(r"(<body[^>]*>)", r"\1\n" + kpi_html + replenish_kpi_html, rep_html, count=1)
            rep_html = rep_html.replace("</body>", replenish_toolbar_js + "\n</body>", 1)
            with open(replenish_output, "w", encoding="utf-8") as f:
                f.write(rep_html)
            print(f"已生成：{replenish_output}")

            # ── 为汇总（编号 00）生成动销差与库存积压预警 BI 看板 ──────
            alert_filename = filename.replace("_分城市.html", "_动销差与库存积压预警.html")
            alert_output = os.path.join(output_product_dir, alert_filename)

            # 全局均单价（用于零动销品估算积压金额）
            _global_avg_price = unit_price if unit_price else 0

            _alert_items = []
            for _pi, pg in enumerate(product_groups[1:], 1):
                _pg_qty = pg["qty"]
                _pg_stock = pg["stock"]
                _pg_sales = pg["sales"]
                _pg_price = pg["price"]
                _pg_daily = _pg_qty / num_days if num_days else 0
                _pg_est60 = int(_pg_daily * 60)
                _pg_turnover = round(_pg_stock / _pg_daily, 1) if _pg_daily else None

                # 判定是否预警
                _is_zero_sales = (_pg_qty == 0 and _pg_stock > 0)
                _is_slow = (_pg_qty > 0 and _pg_turnover is not None and _pg_turnover > 60)
                if not (_is_zero_sales or _is_slow):
                    continue

                # 积压件数
                if _is_zero_sales:
                    _pg_excess = _pg_stock
                else:
                    _pg_excess = max(0, _pg_stock - _pg_est60)

                # 积压金额：优先用自身单价，零动销无单价则用全局均价
                _pg_unit_price = _pg_price if _pg_price else _global_avg_price
                _pg_excess_value = round(_pg_excess * _pg_unit_price, 2)

                # 状态分档
                if _is_zero_sales:
                    _pg_status = "零动销"
                elif _pg_turnover is not None and _pg_turnover > 180:
                    _pg_status = "严重积压"
                else:
                    _pg_status = "库存积压"

                _alert_items.append({
                    "orig_idx": _pi,
                    "name": pg["name"],
                    "sales": _pg_sales,
                    "qty": _pg_qty,
                    "stock": _pg_stock,
                    "est60": _pg_est60,
                    "turnover": _pg_turnover,
                    "excess": _pg_excess,
                    "excess_value": _pg_excess_value,
                    "price": _pg_price,
                    "status": _pg_status,
                })

            # 按积压件数降序排序
            _alert_items.sort(key=lambda x: x["excess"], reverse=True)

            # 生成 label：【文件编号】商品名【顺序编号】
            for _oi, it in enumerate(_alert_items, 1):
                it["label"] = f"【{it['orig_idx']:02d}】{it['name']}【{_oi:02d}】"

            if len(_alert_items) == 0:
                # 无预警商品，仅生成一个空提示页
                print(f"跳过：无动销差/库存积压商品，未生成 {alert_output}")
            else:
                _alert_names = [it["label"] for it in _alert_items]
                _alert_est60_list = [it["est60"] for it in _alert_items]
                _alert_stock_list = [it["stock"] for it in _alert_items]
                _alert_excess_list = [it["excess"] for it in _alert_items]

                _alert_tooltip = {}
                for it in _alert_items:
                    _alert_tooltip[it["label"]] = {
                        "销售额": round(float(it["sales"]), 2),
                        "销售量": it["qty"],
                        "预估60天销量": it["est60"],
                        "总库存": it["stock"],
                        "单价": it["price"],
                        "周转天数": it["turnover"],
                        "积压件数": it["excess"],
                        "积压金额": it["excess_value"],
                        "状态": it["status"],
                    }

                # 预警汇总 KPI
                _alert_total_count = len(_alert_items)
                _alert_zero_count = sum(1 for it in _alert_items if it["status"] == "零动销")
                _alert_severe_count = sum(1 for it in _alert_items if it["status"] == "严重积压")
                _alert_total_excess_qty = sum(it["excess"] for it in _alert_items)
                _alert_total_excess_value = sum(it["excess_value"] for it in _alert_items)

                _alert_h = f"{max(600, len(_alert_names) * 32 + 80)}px"
                chart_alert_bar = (
                    Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height=_alert_h))
                    .add_xaxis(_alert_names[::-1])
                    .add_yaxis("预估60天销量", _normalize_series(_alert_est60_list[::-1]),
                               label_opts=_norm_label)
                    .add_yaxis("总库存", _normalize_series(_alert_stock_list[::-1]),
                               label_opts=_norm_label)
                    .add_yaxis("积压件数", _normalize_series(_alert_excess_list[::-1]),
                               label_opts=_norm_label,
                               itemstyle_opts=opts.ItemStyleOpts(color="#c0392b"))
                    .reversal_axis()
                    .set_global_opts(
                        title_opts=opts.TitleOpts(title="商品动销差与库存积压预警（60天卖不完）"),
                        xaxis_opts=opts.AxisOpts(name="相对比例 (%)", max_=100),
                        legend_opts=opts.LegendOpts(pos_top="40px", pos_left="center",
                            selected_map={"预估60天销量": True, "总库存": True, "积压件数": True}),
                        tooltip_opts=opts.TooltipOpts(trigger="axis", formatter=JsCode(
                            "function(ps){var p=ps[0],d=ALERT_DATA[p.name]||{};"
                            "return p.name"
                            "+'<br/><b>状态: '+(d.状态!=null?d.状态:'-')+'</b>'"
                            "+'<br/>销售额: '+(d.销售额!=null?'¥'+d.销售额+' 元':'-')"
                            "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                            "+'<br/>预估60天销量: '+(d.预估60天销量!=null?d.预估60天销量+' 件':'-')"
                            "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                            "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-')"
                            "+'<br/>周转天数: '+(d.周转天数!=null?d.周转天数+' 天':'-')"
                            "+'<br/><b>积压件数: '+(d.积压件数!=null?d.积压件数+' 件':'-')+'</b>'"
                            "+'<br/><b>积压金额: '+(d.积压金额!=null?'¥'+d.积压金额+' 元':'-')+'</b>';}"
                        )),
                    )
                )

                alert_page = Page(layout=Page.SimplePageLayout)
                alert_page.add(chart_alert_bar)
                alert_page.render(alert_output)

                alert_js_data = (
                    "<script>\n"
                    f"var ALERT_DATA={json.dumps(_alert_tooltip, ensure_ascii=False)};\n"
                    "</script>"
                )

                # 预警专用 KPI 卡片
                alert_kpi_html = f"""
<div style="font-family:'Microsoft YaHei',sans-serif;background:#fff5f5;padding:14px 24px 12px;border-bottom:2px solid #f0d0d0;margin-bottom:4px;">
  <div style="display:flex;justify-content:center;gap:20px;flex-wrap:wrap;">
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(192,57,43,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">预警商品总数</div>
      <div style="color:#c0392b;font-size:26px;font-weight:bold;">{_alert_total_count}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">个</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(192,57,43,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">零动销商品</div>
      <div style="color:#7f8c8d;font-size:26px;font-weight:bold;">{_alert_zero_count}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">近{num_days}天无销量</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(192,57,43,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">严重积压商品</div>
      <div style="color:#d35400;font-size:26px;font-weight:bold;">{_alert_severe_count}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">周转&gt;180天</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(192,57,43,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">总积压件数</div>
      <div style="color:#e74c3c;font-size:26px;font-weight:bold;">{_alert_total_excess_qty:,}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">件</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(192,57,43,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">总积压金额</div>
      <div style="color:#c0392b;font-size:26px;font-weight:bold;">¥ {_alert_total_excess_value:,.0f}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">元</div>
    </div>
  </div>
</div>
"""

                # 工具按钮 JS（复用补货建议的模板，仅替换 names）
                alert_toolbar_js = replenish_toolbar_js.replace(
                    "var names=['预估60天销量','总库存','补货建议'];",
                    "var names=['预估60天销量','总库存','积压件数'];"
                )

                with open(alert_output, "r", encoding="utf-8") as f:
                    alert_html = f.read()
                alert_html = alert_html.replace("</head>", alert_js_data + "\n</head>", 1)
                alert_html = re.sub(r"(<body[^>]*>)", r"\1\n" + kpi_html + alert_kpi_html, alert_html, count=1)
                alert_html = alert_html.replace("</body>", alert_toolbar_js + "\n</body>", 1)
                with open(alert_output, "w", encoding="utf-8") as f:
                    f.write(alert_html)
                print(f"已生成：{alert_output}")

            # ── 为汇总（编号 00）生成爆品扩城市推荐 BI 看板 ──────────
            expand_filename = filename.replace("_分城市.html", "_爆品扩城市推荐.html")
            expand_output = os.path.join(output_product_dir, expand_filename)

            # 全局 商品×城市 聚合（排除"全国"行）
            _df_exp = df[df["城市_清洗"] != "全国"].copy()
            _df_pc = (_df_exp.groupby(["商品名称", "城市_清洗"])
                             .agg(pc_sales=("商品销售额", "sum"),
                                  pc_qty=("商品销售量", "sum"))
                             .reset_index())

            # 全国运营城市集：任何商品有销量的城市并集
            _global_cities = set(_df_pc[_df_pc["pc_qty"] > 0]["城市_清洗"].unique())
            _total_city_count = len(_global_cities)

            # 爆品筛选：销售量>0 且 周转天数<=60
            _expand_all = []
            for _pi, pg in enumerate(product_groups[1:], 1):
                _pg_qty = pg["qty"]
                _pg_stock = pg["stock"]
                _pg_sales = pg["sales"]
                _pg_daily = _pg_qty / num_days if num_days else 0
                _pg_turnover = round(_pg_stock / _pg_daily, 1) if _pg_daily else None

                if not (_pg_qty > 0 and _pg_turnover is not None and _pg_turnover <= 60):
                    continue

                _prod_cities = set(_df_pc[(_df_pc["商品名称"] == pg["name"]) &
                                          (_df_pc["pc_qty"] > 0)]["城市_清洗"].unique())
                _missing_cities = sorted(_global_cities - _prod_cities)
                _covered_count = len(_prod_cities)
                _missing_count = len(_missing_cities)
                _avg_city_sales = round(_pg_sales / _covered_count, 2) if _covered_count else 0
                _potential = round(_avg_city_sales * _missing_count, 2)

                _expand_all.append({
                    "orig_idx": _pi,
                    "name": pg["name"],
                    "sales": _pg_sales,
                    "qty": _pg_qty,
                    "stock": _pg_stock,
                    "turnover": _pg_turnover,
                    "price": pg["price"],
                    "covered_count": _covered_count,
                    "missing_count": _missing_count,
                    "missing_cities": _missing_cities,
                    "avg_city_sales": _avg_city_sales,
                    "potential": _potential,
                })

            # 按潜力扩城销售额降序，取 top 30
            _expand_all.sort(key=lambda x: x["potential"], reverse=True)
            TOP_N_EXPAND = 30
            _expand_items = _expand_all[:TOP_N_EXPAND]

            if len(_expand_items) == 0 or _total_city_count == 0:
                print(f"跳过：无爆品或无运营城市，未生成 {expand_output}")
            else:
                # 生成 label：【文件编号】商品名【顺序编号】
                for _oi, it in enumerate(_expand_items, 1):
                    it["label"] = f"【{it['orig_idx']:02d}】{it['name']}【{_oi:02d}】"

                _exp_names = [it["label"] for it in _expand_items]
                _exp_sales_list = [it["sales"] for it in _expand_items]
                _exp_avg_city_list = [it["avg_city_sales"] for it in _expand_items]
                _exp_potential_list = [it["potential"] for it in _expand_items]
                _exp_covered_list = [it["covered_count"] for it in _expand_items]
                _exp_missing_list = [it["missing_count"] for it in _expand_items]

                _exp_tooltip = {}
                for it in _expand_items:
                    _exp_tooltip[it["label"]] = {
                        "销售额": round(float(it["sales"]), 2),
                        "销售量": it["qty"],
                        "周转天数": it["turnover"],
                        "单价": it["price"],
                        "已覆盖城市数": it["covered_count"],
                        "缺失城市数": it["missing_count"],
                        "单城市均销": it["avg_city_sales"],
                        "潜力扩城销售额": it["potential"],
                    }

                # KPI 汇总
                _exp_count = len(_expand_items)
                _exp_total_potential = sum(it["potential"] for it in _expand_items)
                _exp_avg_coverage = (
                    sum(it["covered_count"] for it in _expand_items) /
                    (_exp_count * _total_city_count) * 100
                ) if _exp_count and _total_city_count else 0

                _expand_h = f"{max(600, len(_exp_names) * 36 + 80)}px"
                chart_expand_bar = (
                    Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height=_expand_h))
                    .add_xaxis(_exp_names[::-1])
                    .add_yaxis("总销售额", _normalize_series(_exp_sales_list[::-1]),
                               label_opts=_norm_label)
                    .add_yaxis("单城市均销", _normalize_series(_exp_avg_city_list[::-1]),
                               label_opts=_norm_label)
                    .add_yaxis("潜力扩城销售额", _normalize_series(_exp_potential_list[::-1]),
                               label_opts=_norm_label,
                               itemstyle_opts=opts.ItemStyleOpts(color="#e74c3c"))
                    .add_yaxis("已覆盖城市数", _normalize_series(_exp_covered_list[::-1]),
                               label_opts=_norm_label)
                    .add_yaxis("缺失城市数", _normalize_series(_exp_missing_list[::-1]),
                               label_opts=_norm_label,
                               itemstyle_opts=opts.ItemStyleOpts(color="#f39c12"))
                    .reversal_axis()
                    .set_global_opts(
                        title_opts=opts.TitleOpts(title=f"爆品扩城市推荐（动销好的商品 × 缺失城市，全国共{_total_city_count}个运营城市）"),
                        xaxis_opts=opts.AxisOpts(name="相对比例 (%)", max_=100),
                        legend_opts=opts.LegendOpts(pos_top="40px", pos_left="center",
                            selected_map={"总销售额": False, "单城市均销": False,
                                          "潜力扩城销售额": True, "已覆盖城市数": False, "缺失城市数": True}),
                        tooltip_opts=opts.TooltipOpts(trigger="axis", formatter=JsCode(
                            "function(ps){var p=ps[0],d=EXPAND_DATA[p.name]||{};"
                            "return p.name"
                            "+'<br/>销售额: '+(d.销售额!=null?'¥'+d.销售额+' 元':'-')"
                            "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                            "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-')"
                            "+'<br/>周转天数: '+(d.周转天数!=null?d.周转天数+' 天':'-')"
                            "+'<br/>已覆盖城市数: '+(d.已覆盖城市数!=null?d.已覆盖城市数+' 个':'-')"
                            "+'<br/><b>缺失城市数: '+(d.缺失城市数!=null?d.缺失城市数+' 个':'-')+'</b>'"
                            "+'<br/>单城市均销: '+(d.单城市均销!=null?'¥'+d.单城市均销:'-')"
                            "+'<br/><b>潜力扩城销售额: '+(d.潜力扩城销售额!=null?'¥'+d.潜力扩城销售额+' 元':'-')+'</b>';}"
                        )),
                    )
                )

                expand_page = Page(layout=Page.SimplePageLayout)
                expand_page.add(chart_expand_bar)
                expand_page.render(expand_output)

                expand_js_data = (
                    "<script>\n"
                    f"var EXPAND_DATA={json.dumps(_exp_tooltip, ensure_ascii=False)};\n"
                    "</script>"
                )

                # 扩城专用 KPI 卡片
                expand_kpi_html = f"""
<div style="font-family:'Microsoft YaHei',sans-serif;background:#f0fff4;padding:14px 24px 12px;border-bottom:2px solid #c8e6c9;margin-bottom:4px;">
  <div style="display:flex;justify-content:center;gap:20px;flex-wrap:wrap;">
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(39,174,96,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">爆品数量</div>
      <div style="color:#27ae60;font-size:26px;font-weight:bold;">{_exp_count}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">动销&lt;60天&取Top{TOP_N_EXPAND}</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(39,174,96,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">全国运营城市数</div>
      <div style="color:#2c7be5;font-size:26px;font-weight:bold;">{_total_city_count}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">个</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(39,174,96,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">平均覆盖率</div>
      <div style="color:#8e44ad;font-size:26px;font-weight:bold;">{_exp_avg_coverage:.1f}%</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">爆品平均铺货率</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(39,174,96,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">总潜力扩城销售额</div>
      <div style="color:#e74c3c;font-size:26px;font-weight:bold;">¥ {_exp_total_potential:,.0f}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">元（均销×缺失数）</div>
    </div>
  </div>
</div>
"""

                # 底部卡片列表：每个爆品一张卡片 + 缺失城市标签墙
                _cards_parts = ['<div style="font-family:\'Microsoft YaHei\',sans-serif;background:#fafbfd;padding:20px;">']
                _cards_parts.append('<h2 style="text-align:center;color:#2c3e50;margin:10px 0 20px;">📍 爆品扩城市推荐详单（按潜力降序）</h2>')
                for it in _expand_items:
                    _missing_tags = "".join(
                        f'<span style="display:inline-block;background:#fff3e0;color:#e67e22;border:1px solid #ffcc80;'
                        f'border-radius:14px;padding:4px 12px;margin:3px;font-size:12px;">{c}</span>'
                        for c in it["missing_cities"]
                    )
                    if not _missing_tags:
                        _missing_tags = '<span style="color:#27ae60;">✔ 已全量覆盖</span>'
                    _cards_parts.append(f"""
<div style="background:#fff;border-radius:10px;padding:16px 20px;margin-bottom:14px;box-shadow:0 2px 10px rgba(0,0,0,0.06);border-left:5px solid #e74c3c;">
  <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;margin-bottom:10px;">
    <div style="font-size:16px;font-weight:bold;color:#2c3e50;">{it['label']}</div>
    <div style="font-size:13px;color:#555;">
      <span style="margin-right:16px;">💰 销售额 <b style="color:#2c7be5;">¥{it['sales']:,.0f}</b></span>
      <span style="margin-right:16px;">🏙 已覆盖 <b style="color:#27ae60;">{it['covered_count']}</b>/{_total_city_count}</span>
      <span style="margin-right:16px;">❌ 缺失 <b style="color:#e67e22;">{it['missing_count']}</b> 城</span>
      <span style="margin-right:16px;">📊 单城市均销 <b>¥{it['avg_city_sales']:,.0f}</b></span>
      <span>🚀 潜力 <b style="color:#e74c3c;">¥{it['potential']:,.0f}</b></span>
    </div>
  </div>
  <div style="padding-top:8px;border-top:1px dashed #eee;">
    <div style="font-size:12px;color:#888;margin-bottom:6px;">建议开拓城市：</div>
    <div>{_missing_tags}</div>
  </div>
</div>
""")
                _cards_parts.append('</div>')
                expand_cards_html = "".join(_cards_parts)

                # 工具按钮 JS（复用补货建议模板，替换 names，关闭默认实际数值）
                expand_toolbar_js = replenish_toolbar_js.replace(
                    "var names=['预估60天销量','总库存','补货建议'];",
                    "var names=['总销售额','单城市均销','潜力扩城销售额','已覆盖城市数','缺失城市数'];"
                ).replace(
                    "// 默认开启【实际数值】\n        btn4.onclick();",
                    "// 扩城看板保持归一化模式"
                )

                with open(expand_output, "r", encoding="utf-8") as f:
                    exp_html = f.read()
                exp_html = exp_html.replace("</head>", expand_js_data + "\n</head>", 1)
                exp_html = re.sub(r"(<body[^>]*>)", r"\1\n" + kpi_html + expand_kpi_html, exp_html, count=1)
                exp_html = exp_html.replace("</body>", expand_cards_html + "\n" + expand_toolbar_js + "\n</body>", 1)
                with open(expand_output, "w", encoding="utf-8") as f:
                    f.write(exp_html)
                print(f"已生成：{expand_output}")

            # ── 为汇总（编号 00）生成库存结构分布 BI 看板 ──────────
            try:
                structure_filename = filename.replace("_分城市.html", "_库存结构分析.html")
                structure_output = os.path.join(output_product_dir, structure_filename)

                _latest_date_s = df["日期"].astype(str).str.strip().max()
                _df_latest = df[(df["日期"].astype(str).str.strip() == _latest_date_s) &
                                (df["城市_清洗"] != "全国")].copy()
                for _c in ["供应商到大仓在途数量", "大仓库存数量", "大仓到门店在途数量", "前置站点库存数量"]:
                    _df_latest[_c] = pd.to_numeric(_df_latest[_c], errors="coerce").fillna(0)

                _df_stock_struct = _df_latest.groupby("商品名称").agg(
                    供应商在途=("供应商到大仓在途数量", "sum"),
                    大仓库存=("大仓库存数量", "sum"),
                    大仓到门店在途=("大仓到门店在途数量", "sum"),
                    前置站点=("前置站点库存数量", "sum"),
                ).reset_index()
                _df_stock_struct["总库存"] = (_df_stock_struct["供应商在途"] +
                                              _df_stock_struct["大仓库存"] +
                                              _df_stock_struct["大仓到门店在途"] +
                                              _df_stock_struct["前置站点"])

                # 映射 orig_idx
                _name_to_idx = {pg["name"]: _i for _i, pg in enumerate(product_groups[1:], 1)}

                _struct_items = []
                for _, row in _df_stock_struct.iterrows():
                    if row["总库存"] <= 0:
                        continue
                    _n = row["商品名称"]
                    _oi = _name_to_idx.get(_n, 0)
                    _supplier = int(row["供应商在途"])
                    _warehouse = int(row["大仓库存"])
                    _transit = int(row["大仓到门店在途"])
                    _frontier = int(row["前置站点"])
                    _total = int(row["总库存"])
                    _is_empty_frontier = (_frontier == 0 and _total > 0)
                    _transit_ratio = (_supplier + _transit) / _total if _total else 0
                    _is_stuck = (_transit_ratio > 0.5)
                    _tags = []
                    if _is_empty_frontier:
                        _tags.append("⚠前置空仓")
                    if _is_stuck:
                        _tags.append("⚠在途滞留")
                    _struct_items.append({
                        "orig_idx": _oi,
                        "name": _n,
                        "supplier": _supplier,
                        "warehouse": _warehouse,
                        "transit": _transit,
                        "frontier": _frontier,
                        "total": _total,
                        "frontier_ratio": round(_frontier / _total * 100, 1) if _total else 0,
                        "transit_ratio": round(_transit_ratio * 100, 1),
                        "tags": "/".join(_tags) if _tags else "健康",
                    })

                _struct_items.sort(key=lambda x: x["total"], reverse=True)

                if len(_struct_items) == 0:
                    print(f"跳过：无库存数据，未生成 {structure_output}")
                else:
                    for _oi, it in enumerate(_struct_items, 1):
                        it["label"] = f"【{it['orig_idx']:02d}】{it['name']}【{_oi:02d}】"

                    _struct_names = [it["label"] for it in _struct_items]
                    _struct_tooltip = {}
                    for it in _struct_items:
                        _struct_tooltip[it["label"]] = {
                            "供应商在途": it["supplier"],
                            "大仓库存": it["warehouse"],
                            "大仓到门店在途": it["transit"],
                            "前置站点": it["frontier"],
                            "总库存": it["total"],
                            "前置占比": it["frontier_ratio"],
                            "在途占比": it["transit_ratio"],
                            "状态": it["tags"],
                        }

                    _empty_frontier_count = sum(1 for it in _struct_items if it["frontier"] == 0)
                    _stuck_count = sum(1 for it in _struct_items if it["transit_ratio"] > 50)
                    _healthy_count = sum(1 for it in _struct_items if it["tags"] == "健康")
                    _total_all = sum(it["total"] for it in _struct_items)
                    _avg_supplier = sum(it["supplier"] for it in _struct_items) / _total_all * 100 if _total_all else 0
                    _avg_warehouse = sum(it["warehouse"] for it in _struct_items) / _total_all * 100 if _total_all else 0
                    _avg_transit = sum(it["transit"] for it in _struct_items) / _total_all * 100 if _total_all else 0
                    _avg_frontier = sum(it["frontier"] for it in _struct_items) / _total_all * 100 if _total_all else 0

                    _struct_h = f"{max(600, len(_struct_names) * 36 + 80)}px"
                    chart_struct_bar = (
                        Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height=_struct_h))
                        .add_xaxis(_struct_names[::-1])
                        .add_yaxis("供应商在途", [it["supplier"] for it in _struct_items][::-1],
                                   stack="stock", itemstyle_opts=opts.ItemStyleOpts(color="#3498db"))
                        .add_yaxis("大仓库存", [it["warehouse"] for it in _struct_items][::-1],
                                   stack="stock", itemstyle_opts=opts.ItemStyleOpts(color="#27ae60"))
                        .add_yaxis("大仓到门店在途", [it["transit"] for it in _struct_items][::-1],
                                   stack="stock", itemstyle_opts=opts.ItemStyleOpts(color="#f39c12"))
                        .add_yaxis("前置站点", [it["frontier"] for it in _struct_items][::-1],
                                   stack="stock", itemstyle_opts=opts.ItemStyleOpts(color="#e74c3c"))
                        .reversal_axis()
                        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
                        .set_global_opts(
                            title_opts=opts.TitleOpts(title="库存结构分布（四段堆叠）"),
                            xaxis_opts=opts.AxisOpts(name="库存件数"),
                            legend_opts=opts.LegendOpts(pos_top="40px", pos_left="center"),
                            tooltip_opts=opts.TooltipOpts(trigger="axis", formatter=JsCode(
                                "function(ps){var p=ps[0],d=STRUCT_DATA[p.name]||{};"
                                "return p.name"
                                "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                                "+'<br/>供应商在途: '+(d.供应商在途!=null?d.供应商在途+' 件':'-')"
                                "+'<br/>大仓库存: '+(d.大仓库存!=null?d.大仓库存+' 件':'-')"
                                "+'<br/>大仓到门店在途: '+(d.大仓到门店在途!=null?d.大仓到门店在途+' 件':'-')"
                                "+'<br/>前置站点: '+(d.前置站点!=null?d.前置站点+' 件':'-')"
                                "+'<br/>前置占比: '+(d.前置占比!=null?d.前置占比+' %':'-')"
                                "+'<br/>在途占比: '+(d.在途占比!=null?d.在途占比+' %':'-')"
                                "+'<br/><b>状态: '+(d.状态!=null?d.状态:'-')+'</b>';}"
                            )),
                        )
                    )

                    struct_page = Page(layout=Page.SimplePageLayout)
                    struct_page.add(chart_struct_bar)
                    struct_page.render(structure_output)

                    struct_js_data = (
                        "<script>\n"
                        f"var STRUCT_DATA={json.dumps(_struct_tooltip, ensure_ascii=False)};\n"
                        "</script>"
                    )

                    struct_kpi_html = f"""
<div style="font-family:'Microsoft YaHei',sans-serif;background:#f3f0ff;padding:14px 24px 12px;border-bottom:2px solid #d8d0f0;margin-bottom:4px;">
  <div style="display:flex;justify-content:center;gap:20px;flex-wrap:wrap;">
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(142,68,173,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">前置站点空仓</div>
      <div style="color:#e74c3c;font-size:26px;font-weight:bold;">{_empty_frontier_count}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">个商品前置=0</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(142,68,173,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">在途滞留商品</div>
      <div style="color:#f39c12;font-size:26px;font-weight:bold;">{_stuck_count}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">在途占比&gt;50%</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(142,68,173,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">结构健康商品</div>
      <div style="color:#27ae60;font-size:26px;font-weight:bold;">{_healthy_count}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">个</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(142,68,173,0.12);text-align:center;min-width:180px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">四段平均占比</div>
      <div style="color:#555;font-size:14px;font-weight:bold;line-height:1.5;">
        <span style="color:#3498db;">供应商 {_avg_supplier:.1f}%</span> /
        <span style="color:#27ae60;">大仓 {_avg_warehouse:.1f}%</span><br/>
        <span style="color:#f39c12;">在途 {_avg_transit:.1f}%</span> /
        <span style="color:#e74c3c;">前置 {_avg_frontier:.1f}%</span>
      </div>
    </div>
  </div>
</div>
"""

                    struct_toolbar_js = replenish_toolbar_js.replace(
                        "var names=['预估60天销量','总库存','补货建议'];",
                        "var names=['供应商在途','大仓库存','大仓到门店在途','前置站点'];"
                    ).replace(
                        "// 默认开启【实际数值】\n        btn4.onclick();",
                        "// 堆叠柱状本身就是实际数值"
                    )

                    with open(structure_output, "r", encoding="utf-8") as f:
                        struct_html = f.read()
                    struct_html = struct_html.replace("</head>", struct_js_data + "\n</head>", 1)
                    struct_html = re.sub(r"(<body[^>]*>)", r"\1\n" + kpi_html + struct_kpi_html, struct_html, count=1)
                    struct_html = struct_html.replace("</body>", struct_toolbar_js + "\n</body>", 1)
                    with open(structure_output, "w", encoding="utf-8") as f:
                        f.write(struct_html)
                    print(f"已生成：{structure_output}")
            except Exception as _e:
                print(f"库存结构看板生成失败：{_e}")

            # ── 为汇总（编号 00）生成日销趋势 BI 看板 ──────────────
            try:
                if num_days <= 1:
                    print(f"跳过：单日数据无趋势，未生成日销趋势看板")
                else:
                    trend_filename = filename.replace("_分城市.html", "_日销趋势.html")
                    trend_output = os.path.join(output_product_dir, trend_filename)

                    _df_trend_base = df[df["城市_清洗"] != "全国"].copy()
                    _df_trend_base["日期_str"] = _df_trend_base["日期"].astype(str).str.strip()
                    _df_trend_all = _df_trend_base.groupby("日期_str").agg(
                        总销售额=("商品销售额", "sum"),
                        总销售量=("商品销售量", "sum"),
                    ).reset_index().sort_values("日期_str")

                    def _fmt_date(s):
                        s = str(s).strip()
                        if len(s) == 8 and s.isdigit():
                            return f"{s[4:6]}/{s[6:8]}"
                        return s

                    _trend_dates = [_fmt_date(d) for d in _df_trend_all["日期_str"].tolist()]
                    _trend_sales = [round(float(v), 2) for v in _df_trend_all["总销售额"].tolist()]
                    _trend_qty = [int(v) for v in _df_trend_all["总销售量"].tolist()]

                    chart_trend_all = (
                        Line(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="450px"))
                        .add_xaxis(_trend_dates)
                        .add_yaxis("日销售额(元)", _trend_sales, yaxis_index=0,
                                   is_smooth=True,
                                   areastyle_opts=opts.AreaStyleOpts(opacity=0.3),
                                   itemstyle_opts=opts.ItemStyleOpts(color="#2c7be5"))
                        .add_yaxis("日销售量(件)", _trend_qty, yaxis_index=1,
                                   is_smooth=True,
                                   itemstyle_opts=opts.ItemStyleOpts(color="#27ae60"))
                        .extend_axis(yaxis=opts.AxisOpts(name="销售量(件)", position="right"))
                        .set_global_opts(
                            title_opts=opts.TitleOpts(title="全国日销趋势"),
                            xaxis_opts=opts.AxisOpts(name="日期", type_="category"),
                            yaxis_opts=opts.AxisOpts(name="销售额(元)", position="left"),
                            tooltip_opts=opts.TooltipOpts(trigger="axis"),
                            legend_opts=opts.LegendOpts(pos_top="40px"),
                            datazoom_opts=[opts.DataZoomOpts(type_="inside"), opts.DataZoomOpts(type_="slider")],
                        )
                    )

                    # Top 10 商品日销
                    _top10_pgs = product_groups[1:11]
                    chart_trend_top = (
                        Line(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="500px"))
                        .add_xaxis(_trend_dates)
                    )
                    for _pi, pg in enumerate(_top10_pgs, 1):
                        _pg_df = pg["df"][pg["df"]["城市_清洗"] != "全国"].copy()
                        _pg_df["日期_str"] = _pg_df["日期"].astype(str).str.strip()
                        _pg_trend_map = _pg_df.groupby("日期_str")["商品销售额"].sum().to_dict()
                        _pg_trend_values = [round(float(_pg_trend_map.get(d, 0)), 2)
                                            for d in _df_trend_all["日期_str"].tolist()]
                        _short_name = pg["name"][:18] + ("..." if len(pg["name"]) > 18 else "")
                        chart_trend_top.add_yaxis(f"【{_pi:02d}】{_short_name}", _pg_trend_values,
                                                  is_smooth=True, symbol_size=6,
                                                  label_opts=opts.LabelOpts(is_show=False))
                    chart_trend_top.set_global_opts(
                        title_opts=opts.TitleOpts(title="Top 10 商品日销售额趋势对比"),
                        xaxis_opts=opts.AxisOpts(name="日期", type_="category"),
                        yaxis_opts=opts.AxisOpts(name="销售额(元)"),
                        tooltip_opts=opts.TooltipOpts(trigger="axis"),
                        legend_opts=opts.LegendOpts(pos_top="40px", type_="scroll"),
                        datazoom_opts=[opts.DataZoomOpts(type_="inside"), opts.DataZoomOpts(type_="slider")],
                    )

                    trend_page = Page(layout=Page.SimplePageLayout)
                    trend_page.add(chart_trend_all, chart_trend_top)
                    trend_page.render(trend_output)

                    # KPI
                    _trend_total_days = len(_trend_dates)
                    _trend_avg_sales = sum(_trend_sales) / _trend_total_days if _trend_total_days else 0
                    _trend_max_idx = _trend_sales.index(max(_trend_sales)) if _trend_sales else 0
                    _trend_min_idx = _trend_sales.index(min(_trend_sales)) if _trend_sales else 0
                    _trend_max_date = _trend_dates[_trend_max_idx] if _trend_dates else "-"
                    _trend_min_date = _trend_dates[_trend_min_idx] if _trend_dates else "-"
                    _trend_max_val = max(_trend_sales) if _trend_sales else 0
                    _trend_min_val = min(_trend_sales) if _trend_sales else 0
                    if len(_trend_sales) >= 2 and _trend_sales[0] > 0:
                        _trend_growth = (_trend_sales[-1] - _trend_sales[0]) / _trend_sales[0] * 100
                        _trend_growth_str = f"{_trend_growth:+.1f}%"
                        _trend_growth_color = "#27ae60" if _trend_growth >= 0 else "#e74c3c"
                    else:
                        _trend_growth_str = "-"
                        _trend_growth_color = "#888"

                    trend_kpi_html = f"""
<div style="font-family:'Microsoft YaHei',sans-serif;background:#fffcf0;padding:14px 24px 12px;border-bottom:2px solid #f0e8c0;margin-bottom:4px;">
  <div style="display:flex;justify-content:center;gap:20px;flex-wrap:wrap;">
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(241,196,15,0.15);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">统计天数</div>
      <div style="color:#f39c12;font-size:26px;font-weight:bold;">{_trend_total_days}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">天</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(241,196,15,0.15);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">日均销售额</div>
      <div style="color:#2c7be5;font-size:26px;font-weight:bold;">¥ {_trend_avg_sales:,.0f}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">元/天</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(241,196,15,0.15);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">最高日</div>
      <div style="color:#27ae60;font-size:20px;font-weight:bold;">¥ {_trend_max_val:,.0f}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">{_trend_max_date}</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(241,196,15,0.15);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">最低日</div>
      <div style="color:#e67e22;font-size:20px;font-weight:bold;">¥ {_trend_min_val:,.0f}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">{_trend_min_date}</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(241,196,15,0.15);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">首末日环比</div>
      <div style="color:{_trend_growth_color};font-size:26px;font-weight:bold;">{_trend_growth_str}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">末日 vs 首日</div>
    </div>
  </div>
</div>
"""

                    with open(trend_output, "r", encoding="utf-8") as f:
                        trend_html = f.read()
                    trend_html = re.sub(r"(<body[^>]*>)", r"\1\n" + kpi_html + trend_kpi_html, trend_html, count=1)
                    with open(trend_output, "w", encoding="utf-8") as f:
                        f.write(trend_html)
                    print(f"已生成：{trend_output}")
            except Exception as _e:
                print(f"日销趋势看板生成失败：{_e}")

            # ── 为汇总（编号 00）生成城市销售健康度评分 BI 看板 ──────
            try:
                health_filename = filename.replace("_分城市.html", "_城市健康度.html")
                health_output = os.path.join(output_product_dir, health_filename)

                # 全国 SKU 总数
                _global_sku_count = df[df["城市_清洗"] != "全国"]["商品名称"].nunique()

                # 按城市聚合
                _df_city_base = df[df["城市_清洗"] != "全国"].copy()
                _df_city_latest = _df_city_base[_df_city_base["日期"].astype(str).str.strip() == _latest_date_s]
                _city_stock_map = _df_city_latest.groupby("城市_清洗").apply(
                    lambda g: (pd.to_numeric(g["供应商到大仓在途数量"], errors="coerce").fillna(0).sum() +
                               pd.to_numeric(g["大仓库存数量"], errors="coerce").fillna(0).sum() +
                               pd.to_numeric(g["大仓到门店在途数量"], errors="coerce").fillna(0).sum() +
                               pd.to_numeric(g["前置站点库存数量"], errors="coerce").fillna(0).sum())
                ).to_dict()

                _city_agg = _df_city_base.groupby("城市_清洗").agg(
                    城市销售额=("商品销售额", "sum"),
                    城市销售量=("商品销售量", "sum"),
                    SKU数=("商品名称", "nunique"),
                ).reset_index()

                _city_agg["城市总库存"] = _city_agg["城市_清洗"].map(_city_stock_map).fillna(0)
                _city_agg["均单价"] = _city_agg["城市销售额"] / _city_agg["城市销售量"].replace(0, float("nan"))
                _city_agg["单日销量"] = _city_agg["城市销售量"] / num_days
                _city_agg["周转天数"] = _city_agg["城市总库存"] / _city_agg["单日销量"].replace(0, float("nan"))

                import math as _math
                _max_sales = float(_city_agg["城市销售额"].max()) if len(_city_agg) else 1
                _max_price = float(_city_agg["均单价"].max()) if len(_city_agg) else 1

                _health_items = []
                for _, row in _city_agg.iterrows():
                    _city = row["城市_清洗"]
                    _sales = float(row["城市销售额"])
                    _qty = int(row["城市销售量"])
                    _sku = int(row["SKU数"])
                    _price = float(row["均单价"]) if pd.notna(row["均单价"]) else 0
                    _turn = float(row["周转天数"]) if pd.notna(row["周转天数"]) else None

                    _sku_score = _sku / _global_sku_count * 100 if _global_sku_count else 0
                    _sales_score = (_math.log(_sales + 1) / _math.log(_max_sales + 1) * 100) if _max_sales > 0 else 0
                    _price_score = (_price / _max_price * 100) if _max_price > 0 else 0
                    if _turn is None:
                        _turn_score = 0
                    else:
                        _turn_score = max(0.0, min(100.0, (1 - (_turn - 30) / 150) * 100))
                    _total_score = round(_sku_score * 0.30 + _sales_score * 0.25 +
                                         _price_score * 0.15 + _turn_score * 0.30, 1)

                    if _total_score >= 80:
                        _level = "标杆"
                    elif _total_score >= 60:
                        _level = "健康"
                    elif _total_score >= 40:
                        _level = "待改善"
                    else:
                        _level = "问题"

                    _health_items.append({
                        "city": _city,
                        "sales": round(_sales, 2),
                        "qty": _qty,
                        "sku": _sku,
                        "price": round(_price, 2),
                        "turnover": round(_turn, 1) if _turn is not None else None,
                        "sku_score": round(_sku_score, 1),
                        "sales_score": round(_sales_score, 1),
                        "price_score": round(_price_score, 1),
                        "turn_score": round(_turn_score, 1),
                        "total_score": _total_score,
                        "level": _level,
                    })

                _health_items.sort(key=lambda x: x["total_score"], reverse=True)

                if len(_health_items) == 0:
                    print(f"跳过：无城市数据，未生成 {health_output}")
                else:
                    _health_labels = [f"{it['city']}（{it['level']}）" for it in _health_items]
                    _health_tooltip = {}
                    for i, it in enumerate(_health_items):
                        _health_tooltip[_health_labels[i]] = {
                            "销售额": it["sales"],
                            "销售量": it["qty"],
                            "SKU数": it["sku"],
                            "均单价": it["price"],
                            "周转天数": it["turnover"],
                            "SKU覆盖分": it["sku_score"],
                            "销售规模分": it["sales_score"],
                            "单价水平分": it["price_score"],
                            "周转健康分": it["turn_score"],
                            "总分": it["total_score"],
                            "分级": it["level"],
                        }

                    _benchmark_count = sum(1 for it in _health_items if it["total_score"] >= 80)
                    _problem_count = sum(1 for it in _health_items if it["total_score"] < 40)
                    _avg_health = sum(it["total_score"] for it in _health_items) / len(_health_items)
                    _top_city = _health_items[0]

                    _health_h = f"{max(600, len(_health_labels) * 28 + 80)}px"
                    chart_health_bar = (
                        Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height=_health_h))
                        .add_xaxis(_health_labels[::-1])
                        .add_yaxis("总分", [it["total_score"] for it in _health_items][::-1],
                                   label_opts=opts.LabelOpts(position="right"),
                                   itemstyle_opts=opts.ItemStyleOpts(color=JsCode(
                                       "function(p){var v=p.value;"
                                       "if(v>=80)return '#27ae60';"
                                       "if(v>=60)return '#2c7be5';"
                                       "if(v>=40)return '#f39c12';"
                                       "return '#e74c3c';}"
                                   )))
                        .add_yaxis("SKU覆盖分", [it["sku_score"] for it in _health_items][::-1],
                                   label_opts=opts.LabelOpts(position="right"))
                        .add_yaxis("销售规模分", [it["sales_score"] for it in _health_items][::-1],
                                   label_opts=opts.LabelOpts(position="right"))
                        .add_yaxis("单价水平分", [it["price_score"] for it in _health_items][::-1],
                                   label_opts=opts.LabelOpts(position="right"))
                        .add_yaxis("周转健康分", [it["turn_score"] for it in _health_items][::-1],
                                   label_opts=opts.LabelOpts(position="right"))
                        .reversal_axis()
                        .set_global_opts(
                            title_opts=opts.TitleOpts(title="城市销售健康度评分（权重 SKU30%+销售25%+单价15%+周转30%）"),
                            xaxis_opts=opts.AxisOpts(name="分数", max_=100),
                            legend_opts=opts.LegendOpts(pos_top="40px", pos_left="center",
                                selected_map={"总分": True, "SKU覆盖分": False, "销售规模分": False,
                                              "单价水平分": False, "周转健康分": False}),
                            tooltip_opts=opts.TooltipOpts(trigger="axis", formatter=JsCode(
                                "function(ps){var p=ps[0],d=HEALTH_DATA[p.name]||{};"
                                "return p.name"
                                "+'<br/><b>总分: '+(d.总分!=null?d.总分:'-')+' ('+(d.分级!=null?d.分级:'-')+')</b>'"
                                "+'<br/>销售额: '+(d.销售额!=null?'¥'+d.销售额+' 元':'-')"
                                "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                                "+'<br/>SKU数: '+(d.SKU数!=null?d.SKU数+' 个':'-')"
                                "+'<br/>均单价: '+(d.均单价!=null?'¥'+d.均单价:'-')"
                                "+'<br/>周转天数: '+(d.周转天数!=null?d.周转天数+' 天':'-')"
                                "+'<br/>—— 分项 ——'"
                                "+'<br/>SKU覆盖分: '+(d.SKU覆盖分!=null?d.SKU覆盖分:'-')"
                                "+'<br/>销售规模分: '+(d.销售规模分!=null?d.销售规模分:'-')"
                                "+'<br/>单价水平分: '+(d.单价水平分!=null?d.单价水平分:'-')"
                                "+'<br/>周转健康分: '+(d.周转健康分!=null?d.周转健康分:'-');}"
                            )),
                        )
                    )

                    health_page = Page(layout=Page.SimplePageLayout)
                    health_page.add(chart_health_bar)
                    health_page.render(health_output)

                    health_js_data = (
                        "<script>\n"
                        f"var HEALTH_DATA={json.dumps(_health_tooltip, ensure_ascii=False)};\n"
                        "</script>"
                    )

                    health_kpi_html = f"""
<div style="font-family:'Microsoft YaHei',sans-serif;background:#e8f4fd;padding:14px 24px 12px;border-bottom:2px solid #b8d9f0;margin-bottom:4px;">
  <div style="display:flex;justify-content:center;gap:20px;flex-wrap:wrap;">
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(44,123,229,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">标杆城市数</div>
      <div style="color:#27ae60;font-size:26px;font-weight:bold;">{_benchmark_count}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">总分≥80</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(44,123,229,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">问题城市数</div>
      <div style="color:#e74c3c;font-size:26px;font-weight:bold;">{_problem_count}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">总分&lt;40</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(44,123,229,0.12);text-align:center;min-width:150px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">平均健康度</div>
      <div style="color:#2c7be5;font-size:26px;font-weight:bold;">{_avg_health:.1f}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">分</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:14px 26px;box-shadow:0 2px 12px rgba(44,123,229,0.12);text-align:center;min-width:180px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">最高分城市</div>
      <div style="color:#8e44ad;font-size:20px;font-weight:bold;">{_top_city['city']}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">{_top_city['total_score']} 分 · {_top_city['level']}</div>
    </div>
  </div>
</div>
"""

                    health_toolbar_js = replenish_toolbar_js.replace(
                        "var names=['预估60天销量','总库存','补货建议'];",
                        "var names=['总分','SKU覆盖分','销售规模分','单价水平分','周转健康分'];"
                    ).replace(
                        "// 默认开启【实际数值】\n        btn4.onclick();",
                        "// 评分为0-100分，保持归一化默认"
                    )

                    with open(health_output, "r", encoding="utf-8") as f:
                        health_html = f.read()
                    health_html = health_html.replace("</head>", health_js_data + "\n</head>", 1)
                    health_html = re.sub(r"(<body[^>]*>)", r"\1\n" + kpi_html + health_kpi_html, health_html, count=1)
                    health_html = health_html.replace("</body>", health_toolbar_js + "\n</body>", 1)
                    with open(health_output, "w", encoding="utf-8") as f:
                        f.write(health_html)
                    print(f"已生成：{health_output}")
            except Exception as _e:
                print(f"城市健康度看板生成失败：{_e}")

            # ── 为汇总（编号 00_0）生成看板说明文档 ──────────────────
            try:
                readme_filename = "00_0【说明】看板说明.html"
                readme_output = os.path.join(output_product_dir, readme_filename)
                readme_html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>BI 看板说明文档</title>
<style>
  body {{ font-family: 'Microsoft YaHei', 'PingFang SC', sans-serif; background: #f5f7fa; color: #2c3e50; line-height: 1.7; margin: 0; padding: 30px; }}
  .container {{ max-width: 1100px; margin: 0 auto; background: #fff; border-radius: 14px; box-shadow: 0 4px 24px rgba(0,0,0,0.06); padding: 40px 50px; }}
  h1 {{ color: #2c7be5; border-bottom: 3px solid #2c7be5; padding-bottom: 12px; margin-top: 0; }}
  h2 {{ color: #2c3e50; border-left: 5px solid #2c7be5; padding-left: 14px; margin-top: 36px; background: #f0f5ff; padding: 10px 14px; border-radius: 0 6px 6px 0; }}
  h3 {{ color: #27ae60; margin-top: 22px; }}
  .meta {{ color: #888; font-size: 13px; margin-bottom: 20px; }}
  .nav {{ background: #f0f5ff; padding: 16px 20px; border-radius: 8px; margin-bottom: 28px; border: 1px solid #d0dff5; }}
  .nav a {{ display: inline-block; margin: 4px 10px 4px 0; padding: 4px 12px; background: #fff; color: #2c7be5; text-decoration: none; border-radius: 14px; font-size: 13px; border: 1px solid #b8d9f0; }}
  .nav a:hover {{ background: #2c7be5; color: #fff; }}
  .card {{ background: #fafbfd; border-radius: 10px; padding: 20px 24px; margin: 16px 0; border-left: 5px solid #2c7be5; box-shadow: 0 2px 10px rgba(0,0,0,0.03); }}
  .card.hot {{ border-left-color: #e74c3c; }}
  .card.warn {{ border-left-color: #f39c12; }}
  .card.good {{ border-left-color: #27ae60; }}
  .card.info {{ border-left-color: #8e44ad; }}
  .label {{ display: inline-block; background: #2c7be5; color: #fff; padding: 2px 10px; border-radius: 10px; font-size: 12px; margin-right: 8px; }}
  .label.hot {{ background: #e74c3c; }}
  .label.warn {{ background: #f39c12; }}
  .label.good {{ background: #27ae60; }}
  .label.info {{ background: #8e44ad; }}
  table {{ width: 100%; border-collapse: collapse; margin: 14px 0; font-size: 14px; }}
  th, td {{ padding: 10px 12px; text-align: left; border-bottom: 1px solid #eee; }}
  th {{ background: #f0f5ff; color: #2c3e50; font-weight: bold; }}
  code {{ background: #f4f4f4; padding: 2px 6px; border-radius: 4px; color: #c0392b; font-family: Consolas, Monaco, monospace; font-size: 13px; }}
  .formula {{ background: #fffef0; border: 1px dashed #f0d066; padding: 10px 16px; border-radius: 6px; margin: 10px 0; font-family: Consolas, monospace; color: #b26a00; }}
  .tip {{ background: #e8f4fd; border-left: 4px solid #2c7be5; padding: 10px 14px; margin: 12px 0; border-radius: 0 4px 4px 0; color: #1a5490; }}
  .warn-box {{ background: #fff5f5; border-left: 4px solid #e74c3c; padding: 10px 14px; margin: 12px 0; border-radius: 0 4px 4px 0; color: #8b2a2a; }}
  .good-box {{ background: #f0fff4; border-left: 4px solid #27ae60; padding: 10px 14px; margin: 12px 0; border-radius: 0 4px 4px 0; color: #1e6b3a; }}
  .toc {{ font-size: 13px; color: #666; }}
  ul {{ padding-left: 22px; }}
  li {{ margin: 4px 0; }}
</style>
</head>
<body>
<div class="container">
  <h1>📊 编号00 全部商品汇总看板说明文档</h1>
  <div class="meta">
    数据源：<code>{os.path.basename(file_path)}</code> ｜ 统计天数：<b>{num_days}</b> 天 ｜ 统计日期（最新快照）：<b>{_latest_date}</b>
  </div>

  <div class="nav">
    <b style="color:#2c3e50;">📑 快速导航：</b><br/>
    <a href="#common">通用规则</a>
    <a href="#f1">①分城市</a>
    <a href="#f2">②分单品xlsx</a>
    <a href="#f3">③分单品html</a>
    <a href="#f4">④补货建议</a>
    <a href="#f5">⑤动销差预警</a>
    <a href="#f6">⑥爆品扩城市</a>
    <a href="#f7">⑦库存结构</a>
    <a href="#f8">⑧日销趋势</a>
    <a href="#f9">⑨城市健康度</a>
    <a href="#single">⑩单品明细页</a>
    <a href="#tips">使用建议</a>
  </div>

  <h2 id="common">🔑 通用规则与关键概念</h2>

  <div class="card info">
    <h3>1. 数据清洗</h3>
    <ul>
      <li><b>销量修正</b>：原始数据中偶尔出现"销售额>0 但销售量=0"的情况（业务端销量为1时偶尔不计入），清洗时自动修正为销售量=1</li>
      <li><b>"全国"行排除</b>：源数据中"全国"行是汇总行，所有统计排除该行避免翻倍</li>
      <li><b>城市名清洗</b>：去除括号内容（如 <code>北京（华北）</code> → <code>北京</code>）</li>
    </ul>
  </div>

  <div class="card info">
    <h3>2. 库存快照策略</h3>
    <div class="tip">
      库存是<b>状态量</b>（某时刻的快照），销售是<b>流量</b>（一段时间的累计）。为避免多日库存累加导致虚高，<b>库存仅保留最新日期的值</b>（当前：<code>{_latest_date}</code>）。
    </div>
    <p>总库存 = 供应商到大仓在途 + 大仓库存 + 大仓到门店在途 + 前置站点库存（4 段汇总）</p>
  </div>

  <div class="card info">
    <h3>3. 核心公式</h3>
    <div class="formula">
      单日销量 = 商品销售量 ÷ 统计天数（<b>{num_days}</b>天）<br/>
      预估60天销量 = 单日销量 × 60<br/>
      周转天数 = 总库存 ÷ 单日销量<br/>
      单价 = 销售额 ÷ 销售量
    </div>
    <p><b>周转天数</b>：数值越小说明库存周转越快（越好），越大说明库存积压越严重。一般阈值参考：≤30 天极健康，30-60 健康，60-180 积压，>180 严重积压。</p>
  </div>

  <div class="card info">
    <h3>4. 双编号约定</h3>
    <p>补货建议、动销差预警、爆品扩城市、库存结构四个看板中商品名格式为：</p>
    <div class="formula">【文件编号】商品名【顺序编号】</div>
    <ul>
      <li><b>文件编号</b>：对应单品 HTML 文件名前缀（按全站销售额降序的原始排名）。看到 <code>【05】</code> 就知道是销售额第 5 名的商品，可直接翻到 <code>05【...】xxx.html</code> 查阅明细。</li>
      <li><b>顺序编号</b>：当前看板按自身核心指标排序后的位次。</li>
    </ul>
  </div>

  <h2 id="f1">① 全部商品汇总_分城市.html</h2>
  <div class="card">
    <p><span class="label">地图/排行榜</span>省份 + 城市两级销售地理分布</p>
    <h3>图表内容</h3>
    <ul>
      <li>省份销售额地图（中国地图，颜色深浅=销售额）</li>
      <li>省份销售额排行榜（横向柱状图）</li>
      <li>城市销售分布散点地图</li>
      <li>城市销售额排行榜（横向柱状图）</li>
    </ul>
    <h3>用途</h3>
    <p>快速看全国销售分布、识别主力市场、发现销售薄弱省份城市。</p>
  </div>

  <h2 id="f2">② 全部商品汇总_分单品.xlsx</h2>
  <div class="card info">
    <p><span class="label info">数据表</span>聚合后的原始明细表</p>
    <h3>包含 Sheet</h3>
    <ul>
      <li><b>商品汇总</b>：每个商品的销售额、销售量、预估60天销量、总库存、单价、周转天数</li>
      <li><b>省份汇总</b>：每个省份的聚合数据</li>
      <li><b>城市汇总</b>：每个城市的聚合数据</li>
    </ul>
    <h3>用途</h3>
    <p>数据兜底，需要自行在 Excel 中做筛选、透视、二次加工时使用。</p>
  </div>

  <h2 id="f3">③ 全部商品汇总_分单品.html</h2>
  <div class="card">
    <p><span class="label">排行榜</span>所有商品的销售额排行榜</p>
    <h3>图表内容</h3>
    <p>横向柱状图（所有商品按销售额降序），5 个指标可切换：总销售额 / 销售量 / 预估60天销量 / 总库存 / 周转天数。</p>
    <h3>工具按钮</h3>
    <ul>
      <li><b>全部不选</b>：清空所有 legend，用于单独勾选某一指标</li>
      <li><b>重置排名</b>：按当前勾选的指标之和重新排序</li>
      <li><b>单选模式</b>：点击 legend 时自动只保留被点的指标</li>
      <li><b>实际数值</b>：默认归一化到相对比例，开启后显示真实数值</li>
    </ul>
  </div>

  <h2 id="f4">④ 全部商品汇总_补货建议.html</h2>
  <div class="card hot">
    <p><span class="label hot">决策看板</span>缺多少补多少</p>
    <h3>核心公式</h3>
    <div class="formula">
      补货建议 = max(0, 预估60天销量 − 总库存)<br/>
      库存覆盖率 = 总库存 ÷ 预估60天销量 × 100%
    </div>
    <h3>紧急度分档</h3>
    <table>
      <tr><th>覆盖率</th><th>分档</th><th>含义</th></tr>
      <tr><td>&lt;30%</td><td style="color:#e74c3c;"><b>紧急</b></td><td>库存严重不足，立即补货</td></tr>
      <tr><td>30%-70%</td><td style="color:#f39c12;">建议</td><td>建议近期补货</td></tr>
      <tr><td>70%-100%</td><td style="color:#f1c40f;">关注</td><td>可延后补货</td></tr>
      <tr><td>≥100%</td><td style="color:#27ae60;">充足</td><td>库存已覆盖60天，无需补货</td></tr>
    </table>
    <p><b>默认开启实际数值模式</b>（补货件数需要具体数字）。排序按补货建议件数降序。</p>
    <div class="tip">无销售历史（销售量=0）的商品不参与排序，显示在末尾。</div>
  </div>

  <h2 id="f5">⑤ 全部商品汇总_动销差与库存积压预警.html</h2>
  <div class="card warn">
    <p><span class="label warn">预警看板</span>卖不出的都在这里</p>
    <h3>筛选条件（只显示问题商品）</h3>
    <ul>
      <li><b>零动销</b>：销售量 = 0 且 库存 > 0 → 完全卖不动</li>
      <li><b>动销差</b>：有销售但周转天数 > 60 天 → 卖不完60天库存</li>
    </ul>
    <h3>核心公式</h3>
    <div class="formula">
      积压件数 = 零动销品: 全部库存<br/>
                = 有销售品: max(0, 库存 − 预估60天销量)<br/>
      积压金额 = 积压件数 × 单价（零动销无自身单价时用全局均单价）
    </div>
    <h3>状态分档</h3>
    <table>
      <tr><th>状态</th><th>判定</th><th>建议</th></tr>
      <tr><td style="color:#7f8c8d;">零动销</td><td>销量=0, 库存>0</td><td>调拨、促销、退货或淘汰</td></tr>
      <tr><td style="color:#d35400;">严重积压</td><td>周转天数 > 180天</td><td>降价促销、停止补货</td></tr>
      <tr><td style="color:#e74c3c;">库存积压</td><td>周转 60-180 天</td><td>控制补货节奏，观察</td></tr>
    </table>
    <p>按积压件数降序排序，默认开启实际数值模式。</p>
  </div>

  <h2 id="f6">⑥ 全部商品汇总_爆品扩城市推荐.html</h2>
  <div class="card good">
    <p><span class="label good">机会看板</span>把卖得好的推到没卖的城市去</p>
    <h3>爆品定义</h3>
    <ul>
      <li>销售量 > 0（有销售历史）</li>
      <li>周转天数 ≤ 60 天（动销健康，不积压）</li>
    </ul>
    <h3>核心公式</h3>
    <div class="formula">
      全国运营城市集 = 所有商品有销量的城市并集<br/>
      某爆品缺失城市 = 全国运营城市集 − 该商品已有销量的城市集<br/>
      单城市均销 = 销售额 ÷ 已覆盖城市数<br/>
      潜力扩城销售额 = 单城市均销 × 缺失城市数
    </div>
    <h3>视图</h3>
    <ul>
      <li><b>顶部柱状图</b>：Top30 爆品（按潜力降序），默认只显示"潜力扩城销售额"和"缺失城市数"</li>
      <li><b>底部卡片列表</b>：每个爆品一张卡片，用橙色标签墙直接列出缺失城市名，可直接抄给业务</li>
    </ul>
    <div class="tip">
      "潜力"是基于"已覆盖城市的平均表现 × 缺失数"的<b>粗估</b>，真实扩城效果还取决于城市消费力、渠道基础等因素。
    </div>
  </div>

  <h2 id="f7">⑦ 全部商品汇总_库存结构分析.html</h2>
  <div class="card info">
    <p><span class="label info">结构诊断</span>库存在供应链哪一段</p>
    <h3>四段库存</h3>
    <table>
      <tr><th>段位</th><th>含义</th><th>颜色</th></tr>
      <tr><td>供应商到大仓在途</td><td>供应商已发货，未入大仓</td><td style="color:#3498db;">蓝色</td></tr>
      <tr><td>大仓库存</td><td>已入大仓，未分拨</td><td style="color:#27ae60;">绿色</td></tr>
      <tr><td>大仓到门店在途</td><td>已分拨，未到门店</td><td style="color:#f39c12;">橙色</td></tr>
      <tr><td>前置站点库存</td><td>门店/前置站点可售</td><td style="color:#e74c3c;">红色</td></tr>
    </table>
    <h3>异常判定</h3>
    <ul>
      <li><b>⚠ 前置空仓</b>：前置站点 = 0 且总库存 > 0 → 货在大仓/在途，卖不到消费者手上</li>
      <li><b>⚠ 在途滞留</b>：(供应商在途 + 大仓到门店在途) / 总库存 > 50% → 大量库存卡在运输环节</li>
    </ul>
    <p>Top30 商品按总库存降序展示，堆叠柱状图直接可见每段占比。</p>
    <div class="warn-box">
      <b>典型场景</b>：总库存显示充足但前置空仓 → 分城市看板中该商品仍会出现缺货，需要<b>加速分拨</b>而不是<b>补进货</b>。
    </div>
  </div>

  <h2 id="f8">⑧ 全部商品汇总_日销趋势.html</h2>
  <div class="card">
    <p><span class="label">时间维度</span>按日期看走势</p>
    <h3>前提</h3>
    <p>仅在<b>多日数据</b>（num_days > 1）时生成。单日文件会自动跳过。</p>
    <h3>图表内容</h3>
    <ul>
      <li><b>图1</b>：全国日销售额（左轴）+ 日销售量（右轴），双轴折线 + 面积填充</li>
      <li><b>图2</b>：Top 10 商品日销售额折线对比（商品名前已加顺序编号）</li>
    </ul>
    <h3>KPI</h3>
    <p>统计天数、日均销售额、最高/最低日及日期、首末日环比（正负带颜色）</p>
    <div class="tip">图表底部有 datazoom 滑块，可缩放查看特定日期范围。</div>
  </div>

  <h2 id="f9">⑨ 全部商品汇总_城市健康度.html</h2>
  <div class="card good">
    <p><span class="label good">综合评分</span>城市运营质量打分</p>
    <h3>四维度加权评分</h3>
    <table>
      <tr><th>维度</th><th>公式</th><th>权重</th></tr>
      <tr><td>SKU 覆盖率</td><td>城市 SKU 数 ÷ 全国 SKU 总数</td><td>30%</td></tr>
      <tr><td>销售规模</td><td>log(销售额+1) ÷ log(最大销售额+1)</td><td>25%</td></tr>
      <tr><td>单价水平</td><td>城市均单价 ÷ 最大均单价</td><td>15%</td></tr>
      <tr><td>周转健康度</td><td>30天=满分，180天=0分 线性递减</td><td>30%</td></tr>
    </table>
    <p>各维度归一化到 0-100 分后加权，总分 0-100。</p>
    <h3>分级（柱状图自动染色）</h3>
    <table>
      <tr><th>分段</th><th>分级</th><th>颜色</th></tr>
      <tr><td>≥80</td><td><b>标杆</b></td><td style="color:#27ae60;">绿色</td></tr>
      <tr><td>60-80</td><td>健康</td><td style="color:#2c7be5;">蓝色</td></tr>
      <tr><td>40-60</td><td>待改善</td><td style="color:#f39c12;">橙色</td></tr>
      <tr><td>&lt;40</td><td>问题</td><td style="color:#e74c3c;">红色</td></tr>
    </table>
    <div class="tip">销售规模用<b>对数</b>而非线性，避免一线城市拉爆所有维度；权重侧重"面"（SKU 覆盖 30% + 周转 30%）而非"量"（销售规模 25%）。</div>
  </div>

  <h2 id="single">⑩ 单品明细页 XX【...】商品名.html</h2>
  <div class="card">
    <p><span class="label">单品钻取</span>每个商品一张页</p>
    <h3>文件名含义</h3>
    <div class="formula">
      XX【销售额 销售量 总库存 单价 周转天数】商品名.html
    </div>
    <p>XX 是销售额降序的排名（00 是全部商品汇总，01 是第一名）。数字部分是简写（w=万，k=千）。</p>
    <h3>页面内容</h3>
    <ul>
      <li>顶部 KPI：全国总销售额/销售量/库存/均单价/周转天数</li>
      <li>省份销售额地图</li>
      <li>省份销售额排行榜（可切换销量/库存/周转等 5 指标）</li>
      <li>城市销售分布散点地图</li>
      <li>城市销售额排行榜（5 指标切换）</li>
    </ul>
    <p>从补货建议/动销差预警等看板看到 <code>【05】</code> 时，对应 <code>05【...】商品名.html</code> 查钻明细。</p>
  </div>

  <h2 id="tips">💡 使用建议</h2>

  <div class="card good">
    <h3>每日/每周例行检查流程</h3>
    <ol>
      <li><b>先看问题</b>：打开 <code>⑤动销差预警</code> → 锁定要清仓/调拨的积压品</li>
      <li><b>再看机会</b>：打开 <code>④补货建议</code> → 紧急档商品立即补货；打开 <code>⑥爆品扩城市</code> → 给业务下发扩城清单</li>
      <li><b>结构诊断</b>：打开 <code>⑦库存结构</code> → 检查是否有"货在大仓卖不到门店"的结构问题</li>
      <li><b>城市层面</b>：打开 <code>⑨城市健康度</code> → 抓问题城市（&lt;40分）重点运营</li>
      <li><b>趋势观察</b>：打开 <code>⑧日销趋势</code> → 看整体走势是否健康</li>
    </ol>
  </div>

  <div class="card warn">
    <h3>常见误读避免</h3>
    <ul>
      <li><b>周转天数虚高</b>：当单日销量极小时（比如 0.1 件/天），周转会被算得极大（几千天）。这种情况请结合"总库存件数"判断是否真积压。</li>
      <li><b>新品问题</b>：新上架商品销售历史短，预估60天销量可能低估或高估。参考时需结合自身业务判断。</li>
      <li><b>补货建议的前提是未来60天销量 ≈ 历史日销 × 60</b>。季节性商品、促销商品、新品需人工调整。</li>
      <li><b>爆品扩城潜力是上界粗估</b>，真实扩城效果取决于当地消费力、渠道、物流等多重因素，是决策参考而非承诺值。</li>
    </ul>
  </div>

  <div class="card info">
    <h3>颜色编码约定</h3>
    <ul>
      <li>🟢 绿色：健康 / 达标 / 正面</li>
      <li>🔵 蓝色：主要指标 / 中性</li>
      <li>🟡 黄色 / 🟠 橙色：关注 / 警告</li>
      <li>🔴 红色：紧急 / 异常 / 核心指标（补货件数/积压件数）</li>
      <li>🟣 紫色：特殊维度（库存结构 / 最高分标记等）</li>
    </ul>
  </div>

  <div class="meta" style="margin-top:40px;text-align:center;border-top:1px solid #eee;padding-top:20px;">
    本说明文档自动生成 · 如对某个计算规则有异议可反馈调整
  </div>
</div>
</body>
</html>
"""
                with open(readme_output, "w", encoding="utf-8") as f:
                    f.write(readme_html)
                print(f"已生成：{readme_output}")
            except Exception as _e:
                print(f"看板说明文档生成失败：{_e}")


# ── 批量处理 02 生成\02 表 目录 ──────────────────────────
TABLE_DIR  = os.path.join(INPUT_DIR, "02 表")
OUTPUT_DIR = os.path.join(INPUT_DIR, "01 BI")
os.makedirs(OUTPUT_DIR, exist_ok=True)

xlsx_files = glob.glob(os.path.join(TABLE_DIR, "*.xlsx"))
if not xlsx_files:
    print(f"未找到任何 xlsx 文件：{TABLE_DIR}")
else:
    for fp in xlsx_files:
        try:
            process_file(fp, OUTPUT_DIR)
        except Exception as e:
            print(f"处理失败 [{fp}]: {e}")
            import traceback
            traceback.print_exc()
