import glob
import json
import os
import re

import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar, Geo, Map, Page
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
                _label = f"【{_pi:02d}】{pg['name']}"
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
                    "label": _label,
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

            # 重新编号（显示在商品名前）
            for _ai, it in enumerate(_alert_items, 1):
                it["label"] = f"【{_ai:02d}】{it['name']}"

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
                # 重新编号（按潜力排序）
                for _ei, it in enumerate(_expand_items, 1):
                    it["label"] = f"【{_ei:02d}】{it['name']}"

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
