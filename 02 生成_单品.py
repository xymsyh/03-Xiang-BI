import glob
import json
import os
import re

import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar, Geo, Map, Page
from pyecharts.commons.utils import JsCode
from pyecharts.globals import ThemeType

# ── 配置参数 ──────────────────────────────────────────────
ESTIMATED_60DAY_MULTIPLIER = 2  # 预估60天销量 = 销量 × 该倍数

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
    df["总库存"]     = (
        pd.to_numeric(df["供应商到大仓在途数量"], errors="coerce").fillna(0) +
        pd.to_numeric(df["大仓库存数量"],         errors="coerce").fillna(0) +
        pd.to_numeric(df["大仓到门店在途数量"],   errors="coerce").fillna(0) +
        pd.to_numeric(df["前置站点库存数量"],     errors="coerce").fillna(0)
    )
    df["城市_清洗"]  = df["城市"].astype(str).apply(lambda x: re.sub(r"（.*?）", "", x))

    # ── 按商品名称分组并按销售额排序 ────────────────────────
    product_groups = []
    for product_name, df_product in df.groupby("商品名称"):
        # 计算汇总数据用于排序
        total_sales = float(df_product["商品销售额"].sum())
        total_qty = int(df_product["商品销售量"].sum())
        total_stock = int(df_product["总库存"].sum())
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

    # ── 按商品名称分组 ────────────────────────────────────
    for idx, product_info in enumerate(product_groups, 1):
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
        df_city["预估60天销量"] = (df_city["商品销售量"] * ESTIMATED_60DAY_MULTIPLIER).astype(int)
        df_city["周转周期"] = (df_city["总库存"] / df_city["预估60天销量"].replace(0, float("nan"))).round(2)

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
        df_province["预估60天销量"] = (df_province["商品销售量"] * ESTIMATED_60DAY_MULTIPLIER).astype(int)
        df_province["周转周期"] = (df_province["总库存"] / df_province["预估60天销量"].replace(0, float("nan"))).round(2)

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
                "周转周期": round(float(row["周转周期"]), 2) if pd.notna(row["周转周期"]) else None,
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
                    "+'<br/>周转周期: '+(d.周转周期!=null?d.周转周期+' 次':'-');}"
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
            .add_yaxis("周转周期(次)", _normalize_series(province_rank_data["周转周期"].tolist()[::-1]),
                       label_opts=_norm_label)
            .reversal_axis()
            .set_global_opts(
                title_opts=opts.TitleOpts(title="省份销售额排行榜（全部）"),
                xaxis_opts=opts.AxisOpts(name="相对比例 (%)", max_=100),
                legend_opts=opts.LegendOpts(pos_top="40px", pos_left="center",
                    selected_map={"总销售额": True, "销售量": False, "预估60天销量": False, "总库存": True, "周转周期(次)": False}),
                tooltip_opts=opts.TooltipOpts(trigger="axis", formatter=JsCode(
                    "function(ps){var p=ps[0],d=PROVINCE_DATA[p.name]||{};"
                    "return p.name"
                    "+'<br/>销售额: '+(d.销售额!=null?'¥'+d.销售额+' 元':'-')"
                    "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                    "+'<br/>预估60天销量: '+(d.预估60天销量!=null?d.预估60天销量+' 件':'-')"
                    "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                    "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-')"
                    "+'<br/>周转周期: '+(d.周转周期!=null?d.周转周期+' 次':'-');}"
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
                    "+'<br/>周转周期: '+(d.周转周期!=null?d.周转周期+' 次':'-');}"
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
            .add_yaxis("周转周期(次)", _normalize_series(city_rank_data["周转周期"].tolist()[::-1]),
                       label_opts=_norm_label)
            .reversal_axis()
            .set_global_opts(
                title_opts=opts.TitleOpts(title="城市销售额排行榜（全部）"),
                xaxis_opts=opts.AxisOpts(name="相对比例 (%)", max_=100),
                legend_opts=opts.LegendOpts(pos_top="40px", pos_left="center",
                    selected_map={"总销售额": True, "销售量": False, "预估60天销量": False, "总库存": True, "周转周期(次)": False}),
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
                    "+'<br/>周转周期: '+(d.周转周期!=null?d.周转周期+' 次':'-');}"
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
        filename = f"{idx:02d}【{sales_fmt}  {qty_fmt}  {stock_fmt}  {price_fmt}】{safe_product_name}.html"
        output_file = os.path.join(output_product_dir, filename)

        # ── 渲染输出 ──────────────────────────────────────────────
        page = Page(layout=Page.SimplePageLayout)
        page.add(chart_province_map, chart_province_bar, chart_city_map, chart_city_bar)
        page.render(output_file)

        # ── 注入全国总计 KPI 卡片 ─────────────────────────────────
        _price_str = f"¥ {national_price:,.2f}" if national_price else "-"
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
    var names=['总销售额','销售量','预估60天销量','总库存','周转周期(次)'];
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

        wrap.appendChild(btn1);
        wrap.appendChild(btn2);
        wrap.appendChild(btn3);
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
