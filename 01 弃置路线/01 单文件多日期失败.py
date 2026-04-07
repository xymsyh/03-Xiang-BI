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
file_path    = os.path.join(BASE_DIR, "01 商品明细_2026-04-07.xlsx")
mapping_file = os.path.join(BASE_DIR, "01b 映射.json")
output_file  = "xiaoxiang_bi_fixed.html"
tmp_file     = os.path.join(BASE_DIR, "_tmp_chart.html")

# ── 加载映射表 ────────────────────────────────────────────
with open(mapping_file, encoding="utf-8") as f:
    _mapping = json.load(f)
city_to_province    = _mapping["city_to_province"]
province_full_names = _mapping["province_full_names"]

def to_echarts_province(name):
    return province_full_names.get(name, name + "省")

# ── 读取数据 ──────────────────────────────────────────────
try:
    df_raw = pd.read_excel(file_path) if file_path.endswith(".xlsx") else pd.read_csv(file_path)
except Exception as e:
    print(f"数据读取出错: {e}")
    exit()

# A 列为日期（格式 20260401），统一转字符串
date_col = df_raw.columns[0]
df_raw[date_col] = df_raw[date_col].astype(str).str.strip()
dates = sorted(df_raw[date_col].unique())
print(f"发现日期共 {len(dates)} 个：{dates}")

# ── 单日数据清洗与聚合 ────────────────────────────────────
def process_date(df_d):
    df = df_d.copy()
    df["商品销售额"] = pd.to_numeric(df["商品销售额"], errors="coerce").fillna(0)
    df["商品销售量"] = pd.to_numeric(df["商品销售量"], errors="coerce").fillna(0)
    df["总库存"] = (
        pd.to_numeric(df["供应商到大仓在途数量"], errors="coerce").fillna(0) +
        pd.to_numeric(df["大仓库存数量"],         errors="coerce").fillna(0) +
        pd.to_numeric(df["大仓到门店在途数量"],   errors="coerce").fillna(0) +
        pd.to_numeric(df["前置站点库存数量"],     errors="coerce").fillna(0)
    )
    df["城市_清洗"] = df["城市"].astype(str).apply(lambda x: re.sub(r"（.*?）", "", x))

    df_city = (
        df.groupby("城市_清洗")
          .agg(商品销售额=("商品销售额", "sum"),
               商品销售量=("商品销售量", "sum"),
               总库存    =("总库存",     "sum"))
          .round(2).reset_index()
    )
    df_city["单价"]     = (df_city["商品销售额"] / df_city["商品销售量"].replace(0, float("nan"))).round(2)
    df_city["城市_标准"] = df_city["城市_清洗"].apply(lambda x: re.sub(r"市$", "", x))
    df_city["省份"]     = df_city["城市_标准"].map(city_to_province)

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

    national_row = df_city[df_city["城市_清洗"] == "全国"]
    city_only    = df_city[df_city["城市_清洗"] != "全国"].sort_values("商品销售额", ascending=False)

    if not national_row.empty:
        nt = float(national_row["商品销售额"].iloc[0])
        nq = int(national_row["商品销售量"].iloc[0])
        ns = int(national_row["总库存"].iloc[0])
    else:
        nt = float(city_only["商品销售额"].sum())
        nq = int(city_only["商品销售量"].sum())
        ns = int(city_only["总库存"].sum())
    np_ = round(nt / nq, 2) if nq else None

    return df_city, df_province, city_only, nt, nq, ns, np_

# ── 构建单日4张图表 ───────────────────────────────────────
def build_charts(df_city, df_province, city_only, d):
    """d 为日期字符串，如 '20260401'，用于 chart_id 和 JS 变量名"""

    def _row_dict(row):
        return {
            "销售量": int(row["商品销售量"]),
            "总库存": int(row["总库存"]),
            "单价":   round(float(row["单价"]), 2) if pd.notna(row["单价"]) else None,
        }

    prov_dict = {}
    for _, row in df_province.iterrows():
        rd = _row_dict(row)
        prov_dict[row["省份"]]                        = rd
        prov_dict[to_echarts_province(row["省份"])]   = rd
    city_dict = {row["城市_清洗"]: _row_dict(row) for _, row in df_city.iterrows()}

    vp = f"PD_{d}"   # JS 变量名：省份数据
    vc = f"CD_{d}"   # JS 变量名：城市数据

    province_map_data = [[to_echarts_province(p), float(v)]
                         for p, v in zip(df_province["省份"], df_province["商品销售额"])]

    # 省份地图
    chart_province_map = (
        Map(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="500px",
                                    chart_id=f"pmap_{d}"))
        .add("省份总销售额(元)", province_map_data, "china")
        .set_global_opts(
            title_opts=opts.TitleOpts(title="省份销售额分布"),
            visualmap_opts=opts.VisualMapOpts(max_=float(df_province["商品销售额"].max()), is_piecewise=False),
            tooltip_opts=opts.TooltipOpts(formatter=JsCode(
                f"function(p){{var d={vp}[p.name]||{{}};"
                "return p.name+'<br/>销售额: ¥'+p.value+' 元'"
                "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-');}"
            )),
        )
    )

    # 省份排行榜
    _ph = f"{max(500, len(df_province) * 32 + 80)}px"
    chart_province_bar = (
        Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height=_ph,
                                    chart_id=f"pbar_{d}"))
        .add_xaxis(df_province["省份"].tolist()[::-1])
        .add_yaxis("总销售额", df_province["商品销售额"].tolist()[::-1],
                   label_opts=opts.LabelOpts(position="right"))
        .add_yaxis("总库存",   df_province["总库存"].tolist()[::-1],
                   label_opts=opts.LabelOpts(position="right"))
        .reversal_axis()
        .set_global_opts(
            title_opts=opts.TitleOpts(title="省份销售额排行榜（全部）"),
            xaxis_opts=opts.AxisOpts(name="金额 / 库存"),
            legend_opts=opts.LegendOpts(pos_top="35px", pos_left="center"),
            tooltip_opts=opts.TooltipOpts(trigger="axis", formatter=JsCode(
                f"function(ps){{var p=ps[0],d={vp}[p.name]||{{}};"
                "return p.name+'<br/>销售额: ¥'+p.value+' 元'"
                "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-');}"
            )),
        )
    )

    # 城市地图
    city_map_data = [list(z) for z in zip(city_only["城市_清洗"], city_only["商品销售额"])]
    chart_city_map = (
        Geo(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height="500px",
                                    chart_id=f"cmap_{d}"))
        .add_schema(maptype="china")
        .add("总销售额(元)", city_map_data, type_="scatter")
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(
            title_opts=opts.TitleOpts(title="城市销售分布图"),
            visualmap_opts=opts.VisualMapOpts(max_=float(city_only["商品销售额"].max()), is_piecewise=False),
            tooltip_opts=opts.TooltipOpts(formatter=JsCode(
                f"function(p){{var d={vc}[p.name]||{{}};"
                "return p.name+'<br/>销售额: ¥'+p.value[2]+' 元'"
                "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-');}"
            )),
        )
    )

    # 城市排行榜
    city_rank = city_only.copy()
    city_rank["城市_轴标"] = city_rank.apply(
        lambda r: f"（{r['省份']}）{r['城市_清洗']}" if pd.notna(r.get("省份")) else r["城市_清洗"],
        axis=1,
    )
    _ch = f"{max(500, len(city_rank) * 32 + 80)}px"
    chart_city_bar = (
        Bar(init_opts=opts.InitOpts(theme=ThemeType.MACARONS, width="100%", height=_ch,
                                    chart_id=f"cbar_{d}"))
        .add_xaxis(city_rank["城市_轴标"].tolist()[::-1])
        .add_yaxis("总销售额", city_rank["商品销售额"].tolist()[::-1],
                   label_opts=opts.LabelOpts(position="right"))
        .add_yaxis("总库存",   city_rank["总库存"].tolist()[::-1],
                   label_opts=opts.LabelOpts(position="right"))
        .reversal_axis()
        .set_global_opts(
            title_opts=opts.TitleOpts(title="城市销售额排行榜（全部）"),
            xaxis_opts=opts.AxisOpts(name="金额 / 库存"),
            legend_opts=opts.LegendOpts(pos_top="35px", pos_left="center"),
            tooltip_opts=opts.TooltipOpts(trigger="axis", formatter=JsCode(
                f"function(ps){{"
                "var p=ps[0];"
                "var m=p.name.match(/^（(.+)）(.+)$/);"
                "var city=m?m[2]:p.name,prov=m?m[1]:'';"
                "var label=m?city+'（'+prov+'）':city;"
                f"var d={vc}[city]||{{}};"
                "return label+'<br/>销售额: ¥'+p.value+' 元'"
                "+'<br/>销售量: '+(d.销售量!=null?d.销售量+' 件':'-')"
                "+'<br/>总库存: '+(d.总库存!=null?d.总库存+' 件':'-')"
                "+'<br/>单价: '+(d.单价!=null?'¥'+d.单价:'-');}"
            )),
        )
    )

    return chart_province_map, chart_province_bar, chart_city_map, chart_city_bar, prov_dict, city_dict

# ── 渲染每日 Page 到临时文件，提取 head / body ─────────────
def render_page_fragments(charts):
    page = Page(layout=Page.SimplePageLayout)
    page.add(*charts)
    page.render(tmp_file)
    with open(tmp_file, encoding="utf-8") as f:
        html = f.read()
    hs = html.find("<head>"); he = html.find("</head>")
    bs = html.find("<body>"); be = html.rfind("</body>")
    head = html[hs + 6:he] if hs != -1 and he != -1 else ""
    body = html[bs + 6:be] if bs != -1 and be != -1 else ""
    return head, body

# ── KPI 卡片 ──────────────────────────────────────────────
def make_kpi(nt, nq, ns, np_):
    price_str = f"¥ {np_:,.2f}" if np_ else "-"
    return f"""
<div style="font-family:'Microsoft YaHei',sans-serif;background:#f0f4ff;padding:18px 24px 14px;border-bottom:2px solid #d0d8f0;margin-bottom:4px;">
  <div style="display:flex;justify-content:center;gap:20px;flex-wrap:wrap;">
    <div style="background:#fff;border-radius:12px;padding:16px 28px;box-shadow:0 2px 12px rgba(44,123,229,0.11);text-align:center;min-width:160px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">全国总销售额</div>
      <div style="color:#2c7be5;font-size:28px;font-weight:bold;">¥ {nt:,.2f}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">元</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:16px 28px;box-shadow:0 2px 12px rgba(44,123,229,0.11);text-align:center;min-width:160px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">全国总销售量</div>
      <div style="color:#27ae60;font-size:28px;font-weight:bold;">{nq:,}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">件</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:16px 28px;box-shadow:0 2px 12px rgba(44,123,229,0.11);text-align:center;min-width:160px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">全国总库存</div>
      <div style="color:#e67e22;font-size:28px;font-weight:bold;">{ns:,}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">件</div>
    </div>
    <div style="background:#fff;border-radius:12px;padding:16px 28px;box-shadow:0 2px 12px rgba(44,123,229,0.11);text-align:center;min-width:160px;">
      <div style="color:#888;font-size:12px;letter-spacing:1px;margin-bottom:6px;">全国平均单价</div>
      <div style="color:#8e44ad;font-size:28px;font-weight:bold;">{price_str}</div>
      <div style="color:#bbb;font-size:11px;margin-top:3px;">元/件</div>
    </div>
  </div>
</div>"""

# ── 主流程：处理所有日期 ──────────────────────────────────
common_head  = None
date_panels  = {}   # d -> {"kpi": str, "body": str, "prov_dict": dict, "city_dict": dict}

for d in dates:
    df_d = df_raw[df_raw[date_col] == d]
    try:
        df_city, df_province, city_only, nt, nq, ns, np_ = process_date(df_d)
        pmap, pbar, cmap, cbar, prov_dict, city_dict = build_charts(df_city, df_province, city_only, d)
        head, body = render_page_fragments([pmap, pbar, cmap, cbar])
        if common_head is None:
            common_head = head
        date_panels[d] = {
            "kpi":       make_kpi(nt, nq, ns, np_),
            "body":      body,
            "prov_dict": prov_dict,
            "city_dict": city_dict,
        }
        print(f"  ✓ {d}")
    except Exception as e:
        print(f"  ✗ {d} 出错: {e}")

if os.path.exists(tmp_file):
    os.remove(tmp_file)

# ── 日期切换按钮 ──────────────────────────────────────────
def fmt_date(d):
    return f"{d[:4]}-{d[4:6]}-{d[6:]}" if len(d) == 8 else d

btn_html = ""
for i, d in enumerate(date_panels):
    active_style = "background:#2c7be5;color:#fff;" if i == 0 else "background:#fff;color:#2c7be5;"
    btn_html += (
        f'<button id="btn_{d}" onclick="switchDate(\'{d}\')" '
        f'style="border:2px solid #2c7be5;border-radius:8px;padding:8px 18px;cursor:pointer;'
        f'font-size:14px;font-family:\'Microsoft YaHei\',sans-serif;{active_style}transition:all 0.2s;">'
        f'{fmt_date(d)}</button>\n'
    )

switcher_html = f"""
<div id="date-switcher" style="font-family:'Microsoft YaHei',sans-serif;background:#fff;
  padding:16px 24px;border-bottom:2px solid #e8eef8;text-align:center;
  position:sticky;top:0;z-index:999;box-shadow:0 2px 8px rgba(0,0,0,0.08);">
  <span style="color:#555;font-size:13px;margin-right:12px;">选择日期：</span>
  <div style="display:inline-flex;gap:10px;flex-wrap:wrap;justify-content:center;">
{btn_html}  </div>
</div>"""

# ── JS 数据 + 切换逻辑 ─────────────────────────────────────
js_vars = ""
for d, res in date_panels.items():
    js_vars += (
        f"var PD_{d}={json.dumps(res['prov_dict'], ensure_ascii=False)};\n"
        f"var CD_{d}={json.dumps(res['city_dict'], ensure_ascii=False)};\n"
    )

chart_prefixes = json.dumps(["pmap_", "pbar_", "cmap_", "cbar_"])
first_date = list(date_panels.keys())[0] if date_panels else ""

switch_js = f"""
<script>
{js_vars}
var _activeDate = '{first_date}';

function switchDate(d) {{
  if (d === _activeDate) return;
  _activeDate = d;

  // 单选：隐藏所有面板，显示选中
  document.querySelectorAll('.date-panel').forEach(function(el) {{
    el.style.display = 'none';
  }});
  var panel = document.getElementById('panel_' + d);
  if (panel) panel.style.display = 'block';

  // 按钮高亮（单选）
  document.querySelectorAll('[id^="btn_"]').forEach(function(btn) {{
    btn.style.background = '#fff';
    btn.style.color = '#2c7be5';
  }});
  var activeBtn = document.getElementById('btn_' + d);
  if (activeBtn) {{ activeBtn.style.background = '#2c7be5'; activeBtn.style.color = '#fff'; }}

  // 显示后 resize，修复隐藏期间尺寸为 0 的问题
  {chart_prefixes}.forEach(function(prefix) {{
    var el = document.getElementById(prefix + d);
    if (el) {{
      var inst = echarts.getInstanceByDom(el);
      if (inst) inst.resize();
    }}
  }});
}}

// ECharts 在可见容器上完成初始化后，再隐藏非当前面板
window.addEventListener('load', function() {{
  setTimeout(function() {{
    document.querySelectorAll('.date-panel').forEach(function(el) {{
      if (el.id !== 'panel_{first_date}') el.style.display = 'none';
    }});
    // resize 当前面板图表
    {chart_prefixes}.forEach(function(prefix) {{
      var el = document.getElementById(prefix + '{first_date}');
      if (el) {{
        var inst = echarts.getInstanceByDom(el);
        if (inst) inst.resize();
      }}
    }});
  }}, 300);
}});
</script>"""

# ── 拼装面板（初始全部可见，让 ECharts 正常初始化，load 后再隐藏）──
panels_html = ""
for d, res in date_panels.items():
    panels_html += (
        f'<div class="date-panel" id="panel_{d}" style="display:block;">\n'
        f'{res["kpi"]}\n{res["body"]}\n</div>\n'
    )

# ── 输出最终 HTML ─────────────────────────────────────────
final_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
{common_head or ""}
</head>
<body>
{switcher_html}
{panels_html}
{switch_js}
</body>
</html>"""

with open(output_file, "w", encoding="utf-8") as f:
    f.write(final_html)

print(f"\n看板已生成：{output_file}（共 {len(date_panels)} 个日期）")
os.system(f'start msedge "{os.path.abspath(output_file)}"')
