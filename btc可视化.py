import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go


EXCEL_PATH = r"D:\30y BTC.xlsx"
OUTPUT_HTML = r"D:\btc_dca_okx_dashboard.html"

# 你的 Excel 里 BTC 价格是 USDT，投入是 RMB。
# 市值和收益曲线需要一个 RMB/USDT 估算汇率。
CNY_PER_USDT = 7.10

TARGET_BTC = 0.01


def normalize_column_name(name):
    if pd.isna(name):
        return ""
    text = str(name).strip()
    text = text.replace("\n", "")
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"\s+", "", text)
    return text


def parse_date(value):
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value

    text = str(value).strip()

    # 兼容 2026.6.12
    match = re.match(r"^(\d{4})\.(\d{1,2})\.(\d{1,2})$", text)
    if match:
        year, month, day = match.groups()
        return pd.Timestamp(int(year), int(month), int(day))

    parsed = pd.to_datetime(text, errors="coerce")
    return parsed


def to_number(value):
    if pd.isna(value):
        return np.nan

    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)

    text = str(value).strip()

    if not text or text in {"—", "-", "——", "nan", "None"}:
        return np.nan

    text = text.replace(",", "")
    text = text.replace("￥", "")
    text = text.replace("RMB", "")
    text = text.replace("rmb", "")
    text = text.replace("USDT", "")
    text = text.replace("usdt", "")
    text = text.replace("USD", "")
    text = text.replace("usd", "")
    text = text.replace("BTC", "")
    text = text.replace("btc", "")
    text = text.replace("U", "u")

    # 处理 678(100u)，只取前面的人民币金额
    match = re.match(r"^([0-9.]+)\(.*\)$", text)
    if match:
        text = match.group(1)

    # 处理百分比
    if text.endswith("%"):
        try:
            return float(text[:-1]) / 100
        except ValueError:
            return np.nan

    try:
        return float(text)
    except ValueError:
        return np.nan


def find_column(columns, keywords):
    for col in columns:
        for keyword in keywords:
            if keyword in col:
                return col
    return None


def find_header_row(raw):
    for idx in range(min(30, len(raw))):
        row_text = "".join(str(x) for x in raw.iloc[idx].dropna().tolist())
        if "日期" in row_text and ("本周价格" in row_text or "价格" in row_text) and "累计" in row_text:
            return idx

    raise ValueError("没有找到表头行。请确认第一个 sheet 里有 日期、本周价格、累计持有 等列。")


def load_data(excel_path):
    raw = pd.read_excel(excel_path, sheet_name=0, header=None, dtype=object)
    header_row = find_header_row(raw)

    df = pd.read_excel(excel_path, sheet_name=0, header=header_row, dtype=object)
    df.columns = [normalize_column_name(c) for c in df.columns]
    df = df.loc[:, [c for c in df.columns if c]]

    columns = list(df.columns)

    col_no = find_column(columns, ["次数"])
    col_date = find_column(columns, ["日期"])
    col_price = find_column(columns, ["本周价格", "价格"])
    col_change = find_column(columns, ["涨跌幅"])
    col_invest = find_column(columns, ["投入金额", "投入"])
    col_rule = find_column(columns, ["触发规则", "规则"])
    col_buy_btc = find_column(columns, ["买入数量"])
    col_total_btc = find_column(columns, ["累计持有", "累计持仓"])
    col_total_cost = find_column(columns, ["总成本"])
    col_avg_cost = find_column(columns, ["平均成本"])
    col_note = find_column(columns, ["备注"])

    required = {
        "日期": col_date,
        "本周价格": col_price,
        "投入金额": col_invest,
        "买入数量": col_buy_btc,
        "累计持有": col_total_btc,
        "总成本": col_total_cost,
        "平均成本": col_avg_cost,
    }

    missing = [name for name, col in required.items() if col is None]
    if missing:
        raise ValueError(f"缺少必要列：{missing}")

    clean = pd.DataFrame()
    clean["次数"] = df[col_no].apply(to_number) if col_no else np.nan
    clean["日期"] = df[col_date].apply(parse_date)
    clean["本周价格_USDT"] = df[col_price].apply(to_number)
    clean["涨跌幅"] = df[col_change].apply(to_number) if col_change else np.nan
    clean["投入金额_RMB"] = df[col_invest].apply(to_number)
    clean["触发规则"] = df[col_rule].astype(str).str.strip() if col_rule else ""
    clean["买入数量_BTC"] = df[col_buy_btc].apply(to_number)
    clean["累计持有_BTC"] = df[col_total_btc].apply(to_number)
    clean["总成本_RMB"] = df[col_total_cost].apply(to_number)
    clean["平均成本_USDT"] = df[col_avg_cost].apply(to_number)
    clean["备注"] = df[col_note].astype(str).str.strip() if col_note else ""

    clean = clean.dropna(subset=["日期", "本周价格_USDT", "累计持有_BTC", "总成本_RMB"])
    clean = clean.sort_values("日期").reset_index(drop=True)

    clean["市值_RMB"] = clean["累计持有_BTC"] * clean["本周价格_USDT"] * CNY_PER_USDT
    clean["浮盈_RMB"] = clean["市值_RMB"] - clean["总成本_RMB"]
    clean["收益率"] = clean["浮盈_RMB"] / clean["总成本_RMB"]
    clean["0.01进度"] = clean["累计持有_BTC"] / TARGET_BTC
    clean["距离0.01_BTC"] = TARGET_BTC - clean["累计持有_BTC"]

    clean["日期文本"] = clean["日期"].dt.strftime("%Y-%m-%d")
    clean["涨跌幅百分比"] = clean["涨跌幅"] * 100

    return clean


def rule_color(rule):
    rule = str(rule).upper()

    if "生日" in rule or "礼炮" in rule:
        return "#F0B90B"
    if "戒赌" in rule:
        return "#F6465D"
    if "C" in rule:
        return "#1E88E5"
    if "B" in rule:
        return "#5DADE2"
    if "A" in rule:
        return "#2EBD85"
    if "首次" in rule or "第一次" in rule:
        return "#848E9C"

    return "#6E7781"


def change_color(value):
    if pd.isna(value):
        return "#848E9C"
    if value >= 0:
        return "#F6465D"
    return "#2EBD85"


def safe_text(value):
    if pd.isna(value):
        return ""
    text = str(value)
    if text.lower() == "nan":
        return ""
    return text


def fig_to_html(fig):
    return fig.to_html(
        full_html=False,
        include_plotlyjs=False,
        config={
            "displaylogo": False,
            "responsive": True,
            "modeBarButtonsToRemove": [
                "lasso2d",
                "select2d",
                "autoScale2d"
            ]
        }
    )


def make_price_chart(df):
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df["日期"],
        y=df["本周价格_USDT"],
        mode="lines+markers",
        name="BTC 当周价格",
        line=dict(width=3, color="#F0B90B"),
        marker=dict(size=7),
        customdata=np.stack([
            df["日期文本"],
            df["买入数量_BTC"],
            df["投入金额_RMB"],
            df["触发规则"],
            df["备注"]
        ], axis=-1),
        hovertemplate=(
            "<b>%{customdata[0]}</b><br>"
            "BTC价格：%{y:,.2f} USDT<br>"
            "本次买入：%{customdata[1]:.8f} BTC<br>"
            "投入金额：¥%{customdata[2]:,.2f}<br>"
            "触发规则：%{customdata[3]}<br>"
            "备注：%{customdata[4]}<extra></extra>"
        )
    ))

    fig.add_trace(go.Scatter(
        x=df["日期"],
        y=df["平均成本_USDT"],
        mode="lines+markers",
        name="持仓平均成本",
        line=dict(width=3, color="#2EBD85"),
        marker=dict(size=7),
        hovertemplate="<b>%{x|%Y-%m-%d}</b><br>平均成本：%{y:,.2f} USDT<extra></extra>"
    ))

    fig.update_layout(
        title="BTC 价格 vs 持仓平均成本",
        height=430,
        template="plotly_dark",
        paper_bgcolor="#0B0E11",
        plot_bgcolor="#0B0E11",
        font=dict(color="#EAECEF"),
        margin=dict(l=40, r=25, t=65, b=40),
        legend=dict(orientation="h", y=1.08, x=0),
        hovermode="x unified"
    )

    fig.update_xaxes(gridcolor="#1E2329", zerolinecolor="#1E2329")
    fig.update_yaxes(gridcolor="#1E2329", zerolinecolor="#1E2329", tickformat=",.0f")

    return fig


def make_assets_chart(df):
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df["日期"],
        y=df["累计持有_BTC"],
        mode="lines+markers",
        name="累计持有 BTC",
        fill="tozeroy",
        line=dict(width=3, color="#F0B90B"),
        marker=dict(size=7),
        hovertemplate="<b>%{x|%Y-%m-%d}</b><br>累计持有：%{y:.8f} BTC<extra></extra>"
    ))

    fig.add_hline(
        y=0.005,
        line_dash="dash",
        line_color="#848E9C",
        annotation_text="0.005 BTC",
        annotation_font_color="#848E9C"
    )

    fig.add_hline(
        y=0.01,
        line_dash="dash",
        line_color="#F0B90B",
        annotation_text="0.01 BTC 提币目标",
        annotation_font_color="#F0B90B"
    )

    fig.update_layout(
        title="累计 BTC 持仓进度",
        height=430,
        template="plotly_dark",
        paper_bgcolor="#0B0E11",
        plot_bgcolor="#0B0E11",
        font=dict(color="#EAECEF"),
        margin=dict(l=40, r=25, t=65, b=40),
        hovermode="x unified",
        showlegend=False
    )

    fig.update_xaxes(gridcolor="#1E2329", zerolinecolor="#1E2329")
    fig.update_yaxes(gridcolor="#1E2329", zerolinecolor="#1E2329")

    return fig


def make_pnl_chart(df):
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df["日期"],
        y=df["总成本_RMB"],
        mode="lines+markers",
        name="累计投入",
        line=dict(width=3, color="#848E9C"),
        marker=dict(size=6),
        hovertemplate="<b>%{x|%Y-%m-%d}</b><br>累计投入：¥%{y:,.2f}<extra></extra>"
    ))

    fig.add_trace(go.Scatter(
        x=df["日期"],
        y=df["市值_RMB"],
        mode="lines+markers",
        name="估算市值",
        line=dict(width=3, color="#F0B90B"),
        marker=dict(size=6),
        hovertemplate="<b>%{x|%Y-%m-%d}</b><br>估算市值：¥%{y:,.2f}<extra></extra>"
    ))

    fig.add_trace(go.Scatter(
        x=df["日期"],
        y=df["浮盈_RMB"],
        mode="lines+markers",
        name="浮盈/浮亏",
        line=dict(width=3, color="#2EBD85"),
        marker=dict(size=6),
        hovertemplate="<b>%{x|%Y-%m-%d}</b><br>浮盈/浮亏：¥%{y:,.2f}<extra></extra>"
    ))

    fig.add_hline(y=0, line_dash="dash", line_color="#848E9C")

    fig.update_layout(
        title="成本、市值与浮盈曲线",
        height=430,
        template="plotly_dark",
        paper_bgcolor="#0B0E11",
        plot_bgcolor="#0B0E11",
        font=dict(color="#EAECEF"),
        margin=dict(l=40, r=25, t=65, b=40),
        legend=dict(orientation="h", y=1.08, x=0),
        hovermode="x unified"
    )

    fig.update_xaxes(gridcolor="#1E2329", zerolinecolor="#1E2329")
    fig.update_yaxes(gridcolor="#1E2329", zerolinecolor="#1E2329", tickprefix="¥", tickformat=",.0f")

    return fig


def make_return_chart(df):
    colors = [change_color(v) for v in df["收益率"]]

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=df["日期"],
        y=df["收益率"] * 100,
        name="收益率",
        marker=dict(color=colors),
        hovertemplate="<b>%{x|%Y-%m-%d}</b><br>收益率：%{y:.2f}%<extra></extra>"
    ))

    fig.add_hline(y=0, line_dash="dash", line_color="#848E9C")

    fig.update_layout(
        title="收益率变化",
        height=380,
        template="plotly_dark",
        paper_bgcolor="#0B0E11",
        plot_bgcolor="#0B0E11",
        font=dict(color="#EAECEF"),
        margin=dict(l=40, r=25, t=65, b=40),
        showlegend=False
    )

    fig.update_xaxes(gridcolor="#1E2329", zerolinecolor="#1E2329")
    fig.update_yaxes(gridcolor="#1E2329", zerolinecolor="#1E2329", ticksuffix="%")

    return fig


def make_investment_chart(df):
    colors = [rule_color(rule) for rule in df["触发规则"]]

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=df["日期"],
        y=df["投入金额_RMB"],
        marker=dict(color=colors),
        customdata=np.stack([
            df["日期文本"],
            df["触发规则"],
            df["涨跌幅百分比"],
            df["买入数量_BTC"],
            df["备注"]
        ], axis=-1),
        hovertemplate=(
            "<b>%{customdata[0]}</b><br>"
            "投入：¥%{y:,.2f}<br>"
            "触发规则：%{customdata[1]}<br>"
            "周涨跌：%{customdata[2]:.2f}%<br>"
            "买入 BTC：%{customdata[3]:.8f}<br>"
            "备注：%{customdata[4]}<extra></extra>"
        )
    ))

    fig.update_layout(
        title="每周投入金额与触发规则",
        height=380,
        template="plotly_dark",
        paper_bgcolor="#0B0E11",
        plot_bgcolor="#0B0E11",
        font=dict(color="#EAECEF"),
        margin=dict(l=40, r=25, t=65, b=40),
        showlegend=False
    )

    fig.update_xaxes(gridcolor="#1E2329", zerolinecolor="#1E2329")
    fig.update_yaxes(gridcolor="#1E2329", zerolinecolor="#1E2329", tickprefix="¥")

    return fig


def make_weekly_change_chart(df):
    colors = [change_color(v) for v in df["涨跌幅"]]

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=df["日期"],
        y=df["涨跌幅"] * 100,
        marker=dict(color=colors),
        hovertemplate="<b>%{x|%Y-%m-%d}</b><br>周涨跌幅：%{y:.2f}%<extra></extra>"
    ))

    fig.add_hline(y=0, line_dash="dash", line_color="#848E9C")
    fig.add_hline(y=-3, line_dash="dot", line_color="#5DADE2", annotation_text="B 档线 -3%")
    fig.add_hline(y=-10, line_dash="dot", line_color="#1E88E5", annotation_text="C 档线 -10%")

    fig.update_layout(
        title="每周 BTC 涨跌幅",
        height=380,
        template="plotly_dark",
        paper_bgcolor="#0B0E11",
        plot_bgcolor="#0B0E11",
        font=dict(color="#EAECEF"),
        margin=dict(l=40, r=25, t=65, b=40),
        showlegend=False
    )

    fig.update_xaxes(gridcolor="#1E2329", zerolinecolor="#1E2329")
    fig.update_yaxes(gridcolor="#1E2329", zerolinecolor="#1E2329", ticksuffix="%")

    return fig


def make_dca_contribution_map(df):
    rows = []
    total = len(df)

    for _, row in df.iterrows():
        rule = safe_text(row["触发规则"])
        note = safe_text(row["备注"])
        color = rule_color(rule)

        date_text = row["日期"].strftime("%Y-%m-%d")
        price = row["本周价格_USDT"]
        invest = row["投入金额_RMB"]
        buy_btc = row["买入数量_BTC"]
        total_btc = row["累计持有_BTC"]
        change = row["涨跌幅"] * 100 if not pd.isna(row["涨跌幅"]) else 0

        tooltip = (
            f"{date_text}\n"
            f"规则：{rule}\n"
            f"BTC价格：{price:,.2f} USDT\n"
            f"周涨跌：{change:.2f}%\n"
            f"投入：¥{invest:,.0f}\n"
            f"买入：{buy_btc:.8f} BTC\n"
            f"累计：{total_btc:.8f} BTC"
        )

        if note:
            tooltip += f"\n备注：{note}"

        rows.append(f"""
            <div
                class="dca-cell"
                style="background:{color};"
                title="{tooltip}"
                data-rule="{rule}"
            ></div>
        """)

    cells_html = "\n".join(rows)

    rule_series = df["触发规则"].astype(str).str.strip()

    stats = []

    def add_stat(label, matcher, color):
        count = int(rule_series.apply(matcher).sum())
        percent = count / total * 100 if total else 0
        stats.append((label, count, percent, color))

    add_stat("A 基础定投", lambda x: "A" in x.upper(), "#2EBD85")
    add_stat("B 小跌加仓", lambda x: "B" in x.upper(), "#5DADE2")
    add_stat("C 大跌加仓", lambda x: "C" in x.upper(), "#1E88E5")
    add_stat("戒赌资金", lambda x: "戒赌" in x, "#F6465D")
    add_stat("生日礼炮", lambda x: ("生日" in x or "礼炮" in x), "#F0B90B")
    add_stat(
        "其他",
        lambda x: not (
            "A" in x.upper()
            or "B" in x.upper()
            or "C" in x.upper()
            or "戒赌" in x
            or "生日" in x
            or "礼炮" in x
        ),
        "#6E7781"
    )

    stat_cards = "\n".join(
        f"""
        <div class="rule-stat">
            <div class="rule-dot" style="background:{color};"></div>
            <div>
                <div class="rule-name">{label}</div>
                <div class="rule-num">{count} 次 · {percent:.1f}%</div>
            </div>
        </div>
        """
        for label, count, percent, color in stats
        if count > 0
    )

    special_count = int(rule_series.apply(lambda x: "戒赌" in x or "生日" in x or "礼炮" in x).sum())
    normal_count = total - special_count

    return f"""
    <div class="contribution-panel">
        <div class="contribution-head">
            <div>
                <h2>{total} contributions in this BTC plan</h2>
                <div class="contribution-subtitle">
                    每个格子是一笔买入记录，颜色代表触发规则；悬停可查看详情。
                </div>
            </div>
            <div class="contribution-summary">
                <span>{normal_count} 次规则买入</span>
                <span>{special_count} 次特殊买入</span>
            </div>
        </div>

        <div class="dca-map-wrap">
            <div class="dca-map">
                {cells_html}
            </div>
        </div>

        <div class="legend-row">
            <span>Base</span>
            <span class="legend-box" style="background:#2EBD85;"></span>
            <span class="legend-box" style="background:#5DADE2;"></span>
            <span class="legend-box" style="background:#1E88E5;"></span>
            <span class="legend-box" style="background:#F6465D;"></span>
            <span class="legend-box" style="background:#F0B90B;"></span>
            <span>Special</span>
        </div>

        <div class="rule-stats">
            {stat_cards}
        </div>
    </div>
    """


def make_recent_table(df, rows=8):
    recent = df.tail(rows).copy()
    recent = recent.sort_values("日期", ascending=False)

    table_rows = []

    for _, row in recent.iterrows():
        rule = safe_text(row["触发规则"])
        color = rule_color(rule)
        change = row["涨跌幅"] * 100 if not pd.isna(row["涨跌幅"]) else 0
        change_class = "up" if change >= 0 else "down"

        table_rows.append(f"""
            <tr>
                <td>{row["日期"].strftime("%Y-%m-%d")}</td>
                <td>{row["本周价格_USDT"]:,.2f}</td>
                <td class="{change_class}">{change:.2f}%</td>
                <td>¥{row["投入金额_RMB"]:,.0f}</td>
                <td><span class="tag" style="border-color:{color}; color:{color};">{rule}</span></td>
                <td>{row["买入数量_BTC"]:.8f}</td>
                <td>{row["累计持有_BTC"]:.8f}</td>
            </tr>
        """)

    return "\n".join(table_rows)


def make_html(df):
    latest = df.iloc[-1]

    price_fig = make_price_chart(df)
    assets_fig = make_assets_chart(df)
    pnl_fig = make_pnl_chart(df)
    return_fig = make_return_chart(df)
    investment_fig = make_investment_chart(df)
    change_fig = make_weekly_change_chart(df)

    progress = latest["0.01进度"] * 100
    remaining = max(latest["距离0.01_BTC"], 0)
    latest_return = latest["收益率"] * 100
    latest_change = latest["涨跌幅"] * 100 if not pd.isna(latest["涨跌幅"]) else 0

    return_class = "up" if latest_return >= 0 else "down"
    change_class = "up" if latest_change >= 0 else "down"

    table_rows = make_recent_table(df)
    contribution_map = make_dca_contribution_map(df)

    plotly_js = """
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
"""

    html = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>30y BTC Dashboard</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
{plotly_js}
<style>
    :root {{
        --bg: #0B0E11;
        --panel: #11161C;
        --panel2: #181E25;
        --line: #1E2329;
        --text: #EAECEF;
        --muted: #848E9C;
        --yellow: #F0B90B;
        --green: #2EBD85;
        --red: #F6465D;
        --blue: #1E88E5;
    }}

    * {{
        box-sizing: border-box;
    }}

    body {{
        margin: 0;
        background:
            radial-gradient(circle at top left, rgba(240,185,11,0.12), transparent 28%),
            radial-gradient(circle at bottom right, rgba(30,136,229,0.10), transparent 32%),
            var(--bg);
        color: var(--text);
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", Arial, sans-serif;
    }}

    .container {{
        max-width: 1480px;
        margin: 0 auto;
        padding: 28px;
    }}

    .topbar {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 26px;
    }}

    .brand {{
        display: flex;
        align-items: center;
        gap: 14px;
    }}

    .logo {{
        width: 42px;
        height: 42px;
        border-radius: 12px;
        background: var(--yellow);
        color: #000;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 900;
        letter-spacing: -1px;
    }}

    .title h1 {{
        margin: 0;
        font-size: 26px;
        letter-spacing: 0.2px;
    }}

    .title div {{
        color: var(--muted);
        margin-top: 4px;
        font-size: 13px;
    }}

    .pill {{
        border: 1px solid var(--line);
        background: rgba(17,22,28,0.8);
        border-radius: 999px;
        padding: 9px 14px;
        color: var(--muted);
        font-size: 13px;
    }}

    .hero {{
        display: grid;
        grid-template-columns: 1.25fr 0.75fr;
        gap: 18px;
        margin-bottom: 18px;
    }}

    .price-card {{
        background: linear-gradient(135deg, rgba(17,22,28,0.96), rgba(24,30,37,0.86));
        border: 1px solid var(--line);
        border-radius: 22px;
        padding: 24px;
        min-height: 230px;
        position: relative;
        overflow: hidden;
    }}

    .price-card::after {{
        content: "";
        position: absolute;
        right: -70px;
        top: -70px;
        width: 200px;
        height: 200px;
        background: rgba(240,185,11,0.10);
        border-radius: 50%;
    }}

    .pair {{
        color: var(--muted);
        font-size: 14px;
        margin-bottom: 12px;
    }}

    .big-price {{
        font-size: 56px;
        font-weight: 800;
        line-height: 1;
        color: var(--yellow);
        letter-spacing: -1px;
    }}

    .subline {{
        margin-top: 12px;
        color: var(--muted);
        font-size: 14px;
    }}

    .progress-wrap {{
        margin-top: 24px;
    }}

    .progress-head {{
        display: flex;
        justify-content: space-between;
        font-size: 13px;
        color: var(--muted);
        margin-bottom: 8px;
    }}

    .progress-bar {{
        height: 12px;
        background: #242A32;
        border-radius: 999px;
        overflow: hidden;
    }}

    .progress-fill {{
        height: 100%;
        width: {min(progress, 100):.2f}%;
        background: linear-gradient(90deg, var(--yellow), #FFE08A);
        border-radius: 999px;
    }}

    .cards {{
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 14px;
    }}

    .card {{
        background: rgba(17,22,28,0.92);
        border: 1px solid var(--line);
        border-radius: 18px;
        padding: 18px;
    }}

    .label {{
        color: var(--muted);
        font-size: 13px;
        margin-bottom: 8px;
    }}

    .value {{
        font-size: 24px;
        font-weight: 800;
    }}

    .up {{
        color: var(--red);
    }}

    .down {{
        color: var(--green);
    }}

    .contribution-panel {{
        background: rgba(17,22,28,0.92);
        border: 1px solid var(--line);
        border-radius: 22px;
        padding: 22px;
        margin-bottom: 18px;
    }}

    .contribution-head {{
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 16px;
        margin-bottom: 18px;
    }}

    .contribution-head h2 {{
        margin: 0;
        font-size: 20px;
        letter-spacing: 0.2px;
    }}

    .contribution-subtitle {{
        color: var(--muted);
        font-size: 13px;
        margin-top: 6px;
    }}

    .contribution-summary {{
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
        justify-content: flex-end;
    }}

    .contribution-summary span {{
        border: 1px solid var(--line);
        background: rgba(255,255,255,0.03);
        border-radius: 999px;
        padding: 7px 11px;
        color: var(--muted);
        font-size: 12px;
    }}

    .dca-map-wrap {{
        overflow-x: auto;
        padding: 4px 0 10px 0;
    }}

    .dca-map {{
        display: grid;
        grid-template-rows: repeat(5, 14px);
        grid-auto-flow: column;
        grid-auto-columns: 14px;
        gap: 5px;
        min-height: 90px;
        align-content: start;
        justify-content: start;
    }}

    .dca-cell {{
        width: 14px;
        height: 14px;
        border-radius: 4px;
        opacity: 0.92;
        border: 1px solid rgba(255,255,255,0.05);
        cursor: pointer;
        transition: transform 0.12s ease, box-shadow 0.12s ease, opacity 0.12s ease;
    }}

    .dca-cell:hover {{
        transform: scale(1.35);
        opacity: 1;
        box-shadow: 0 0 0 2px rgba(240,185,11,0.35);
        z-index: 3;
    }}

    .legend-row {{
        display: flex;
        align-items: center;
        gap: 7px;
        color: var(--muted);
        font-size: 12px;
        margin-top: 4px;
        margin-bottom: 18px;
    }}

    .legend-box {{
        width: 13px;
        height: 13px;
        display: inline-block;
        border-radius: 3px;
        border: 1px solid rgba(255,255,255,0.05);
    }}

    .rule-stats {{
        display: grid;
        grid-template-columns: repeat(6, minmax(120px, 1fr));
        gap: 12px;
    }}

    .rule-stat {{
        display: flex;
        align-items: center;
        gap: 10px;
        background: rgba(255,255,255,0.03);
        border: 1px solid var(--line);
        border-radius: 14px;
        padding: 12px;
    }}

    .rule-dot {{
        width: 12px;
        height: 12px;
        border-radius: 4px;
        flex-shrink: 0;
    }}

    .rule-name {{
        color: var(--text);
        font-size: 13px;
        font-weight: 700;
    }}

    .rule-num {{
        color: var(--muted);
        font-size: 12px;
        margin-top: 3px;
    }}

    .grid {{
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 18px;
    }}

    .panel {{
        background: rgba(17,22,28,0.92);
        border: 1px solid var(--line);
        border-radius: 22px;
        padding: 14px;
        overflow: hidden;
    }}

    .panel.wide {{
        grid-column: span 2;
    }}

    .recent {{
        background: rgba(17,22,28,0.92);
        border: 1px solid var(--line);
        border-radius: 22px;
        padding: 20px;
        margin-top: 18px;
    }}

    .recent h2 {{
        margin: 0 0 14px 0;
        font-size: 18px;
    }}

    table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
    }}

    th, td {{
        padding: 12px 10px;
        border-bottom: 1px solid var(--line);
        text-align: right;
        white-space: nowrap;
    }}

    th:first-child, td:first-child {{
        text-align: left;
    }}

    th {{
        color: var(--muted);
        font-weight: 600;
    }}

    .tag {{
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border: 1px solid;
        border-radius: 999px;
        padding: 3px 9px;
        font-size: 12px;
        font-weight: 700;
        background: rgba(255,255,255,0.03);
    }}

    .footer {{
        color: var(--muted);
        font-size: 12px;
        line-height: 1.7;
        margin-top: 18px;
        padding: 16px 4px;
    }}

    @media (max-width: 1100px) {{
        .hero {{
            grid-template-columns: 1fr;
        }}

        .grid {{
            grid-template-columns: 1fr;
        }}

        .panel.wide {{
            grid-column: span 1;
        }}

        .cards {{
            grid-template-columns: 1fr 1fr;
        }}

        .rule-stats {{
            grid-template-columns: repeat(3, 1fr);
        }}
    }}

    @media (max-width: 720px) {{
        .container {{
            padding: 16px;
        }}

        .topbar {{
            align-items: flex-start;
            gap: 12px;
            flex-direction: column;
        }}

        .big-price {{
            font-size: 40px;
        }}

        .cards {{
            grid-template-columns: 1fr;
        }}

        .contribution-head {{
            flex-direction: column;
        }}

        .contribution-summary {{
            justify-content: flex-start;
        }}

        .rule-stats {{
            grid-template-columns: repeat(2, 1fr);
        }}

        .recent {{
            overflow-x: auto;
        }}
    }}
</style>
</head>
<body>
<div class="container">
    <div class="topbar">
        <div class="brand">
            <div class="logo">₿</div>
            <div class="title">
                <h1>30y BTC Dashboard</h1>
                <div>长期现货 · 按表执行 · 不碰杠杆 · 慢慢变富</div>
            </div>
        </div>
        <div class="pill">数据源：D:\\30y BTC.xlsx ｜ 最新：{latest["日期"].strftime("%Y-%m-%d")}</div>
    </div>

    <div class="hero">
        <div class="price-card">
            <div class="pair">BTC / USDT · Weekly DCA</div>
            <div class="big-price">{latest["本周价格_USDT"]:,.1f}</div>
            <div class="subline">
                周涨跌：
                <span class="{change_class}">{latest_change:.2f}%</span>
                ｜ 平均成本：{latest["平均成本_USDT"]:,.2f} USDT
            </div>

            <div class="progress-wrap">
                <div class="progress-head">
                    <span>首次 0.01 BTC 提币进度</span>
                    <span>{progress:.2f}%</span>
                </div>
                <div class="progress-bar">
                    <div class="progress-fill"></div>
                </div>
                <div class="subline">还差 {remaining:.8f} BTC 到达 0.01 BTC</div>
            </div>
        </div>

        <div class="cards">
            <div class="card">
                <div class="label">累计持有</div>
                <div class="value">{latest["累计持有_BTC"]:.8f} BTC</div>
            </div>
            <div class="card">
                <div class="label">累计投入</div>
                <div class="value">¥{latest["总成本_RMB"]:,.0f}</div>
            </div>
            <div class="card">
                <div class="label">估算市值</div>
                <div class="value">¥{latest["市值_RMB"]:,.0f}</div>
            </div>
            <div class="card">
                <div class="label">浮盈 / 浮亏</div>
                <div class="value {return_class}">¥{latest["浮盈_RMB"]:,.0f}</div>
            </div>
            <div class="card">
                <div class="label">收益率</div>
                <div class="value {return_class}">{latest_return:.2f}%</div>
            </div>
            <div class="card">
                <div class="label">本周投入</div>
                <div class="value">¥{latest["投入金额_RMB"]:,.0f}</div>
            </div>
        </div>
    </div>

    {contribution_map}

    <div class="grid">
        <div class="panel wide">
            {fig_to_html(price_fig)}
        </div>

        <div class="panel">
            {fig_to_html(assets_fig)}
        </div>

        <div class="panel">
            {fig_to_html(pnl_fig)}
        </div>

        <div class="panel">
            {fig_to_html(investment_fig)}
        </div>

        <div class="panel">
            {fig_to_html(change_fig)}
        </div>

        <div class="panel wide">
            {fig_to_html(return_fig)}
        </div>
    </div>

    <div class="recent">
        <h2>最近定投记录</h2>
        <table>
            <thead>
                <tr>
                    <th>日期</th>
                    <th>BTC价格</th>
                    <th>周涨跌</th>
                    <th>投入</th>
                    <th>规则</th>
                    <th>买入BTC</th>
                    <th>累计BTC</th>
                </tr>
            </thead>
            <tbody>
                {table_rows}
            </tbody>
        </table>
    </div>

    <div class="footer">
        说明：市值与收益按 CNY/USDT = {CNY_PER_USDT:.2f} 估算。
        这里是 30 年 BTC 定投航海日志。
    </div>
</div>
</body>
</html>
"""

    return html


def main():
    excel_path = Path(EXCEL_PATH)
    output_html = Path(OUTPUT_HTML)

    if not excel_path.exists():
        raise FileNotFoundError(f"找不到文件：{excel_path}")

    df = load_data(excel_path)

    if df.empty:
        raise ValueError("没有读取到有效数据。")

    html = make_html(df)
    output_html.write_text(html, encoding="utf-8")

    latest = df.iloc[-1]

    print("=" * 70)
    print("OKX 风格 BTC 定投仪表盘生成完成")
    print("=" * 70)
    print(f"输入文件：{excel_path}")
    print(f"输出文件：{output_html}")
    print("-" * 70)
    print(f"最新日期：{latest['日期'].strftime('%Y-%m-%d')}")
    print(f"BTC价格：{latest['本周价格_USDT']:,.2f} USDT")
    print(f"累计持有：{latest['累计持有_BTC']:.8f} BTC")
    print(f"0.01 BTC 进度：{latest['0.01进度'] * 100:.2f}%")
    print(f"累计投入：¥{latest['总成本_RMB']:,.2f}")
    print(f"估算市值：¥{latest['市值_RMB']:,.2f}")
    print(f"浮盈/浮亏：¥{latest['浮盈_RMB']:,.2f}")
    print("=" * 70)

    try:
        import os
        os.startfile(output_html)
    except Exception:
        pass


if __name__ == "__main__":
    main()