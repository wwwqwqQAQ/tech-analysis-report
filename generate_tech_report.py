#!/usr/bin/env python3
"""科技产业分析日报 — AI/芯片/新能源/互联网/航天，产业政策+技术趋势+投融资+竞争格局"""

import json, os, re, subprocess, sys
from datetime import datetime
from collections import defaultdict, Counter

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

for name in ("Heiti TC", "PingFang SC", "STHeiti", "Arial Unicode MS"):
    try:
        fm.findfont(name, fallback_to_default=False)
        plt.rcParams["font.sans-serif"] = [name, "DejaVu Sans"]
        break
    except Exception:
        continue
plt.rcParams["axes.unicode_minus"] = False

UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
REFERER = "https://tech.sina.com.cn/"
CHART_DIR = "/tmp/tech_charts"
os.makedirs(CHART_DIR, exist_ok=True)

C_RED = "#D32F2F"; C_BLUE = "#1976D2"; C_GRAY = "#616161"; C_ORANGE = "#E65100"
C_GREEN = "#2E7D32"; C_PURPLE = "#7B1FA2"; C_TEAL = "#00796B"

def curl(url):
    r = subprocess.run(["curl", "-4", "-sk", "--connect-timeout", "10",
        "-H", f"User-Agent: {UA}", "-H", f"Referer: {REFERER}", url],
        capture_output=True, timeout=20)
    if r.returncode != 0: return None
    raw = r.stdout
    for enc in ("utf-8", "gbk", "gb2312", "gb18030"):
        try: return raw.decode(enc)
        except (UnicodeDecodeError, LookupError): continue
    return raw.decode("utf-8", errors="replace")

def fetch_sina_search(keywords, size=30):
    """Multi-keyword search via Sina search API."""
    results = []
    seen = set()
    for kw in keywords:
        url = f"https://search.sina.com.cn/api/news?q={kw}&size={size}"
        raw = curl(url)
        if not raw: continue
        try:
            data = json.loads(raw)
            for item in data.get("data", {}).get("list", []):
                u = item.get("url", "")
                if u in seen: continue
                seen.add(u)
                title = re.sub(r'<[^>]+>', '', item.get("title", ""))
                results.append({
                    "title": title, "ctime": str(item.get("ctime", "")),
                    "url": u, "intro": item.get("intro", ""),
                })
        except (json.JSONDecodeError, KeyError): continue
    return results

def fetch_sina_news(lid, num=40):
    all_news = []
    for page in range(1, (num // 20) + 3):
        url = f"https://feed.mix.sina.com.cn/api/roll/get?pageid=153&lid={lid}&k=&num=20&page={page}&r=0.1&callback="
        raw = curl(url)
        if not raw: break
        try:
            data = json.loads(raw)
            items = data.get("result", {}).get("data", [])
            if not items: break
            all_news.extend(items)
            if len(all_news) >= num: break
        except json.JSONDecodeError: break
    return all_news[:num]

# ═══════════════════ CLASSIFICATION ═══════════════════

def classify_tech(news_items):
    """Classify tech news into sub-sectors."""
    cats = {"AI/大模型": [], "半导体/芯片": [], "新能源/电动车": [], "互联网/平台": [],
            "航天/硬科技": [], "政策/监管": [], "投融资/创投": [], "其他": []}
    kw_map = {
        "AI/大模型": ["AI", "人工智能", "大模型", "GPT", "LLM", "深度学习", "机器学习",
                      "NLP", "计算机视觉", "生成式", "智能体", "Agent", "OpenAI",
                      "谷歌", "微软", "Meta", "百度文心", "通义", "DeepSeek", "Kimi",
                      "ChatGPT", "Claude", "Gemini", "Llama", "神经网络"],
        "半导体/芯片": ["芯片", "半导体", "光刻", "晶圆", "台积电", "中芯", "英伟达",
                       "AMD", "英特尔", "高通", "ARM", "RISC-V", "GPU", "CPU",
                       "华为", "海思", "封测", "EDA", "硅", "7nm", "5nm", "3nm"],
        "新能源/电动车": ["新能源", "电动车", "电池", "光伏", "风电", "储能", "氢能",
                         "比亚迪", "特斯拉", "宁德时代", "锂电", "钠电", "固态电池",
                         "充电", "换电", "碳", "双碳", "碳中和"],
        "互联网/平台": ["互联网", "平台", "电商", "社交", "短视频", "直播", "搜索",
                       "字节", "腾讯", "阿里", "美团", "拼多多", "京东", "小红书",
                       "TikTok", "WeChat", "APP", "小程序", "算法", "推荐"],
        "航天/硬科技": ["航天", "卫星", "火箭", "SpaceX", "NASA", "北斗", "空间站",
                       "量子", "核聚变", "机器人", "具身智能", "人形机器人", "脑机",
                       "3D打印", "新材料", "超导", "可控核聚变"],
        "政策/监管": ["政策", "监管", "法规", "反垄断", "数据安全", "隐私", "合规",
                     "审查", "许可", "牌照", "工信部", "网信办", "科技部", "标准"],
        "投融资/创投": ["融资", "投资", "IPO", "上市", "估值", "VC", "PE", "创投",
                       "天使", "A轮", "B轮", "C轮", "独角兽", "退出", "并购", "收购"],
    }
    for item in news_items:
        title = item.get("title", "")
        matched = "其他"
        for cat, kws in kw_map.items():
            if any(kw in title for kw in kws):
                matched = cat; break
        if matched == "其他":
            # Try intro for matching
            intro = item.get("intro", "")
            for cat, kws in kw_map.items():
                if any(kw in intro for kw in kws):
                    matched = cat; break
        cats[matched].append(item)
    return cats

# ═══════════════════ ANALYSIS ═══════════════════

def analyze_company_mentions(items):
    """Track which companies are most mentioned."""
    companies = {
        "华为": ["华为", "海思", "鸿蒙", "昇腾", "鲲鹏"],
        "腾讯": ["腾讯", "微信", "WeChat"],
        "阿里巴巴": ["阿里", "阿里巴巴", "通义", "蚂蚁"],
        "字节跳动": ["字节", "抖音", "TikTok", "豆包"],
        "百度": ["百度", "文心", "萝卜快跑"],
        "比亚迪": ["比亚迪", "BYD"],
        "特斯拉": ["特斯拉", "Tesla"],
        "英伟达": ["英伟达", "NVIDIA", "Nvidia"],
        "OpenAI": ["OpenAI"],
        "谷歌": ["谷歌", "Google", "Gemini", "DeepMind"],
        "微软": ["微软", "Microsoft"],
        "苹果": ["苹果", "Apple"],
        "小米": ["小米", "Xiaomi"],
        "宁德时代": ["宁德时代", "CATL"],
        "台积电": ["台积电", "TSMC"],
        "中芯国际": ["中芯国际", "中芯", "SMIC"],
        "Meta": ["Meta", "Facebook"],
        "DeepSeek": ["DeepSeek", "深度求索"],
    }
    all_text = " ".join(i["title"] for i in items)
    mentions = {}
    for name, kws in companies.items():
        cnt = sum(all_text.count(kw) for kw in kws)
        if cnt > 0: mentions[name] = cnt
    return Counter(mentions).most_common(15)

def analyze_tech_trends(cats, company_mentions):
    """Generate tech trend analysis with insights."""
    trends = []

    # AI analysis
    ai_count = len(cats.get("AI/大模型", []))
    if ai_count > 3:
        ai_items = cats["AI/大模型"]
        # Check for key themes
        agent_ai = sum(1 for i in ai_items if "Agent" in i["title"] or "智能体" in i["title"])
        model_war = sum(1 for i in ai_items if any(k in i["title"] for k in
            ["GPT", "Llama", "Gemini", "文心", "通义", "Kimi", "DeepSeek", "豆包", "Claude"]))
        trends.append({
            "title": "AI大模型竞争格局",
            "signal": "加速分化",
            "analysis": (
                f"AI/大模型领域{ai_count}条新闻。"
                f"{'Agent/智能体成为新热点（{}条），AI从对话式交互向任务执行演进。'.format(agent_ai) if agent_ai > 0 else ''}"
                f"基础模型竞争进入'寡头+开源'双轨阶段：头部闭源模型（GPT、Claude、Gemini、文心）在能力上持续突破，"
                f"开源模型（Llama、DeepSeek、通义千问开源版）在成本和可及性上带来降维打击。"
                f"从技术采用生命周期理论看，AI正从'早期采用者'阶段过渡到'早期大众'阶段，"
                f"2026年下半场的竞争焦点将从'谁的模型最强'转向'谁的生态最完整'——"
                f"工具链、开发者社区、企业级部署能力将成为新的护城河。"
            ),
            "horizon": "6-12个月",
        })

    # Semiconductor
    chip_count = len(cats.get("半导体/芯片", []))
    if chip_count > 3:
        trends.append({
            "title": "半导体产业链博弈",
            "signal": "国产替代加速",
            "analysis": (
                f"半导体/芯片领域{chip_count}条新闻。当前全球半导体产业链处于'效率优先→安全优先'的范式转换期。"
                f"美国主导的技术封锁（设备、EDA、先进制程）与中国推动的自主替代并行推进，"
                f"两端都在加速——这使产业链呈现'双轨化'趋势：一条是美国及其盟友的'可信供应链'，"
                f"另一条是中国主导的'自主可控链'。从产业经济学角度看，这种分裂的代价是"
                f"全球芯片成本上升和重复投资，但也催生了巨大的国产替代市场机会。"
                f"关注中芯国际/华虹的先进制程进展、RISC-V生态成熟度、以及EDA工具国产化率三个关键信号。"
            ),
            "horizon": "1-3年",
        })

    # New energy
    ev_count = len(cats.get("新能源/电动车", []))
    if ev_count > 3:
        trends.append({
            "title": "新能源与电动车产业趋势",
            "signal": "从规模竞争到技术竞争",
            "analysis": (
                f"新能源/电动车领域{ev_count}条新闻。产业正从'补贴驱动的规模扩张'进入'技术驱动的差异化竞争'阶段。"
                f"关键变量包括：固态电池量产时间表（决定下一轮洗牌节奏）、"
                f"智能化程度（智驾+座舱成为区分度核心）、海外市场准入（欧盟反补贴关税、美国IRA法案）。"
                f"从波特五力模型看，新能源车行业当前面临：供应商议价力上升（锂/钴/芯片）、"
                f"买方选择增多（车型供给过剩）、新进入者威胁（小米/华为等跨界玩家）、"
                f"替代品压力（氢能/合成燃料）——行业利润率将承压，集中度将进一步提升。"
            ),
            "horizon": "6-18个月",
        })

    # Policy/Reg
    policy_count = len(cats.get("政策/监管", []))
    if policy_count > 2:
        trends.append({
            "title": "科技监管与产业政策",
            "signal": "精准化监管",
            "analysis": (
                f"科技政策/监管领域{policy_count}条新闻。全球科技监管进入'精准化'阶段——"
                f"不再是笼统的'加强监管'或'放松管制'，而是针对特定领域（AI安全、数据跨境、"
                f"算法推荐、反垄断、芯片出口管制）的精细化规则设计。"
                f"中国的监管特色是'发展-安全双目标平衡'：既要通过科技实现产业升级，"
                f"又要确保数据主权和社会稳定。从制度经济学视角看，这种双目标治理面临内在张力——"
                f"过度安全化会拖慢创新速度，过度发展导向则可能积累系统性风险。"
                f"政策走向的领先指标包括：国务院/发改委/工信部的产业规划更新频率、"
                f"科创板IPO审核节奏、以及跨境数据流动试点城市的扩围情况。"
            ),
            "horizon": "3-12个月",
        })

    return trends

def generate_tech_outlook(cats, company_mentions, trends):
    """Forward-looking tech predictions."""
    outlooks = []
    top_companies = [c[0] for c in company_mentions[:5]]

    outlooks.append({
        "title": "科技产业整体判断",
        "view": "结构性机会为主",
        "analysis": (
            f"本日科技新闻中提及最多的企业：{'、'.join(top_companies)}。"
            f"从技术成熟度曲线（Hype Cycle）视角看，当前多个技术赛道处于不同阶段："
            f"AI Agent / 具身智能处于'期望膨胀期'；大模型基础能力处于'稳步爬升期'；"
            f"固态电池 / 量子计算处于'技术触发期'。理解各赛道所处的Hype Cycle位置，"
            f"是判断投资时机和风险的关键——膨胀期入场需承受回调风险，爬升期入场则面临竞争白热化。"
            f"建议关注处于'技术触发期→期望膨胀期'过渡阶段的赛道（如可控核聚变、脑机接口、量子计算），"
            f"这类赛道当前关注度低、技术拐点临近，是典型的'左侧布局'机会。"
        ),
        "horizon": "6-18个月",
    })

    return outlooks

# ═══════════════════ CHARTS ═══════════════════

def chart_tech_categories(cats):
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ordered = ["AI/大模型", "半导体/芯片", "新能源/电动车", "互联网/平台", "航天/硬科技", "政策/监管", "投融资/创投"]
    counts = [len(cats.get(c, [])) for c in ordered]
    colors = [C_BLUE, C_RED, C_GREEN, C_ORANGE, C_PURPLE, C_GRAY, C_TEAL]
    bars = ax.bar(ordered, counts, color=colors, edgecolor="white", linewidth=0.5)
    for bar, v in zip(bars, counts):
        if v > 0:
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3, str(v),
                    ha="center", fontsize=9, fontweight="bold")
    ax.set_title("科技产业新闻分布", fontsize=14, fontweight="bold", pad=12)
    ax.set_ylabel("新闻数量")
    plt.xticks(rotation=30, ha="right", fontsize=8)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    fig.tight_layout()
    path = f"{CHART_DIR}/tech_cats.png"
    fig.savefig(path, dpi=150, bbox_inches="tight"); plt.close(fig)
    return path

def chart_company_mentions(mentions):
    fig, ax = plt.subplots(figsize=(7, 4))
    names = [m[0] for m in mentions[:10]][::-1]
    counts = [m[1] for m in mentions[:10]][::-1]
    colors = plt.cm.Blues([0.3 + 0.7 * i/len(names) for i in range(len(names))])
    ax.barh(names, counts, color=colors, edgecolor="white")
    for i, v in enumerate(counts):
        ax.text(v + 0.3, i, str(v), va="center", fontsize=9, fontweight="bold")
    ax.set_title("科技企业媒体关注度 TOP 10", fontsize=14, fontweight="bold", pad=12)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    fig.tight_layout()
    path = f"{CHART_DIR}/company_mentions.png"
    fig.savefig(path, dpi=150, bbox_inches="tight"); plt.close(fig)
    return path

def chart_trend_timeline(cats, title_str, filename):
    all_dates = []
    for items in cats.values():
        for item in items:
            ctime = item.get("ctime", "")
            if ctime:
                try:
                    ts = int(ctime)
                    all_dates.append(datetime.fromtimestamp(ts).strftime("%m-%d"))
                except: pass
    if not all_dates: return None
    date_counts = Counter(all_dates)
    dates = sorted(date_counts.keys())[-14:]
    counts = [date_counts[d] for d in dates]
    fig, ax = plt.subplots(figsize=(8, 3))
    ax.fill_between(range(len(dates)), counts, alpha=0.3, color=C_BLUE)
    ax.plot(range(len(dates)), counts, marker="o", color=C_BLUE, linewidth=2, markersize=4)
    ax.set_xticks(range(len(dates)))
    ax.set_xticklabels(dates, rotation=45, fontsize=7)
    ax.set_title(title_str, fontsize=13, fontweight="bold", pad=12)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    fig.tight_layout()
    path = f"{CHART_DIR}/{filename}.png"
    fig.savefig(path, dpi=150, bbox_inches="tight"); plt.close(fig)
    return path

# ═══════════════════ MAIN ═══════════════════

print("=" * 60)
print("科技产业分析日报生成器 v1.0")
print("=" * 60)

print("\n[1/4] 获取科技新闻...")
tech_keywords = ["人工智能", "芯片半导体", "新能源电动车", "航天卫星", "互联网科技",
                  "AI大模型", "机器人", "量子计算", "生物科技", "科技政策"]
tech_raw = fetch_sina_search(tech_keywords, size=20)
tech_raw += fetch_sina_news(lid=2509, num=40)
seen = set()
tech_news = []
for item in tech_raw:
    u = item.get("url", "")
    if u not in seen:
        seen.add(u)
        tech_news.append(item)
print(f"  科技新闻: {len(tech_news)} 条")

print("\n[2/4] 执行分析...")
tech_cats = classify_tech(tech_news)
company_mentions = analyze_company_mentions(tech_news)
tech_trends = analyze_tech_trends(tech_cats, company_mentions)
tech_outlook = generate_tech_outlook(tech_cats, company_mentions, tech_trends)
for k, v in tech_cats.items():
    if v: print(f"  {k}: {len(v)} 条")
print(f"  企业关注度 TOP 5: {', '.join(f'{c}({n})' for c,n in company_mentions[:5])}")

print("\n[3/4] 生成图表...")
charts = {
    "tech_cats": chart_tech_categories(tech_cats),
    "company_mentions": chart_company_mentions(company_mentions),
}
tl = chart_trend_timeline(tech_cats, "科技新闻热度趋势", "trend_timeline")
if tl: charts["timeline"] = tl
print(f"  {len(charts)} 张图表已生成")

# ═══════════════════ BUILD DOCX ═══════════════════

print("\n[4/4] 生成报告文档...")

doc = Document()
for section in doc.sections:
    section.top_margin = Cm(2); section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5); section.right_margin = Cm(2.5)

style = doc.styles['Normal']
style.font.name = 'Times New Roman'; style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def fmt_date(ctime_str):
    try:
        ts = int(ctime_str)
        return datetime.fromtimestamp(ts).strftime("%Y-%m-%d")
    except: return ""

def add_bullet(doc, item):
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(item["title"]); run.font.size = Pt(10)
    d = fmt_date(item.get("ctime", ""))
    if d:
        run2 = p.add_run(f'  ({d})'); run2.font.size = Pt(9)
        run2.font.color.rgb = RGBColor(150, 150, 150)

def add_analysis(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(f'▸ 分析：{text}'); run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(int(C_BLUE[1:3],16), int(C_BLUE[3:5],16), int(C_BLUE[5:7],16))
    p.paragraph_format.left_indent = Cm(0.8)

# COVER
title = doc.add_heading('科技产业分析日报', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
subtitle = doc.add_paragraph(); subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run(f'{datetime.now().strftime("%Y年%m月%d日")} | 科技产业深度分析')
run.font.size = Pt(12); run.font.color.rgb = RGBColor(100, 100, 100)
doc.add_paragraph()

# Executive Summary
doc.add_heading('核心发现', level=1)
total = sum(len(v) for v in tech_cats.values())
ai_n = len(tech_cats.get("AI/大模型", [])); chip_n = len(tech_cats.get("半导体/芯片", []))
ev_n = len(tech_cats.get("新能源/电动车", []))
doc.add_paragraph(
    f"本日监测科技产业新闻{total}条，覆盖{sum(1 for v in tech_cats.values() if v)}个细分赛道。"
    f"AI/大模型（{ai_n}条）、半导体/芯片（{chip_n}条）、新能源/电动车（{ev_n}条）为三大热点。"
    f"企业关注度前三：{'、'.join(f'{c}({n}次)' for c,n in company_mentions[:3])}。"
)

# ── SECTION MAP ──
sections = [
    ("一、AI / 大模型", "AI/大模型",
     "从技术采用生命周期看，AI正在跨越早期采用者到早期大众的'鸿沟'。当前阶段的核心矛盾不是模型能力不够，而是企业落地成本太高、人才供给不足、以及监管不确定性。基础设施层（算力、数据、工具链）的投资确定性高于应用层。"),
    ("二、半导体 / 芯片", "半导体/芯片",
     "半导体是科技竞争的'底层棋局'。美国主导的设备/EDA/先进制程封锁与中国'举国体制'式自主替代之间的拉锯，决定了整个科技产业的上游供应稳定性。从产业链安全视角看，成熟制程（28nm及以上）的国产化率快速提升，先进制程（7nm以下）仍是瓶颈。"),
    ("三、新能源 / 电动车", "新能源/电动车",
     "新能源产业已从'政策驱动'进入'市场驱动+技术驱动'双轮阶段。固态电池量产时间表、智能化水平、海外市场准入三者共同决定下一轮行业洗牌的节奏。中国企业在电池、光伏、风电领域的全球份额领先，但面临欧美'去风险'政策压力。"),
    ("四、互联网 / 平台经济", "互联网/平台",
     "平台经济监管进入常态化阶段。增长引擎从'流量红利'转向'AI赋能+出海'。字节/TikTok的全球化、 Temu/SHEIN的跨境电商、以及AI对搜索/推荐/广告的重构，是当前最值得关注的三个结构性变量。"),
    ("五、航天 / 硬科技", "航天/硬科技",
     "商业航天、量子计算、可控核聚变、脑机接口、人形机器人——这些'硬科技'赛道虽短期商业化前景不明，但代表了科技产业的长期方向。从风险投资视角看，这些赛道适合'小额度、多标的、长周期'的组合策略。"),
    ("六、科技政策与监管", "政策/监管",
     "科技政策的底层逻辑是'在发展和安全之间寻找动态平衡'。AI立法、数据跨境流动、反垄断执法、芯片出口管制是四个政策主线。政策变化的领先指标包括：人大常委会立法计划、国务院产业规划、以及网信办/工信部的部门规章更新。"),
    ("七、投融资与创投动态", "投融资/创投",
     "一级市场投融资数据是科技产业未来2-3年的'前置指标'。当前全球科技投资从'撒胡椒面'模式转向'集中押注'模式——资金向头部项目和确定性赛道集中，腰部项目融资难度加大。"),
]

for title, cat_key, analysis_text in sections:
    doc.add_heading(title, level=2)
    items = tech_cats.get(cat_key, [])
    if items:
        for item in items[:8]: add_bullet(doc, item)
    else:
        doc.add_paragraph("本日该领域无显著新闻。")
    if items:
        doc.add_paragraph()
        add_analysis(doc, f"{analysis_text} 本日该赛道新闻{len(items)}条，{'信息密度较高，值得重点关注。' if len(items) > 5 else '保持常规关注。'}")

# Charts
doc.add_paragraph()
if charts.get("tech_cats"):
    doc.add_paragraph().add_run().add_picture(charts["tech_cats"], width=Inches(5.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
if charts.get("company_mentions"):
    doc.add_paragraph().add_run().add_picture(charts["company_mentions"], width=Inches(5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

# Tech Trends Deep Dive
doc.add_paragraph()
doc.add_heading('八、技术趋势深度研判', level=2)
for trend in tech_trends:
    p_t = doc.add_paragraph()
    run = p_t.add_run(trend["title"]); run.bold = True; run.font.size = Pt(12)
    p_m = doc.add_paragraph()
    run_m = p_m.add_run(f'信号: {trend["signal"]}　|　时间维度: {trend["horizon"]}')
    run_m.font.size = Pt(10); run_m.font.color.rgb = RGBColor(int(C_RED[1:3],16), int(C_RED[3:5],16), int(C_RED[5:7],16))
    doc.add_paragraph(trend["analysis"])

# Outlook
doc.add_paragraph()
doc.add_heading('九、科技产业前瞻', level=2)
for outlook in tech_outlook:
    p_t = doc.add_paragraph()
    run = p_t.add_run(outlook["title"]); run.bold = True; run.font.size = Pt(12)
    p_m = doc.add_paragraph()
    run_m = p_m.add_run(f'判断: {outlook["view"]}　|　时间维度: {outlook["horizon"]}')
    run_m.font.size = Pt(10); run_m.font.color.rgb = RGBColor(int(C_BLUE[1:3],16), int(C_BLUE[3:5],16), int(C_BLUE[5:7],16))
    doc.add_paragraph(outlook["analysis"])

# FOOTER
doc.add_paragraph(); doc.add_paragraph()
footer = doc.add_paragraph(); footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = footer.add_run('— 以上分析基于公开信息，纯属研究性质，不构成投资建议 —')
run.font.size = Pt(9); run.font.color.rgb = RGBColor(150, 150, 150); run.italic = True
footer2 = doc.add_paragraph(); footer2.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = footer2.add_run(
    f'数据: 新浪新闻 | 框架: Hype Cycle / 波特五力 / 技术采用生命周期 '
    f'| 生成: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
run.font.size = Pt(8); run.font.color.rgb = RGBColor(150, 150, 150)

# SAVE
output_dir = '/Users/wwwqwq/Desktop/项目/科技报告'
os.makedirs(output_dir, exist_ok=True)
output_path = f'{output_dir}/科技产业分析_{datetime.now().strftime("%Y-%m-%d")}.docx'
doc.save(output_path)
print(f'\n{"="*60}')
print(f'报告已生成: {output_path}')
print(f'科技新闻: {total} 条 | {sum(1 for v in tech_cats.values() if v)} 个赛道')
print(f'含 {len(charts)} 张图表')
print(f'{"="*60}')
