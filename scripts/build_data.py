#!/usr/bin/env python3
"""Turn a 腾讯问卷 export into the public, aggregated data.json used by the site.

Usage
    python3 scripts/build_data.py raw/export.xlsx      # real data (xlsx or csv)
    python3 scripts/build_data.py --demo               # fake demo data

Privacy rules enforced here (the raw export must never be committed):
  * only aggregated counts are written; no row-level port/date/subfield combos
  * any category with fewer than MIN_N records is folded into "其他"
  * free-text stories are published only if the moderator marks the row's
    "公开" column with Y / 是 / 1, and only coarse fields travel with them
"""
import argparse
import datetime as dt
import json
import random
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

MIN_N = 3
TREND_WEEKS = 26
OUT = Path(__file__).resolve().parent.parent / "data.json"

PORTS = ["LAX 洛杉矶", "SFO 旧金山", "SJC 圣何塞", "OAK 奥克兰", "SEA 西雅图", "PDX 波特兰", "SAN 圣地亚哥", "SNA 橙县", "ONT 安大略", "SMF 萨克拉门托", "LAS 拉斯维加斯", "PHX 凤凰城", "SLC 盐湖城", "DEN 丹佛", "HNL 檀香山", "ANC 安克雷奇", "ORD 芝加哥", "DTW 底特律", "MSP 明尼阿波利斯", "DFW 达拉斯", "IAH 休斯顿", "AUS 奥斯汀", "SAT 圣安东尼奥", "STL 圣路易斯", "MCI 堪萨斯城", "CLE 克利夫兰", "CMH 哥伦布", "IND 印第安纳波利斯", "JFK 纽约肯尼迪", "EWR 纽瓦克", "BOS 波士顿", "PHL 费城", "IAD 华盛顿杜勒斯", "BWI 巴尔的摩", "PIT 匹兹堡", "RDU 罗利-达勒姆", "CLT 夏洛特", "ATL 亚特兰大", "BNA 纳什维尔", "MIA 迈阿密", "FLL 劳德代尔堡", "MCO 奥兰多", "TPA 坦帕", "预清关 · 温哥华 YVR", "预清关 · 多伦多 YYZ", "预清关 · 蒙特利尔 YUL", "预清关 · 卡尔加里 YYC", "预清关 · 其他加拿大机场", "预清关 · 都柏林 DUB", "预清关 · 阿布扎比 AUH", "陆路口岸（美加/美墨边境）", "其他"]
VISAS = ["F-1", "J-1", "H-1B", "B1/B2", "O-1", "其他"]
VISA_STATUS = ["首次签证", "续签后首次入境", "用旧签证再次入境"]
STAYS = ["1 周以内", "1 到 2 周", "2 到 3 周", "3 周到 1 个月", "1 个月到 1 个半月", "1 个半月到 2 个月", "2 到 3 个月", "3 到 5 个月", "5 个月以上", "首次赴美（不适用）"]
ROLES = ["本科", "硕士", "博士", "博士后", "访问学者", "工作", "其他"]
GRAD_ROLES = {"硕士", "博士", "博士后", "访问学者"}
STEM_SUB = ["计算机科学/软件", "AI/机器学习/数据科学", "网络安全", "电子电气/通信",
            "芯片/半导体/集成电路", "航空航天", "机械/机器人/自动化", "材料科学",
            "核工程/核物理", "物理/天文", "化学/化工", "数学/统计",
            "生物/生物医学/生物工程", "医学/药学/公共卫生", "土木/环境/能源", "其他 STEM"]
NONSTEM_SUB = ["商科/金融/会计", "经济", "法律", "政治/国际关系", "教育", "传媒/新闻",
               "文学/语言/历史/哲学", "社会学/心理学", "艺术/设计/建筑", "其他非 STEM"]
SCHOOLS = ["国防七子 / 军队院校", "其他可能涉及 10043 的高校", "其他国内高校 / 没有在国内读过大学"]
OUTCOMES = ["pass", "held", "denied", "revoked"]
DURATIONS = ["30 分钟以内", "30 分钟到 1 小时", "1 到 3 小时", "3 到 6 小时", "6 小时以上"]
QUESTIONS = ["研究方向", "导师背景", "国内院校与科研经历", "资助来源（如 CSC）",
             "与军方或军工单位的联系", "本人或家人是否是党员", "家人职业", "毕业后是否回国",
             "之前的出入境记录", "查看微信或社交媒体", "和某人或某机构的关系", "其他"]
DEVICES = ["没有", "被要求解锁并当面翻看", "被拿走单独检查", "不确定"]

MULTI_SEP = re.compile(r"[┋|;；,，、\n]+")


# ---------------------------------------------------------------- normalising

# 国内学校名称归一：canonical -> 常见简称/英文缩写（大小写不敏感，整词匹配）
UNI_ALIASES = {
    "北京航空航天大学": ["北航", "北京航天航空大学", "BUAA"], "北京理工大学": ["北理工", "北理", "BIT"],
    "哈尔滨工业大学": ["哈工大", "哈尔滨工业大学（深圳）", "哈尔滨工业大学（威海）", "HIT"], "哈尔滨工程大学": ["哈工程", "HEU", "HRBEU"],
    "西北工业大学": ["西工大", "NPU", "NWPU"], "南京航空航天大学": ["南航", "NUAA"], "南京理工大学": ["南理工", "NJUST"],
    "国防科技大学": ["国防科大", "NUDT"], "清华大学": ["清华", "THU", "Tsinghua"], "北京大学": ["北大", "PKU"],
    "浙江大学": ["浙大", "ZJU"], "上海交通大学": ["上交", "上海交大", "SJTU"], "复旦大学": ["复旦", "FDU"],
    "中国科学技术大学": ["中科大", "USTC"], "南京大学": ["南大", "NJU"], "武汉大学": ["武大", "WHU"],
    "华中科技大学": ["华科", "华中科大", "HUST"], "西安交通大学": ["西交", "西安交大", "XJTU"], "中山大学": ["中大", "SYSU"],
    "同济大学": ["同济", "Tongji"], "四川大学": ["川大", "SCU"], "电子科技大学": ["电子科大", "成电", "UESTC"],
    "西安电子科技大学": ["西电", "XDU", "Xidian"], "中国人民大学": ["人大", "RUC"], "北京邮电大学": ["北邮", "BUPT"],
    "天津大学": ["天大", "TJU"], "南开大学": ["南开"], "东南大学": ["东南", "SEU"], "山东大学": ["山大", "SDU"],
    "厦门大学": ["厦大", "XMU"], "中南大学": ["中南", "CSU"], "湖南大学": ["湖大", "HNU"], "大连理工大学": ["大工", "大连理工", "DUT"],
    "吉林大学": ["吉大", "JLU"], "东北大学": ["东北大学", "NEU"], "北京师范大学": ["北师大", "BNU"], "华东师范大学": ["华东师大", "华师大", "ECNU"],
    "重庆大学": ["重大", "CQU"], "华南理工大学": ["华工", "华南理工", "SCUT"], "中国科学院大学": ["国科大", "中科院", "UCAS"],
    "北京科技大学": ["北科", "北科大", "USTB"], "北京交通大学": ["北交", "北交大", "BJTU"], "北京化工大学": ["北化", "BUCT"],
    "中国农业大学": ["中农", "CAU"], "南方科技大学": ["南科大", "SUSTech"], "上海科技大学": ["上科大", "ShanghaiTech"],
    "西南交通大学": ["西南交大", "SWJTU"], "武汉理工大学": ["武理", "WHUT"], "中国海洋大学": ["中国海大", "OUC"],
}
_UNI = {}
for canon, al in UNI_ALIASES.items():
    for a in [canon] + al:
        _UNI[a.lower()] = canon


def norm_unis(text):
    """'北航、清华 / ZJU' -> ['北京航空航天大学', '清华大学', '浙江大学']"""
    out = []
    for part in re.split(r"[、,，;；/|\n]+|\s{2,}| 和 |和", str(text or "")):
        x = part.strip().strip("。.")
        if not x or x.lower() in ("nan", "无", "none", "-"):
            continue
        canon = _UNI.get(x.lower())
        if canon is None:
            canon = next((c for c in UNI_ALIASES if c in x), x)  # e.g. "北京航空航天大学本科" -> canonical
        if canon not in out:
            out.append(canon)
    return out


def norm_outcome(s):
    s = str(s or "")
    if "取消" in s:
        return "revoked"
    if "拒" in s or "遣返" in s:
        return "denied"
    if "小黑屋" in s or "放行" in s:
        return "held"
    if "直接" in s or "通过" in s:
        return "pass"
    return None


def pick(value, options, other="其他"):
    """Map free/partial text onto a canonical option (substring match)."""
    v = str(value or "").strip()
    if not v or v.lower() == "nan":
        return None
    if v in options:
        return v
    for o in options:            # option text inside the raw answer
        if o in v:
            return o
    for o in options:            # e.g. "LAX" -> "LAX 洛杉矶"; avoid tiny/ambiguous fragments
        if len(v) >= 3 and not v.startswith("其他") and v in o:
            return o
    return other


def parse_date(v):
    if v is None or str(v).strip() in ("", "nan", "NaT"):
        return None
    if isinstance(v, (dt.date, dt.datetime)):
        return v if isinstance(v, dt.date) and not isinstance(v, dt.datetime) else v.date()
    m = re.search(r"(\d{4})\D(\d{1,2})\D(\d{1,2})", str(v))
    if not m:
        return None
    try:
        return dt.date(*map(int, m.groups()))
    except ValueError:
        return None


def find_col(cols, *keys):
    for c in cols:
        if all(k in str(c) for k in keys):
            return c
    return None


def read_export(path):
    import pandas as pd
    path = Path(path)
    df = pd.read_csv(path, dtype=str) if path.suffix.lower() == ".csv" else pd.read_excel(path)
    cols = list(df.columns)

    def multi(row, key):
        """Multi-select: either one column with separators, or one column per option."""
        sub = [c for c in cols if key in str(c) and re.search(r"[:：\-—_]\s*\S", str(c).split(key, 1)[1])]
        if sub:
            out = []
            for c in sub:
                val = str(row[c]).strip()
                if val and val.lower() != "nan" and val not in ("0", "未选中", "否"):
                    out.append(re.split(r"[:：\-—_]", str(c).split(key, 1)[1], 1)[-1].strip())
            return out
        c = find_col(cols, key)
        val = str(row[c]) if c is not None else ""
        return [x.strip() for x in MULTI_SEP.split(val) if x.strip() and x.strip().lower() != "nan"]

    c = {k: find_col(cols, k) for k in
         ["入境日期", "入境机场", "签证类型", "签证情况", "当前身份", "专业大类",
          "入境结果", "待了多久", "手机", "补充", "停留"]}
    c["公开"] = next((x for x in cols if str(x).strip() == "公开"), None)
    if c["签证情况"] is None:
        c["签证情况"] = find_col(cols, "签证") if find_col(cols, "签证") != c["签证类型"] else None
    c_uni = find_col(cols, "学校名称")
    c_stem = find_col(cols, "STEM", "小类")
    c_non = find_col(cols, "非 STEM", "小类") or find_col(cols, "非STEM", "小类")
    if c_stem == c_non:
        c_stem = next((x for x in cols if "小类" in str(x) and "非" not in str(x)), None)

    # 个人ID：同一 ID 只保留最后一次提交；最后一次勾了“请删除”的整条丢弃
    c_id = find_col(cols, "个人ID") or find_col(cols, "ID")
    c_del = find_col(cols, "删除")
    c_time = find_col(cols, "提交时间") or find_col(cols, "结束答题时间") or find_col(cols, "答题时间")
    df = df.reset_index(drop=True)
    df["_order"] = range(len(df))
    if c_time is not None:
        df["_t"] = pd.to_datetime(df[c_time], errors="coerce")
        df = df.sort_values(["_t", "_order"], na_position="first")
    if c_id is not None:
        ids = df[c_id].astype(str).str.strip().str.lower()
        has_id = ids.ne("") & ids.ne("nan")
        df = pd.concat([df[~has_id], df[has_id].assign(_id=ids[has_id]).drop_duplicates("_id", keep="last")])
    if c_del is not None:
        df = df[~df[c_del].astype(str).str.contains("删除")]

    recs = []
    for _, row in df.iterrows():
        g = lambda k: row[c[k]] if c[k] is not None else None
        cat = "非 STEM" if "非" in str(g("专业大类")) else ("STEM" if "STEM" in str(g("专业大类")).upper() else None)
        sub_raw = row[c_non] if cat == "非 STEM" and c_non is not None else (row[c_stem] if c_stem is not None else None)
        recs.append({
            "date": parse_date(g("入境日期")),
            "port": pick(g("入境机场") if c["入境机场"] is not None else None, PORTS),
            "visa": pick(g("签证类型"), VISAS),
            "visa_status": pick(g("签证情况"), VISA_STATUS),
            "stay": pick(g("停留"), STAYS, None),
            "role": pick(g("当前身份"), ROLES),
            "major_cat": cat,
            "major_sub": pick(sub_raw, NONSTEM_SUB if cat == "非 STEM" else STEM_SUB,
                              "其他非 STEM" if cat == "非 STEM" else "其他 STEM"),
            "schools": [pick(x, SCHOOLS) for x in multi(row, "就读过的院校")],
            "unis": norm_unis(row[c_uni]) if c_uni is not None else [],
            "outcome": norm_outcome(g("入境结果")),
            "duration": pick(g("待了多久"), DURATIONS, None),
            "questions": [pick(x, QUESTIONS) for x in multi(row, "哪些问题")],
            "device": pick(g("手机"), DEVICES, "不确定"),
            "story": str(g("补充") or "").strip(),
            "publish": str(g("公开") or "").strip().upper() in ("Y", "YES", "是", "1", "TRUE"),
        })
    return recs


# ---------------------------------------------------------------- demo data

DEMO_UNIS = {
    "国防七子 / 军队院校": ["北京航空航天大学", "北京理工大学", "哈尔滨工业大学", "哈尔滨工程大学", "西北工业大学", "南京航空航天大学", "南京理工大学", "国防科技大学"],
    "其他可能涉及 10043 的高校": ["电子科技大学", "四川大学", "中国人民大学", "西安电子科技大学"],
    "其他国内高校 / 没有在国内读过大学": ["清华大学", "北京大学", "浙江大学", "上海交通大学", "复旦大学", "中国科学技术大学", "南京大学", "武汉大学", "华中科技大学", "中山大学", "同济大学", "厦门大学", "山东大学", "北京邮电大学"],
}


def fake_records(n=420, seed=7):
    rnd = random.Random(seed)
    today = dt.date.today()
    stories = {  # demo only
        "pass": ["很顺利，只问了学校和专业，不到一分钟。", "排队比较久，但边检官员态度很好，直接通过。"],
        "held": ["被问了很久研究方向和导师的项目经费来源，最后放行。",
                 "海关让解锁手机看了微信聊天记录，问了本科学校，大概两个小时后放行。"],
        "denied": ["被问是否是党员、家里人做什么工作，最后被拒入境并原机返回。"],
        "revoked": ["问了国内实验室和导师的项目，最后签证被当场取消。"],
    }
    recs = []
    for _ in range(n):
        role = rnd.choices(ROLES, [22, 30, 30, 8, 5, 4, 1])[0]
        cat = rnd.choices(["STEM", "非 STEM"], [68, 32])[0]
        sub = rnd.choice(STEM_SUB if cat == "STEM" else NONSTEM_SUB)
        school = rnd.choices(SCHOOLS, [8, 12, 80])[0]
        risk = 0.08 + (0.10 if cat == "STEM" else 0) + (0.30 if school == SCHOOLS[0] else 0.15 if school == SCHOOLS[1] else 0) \
            + (0.12 if sub in ("芯片/半导体/集成电路", "航空航天", "核工程/核物理", "AI/机器学习/数据科学", "网络安全") else 0) \
            + (0.05 if role in GRAD_ROLES else 0)
        r = rnd.random()
        outcome = "pass" if r > risk else ("held" if r > risk * 0.18 else rnd.choice(["denied", "denied", "revoked"]))
        d = today - dt.timedelta(days=int(rnd.triangular(0, 200, 30)))
        rec = {
            "date": d, "port": rnd.choices(PORTS, [16, 12, 3, 1, 6, 2, 2, 1, 1, 1, 2, 2, 1, 2, 1, 1, 8, 5, 2, 5, 4, 1, 1, 1, 1, 1, 1, 1, 10, 5, 7, 2, 4, 1, 1, 1, 1, 3, 1, 1, 1, 1, 1, 4, 3, 1, 1, 1, 1, 1, 1, 1])[0],
            "visa": rnd.choices(VISAS, [70, 12, 8, 6, 2, 2])[0],
            "visa_status": rnd.choices(VISA_STATUS, [45, 25, 30])[0],
            "stay": rnd.choices(STAYS, [5, 10, 14, 16, 12, 9, 10, 8, 5, 11])[0],
            "role": role, "major_cat": cat, "major_sub": sub, "schools": [school],
            "unis": rnd.sample(DEMO_UNIS[school], 1 if rnd.random() < .7 else 2) if rnd.random() < .75 else [],
            "outcome": outcome, "duration": None, "questions": [], "device": "没有",
            "story": "", "publish": False,
        }
        if outcome != "pass":
            rec["duration"] = rnd.choices(DURATIONS, [15, 25, 35, 15, 10])[0]
            rec["questions"] = rnd.sample(QUESTIONS[:-1], rnd.randint(2, 6))
            rec["device"] = rnd.choices(DEVICES, [40, 40, 15, 5])[0]
        if rnd.random() < (0.02 if outcome == "pass" else 0.08):
            rec["story"], rec["publish"] = "【示例】" + rnd.choice(stories[outcome]), True
        recs.append(rec)
    return recs


# ---------------------------------------------------------------- aggregation

def tally(recs):
    c = Counter(r["outcome"] for r in recs)
    return {"n": len(recs), "pass": c["pass"], "held": c["held"], "denied": c["denied"] + c["revoked"]}


def breakdown(recs, key, order=None, multi=False):
    groups = defaultdict(list)
    for r in recs:
        vals = r[key] if multi else [r[key]]
        for v in vals or []:
            if v:
                groups[v].append(r)
    rows, other = [], []
    for k, rs in groups.items():
        (rows if len(rs) >= MIN_N else other).append((k, rs))
    out = [{"key": k, **tally(rs)} for k, rs in rows]
    if order:
        out.sort(key=lambda x: order.index(x["key"]) if x["key"] in order else 999)
    else:
        out.sort(key=lambda x: -x["n"])
    folded = [r for _, rs in other for r in rs]
    if len(folded) >= MIN_N:
        out.append({"key": "其他（样本较少）", **tally(folded)})
    return out


def counts(items, order):
    c = Counter(i for i in items if i)
    return [{"key": k, "n": c[k]} for k in order if c[k] >= 1]


def school_group(r):
    s = set(r["schools"] or [])
    if s & {"国防七子", "其他军工背景高校"}:
        return "国防七子或军工背景"
    if "其他国内高校" in s:
        return "其他国内高校"
    if "没有在国内读过大学" in s:
        return "没有在国内读过大学"
    return None


def aggregate(recs, demo):
    """Publish de-identified rows (week-level date, no school, no free text) for client-side filtering."""
    today = dt.date.today()
    recs = [r for r in recs if r["outcome"] and r["date"] and dt.date(2020, 1, 1) <= r["date"] <= today]
    recs.sort(key=lambda r: r["date"], reverse=True)
    uni_n = Counter(u for r in recs for u in (r.get("unis") or []))
    rows = []
    for r in recs:
        wk = r["date"] - dt.timedelta(days=r["date"].weekday())
        rows.append({
            "w": wk.isoformat(), "port": r["port"], "visa": r["visa"], "vs": r["visa_status"],
            "stay": r.get("stay"), "role": r["role"], "mc": r["major_cat"], "ms": r["major_sub"],
            "sch": sorted(set(x for x in (r["schools"] or []) if x)), "o": r["outcome"],
            "u": [u for u in (r.get("unis") or []) if uni_n[u] >= MIN_N],
            "dur": r["duration"] if r["outcome"] != "pass" else None,
            "q": sorted(set(x for x in r["questions"] if x)) if r["outcome"] != "pass" else [],
            "dev": r["device"] if r["outcome"] != "pass" else None,
        })
    stories = [{"month": r["date"].strftime("%Y-%m"), "visa": r["visa"], "role": r["role"],
                "mc": r["major_cat"], "o": r["outcome"], "dur": r["duration"], "text": r["story"][:600]}
               for r in recs if r["publish"] and r["story"]]
    return {"demo": demo, "updated": today.isoformat(), "rows": rows, "min_n": MIN_N, "stories": stories[:60]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("export", nargs="?", help="腾讯问卷 export (.xlsx / .csv)")
    ap.add_argument("--demo", action="store_true", help="generate fake demo data instead")
    ap.add_argument("--empty", action="store_true", help="write an empty dataset (launch state)")
    ap.add_argument("--out", default=str(OUT), help="output path (default data.json)")
    a = ap.parse_args()
    if not (a.demo or a.empty or a.export):
        ap.error("give an export file, --demo or --empty")
    recs = fake_records() if a.demo else ([] if a.empty else read_export(a.export))
    data = aggregate(recs, a.demo)
    out = Path(a.out)
    out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"wrote {out}  records={len(data['rows'])}  stories={len(data['stories'])}{'  (DEMO)' if a.demo else ''}")


if __name__ == "__main__":
    sys.exit(main())
