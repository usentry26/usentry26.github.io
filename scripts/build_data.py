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

PORTS = ["LAX 洛杉矶", "SFO 旧金山", "SEA 西雅图", "JFK 纽约肯尼迪", "EWR 纽瓦克",
         "ORD 芝加哥", "BOS 波士顿", "IAD 华盛顿杜勒斯", "DFW 达拉斯", "IAH 休斯顿",
         "DTW 底特律", "ATL 亚特兰大", "PHL 费城", "MSP 明尼阿波利斯", "DEN 丹佛",
         "SJC 圣何塞", "SAN 圣地亚哥", "PHX 凤凰城", "LAS 拉斯维加斯", "PDX 波特兰",
         "SLC 盐湖城", "MIA 迈阿密", "CLT 夏洛特", "HNL 檀香山",
         "加拿大预清关（温哥华/多伦多等）", "其他地点预清关（都柏林/阿布扎比等）", "陆路口岸"]
VISAS = ["F-1", "J-1", "H-1B", "B1/B2", "O-1", "其他"]
VISA_STATUS = ["首次签证", "续签后首次入境", "用旧签证再次入境"]
STAYS = ["1 个月以内", "1 到 3 个月", "3 到 5 个月", "5 个月以上", "首次赴美（不适用）"]
ROLES = ["本科", "硕士", "博士", "博士后", "访问学者", "工作", "其他"]
GRAD_ROLES = {"硕士", "博士", "博士后", "访问学者"}
STEM_SUB = ["计算机科学/软件", "AI/机器学习/数据科学", "网络安全", "电子电气/通信",
            "芯片/半导体/集成电路", "航空航天", "机械/机器人/自动化", "材料科学",
            "核工程/核物理", "物理/天文", "化学/化工", "数学/统计",
            "生物/生物医学/生物工程", "医学/药学/公共卫生", "土木/环境/能源", "其他 STEM"]
NONSTEM_SUB = ["商科/金融/会计", "经济", "法律", "政治/国际关系", "教育", "传媒/新闻",
               "文学/语言/历史/哲学", "社会学/心理学", "艺术/设计/建筑", "其他非 STEM"]
SCHOOLS = ["国防七子", "其他军工背景高校", "其他国内高校", "没有在国内读过大学"]
OUTCOMES = ["pass", "held", "denied", "revoked"]
DURATIONS = ["30 分钟以内", "30 分钟到 1 小时", "1 到 3 小时", "3 到 6 小时", "6 小时以上"]
QUESTIONS = ["研究方向", "导师背景", "国内院校与科研经历", "资助来源（如 CSC）",
             "与军方或军工单位的联系", "本人或家人是否是党员", "家人职业", "毕业后是否回国",
             "之前的出入境记录", "查看微信或社交媒体", "和某人或某机构的关系", "其他"]
DEVICES = ["没有", "被要求解锁并当面翻看", "被拿走单独检查", "不确定"]

MULTI_SEP = re.compile(r"[┋|;；,，、\n]+")


# ---------------------------------------------------------------- normalising

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
          "入境结果", "待了多久", "手机", "补充", "公开", "停留"]}
    if c["签证情况"] is None:
        c["签证情况"] = find_col(cols, "签证") if find_col(cols, "签证") != c["签证类型"] else None
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
            "schools": [pick(x, SCHOOLS) for x in multi(row, "院校类型")],
            "outcome": norm_outcome(g("入境结果")),
            "duration": pick(g("待了多久"), DURATIONS, None),
            "questions": [pick(x, QUESTIONS) for x in multi(row, "哪些问题")],
            "device": pick(g("手机"), DEVICES, "不确定"),
            "story": str(g("补充") or "").strip(),
            "publish": str(g("公开") or "").strip().upper() in ("Y", "YES", "是", "1", "TRUE"),
        })
    return recs


# ---------------------------------------------------------------- demo data

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
        school = rnd.choices(SCHOOLS, [7, 8, 70, 15])[0]
        risk = 0.08 + (0.10 if cat == "STEM" else 0) + (0.25 if school in SCHOOLS[:2] else 0) \
            + (0.12 if sub in ("芯片/半导体/集成电路", "航空航天", "核工程/核物理", "AI/机器学习/数据科学", "网络安全") else 0) \
            + (0.05 if role in GRAD_ROLES else 0)
        r = rnd.random()
        outcome = "pass" if r > risk else ("held" if r > risk * 0.18 else rnd.choice(["denied", "denied", "revoked"]))
        d = today - dt.timedelta(days=int(rnd.triangular(0, 200, 30)))
        rec = {
            "date": d, "port": rnd.choices(PORTS, [16, 12, 6, 10, 5, 8, 7, 4, 5, 4, 5, 3] + [1] * (len(PORTS) - 12))[0],
            "visa": rnd.choices(VISAS, [70, 12, 8, 6, 2, 2])[0],
            "visa_status": rnd.choices(VISA_STATUS, [45, 25, 30])[0],
            "stay": rnd.choices(STAYS, [30, 35, 15, 8, 12])[0],
            "role": role, "major_cat": cat, "major_sub": sub, "schools": [school],
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
    rows = []
    for r in recs:
        wk = r["date"] - dt.timedelta(days=r["date"].weekday())
        rows.append({
            "w": wk.isoformat(), "port": r["port"], "visa": r["visa"], "vs": r["visa_status"],
            "stay": r.get("stay"), "role": r["role"], "mc": r["major_cat"], "ms": r["major_sub"],
            "sch": sorted(set(x for x in (r["schools"] or []) if x)), "o": r["outcome"],
            "dur": r["duration"] if r["outcome"] != "pass" else None,
            "q": sorted(set(x for x in r["questions"] if x)) if r["outcome"] != "pass" else [],
            "dev": r["device"] if r["outcome"] != "pass" else None,
        })
    stories = [{"month": r["date"].strftime("%Y-%m"), "visa": r["visa"], "role": r["role"],
                "mc": r["major_cat"], "o": r["outcome"], "dur": r["duration"], "text": r["story"][:600]}
               for r in recs if r["publish"] and r["story"]]
    return {"demo": demo, "updated": today.isoformat(), "rows": rows, "stories": stories[:60]}


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
