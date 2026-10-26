#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""读取课表 .xls + 自习室公告，生成次日方案网页 index.html。

用法:
    python build.py                # 默认生成"明天"的方案
    python build.py 2026-10-23     # 指定日期
"""
import datetime
import glob
import io
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsread import read_xls

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEEK1_MONDAY = datetime.date(2026, 8, 31)  # 2026-2027-1 第1周 8/31-9/6

SLOTS = [
    {"id": "1-2", "label": "第1-2节", "start": "07:40", "end": "09:20", "half": "am"},
    {"id": "3-4", "label": "第3-4节", "start": "09:45", "end": "11:25", "half": "am"},
    {"id": "5-6", "label": "第5-6节", "start": "14:30", "end": "16:10", "half": "pm"},
    {"id": "7-8", "label": "第7-8节", "start": "16:35", "end": "18:15", "half": "pm"},
    {"id": "9-11", "label": "第9-11节", "start": "19:20", "end": "21:55", "half": "night"},
]
SLOT_BY_ID = {s["id"]: s for s in SLOTS}
WEEKDAY_CN = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
BUILDING_NAME = {2: "2号教学楼", 3: "3号教学楼", 4: "4号教学楼",
                 5: "5号教学楼", 7: "7号教学楼", 9: "9号教学楼", "信": "信息楼"}


# ---------------------------------------------------------------- 课表
def parse_weeks(line):
    m = re.search(r"([\d,\-，]+)\s*\(周\)", line)
    if not m:
        return set()
    weeks = set()
    for part in m.group(1).replace("，", ",").split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            try:
                weeks.update(range(int(a), int(b) + 1))
            except ValueError:
                pass
        elif part.isdigit():
            weeks.add(int(part))
    return weeks


def parse_place(place):
    s = place.strip()
    s = re.sub(r"^\(海甸\)", "", s).strip()
    m = re.match(r"^(\d+)-(\d+)(.*)$", s)
    if m:
        b = int(m.group(1))
        no = m.group(2)
        extra = m.group(3).strip()
        floor = int(no) // 100 if len(no) >= 3 else int(no)
        room = int(no) % 100
    else:
        m = re.match(r"^信\s*-?\s*(\d+)(.*)$", s)
        if not m:
            return {"raw": s, "building": None, "floor": 0, "room": 0, "extra": ""}
        b, no, extra = "信", m.group(1), m.group(2).strip()
        floor, room = int(no) // 100, int(no) % 100
    return {"raw": s, "building": b, "floor": floor, "room": room, "extra": extra}


def place_text(p):
    name = BUILDING_NAME.get(p["building"], f'{p["building"]}号楼')
    no = f'{p["floor"]}{p["room"]:02d}' if p["building"] in (2, 3, 4, 5, 7) else f'{p["floor"]}{p["room"]:02d}'
    txt = name + no
    if p["extra"]:
        txt += f'（{p["extra"]}）'
    return txt


def parse_cell(cell):
    if not isinstance(cell, str) or not cell.strip():
        return []
    out = []
    for block in re.split(r"\n\s*\n", cell):
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        if len(lines) < 3:
            continue
        wi = next((i for i, l in enumerate(lines) if "(周)" in l and "[" in l), None)
        if wi is None or wi + 1 >= len(lines):
            continue
        name = lines[0]
        teacher = ""
        for l in lines[1:wi]:
            if l.startswith("(") and "分组" in l:
                continue
            teacher = l
            break
        out.append({
            "name": name, "teacher": teacher,
            "weeks": parse_weeks(lines[wi]),
            "place": parse_place(lines[wi + 1]),
            "place_raw": lines[wi + 1].strip(),
        })
    # 同一格重复的条目去重
    seen, uniq = set(), []
    for c in out:
        k = (c["name"], c["teacher"], c["place_raw"])
        if k not in seen:
            seen.add(k)
            uniq.append(c)
    return uniq


def load_courses(target):
    week = (target - WEEK1_MONDAY).days // 7 + 1
    wd = target.weekday()  # 0=Mon
    col = wd + 1

    # 优先用脱敏后的 courses.json（公开仓库不含学号/姓名）
    jf = os.path.join(BASE, "courses.json")
    if os.path.exists(jf):
        with io.open(jf, encoding="utf-8") as fh:
            data = json.load(fh)
        result = {}
        for sid, cols in data.get("slots", {}).items():
            cs = cols.get(str(col), [])
            result[sid] = [dict(c, weeks=set(c.get("weeks") or [])) for c in cs
                           if week in (c.get("weeks") or [])]
        return jf, week, result

    files = sorted(glob.glob(os.path.join(BASE, "*.xls")))
    if not files:
        raise SystemExit("工作区没找到 courses.json 或 .xls 课表")
    sheet_name, cells = read_xls(files[0])[0]
    maxr = max(r for r, c in cells)
    # 行 -> 节次 id
    row_slot = {}
    for r in range(maxr + 1):
        txt = str(cells.get((r, 0), ""))
        m = re.match(r"^(\d+)、(\d+)(?:、(\d+))?节", txt)
        if m:
            row_slot[r] = f"{int(m.group(1))}-{int(m.group(3) or m.group(2))}"
    result = {}
    for r, sid in row_slot.items():
        courses = parse_cell(cells.get((r, col), ""))
        result[sid] = [c for c in courses if week in c["weeks"]]
    return files[0], week, result


# ---------------------------------------------------------------- 公告
def parse_notice(text):
    info = {"date": None, "publish": "", "rooms": {}}
    m = re.search(r"(\d{1,2})月(\d{1,2})日（星期(.)）", text)
    if m:
        info["date"] = (int(m.group(1)), int(m.group(2)))
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日\s+(\d{1,2}:\d{2})", text)
    if m:
        info["publish"] = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d} {m.group(4)}"

    cur = None
    for raw in text.splitlines():
        line = raw.strip().strip("|").strip()
        line = re.sub(r"\*\*", "", line)
        line = re.sub(r"<br\s*/?>", " ", line)
        line = re.sub(r"\s+", " ", line).strip()
        if not line:
            continue
        if line.startswith("观澜湖"):
            cur = None
            continue
        mb = re.match(r"^(\d+)号教学楼", line)
        if mb:
            cur = int(mb.group(1))
            info["rooms"].setdefault(cur, {})
            continue
        if cur is None:
            continue
        mp = re.match(r"^第(\d+)-(\d+)节(.*)$", line)
        if not mp:
            continue
        a, b = int(mp.group(1)), int(mp.group(2))
        body = mp.group(3)
        body = re.split(r"有声教室", body)[0]
        body = re.sub(r"[|]", " ", body).strip()
        rooms = []
        if body and body not in ("无",):
            for tok in re.split(r"[、，,\s]+", body):
                tok = tok.strip()
                if not tok or tok == "无":
                    continue
                mr = re.match(r"^(\d{3})-(\d{3})$", tok)
                if mr:
                    lo, hi = int(mr.group(1)), int(mr.group(2))
                    rooms.extend(range(lo, hi + 1))
                elif tok.isdigit():
                    rooms.append(int(tok))
        key = f"{a}-{b}"
        info["rooms"][cur][key] = rooms
    return info


def rooms_for(info, slot_id):
    """取某节次可用的 3/4 号楼自习室 -> {building: [rooms]}"""
    out = {}
    for b in (3, 4):
        got = []
        tbl = info["rooms"].get(b, {})
        for key, rooms in tbl.items():
            ka, kb = (int(x) for x in key.split("-"))
            sa, sb = (int(x) for x in slot_id.split("-"))
            if ka <= sa and sb <= kb:
                got.extend(rooms)
        out[b] = sorted(set(got))
    return out


# ---------------------------------------------------------------- 推荐
def hdist(building, room_no, anchor):
    if building != anchor["building"]:
        return 0
    if building == 4:  # 南北楼：01-10 / 11-20
        g1 = room_no <= 10
        g2 = anchor["room"] <= 10
        if g1 != g2:
            return 12
    return abs(room_no - anchor["room"])


def recommend(slot_id, info, anchor, busy):
    """anchor: 该半天上课地点(可能 None)；busy: 该半天有课节数"""
    options = rooms_for(info, slot_id)
    cross = anchor is not None and anchor["building"] not in (3, 4)
    scored = []
    for b, rooms in options.items():
        for no in rooms:
            f, n = no // 100, no % 100
            if anchor is None or cross:
                score = (f - 1) * 12 + n * 0.2 + (0 if b == 4 else 2)
                if cross:
                    why = "跨楼 · 选低层少爬楼" if f <= 2 else "跨楼 · 楼层偏高"
                    tag = "跨楼"
                else:
                    why = ("1 层，进出最快" if f <= 1 else
                           "2 层，楼层低" if f == 2 else
                           "楼层偏高，人少安静")
                    tag = "低层"
            else:
                bp = 0 if b == anchor["building"] else 45
                fd = abs(f - anchor["floor"])
                score = bp + fd * 8 + hdist(b, n, anchor) * 1.6
                if b == anchor["building"] and fd == 0:
                    why = "与上课教室同层，出门就到"
                elif b == anchor["building"] and fd == 1:
                    why = "与上课教室上下相邻一层"
                elif b == anchor["building"]:
                    why = "同楼，垂直距离近"
                else:
                    why = "跨楼，步行约 3 分钟"
                if b == 4 and anchor["building"] == 4 and hdist(b, n, anchor) == 12:
                    why += "（跨南北楼）"
                tag = "就近"
            scored.append({"building": b, "no": no, "floor": f, "score": score, "why": why, "tag": tag})
    scored.sort(key=lambda x: x["score"])
    return scored[:3]


# ---------------------------------------------------------------- 页面
CSS = """
*{box-sizing:border-box;margin:0;padding:0}
:root{--bg:#f5f6f8;--card:#fff;--ink:#1b1f24;--dim:#6b7280;--line:#e6e8ec;
--brand:#2f6fed;--free:#0f9d58;--busy:#e8710a;--warn:#b45309;--warnbg:#fff7ed}
body{background:var(--bg);color:var(--ink);font:15px/1.6 -apple-system,"PingFang SC","Microsoft YaHei",system-ui,sans-serif;
padding:18px 14px 40px;max-width:760px;margin:0 auto;-webkit-font-smoothing:antialiased}
header{margin-bottom:14px}
.hd-top{display:flex;align-items:baseline;gap:10px;flex-wrap:wrap}
.date{font-size:30px;font-weight:700;letter-spacing:-.5px}
.wd{font-size:15px;color:var(--dim)}
.week{font-size:12px;background:#eef2ff;color:var(--brand);padding:2px 8px;border-radius:99px;font-weight:600}
.meta{color:var(--dim);font-size:12px;margin-top:6px}
.banner{background:var(--warnbg);border:1px solid #fcd34d;color:var(--warn);border-radius:12px;
padding:11px 14px;font-size:13px;margin:14px 0;line-height:1.65}
.banner b{color:#92400e}
.now{background:linear-gradient(135deg,#2f6fed,#4f8ef7);color:#fff;border-radius:16px;padding:14px 16px;margin-bottom:18px}
.now .k{font-size:12px;opacity:.85}
.now .v{font-size:17px;font-weight:600;margin-top:2px;line-height:1.5}
.now .s{font-size:13px;opacity:.92;margin-top:5px}
label.sum{display:block;font-size:12px;color:var(--dim);margin:0 2px 8px}
.slot{background:var(--card);border-radius:16px;padding:14px 16px 12px;margin-bottom:12px;
box-shadow:0 1px 2px rgba(16,24,40,.05),0 1px 3px rgba(16,24,40,.06)}
.slot-head{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.slot-name{font-weight:700;font-size:16px}
.slot-time{color:var(--dim);font-size:13px;font-variant-numeric:tabular-nums}
.tag{margin-left:auto;font-size:11px;font-weight:600;padding:2px 9px;border-radius:99px}
.tag.free{background:#e8f5ee;color:var(--free)}
.tag.busy{background:#fff1e6;color:var(--busy)}
.course{background:#fbfaf7;border:1px solid #f0ece1;border-radius:10px;padding:8px 11px;margin-bottom:10px;font-size:13.5px}
.course .cn{font-weight:600}
.course .cp{color:#8a5a00;font-weight:600}
.course .ct{color:var(--dim);font-size:12px;margin-left:6px}
.course.empty{color:var(--dim);background:#fafafa;border-color:#eee;font-size:13px}
.recs{display:flex;flex-direction:column;gap:6px}
.rec{display:flex;align-items:center;gap:9px;font-size:13.5px;padding:6px 2px;border-top:1px dashed #f0f1f4}
.rec:first-child{border-top:0}
.num{width:18px;height:18px;border-radius:6px;background:#eef2ff;color:var(--brand);
font-size:11px;font-weight:700;display:grid;place-items:center;flex:none}
.rm{font-weight:600;white-space:nowrap}
.why{color:var(--dim);font-size:12.5px}
.fl{margin-left:auto;font-size:11px;color:#aab1bb;flex:none}
.note{font-size:12px;color:#98a0ab;margin:-5px 0 9px}
footer{margin-top:22px;color:#9aa1ab;font-size:12px;line-height:1.9}
footer a{color:#9aa1ab}
@media(max-width:420px){.date{font-size:26px}.why{font-size:12px}}
"""

JS = """
(function(){
  var D=window.__DATA__, t=new Date();
  var pad=function(n){return (n<10?'0':'')+n};
  var key=t.getFullYear()+'-'+pad(t.getMonth()+1)+'-'+pad(t.getDate());
  var mins=t.getHours()*60+t.getMinutes();
  var toMin=function(s){var a=s.split(':');return (+a[0])*60+(+a[1])};
  var fmt=function(m){if(m<60){return m+' 分钟'}var h=Math.floor(m/60),r=m%60;return h+' 小时'+(r?' '+r+' 分钟':'')};
  var lab=document.getElementById('nowk'),el=document.getElementById('nowv'),sub=document.getElementById('nows');
  var firstRec=function(s){return s.recs&&s.recs.length?s.recs[0]:null};
  var recText=function(s){var r=firstRec(s);return r?r.name+'（'+r.why+'）':'暂无'};

  if(key!==D.date){
    var fc=null;
    for(var i=0;i<D.slots.length;i++){if(D.slots[i].course){fc=D.slots[i];break}}
    lab.textContent=D.dateLabel+' '+D.weekday+' 的方案';
    if(!fc){el.textContent='全天没课，随便安排';sub.textContent='自习首选：'+recText(D.slots[0]);return}
    var gap=toMin(fc.start)-mins;
    var head='第一节课 '+fc.label+'（'+fc.start+'）在 '+fc.course.place;
    el.textContent=gap>0?head+'，还有 '+fmt(gap):head;
    sub.textContent=D.slots[0].course
      ? '早到的话，先在 '+recText(fc)+' 自习'
      : '上午空着，可先去 '+recText(D.slots[0])+'，'+fc.start+' 前 10 分钟再往 '+fc.course.place+' 走';
    return;
  }

  var cur=null,nxt=null;
  for(var j=0;j<D.slots.length;j++){
    var s=D.slots[j];
    if(mins>=toMin(s.start)&&mins<toMin(s.end)){cur=s}
    if(nxt===null&&mins<toMin(s.start)){nxt=s}
  }
  lab.textContent='现在 '+pad(t.getHours())+':'+pad(t.getMinutes());
  if(cur){
    el.textContent=cur.label+'（'+cur.start+'–'+cur.end+'）'+
      (cur.course?' 上《'+cur.course.name+'》· '+cur.course.place:' · 本节没课');
    sub.textContent=(cur.course?'就算不去上课，也可就近去 ':'推荐 ')+recText(cur);
  }else if(nxt){
    el.textContent='距 '+nxt.label+'（'+nxt.start+'）还有 '+fmt(toMin(nxt.start)-mins)+
      (nxt.course?'，下一节在 '+nxt.course.place:'，本节没课');
    sub.textContent='先就近自习：'+recText(nxt);
  }else{
    el.textContent='今天的课和自习时段都结束了';
    sub.textContent='明天见';
  }
})();
"""


def build(target):
    path, week, courses = load_courses(target)
    notice_text = open(os.path.join(BASE, "公告.txt"), encoding="utf-8").read()
    info = parse_notice(notice_text)

    wd_cn = WEEKDAY_CN[target.weekday()]
    date_label = f'{target.month}月{target.day}日'

    am_ids = ["1-2", "3-4"]
    pm_ids = ["5-6", "7-8"]
    am_busy = sum(len(courses.get(s, [])) for s in am_ids)
    pm_busy = sum(len(courses.get(s, [])) for s in pm_ids)

    slots_data = []
    for s in SLOTS:
        sid = s["id"]
        cs = courses.get(sid, [])
        half = s["half"]
        if half == "am":
            busy, ids = am_busy, am_ids
        elif half == "pm":
            busy, ids = pm_busy, pm_ids
        else:
            busy, ids = 0, [sid]

        anchor = None
        if busy >= 1:
            for sid2 in ids:
                if courses.get(sid2):
                    p = courses[sid2][0]["place"]
                    if p["building"]:
                        anchor = p
                        break
        if cs and cs[0]["place"]["building"]:
            anchor = cs[0]["place"]

        recs = recommend(sid, info, anchor, busy)
        if cs:
            note = "不去上课的话，就近自习选这几间"
        elif anchor and anchor["building"] in (3, 4):
            note = "本节没课 · 按本次上课点%s就近推荐" % place_text(anchor)
        elif anchor:
            note = "本节没课 · 上课点在%s，以下是 3/4 号楼可选教室" % place_text(anchor)
        else:
            note = "本半天没课 · 优先低楼层"
        slots_data.append({
            "id": sid, "label": s["label"], "start": s["start"], "end": s["end"],
            "note": note,
            "course": ({"name": cs[0]["name"], "teacher": cs[0]["teacher"],
                        "place": place_text(cs[0]["place"])} if cs else None),
            "recs": [{"name": BUILDING_NAME.get(r["building"], str(r["building"])) + f'{r["floor"]}{r["no"]%100:02d}',
                      "why": r["why"], "floor": r["floor"], "tag": r["tag"]} for r in recs],
        })

    notice_is_target = False
    if info["date"]:
        notice_is_target = (info["date"][0] == target.month and info["date"][1] == target.day)

    banner = ""
    if not info["date"]:
        banner = ('<div class="banner"><b>未能读取自习室公告。</b>下面的自习室推荐为空，'
                  '请把公告正文粘贴到工作区的 <code>公告.txt</code> 后重新生成。</div>')
    elif not notice_is_target:
        banner = ('<div class="banner"><b>10月%s日（%s）的自习室公告还没发布。</b><br>'
                  '目前最新是 <b>%d月%d日</b> 的公告（发布于 %s）。下面的自习室按这份公告推算，'
                  '<b>仅供参考</b>；今晚公告更新后重新生成一次即可变为准确数据。'
                  '</div>') % (target.day, wd_cn, info["date"][0], info["date"][1], info["publish"])

    total = sum(len(courses.get(s["id"], [])) for s in SLOTS)
    gen_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    cards = []
    for sd in slots_data:
        if sd["course"]:
            c = sd["course"]
            course_html = ('<div class="course"><span class="cn">%s</span>'
                           '<span class="ct">%s</span><br><span class="cp">%s</span></div>'
                           % (c["name"], c["teacher"], c["place"]))
            tag = '<span class="tag busy">有课</span>'
        else:
            course_html = '<div class="course empty">这一节没课</div>'
            tag = '<span class="tag free">空</span>'
        note_html = '<div class="note">%s</div>' % sd["note"]
        rec_html = []
        for i, r in enumerate(sd["recs"]):
            rec_html.append('<div class="rec"><span class="num">%d</span>'
                            '<span class="rm">%s</span><span class="why">%s</span>'
                            '<span class="fl">%s</span></div>'
                            % (i + 1, r["name"], r["why"], r["tag"]))
        if not rec_html:
            rec_html.append('<div class="rec"><span class="why">暂无可用自习室数据</span></div>')
        cards.append(
            '<article class="slot"><div class="slot-head"><span class="slot-name">%s</span>'
            '<span class="slot-time">%s–%s</span>%s</div>%s%s<div class="recs">%s</div></article>'
            % (sd["label"], sd["start"], sd["end"], tag, course_html, note_html, "".join(rec_html)))

    data = {"date": target.isoformat(), "dateLabel": date_label, "weekday": wd_cn, "slots": slots_data}

    html = (
        '<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        "<title>%s %s · 课表与自习方案</title><style>%s</style></head><body>"
        '<header><div class="hd-top"><span class="date">%s</span>'
        '<span class="wd">%s</span><span class="week">第 %d 周</span></div>'
        '<div class="meta">明天 %d 节课 · 课表：%s · 生成于 %s</div></header>'
        "%s"
        '<div class="now"><div class="k" id="nowk">现在</div>'
        '<div class="v" id="nowv">—</div><div class="s" id="nows"></div></div>'
        '<label class="sum">各节次课程地点 + 就近自习室（3号/4号教学楼，按距离与楼层排序）</label>'
        "%s"
        '<footer>自习室数据来源：<a href="https://he.hainanu.edu.cn/baozhang/info/1033/15003.htm" target="_blank">海南大学后勤保障部《自习室预告》</a>%s%s（一般每晚 21:30 前后更新次日公告）<br>'
        '规则：半天全空 → 优先低楼层；半天有课 → 优先与上课地点同楼、垂直位置相近的教室；'
        '4号教学楼 01-10 与 11-20 分属南北楼。<br>'
        '更新：把最新公告正文粘贴进 <code>公告.txt</code>，运行 <code>python scripts/build.py</code>。</footer>'
        "<script>window.__DATA__=%s;</script><script>%s</script></body></html>"
    ) % (
        date_label, wd_cn, CSS,
        date_label, wd_cn, week,
        total, os.path.basename(path), gen_at,
        banner,
        "".join(cards),
        ("（%d月%d日公告，发布 %s）" % (info["date"][0], info["date"][1], info["publish"])) if info["date"] else "",
        "；<b>注意：非次日数据</b>" if not notice_is_target else "",
        json.dumps(data, ensure_ascii=False), JS,
    )

    out = os.path.join(BASE, "index.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"生成: {out}")
    print(f"日期: {target} {wd_cn} 第{week}周  课程数={total}  上午{am_busy}节 下午{pm_busy}节")
    print(f"公告: {info['date']} 发布={info['publish']} 是否次日={notice_is_target}")
    for sd in slots_data:
        print(" ", sd["label"], sd["course"]["place"] if sd["course"] else "—",
              "|", ",".join(r["name"] for r in sd["recs"]))


if __name__ == "__main__":
    if len(sys.argv) > 1:
        tgt = datetime.date.fromisoformat(sys.argv[1])
    else:
        tgt = datetime.date.today() + datetime.timedelta(days=1)
    build(tgt)
