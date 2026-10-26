# -*- coding: utf-8 -*-
"""
每晚自动抓取「海大后勤 · 自习室预告」，生成 index.html（明日自习室查询页）。
无第三方依赖；由 Windows 计划任务每天 23:00 调用，失败写日志。
"""
import os
import re
import sys
import html
import json
import datetime
import traceback
import urllib.request

try:
    from zoneinfo import ZoneInfo
    _SH = ZoneInfo("Asia/Shanghai")
except Exception:            # 极老的 Python 没有 zoneinfo
    _SH = None


def now_sh():
    """统一用北京时间，避免在 UTC 机器（GitHub Actions）上时间显示成 16:31。"""
    return datetime.datetime.now(_SH) if _SH else datetime.datetime.now()


BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE, "scripts"))
import build as TB  # 课表解析（含 xlsread）

URL = "https://he.hainanu.edu.cn/baozhang/info/1033/15003.htm"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
PRIOR = ("4号教学楼", "5号教学楼")   # 没课时优先
SLOT_IDS = ["1-2", "3-4", "5-6", "7-8", "9-11"]


# ------------------------------------------------------------------ 抓取
def fetch():
    req = urllib.request.Request(URL, headers=UA)
    raw = urllib.request.urlopen(req, timeout=30).read()
    for enc in ("utf-8", "gb18030"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "ignore")


def txt(x):
    x = re.sub(r"<[^>]+>", " ", x)
    return re.sub(r"\s+", " ", html.unescape(x)).strip()


# ------------------------------------------------------------------ 解析公告
TAG = r"(?:</?[^>]{1,200}>|\s)*"
HEAD_RE = re.compile(r"(\d{1,2})月(\d{1,2})日[（(][^）)]*[）)]" + TAG + r"教学楼" + TAG + r"自习室预告")


def parse_notice(page):
    """返回 {title, publish, days:{"M-D": {楼名:{slot:{rooms,voice,note}}}}}"""
    out = {"title": "", "publish": "", "days": {}}
    p = re.search(r"<ol>([^<]*\d{1,2}:\d{2})</ol>", page)
    if p:
        out["publish"] = txt(p.group(1))
    else:
        p = re.search(r"发布时间[:：]\s*([\d年月日\-/ :]+)", txt(page[:20000]))
        out["publish"] = p.group(1).strip() if p else ""

    heads = [(m.start(), int(m.group(1)), int(m.group(2))) for m in HEAD_RE.finditer(page)]
    for i, (pos, mo, dy) in enumerate(heads):
        end = heads[i + 1][0] if i + 1 < len(heads) else len(page)
        chunk = page[pos:end]
        if not re.search(r"<table", chunk, re.I):
            continue
        site = parse_tables(chunk)
        if site:
            out["days"][f"{mo}-{dy}"] = site
            out["title"] = f"{mo}月{dy}日"
    return out


def parse_tables(chunk):
    site = {}
    cur = None
    rows = []
    for tb in re.findall(r"<table.*?</table>", chunk, re.S | re.I):
        if "教学楼" not in tb:
            continue
        for tr in re.findall(r"<tr.*?</tr>", tb, re.S):
            cells = [txt(td) for td in re.findall(r"<td.*?</td>", tr, re.S)]
            cells = [c for c in cells if c]
            if cells:
                rows.append(cells)
    for cells in rows:
        head = cells[0]
        bm = re.match(r"^([A-Z]?\d+)\s*号?教学楼", head)
        if bm and len(cells) == 1:
            g = bm.group(1)
            cur = (g + "教学楼") if g.startswith("A") else (g + "号教学楼")
            site.setdefault(cur, {})
            continue
        if cur is None or len(cells) < 2:
            continue
        for sid in slots_of(cells[0]):
            rooms, voice = parse_rooms(cells[1])
            site[cur][sid] = {
                "rooms": rooms, "voice": voice,
                "note": "" if rooms or voice else
                {"无": "本时段无空教室", "准备考场": "准备考场，不开放"}.get(cells[1], cells[1]),
            }
    return site


def slots_of(label):
    lab = label.replace("：", ":").strip()
    m = re.match(r"^第\s*(\d+)\s*[-—–]\s*(\d+)\s*节", lab)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        return [s for s in SLOT_IDS if overlaps(s, a, b)]
    m = re.match(r"^(\d{1,2}):(\d{2})\s*[-—–]\s*(\d{1,2}):(\d{2})", lab)
    if m:
        h = int(m.group(1))
        return ["1-2", "3-4"] if h < 12 else ["5-6", "7-8", "9-11"]
    return []


def overlaps(sid, a, b):
    lo, hi = [int(x) for x in sid.split("-")]
    return not (hi < a or lo > b)


ROOM_RE = re.compile(r"[A-Z]?\d{2,3}(?:\s*[-—–]\s*[A-Z]?\d{2,3})?")


def parse_rooms(cell):
    cell = cell.replace("，", "、")
    main, voice = cell, ""
    if "有声教室" in cell:
        main, voice = cell.split("有声教室", 1)
        voice = voice.lstrip("：: ")
    return expand(main), expand(voice)


def expand(s):
    out = []
    for tok in ROOM_RE.findall(s):
        tok = tok.replace(" ", "")
        m = re.match(r"^([A-Z]?)(\d{2,3})[-—–]([A-Z]?)(\d{2,3})$", tok)
        if m and m.group(2) and m.group(4) and int(m.group(4)) > int(m.group(2)):
            p = m.group(1) or m.group(3)
            for n in range(int(m.group(2)), int(m.group(4)) + 1):
                out.append(f"{p}{n}")
        elif not re.fullmatch(r"\d{1,2}", tok):   # 排除 "07:00" 里的数字
            out.append(tok)
    seen, uniq = set(), []
    for r in out:
        if r not in seen:
            seen.add(r)
            uniq.append(r)
    return uniq


def floor_of(room):
    m = re.match(r"^[A-Z]?(\d+)", room)
    return int(m.group(1)) // 100 if m and len(m.group(1)) >= 3 else 0


# ------------------------------------------------------------------ 排序 / 推荐
def room_num(room):
    m = re.match(r"^[A-Z]?(\d+)", room)
    return int(m.group(1)) if m else 0


def pick(site, sid, anchor, free):
    """有课时：只在本教学楼内选，按房间号接近 → 楼层。没课：4/5号楼优先。"""
    if anchor and anchor.get("building"):
        b = f"{anchor['building']}号教学楼" if str(anchor["building"]).isdigit() else "信息楼"
        d = (site.get(b) or {}).get(sid)
        if not d:
            return [], b
        an, af = anchor["room"], anchor["floor"]
        items = sorted((abs(room_num(r) - an), abs(floor_of(r) - af), r)
                       for r in d["rooms"] + d["voice"])
        return [{"room": r, "bname": b, "why": why(an, af, r)} for _, _, r in items[:3]], b
    cand = []
    for bname, slots in site.items():
        dd = slots.get(sid)
        if not dd:
            continue
        for r in dd["rooms"] + dd["voice"]:
            cand.append(((0 if bname in PRIOR else 1), floor_of(r), room_num(r), r, bname))
    cand.sort()
    return ([{"room": r, "bname": bn, "why": "4/5号楼优先" if bn in PRIOR else f"{floor_of(r)}层"}
             for _, _, _, r, bn in cand[:3]], None)


def why(an, af, r):
    dn, df = room_num(r) - an, floor_of(r) - af
    parts = ["同号" if dn == 0 else f"号差{abs(dn)}"]
    parts.append("同层" if df == 0 else (("上" if df > 0 else "下") + f"{abs(df)}层"))
    return " · ".join(parts)


# ------------------------------------------------------------------ 渲染
WD = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]

CSS = """
.daybar{display:flex;gap:8px;overflow-x:auto;padding:0 0 12px;margin:0 0 4px}
.dbtn{flex:none;border:1px solid var(--line);background:#fff;border-radius:12px;padding:7px 12px;cursor:pointer;font:inherit;line-height:1.15;text-align:center;color:var(--ink)}
.dbtn b{display:block;font-size:13.5px;font-weight:600}
.dbtn i{display:block;font-style:normal;font-size:11.5px;color:var(--sub);margin-top:1px}
.dbtn.on{background:var(--acc);border-color:var(--acc);color:#fff}
.dbtn.on i{color:rgba(255,255,255,.85)}
:root{--ink:#1c2128;--sub:#6b7280;--line:#e8eaef;--bg:#f6f7f9;--acc:#0a7d5a;--warn:#fff7e6;--warnd:#d48806}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.6 -apple-system,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif}
.wrap{max-width:860px;margin:0 auto;padding:20px 16px 48px}
header h1{margin:0 0 4px;font-size:22px;letter-spacing:.3px}
header .m{color:var(--sub);font-size:13px}
.warn{margin:12px 0;padding:10px 12px;border-radius:10px;background:var(--warn);border:1px solid #ffe1a8;color:#874d00;font-size:13px}
.tabs{position:sticky;top:0;z-index:5;display:flex;gap:6px;overflow:auto;padding:10px 0;background:linear-gradient(var(--bg) 72%,rgba(246,247,249,0))}
.tab{flex:none;border:1px solid var(--line);background:#fff;border-radius:999px;padding:7px 14px;font-size:13px;cursor:pointer;color:var(--sub);white-space:nowrap}
.tab.on{background:var(--ink);border-color:var(--ink);color:#fff}
.bar{display:flex;gap:8px;align-items:center;margin:6px 0 14px}
.bar input{flex:1;border:1px solid var(--line);border-radius:10px;padding:9px 12px;font-size:14px;background:#fff;outline:none}
.bar input:focus{border-color:#c7ccd6}
.sw{flex:none;border:1px solid var(--line);background:#fff;border-radius:10px;padding:9px 12px;font-size:13px;cursor:pointer;color:var(--sub)}
.sw.on{background:#e8f6f0;border-color:#9fd8bf;color:var(--acc)}
.day{display:none}
.day.on{display:block}
.panel{display:none}
.panel.on{display:block}
.now{display:flex;gap:10px;align-items:baseline;background:#fff;border:1px solid var(--line);border-radius:14px;padding:12px 14px;margin-bottom:12px}
.now .k{font-size:12px;color:var(--sub)}
.now .v{font-size:16px;font-weight:600}
.now .s{font-size:13px;color:var(--sub)}
.course{background:#eef4ff;border:1px solid #d6e2f7;border-radius:12px;padding:10px 12px;font-size:13.5px;margin-bottom:10px}
.course b{font-weight:600}
.recs{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px}
.rec{display:flex;align-items:center;gap:7px;background:#fff;border:1px solid var(--line);border-radius:999px;padding:6px 12px 6px 8px;font-size:13px}
.rec .i{width:18px;height:18px;border-radius:50%;background:var(--ink);color:#fff;font-size:11px;display:grid;place-items:center;flex:none}
.rec .w{color:var(--sub);font-size:12px}
.card{background:#fff;border:1px solid var(--line);border-radius:14px;padding:12px 14px;margin-bottom:10px}
.card.p{border-color:#9fd8bf;box-shadow:0 0 0 3px #f0faf5 inset}
.ch{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.ch .nm{font-weight:600;font-size:14px}
.badge{font-size:11px;padding:2px 7px;border-radius:999px;background:#e8f6f0;color:var(--acc);border:1px solid #c9ecdb}
.ip{font-size:13px;color:var(--sub)}
.chips{display:flex;flex-wrap:wrap;gap:6px}
.chip{border:1px solid var(--line);background:#fafbfc;border-radius:8px;padding:4px 9px;font-size:13px;cursor:pointer;font-variant-numeric:tabular-nums}
.chip:active{background:#eef1f5}
.chip.v{background:#fff8e6;border-color:#ffe1a8;color:#874d00}
.none{color:var(--sub);font-size:13px}
.note{font-size:13px;color:#874d00;background:var(--warn);border:1px solid #ffe1a8;border-radius:10px;padding:8px 11px;margin-bottom:10px}
.sched{background:#fff;border:1px solid var(--line);border-radius:14px;padding:8px 14px;margin:12px 0}
.sched .row{display:flex;gap:8px;align-items:baseline;font-size:13.5px;padding:5px 0;border-bottom:1px dashed #f0f1f4}
.sched .row:last-child{border:0}
.sched .t{color:var(--sub);width:56px;flex:none;font-size:12.5px}
.sched .c{font-weight:600;flex:1 1 auto;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sched .p{color:var(--sub);text-align:right;flex:0 0 auto;max-width:46%;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:12.5px}
.sched .row.free .c{font-weight:400;color:#9aa1ab}
footer{margin-top:22px;color:var(--sub);font-size:12px;line-height:1.8}
footer a{color:var(--sub)}
#toast{position:fixed;left:50%;bottom:32px;transform:translate(-50%,20px);background:#111;color:#fff;padding:9px 16px;border-radius:10px;font-size:13px;opacity:0;transition:.25s;pointer-events:none;z-index:99}
#toast.on{opacity:1;transform:translate(-50%,0)}

/* ---------- 动效层：按下回弹 / 状态过渡 / 层次光影 ---------- */
:root{--spring:cubic-bezier(.34,1.56,.64,1);--ease:cubic-bezier(.32,.72,0,1)}
*{-webkit-tap-highlight-color:transparent}
button{font:inherit;touch-action:manipulation}
.tab,.dbtn,.sw,.chip,.rec,.card,.bar input{
  transition:transform .26s var(--spring),background-color .2s var(--ease),
             border-color .2s var(--ease),box-shadow .3s var(--ease),color .2s var(--ease);
  will-change:transform}
.tab:active,.sw:active{transform:scale(.94)}
.dbtn:active{transform:scale(.94)}
.chip:active{transform:scale(.88)}
.card:active{transform:scale(.986)}
@media (hover:hover){
  .tab:hover,.dbtn:hover,.sw:hover{transform:translateY(-1px)}
  .tab:hover{box-shadow:0 6px 16px -12px rgba(16,24,40,.45)}
  .card:hover{box-shadow:0 10px 26px -20px rgba(16,24,40,.45)}
  .chip:hover{background:#f1f4f8;border-color:#dbe0e8}
}
.tab.on{box-shadow:0 8px 18px -12px rgba(28,33,40,.75)}
.dbtn.on{box-shadow:0 8px 18px -10px rgba(10,125,90,.55)}
.sw.on{box-shadow:0 6px 16px -12px rgba(10,125,90,.5)}
.bar input:focus{box-shadow:0 0 0 3.5px rgba(10,125,90,.13);border-color:#9fd8bf}
.tabs{-webkit-backdrop-filter:saturate(180%) blur(14px);backdrop-filter:saturate(180%) blur(14px);
  background:rgba(246,247,249,.84)}
.chip.pop{animation:chipPop .42s var(--spring)}
@keyframes chipPop{0%{transform:scale(1)}38%{transform:scale(1.16)}100%{transform:scale(1)}}
.panel.on{animation:panelIn .32s var(--ease) both}
@keyframes panelIn{from{opacity:0;transform:translateY(7px)}to{opacity:1;transform:none}}
.panel.on .card,.panel.on .course,.panel.on .rec,.panel.on .note{animation:itemIn .34s var(--ease) both}
.panel.on .course{animation-delay:0s}
.panel.on .rec:nth-child(1){animation-delay:.02s}
.panel.on .rec:nth-child(2){animation-delay:.05s}
.panel.on .rec:nth-child(3){animation-delay:.08s}
.panel.on .card:nth-of-type(1){animation-delay:.04s}
.panel.on .card:nth-of-type(2){animation-delay:.07s}
.panel.on .card:nth-of-type(3){animation-delay:.1s}
.panel.on .card:nth-of-type(4){animation-delay:.13s}
@keyframes itemIn{from{opacity:0;transform:translateY(8px) scale(.995)}to{opacity:1;transform:none}}
.day.on{animation:dayIn .3s var(--ease) both}
@keyframes dayIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
.warn{animation:itemIn .36s var(--ease) both}
.now .v{transition:color .25s var(--ease)}
#toast{transition:opacity .22s var(--ease),transform .45s var(--spring)}
#toast.on{transform:translate(-50%,0)}
@media (prefers-reduced-motion:reduce){
  *,*::before,*::after{animation-duration:.001s!important;animation-delay:0s!important;transition-duration:.001s!important}
}
"""

JS = """
(function(){
  var DAYS=window.__DAYS__||[];
  function pad(n){return (n<10?'0':'')+n}
  function iso(d){return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())}
  var now=new Date(), today=iso(now);
  var tom=new Date(now.getFullYear(),now.getMonth(),now.getDate()+1), want=iso(tom);
  var WEEK=['日','一','二','三','四','五','六'];
  function md(s){return (+s.slice(5,7))+'月'+(+s.slice(8,10))+'日'}
  var toast=document.getElementById('toast'),tt;
  function say(m){if(!toast)return;toast.textContent=m;toast.classList.add('on');
    clearTimeout(tt);tt=setTimeout(function(){toast.classList.remove('on')},1400)}
  if(!DAYS.length)return;

  function bump(el){
    if(!el)return;
    el.classList.remove('pop'); void el.offsetWidth; el.classList.add('pop');
    setTimeout(function(){el.classList.remove('pop')},460);
  }
  function legacyCopy(t,el,done){
    var ta=document.createElement('textarea');
    ta.value=t; ta.setAttribute('readonly','');
    ta.style.cssText='position:fixed;top:-1000px;left:0;opacity:0';
    document.body.appendChild(ta); ta.select();
    var ok=false;
    try{ ok=document.execCommand('copy') }catch(e){}
    document.body.removeChild(ta);
    done(ok?('已复制 '+t):('记一下：'+t));
  }
  function copyRoom(t,el){
    var settled=false;
    var flash=function(msg){ if(settled)return; settled=true; bump(el); say(msg) };
    var fall=function(){ if(settled)return; settled=true;
      legacyCopy(t,el,function(msg){ bump(el); say(msg) }) };
    if(navigator.clipboard&&navigator.clipboard.writeText){
      try{
        navigator.clipboard.writeText(t).then(function(){flash('已复制 '+t)},fall);
        setTimeout(fall,900);
        return;
      }catch(e){}
    }
    fall();
  }
  document.addEventListener('click',function(e){
    var c=e.target.closest&&e.target.closest('.chip'); if(!c)return;
    copyRoom(c.dataset.full||'', c);
  });

  var idx=-1,i;
  for(i=0;i<DAYS.length;i++){ if(DAYS[i].date===want){idx=i;break} }
  if(idx<0){ for(i=0;i<DAYS.length;i++){ if(DAYS[i].date===today){idx=i;break} } }
  if(idx<0)idx=0;

  var wrap=document.querySelector('.wrap'), nav=document.createElement('div');
  nav.className='daybar';
  DAYS.forEach(function(d,k){
    var b=document.createElement('button'); b.className='dbtn'+(k===idx?' on':'');
    b.dataset.date=d.date;
    var tag=(d.date===today?'今天':(d.date===want?'明天':'周'+WEEK[new Date(d.date+'T00:00:00').getDay()]));
    b.innerHTML='<b>'+tag+'</b><i>'+md(d.date)+'</i>';
    b.onclick=function(){show(k)};
    nav.appendChild(b);
  });
  wrap.insertBefore(nav, wrap.firstChild);

  function show(k){
    var D=DAYS[k];
    var box=document.querySelector('.day[data-date="'+D.date+'"]');
    if(!box)return;
    Array.prototype.forEach.call(document.querySelectorAll('.day'),function(x){x.classList.remove('on')});
    box.classList.add('on');
    Array.prototype.forEach.call(nav.querySelectorAll('.dbtn'),function(x){
      x.classList.toggle('on',x.dataset.date===D.date)});
    document.title=md(D.date)+'自习室 · 海大';
    setup(box,D);
    window.scrollTo({top:0,behavior:'smooth'});
  }

  function setup(box,D){
    var $=function(s){return box.querySelector(s)};
    var $$=function(s){return [].slice.call(box.querySelectorAll(s))};
    Array.prototype.forEach.call(box.querySelectorAll('.stale'),function(x){x.remove()});
    if(D.date!==want&&D.date!==today){
      var w=document.createElement('div'); w.className='warn stale';
      w.innerHTML='<b>这是 '+md(D.date)+' 的数据</b>，不是明天（'+md(want)+'）。学校公告还没覆盖到这一天，更新后自动出现。';
      var h=box.querySelector('header'); if(h)h.insertAdjacentElement('afterend',w);
    }
    var q='',only=false;
    function apply(){
      $$('.panel').forEach(function(p){
        Array.prototype.forEach.call(p.querySelectorAll('.card'),function(c){
          var hit=!q||c.dataset.k.indexOf(q)>=0;
          if(only&&c.dataset.p!=='1')hit=false;
          c.style.display=hit?'':'none';
        });
      });
    }
    var qi=$('.q'); if(qi)qi.oninput=function(){q=this.value.trim().toLowerCase();apply()};
    var sw=$('.sw'); if(sw)sw.onclick=function(){only=!only;this.classList.toggle('on',only);apply()};
    var tabs=$$('.tab'), panels=$$('.panel');
    tabs.forEach(function(t){
      t.onclick=function(){
        tabs.forEach(function(x){x.classList.remove('on')});
        panels.forEach(function(x){x.classList.remove('on')});
        t.classList.add('on');
        var p=panels.filter(function(x){return x.dataset.id===t.dataset.id})[0];
        if(p)p.classList.add('on');
      };
    });
    var mins=now.getHours()*60+now.getMinutes(), hit=null, nxt=null;
    (D.slots_meta||[]).forEach(function(m){
      var a=m.start.split(':'), b=m.end.split(':');
      var s=+a[0]*60+ +a[1], e=+b[0]*60+ +b[1];
      if(mins>=s&&mins<e)hit=m.id;
      if(!nxt&&mins<s)nxt=m;
    });
    var label=$('.now .v'), sub=$('.now .s');
    function byId(id){return D.slots.filter(function(x){return x.id===id})[0]||{recs:[]}}
    function rec0(o){var r=(o.recs||[])[0];return r?r.bname+' '+r.room:'暂无'}
    if(D.date!==today){
      var f0=D.slots.filter(function(x){return x.course})[0]||D.slots[0];
      label.textContent=md(D.date);
      sub.textContent='第一节课 '+f0.label+' '+f0.start+(f0.course?' · '+f0.course.place:' · 没课');
    } else if(hit){
      var s0=byId(hit);
      label.textContent=s0.label+'（'+s0.start+'–'+s0.end+'）'+(s0.course?'有课':'没课');
      sub.textContent=(s0.recs&&s0.recs.length)?('去 '+rec0(s0)+' '+(s0.recs[0].why||'')):'本时段无空教室';
    } else if(nxt){
      var gap=+nxt.start.split(':')[0]*60+ +nxt.start.split(':')[1]-mins;
      label.textContent='距 '+nxt.label+' 还有 '+gap+' 分钟';
      sub.textContent='可先去 '+rec0(byId(nxt.id));
    } else {
      label.textContent='今天的自习时段结束了'; sub.textContent='明天见';
    }
    if(hit&&D.date===today){
      var t0=tabs.filter(function(t){return t.dataset.id===hit})[0];
      if(t0)t0.click();
    } else if(tabs.length){ tabs[0].click(); }
  }

  show(idx);
})();
"""


def day_parts(target, notice, week, courses, gen_at):
    wd = WD[target.weekday()]
    key = f"{target.month}-{target.day}"
    site_day = notice["days"].get(key)
    fresh = site_day is not None
    site = site_day or {}
    slots = []
    for s in TB.SLOTS:
        sid = s["id"]
        cs = courses.get(sid, [])
        anchor = None
        if cs and cs[0]["place"]["building"]:
            anchor = cs[0]["place"]
        free = not cs
        ab = cs[0]["place"]["building"] if cs else None
        anchor_b = ((f"{ab}号教学楼" if str(ab).isdigit() else "信息楼") if ab else None)
        blocks = []
        for bname in sorted(site, key=lambda b: (b not in PRIOR, b)):
            d = site[bname].get(sid)
            if not d:
                continue
            blocks.append({"name": bname, "rooms": d["rooms"], "voice": d["voice"],
                           "note": d["note"],
                           "prior": bname in PRIOR and free and bool(d["rooms"] or d["voice"])})
        recs, _rb = pick(site, sid, anchor, free)
        rec_note = ""
        if not fresh:
            rec_note = ""
        elif cs and not recs:
            blk = next((x for x in blocks if x["name"] == anchor_b), None)
            tail = (blk["note"] if blk and blk["note"] else "公告未列入")
            rec_note = f"{anchor_b} 本时段没有空教室（{tail}），按规则不跨楼推荐。"
        elif not recs:
            rec_note = "本时段公告里没有可用的空教室。"
        slots.append({
            "id": sid, "label": s["label"], "start": s["start"], "end": s["end"],
            "course": ({"name": cs[0]["name"], "place": TB.place_text(cs[0]["place"]),
                        "place_s": TB.place_text(cs[0]["place"]).split("（")[0],
                        "teacher": cs[0]["teacher"]} if cs else None),
            "recs": recs, "anchor": anchor_b, "recNote": rec_note,
            "blocks": blocks,
        })

    data = {"date": target.isoformat(), "slots": slots, "slots_meta": TB.SLOTS,
            "gen": gen_at, "fresh": fresh, "week": week}

    avail = []
    for k in notice["days"]:
        mo, dy = k.split("-")
        try:
            avail.append(datetime.date(target.year, int(mo), int(dy)))
        except ValueError:
            pass
    latest = max(avail) if avail else None
    if not notice["days"]:
        banner = '<div class="warn"><b>公告解析失败</b>，页面暂无自习室数据，脚本会自动重试。</div>'
    elif fresh:
        banner = ""
    elif latest and target > latest:
        banner = (f'<div class="warn"><b>{target.month}月{target.day}日 的公告还没发布</b>'
                  f'（最新覆盖到 {latest.month}月{latest.day}日）——自习室信息会在当晚 21:30 后更新，'
                  f'这里先显示当天课程与楼栋。</div>')
    else:
        banner = (f'<div class="warn"><b>没抓到 {target.month}月{target.day}日 的公告</b>'
                  f'——下方按最近一次数据展示，会自动重试。</div>')

    sched = '<div class="sched">' + "".join(
        f'<div class="row{" free" if not sd["course"] else ""}"><span class="t">{sd["label"]}</span>'
        f'<span class="c">{sd["course"]["name"] if sd["course"] else "没课"}</span>'
        f'<span class="p">{sd["course"]["place_s"] if sd["course"] else sd["start"] + "–" + sd["end"]}</span></div>'
        for sd in slots) + '</div>'

    panels = []
    for sd in slots:
        bl = []
        for b in sd["blocks"]:
            chips = "".join(
                f'<span class="chip" data-full="{b["name"]}{r}" title="{b["name"]}{r}">{r}</span>'
                for r in b["rooms"])
            voice = "".join(
                f'<span class="chip v" data-full="{b["name"]}{r}（有声）" title="有声教室">{r}♪</span>'
                for r in b["voice"])
            note = f'<div class="none">{html.escape(b["note"])}</div>' if b["note"] else ""
            inner = chips or voice or note
            if chips and voice:
                inner = (f'<div class="ip">普通</div><div class="chips">{chips}</div>'
                         f'<div class="ip" style="margin-top:6px">有声教室</div><div class="chips">{voice}</div>')
            bl.append(f'<div class="card{" p" if b["prior"] else ""}" data-p="{1 if b["prior"] else 0}" '
                      f'data-k="{(b["name"] + "".join(b["rooms"]) + "".join(b["voice"])).lower()}">'
                      f'<div class="ch"><span class="nm">{b["name"]}</span>'
                      f'{"<span class=badge>优先</span>" if b["prior"] else ""}</div>{inner}</div>')
        if sd["course"]:
            c = sd["course"]
            course = (f'<div class="course"><b>{c["name"]}</b> · {c["place"]}'
                      f'<span style="color:#6b7280">（{c["teacher"]}）</span></div>')
        else:
            course = '<div class="course" style="background:#f7f8fa;border-color:#e8eaef">本时段没课 · 4/5号楼优先</div>'
        rec_html = "".join(
            f'<span class="rec"><span class="i">{i+1}</span>{r["bname"]} {r["room"]}'
            f'<span class="w">{r["why"]}</span></span>'
            for i, r in enumerate(sd["recs"]))
        rec_html = f'<div class="recs">{rec_html}</div>' if rec_html else ""
        if sd.get("recNote"):
            rec_html = f'<div class="note">{html.escape(sd["recNote"])}</div>' + rec_html
        panels.append(f'<section class="panel" data-id="{sd["id"]}">{course}{rec_html}'
                      f'{"".join(bl) or "<div class=none>本时段无数据</div>"}</section>')

    tabs_html = "".join(
        f'<button class="tab" data-id="{s["id"]}">{s["label"]}{"·课" if s["course"] else ""}</button>'
        for s in slots)

    return {
        "date": target.isoformat(),
        "h1": f'{target.month}月{target.day}日{wd} 自习室',
        "meta": (f'第 {week} 周 · 数据更新 {gen_at}'
                 + (f' · 公告 {notice["publish"]}' if notice["publish"] else '')),
        "banner": banner, "sched": sched, "tabs": tabs_html, "panels": "".join(panels),
        "data": data,
    }


def compose(days):
    blocks = "".join(
        f'<div class="day" data-date="{p["date"]}">'
        f'<header><h1>{p["h1"]}</h1><div class="m">{p["meta"]}</div></header>'
        f'{p["banner"]}{p["sched"]}'
        '<div class="now"><div class="k">现在</div><div><div class="v">—</div><div class="s"></div></div></div>'
        f'<div class="tabs">{p["tabs"]}</div>'
        '<div class="bar"><input class="q" placeholder="搜房号或楼名，如 414 / 5号">'
        '<button class="sw">只看4/5号楼</button></div>'
        f'{p["panels"]}</div>'
        for p in days)
    payload = json.dumps([p["data"] for p in days], ensure_ascii=False)
    footer = (
        '<footer>来源：<a href="' + URL + '" target="_blank">海南大学后勤保障部 · 自习室预告</a>'
        '　本页每晚 23:00 自动重抓次日公告，永久同一网址。'
        '<br>规则：没课时优先 4/5 号楼；有课时只在本教学楼内，按房间号接近 → 楼层排序。点房号可复制。</footer>')
    return (
        '<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
        '<meta name="theme-color" content="#0a7d5a">'
        '<meta name="apple-mobile-web-app-capable" content="yes">'
        '<meta name="apple-mobile-web-app-status-bar-style" content="default">'
        '<meta name="apple-mobile-web-app-title" content="自习室">'
        '<link rel="manifest" href="manifest.webmanifest">'
        '<link rel="apple-touch-icon" href="apple-touch-icon.png">'
        '<title>次日自习室 · 海大</title><style>' + CSS + '</style></head><body>'
        '<div class="wrap">' + blocks + footer + '</div><div id="toast"></div>'
        '<script>window.__DAYS__=' + payload + ';' + JS + '</script></body></html>'
    )


# ------------------------------------------------------------------ 主流程
def main():
    today = now_sh().date()
    gen_at = now_sh().strftime("%m-%d %H:%M")
    page = fetch()
    notice = parse_notice(page)
    days = []
    for k in range(0, 8):
        target = today + datetime.timedelta(days=k)
        _, week, courses = TB.load_courses(target)
        days.append(day_parts(target, notice, week, courses, gen_at))
    out = compose(days)
    with open(os.path.join(BASE, "index.html"), "w", encoding="utf-8") as f:
        f.write(out)
    msg = (f'{now_sh():%Y-%m-%d %H:%M:%S} 生成 {len(days)} 天 '
           f'(公告覆盖={"/".join(sorted(notice["days"])) or "无"} 更新={gen_at})')
    with open(os.path.join(BASE, "更新日志.txt"), "a", encoding="utf-8") as f:
        f.write(msg + "\n")
    print(msg)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        with open(os.path.join(BASE, "更新日志.txt"), "a", encoding="utf-8") as f:
            f.write(f'{now_sh():%Y-%m-%d %H:%M:%S} 失败\n{traceback.format_exc()}\n')
        raise SystemExit(1)
