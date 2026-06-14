"""시스템 지도 · 로드맵 — KB의 systemmap.json을 런타임 렌더 + 실제 페이지 바로가기. (ADR 0019)

코드(렌더러)는 공개 app repo, 데이터(systemmap.json)는 private KB repo에서 시크릿 PAT로 런타임 read.
"""
import json
import urllib.request
import urllib.error
import streamlit as st
import streamlit.components.v1 as components

KB_REPO = "OURPROJECTDAO/work-automation-wb"
SM_PATH = "systemmap.json"


def _candidate_pats():
    pats = []
    try:
        p = st.secrets["data"]["pat"]
        if p:
            pats.append(p)
    except Exception:
        pass
    try:
        p = st.secrets.get("GITHUB_PAT", "")
        if p and p not in pats:
            pats.append(p)
    except Exception:
        pass
    return pats


@st.cache_data(ttl=300)
def _fetch_systemmap():
    pats = _candidate_pats()
    if not pats:
        return None, "PAT 시크릿 없음 ([data] pat 또는 GITHUB_PAT)"
    url = "https://api.github.com/repos/%s/contents/%s" % (KB_REPO, SM_PATH)
    last = ""
    for pat in pats:
        req = urllib.request.Request(url, headers={
            "Authorization": "Bearer " + pat,
            "Accept": "application/vnd.github.raw",
            "User-Agent": "wa-app",
        })
        try:
            with urllib.request.urlopen(req, timeout=15) as r:
                txt = r.read().decode("utf-8")
            json.loads(txt)
            return txt, None
        except urllib.error.HTTPError as e:
            last = "HTTP %s" % e.code
        except Exception as e:  # noqa: BLE001
            last = str(e)
    return None, "KB repo(%s) 읽기 실패 (%s) — 시크릿 PAT에 해당 private repo 읽기 권한 필요" % (KB_REPO, last)


TEMPLATE = r'''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>시스템 지도 + 로드맵</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
<style>
  :root{
    --bg:#E9E7DF; --card:#FCFBF8; --ink:#1A2B26; --soft:#586862; --faint:#8C968F;
    --line:#D3D0C5; --petrol:#0E4A43;
    --live:#2E7D5B; --liveb:#E5F1EA;
    --partial:#A9761A; --partialb:#F4ECDA;
    --design:#C5402A; --designb:#FBEDE9;
    --concept:#7A8580; --conceptb:#EDEDE7;
    --hub:#0E4A43; --hubb:#E3EEEC;
    --mono:"JetBrains Mono",ui-monospace,monospace;
    --sans:"Pretendard",-apple-system,system-ui,sans-serif;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  html{-webkit-text-size-adjust:100%}
  body{background:var(--bg);color:var(--ink);font-family:var(--sans);line-height:1.5;
    -webkit-font-smoothing:antialiased;padding:clamp(16px,4vw,40px) clamp(12px,4vw,36px) 70px}
  .wrap{max-width:1180px;margin:0 auto}

  .eyebrow{font-family:var(--mono);font-size:11.5px;font-weight:500;letter-spacing:.18em;
    text-transform:uppercase;color:var(--petrol);display:flex;align-items:center;gap:10px}
  .eyebrow::after{content:"";flex:1;height:1px;background:var(--line)}
  h1{font-size:clamp(26px,5.5vw,42px);font-weight:900;letter-spacing:-.02em;line-height:1.1;margin:14px 0 0}
  .sub{font-size:clamp(14px,3vw,16px);color:var(--soft);margin-top:10px;max-width:62ch}
  .sub b{color:var(--ink);font-weight:600}

  /* segmented toggle */
  .switch{display:inline-flex;background:var(--card);border:1px solid var(--line);border-radius:11px;
    padding:4px;margin-top:20px;gap:3px}
  .switch button{font:inherit;font-weight:700;font-size:13.5px;border:none;background:transparent;color:var(--soft);
    padding:8px 18px;border-radius:8px;cursor:pointer;transition:.15s;display:inline-flex;align-items:center;gap:8px}
  .switch button .cnt{font-family:var(--mono);font-size:10px;font-weight:700;background:#E6E4DB;color:var(--faint);
    border-radius:10px;padding:1px 7px}
  .switch button.active{background:var(--ink);color:#fff}
  .switch button.active .cnt{background:rgba(255,255,255,.2);color:#fff}

  .view{display:none}
  .view.active{display:block;animation:fade .25s ease}
  @keyframes fade{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}

  /* legend */
  .legend{display:flex;flex-wrap:wrap;gap:5px 13px;margin-top:18px}
  .lg{display:inline-flex;align-items:center;gap:6px;font-size:12px;color:var(--soft);font-weight:500}
  .dot{width:9px;height:9px;border-radius:50%;flex:none}
  .dot.live{background:var(--live)} .dot.partial{background:var(--partial)}
  .dot.design{background:var(--design)} .dot.concept{background:var(--concept)}

  /* ===== MAP VIEW ===== */
  .layout{display:grid;grid-template-columns:1fr;gap:18px;margin-top:20px}
  @media (min-width:920px){.layout{grid-template-columns:1fr 350px;align-items:start}}
  .band{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:16px 16px 18px;margin-bottom:16px}
  .band-h{display:flex;align-items:baseline;gap:10px;margin-bottom:12px}
  .band-tag{font-family:var(--mono);font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;
    color:var(--card);background:var(--ink);padding:3px 8px;border-radius:5px}
  .band-h h2{font-size:16px;font-weight:800;letter-spacing:-.01em}
  .band-h .note{font-size:12px;color:var(--faint);margin-left:auto;font-weight:500}
  .assets{display:flex;flex-wrap:wrap;gap:7px}
  .asset{font-family:var(--mono);font-size:12px;font-weight:500;color:var(--petrol);background:#F3F1EB;
    border:1px solid var(--line);border-radius:8px;padding:6px 10px;cursor:pointer;transition:.18s;
    display:inline-flex;align-items:center;gap:7px}
  .asset .ct{font-size:10px;color:var(--faint);background:#E6E4DB;border-radius:10px;padding:1px 6px;font-weight:700}
  .asset.hub{background:var(--hubb);border-color:#B9D2CC;color:var(--hub);font-weight:700}
  .asset.hub .ct{background:#CADFD9;color:var(--hub)}
  .asset:hover{border-color:var(--petrol)}
  .cluster{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:16px;margin-bottom:16px}
  .cl-h{display:flex;align-items:baseline;gap:10px;margin-bottom:14px}
  .cl-num{font-family:var(--mono);font-size:11px;font-weight:700;color:var(--faint)}
  .cl-h h2{font-size:17px;font-weight:800;letter-spacing:-.01em}
  .cl-h .flow{font-size:11.5px;color:var(--faint);margin-left:auto;font-weight:500;font-family:var(--mono)}
  .row{display:flex;flex-wrap:wrap;gap:10px;align-items:stretch}
  .arrow{display:flex;align-items:center;color:var(--faint);font-size:18px;flex:none;align-self:center}
  .node{flex:1 1 165px;min-width:150px;max-width:280px;background:#FBFAF6;border:1.5px solid var(--line);
    border-radius:11px;padding:13px 14px;cursor:pointer;transition:.18s;text-align:left;position:relative;font:inherit;color:inherit}
  .node:hover{border-color:var(--petrol);transform:translateY(-1px)}
  .node .ntop{display:flex;align-items:center;gap:7px;margin-bottom:7px}
  .node .nstat{width:9px;height:9px;border-radius:50%;flex:none}
  .node.s-live .nstat{background:var(--live)} .node.s-partial .nstat{background:var(--partial)}
  .node.s-design .nstat{background:var(--design)} .node.s-concept .nstat{background:var(--concept)}
  .node .ncode{font-family:var(--mono);font-size:11.5px;font-weight:700;color:var(--petrol);line-height:1.2}
  .node.s-design .ncode{color:var(--design)}
  .node .star{margin-left:auto;color:var(--design);font-size:14px;line-height:1}
  .node .nlabel{font-size:14px;font-weight:700;letter-spacing:-.01em;margin-bottom:3px}
  .node .nline{font-size:11.5px;color:var(--soft);line-height:1.35}
  .node .nrm{margin-top:8px;font-family:var(--mono);font-size:10px;font-weight:700;color:var(--faint);
    display:inline-flex;align-items:center;gap:5px}
  .node .nrm .pip{width:6px;height:6px;border-radius:50%;background:var(--design)}
  .node.brain{border-width:2px;border-color:#E0B3AA;background:var(--designb)}
  .node.brain::before{content:"두뇌";position:absolute;top:-9px;right:12px;font-family:var(--mono);
    font-size:9px;font-weight:700;letter-spacing:.08em;background:var(--design);color:#fff;padding:2px 8px;border-radius:20px}
  .node.sel{border-color:var(--ink);box-shadow:0 0 0 3px rgba(26,43,38,.1),0 8px 20px rgba(26,43,38,.1)}
  .dimmable.dim{opacity:.32;filter:saturate(.5)}
  .asset.rel,.node.rel{box-shadow:0 0 0 2px var(--petrol)}
  .asset.sel{box-shadow:0 0 0 2px var(--ink);background:var(--hubb)}

  .detail{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px;position:sticky;top:16px;min-height:200px}
  .detail .empty{color:var(--faint);font-size:13.5px;line-height:1.6}
  .detail .empty b{color:var(--soft)}
  .d-code{font-family:var(--mono);font-size:13px;font-weight:700;color:var(--petrol);word-break:break-all}
  .d-pill{display:inline-block;font-family:var(--mono);font-size:10px;font-weight:700;letter-spacing:.04em;
    padding:3px 9px;border-radius:20px;margin-left:8px;vertical-align:middle}
  .d-pill.live{background:var(--liveb);color:var(--live)} .d-pill.partial{background:var(--partialb);color:var(--partial)}
  .d-pill.design{background:var(--designb);color:var(--design)} .d-pill.concept{background:var(--conceptb);color:var(--concept)}
  .d-label{font-size:18px;font-weight:800;letter-spacing:-.01em;margin:10px 0 6px}
  .d-line{font-size:13.5px;color:var(--soft);line-height:1.5;margin-bottom:6px}
  .d-sec{margin-top:13px}
  .d-sec .h{font-family:var(--mono);font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--faint);margin-bottom:5px}
  .d-sec ul{list-style:none;display:flex;flex-wrap:wrap;gap:5px}
  .d-sec li{font-family:var(--mono);font-size:11px;background:#F3F1EB;border:1px solid var(--line);border-radius:7px;padding:3px 8px;color:var(--petrol)}
  /* roadmap mini-list inside detail */
  .d-rm{margin-top:14px;display:grid;gap:7px}
  .d-rm .rm-i{border-left:2.5px solid var(--line);padding:2px 0 2px 10px}
  .d-rm .rm-i.t-next{border-color:var(--design)} .d-rm .rm-i.t-planned{border-color:var(--petrol)} .d-rm .rm-i.t-later{border-color:var(--concept)}
  .d-rm .rm-t{font-size:12.5px;font-weight:700;display:flex;align-items:center;gap:7px}
  .d-rm .ph{font-family:var(--mono);font-size:9px;font-weight:700;color:var(--faint);background:#EEECE5;border-radius:4px;padding:1px 5px}
  .d-rm .rm-d{font-size:11.5px;color:var(--soft);line-height:1.4;margin-top:2px}
  .d-meta{margin-top:14px;display:flex;flex-wrap:wrap;gap:6px 14px;font-family:var(--mono);font-size:11px;color:var(--faint)}
  .d-meta .pg{color:var(--petrol)}

  /* ===== ROADMAP VIEW ===== */
  .rm-intro{font-size:14px;color:var(--soft);margin:18px 0 6px;max-width:64ch}
  .rmcols{display:grid;grid-template-columns:1fr;gap:14px;margin-top:14px}
  @media (min-width:780px){.rmcols{grid-template-columns:repeat(3,1fr);align-items:start}}
  .rmcol{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:6px 6px 10px;overflow:hidden}
  .rmcol-h{padding:14px 14px 12px;display:flex;align-items:center;gap:9px;border-bottom:1px solid var(--line);margin-bottom:8px}
  .rmcol-h .bar{width:4px;height:18px;border-radius:3px;flex:none}
  .rmcol.next .bar{background:var(--design)} .rmcol.planned .bar{background:var(--petrol)} .rmcol.later .bar{background:var(--concept)}
  .rmcol-h h3{font-size:15px;font-weight:800;letter-spacing:-.01em}
  .rmcol-h .c{font-family:var(--mono);font-size:11px;font-weight:700;color:var(--faint);margin-left:auto}
  .rmcards{display:grid;gap:9px;padding:0 8px}
  .rmcard{background:#FBFAF6;border:1px solid var(--line);border-radius:10px;padding:12px 13px;position:relative}
  .rmcard .chips{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-bottom:7px}
  .chip{font-family:var(--mono);font-size:10px;font-weight:700;border-radius:6px;padding:2px 7px;cursor:pointer;border:1px solid transparent}
  .chip.live{background:var(--liveb);color:var(--live)} .chip.partial{background:var(--partialb);color:var(--partial)}
  .chip.design{background:var(--designb);color:var(--design)} .chip.concept{background:var(--conceptb);color:var(--concept)}
  .chip.sys{background:#ECECE6;color:var(--soft)}
  .chip:hover{border-color:currentColor}
  .rmcard .ph2{font-family:var(--mono);font-size:9px;font-weight:700;color:var(--faint);background:#EEECE5;border-radius:4px;padding:2px 6px}
  .rmcard h4{font-size:14px;font-weight:800;letter-spacing:-.01em;margin-bottom:4px}
  .rmcard h4 .st{color:var(--design)}
  .rmcard p{font-size:12.5px;color:var(--soft);line-height:1.45}
  .rmcard.lift{border-color:#E0B3AA;box-shadow:0 4px 14px rgba(197,64,42,.12)}
  .bl-h{font-family:var(--mono);font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
    color:var(--faint);padding:14px 14px 6px;margin-top:6px;border-top:1px dashed var(--line)}

  footer{margin-top:42px;padding-top:18px;border-top:1px solid var(--line);
    font-family:var(--mono);font-size:11px;color:var(--faint);display:flex;flex-wrap:wrap;gap:6px 16px}
  :focus-visible{outline:2px solid var(--petrol);outline-offset:2px}
  @media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}.node:hover{transform:none}}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div class="eyebrow">업무 자동화 · 시스템 지도 + 로드맵</div>
    <h1>지도와 로드맵,<br>한 소스에서</h1>
    <p class="sub"><b>지도</b>는 지금 무엇이 무엇과 데이터로 이어졌는지, <b>로드맵</b>은 무엇을 어떤 순서로 지을지. 같은 <b>systemmap.json</b>에서 나오고 서로 연결된다.</p>
    <div class="switch" role="tablist">
      <button id="tab-roadmap" class="active" role="tab">로드맵 <span class="cnt" id="rm-count"></span></button>
      <button id="tab-map" role="tab">지도</button>
    </div>
  </header>

  <!-- ===================== ROADMAP VIEW ===================== -->
  <div class="view active" id="view-roadmap">
    <p class="rm-intro">★ = 다음 한 수. 각 카드의 <b>코드 칩</b>을 누르면 지도에서 그 업무로 이동한다. 상태색: 운영 중 · 부분 · 미구현.</p>
    <div class="legend">
      <span class="lg"><span class="dot design"></span>설계확정 · 미구현</span>
      <span class="lg"><span class="dot partial"></span>부분 운영</span>
      <span class="lg"><span class="dot live"></span>운영 중</span>
      <span class="lg"><span class="dot concept"></span>공통 개념</span>
    </div>
    <div class="rmcols">
      <section class="rmcol next">
        <div class="rmcol-h"><span class="bar"></span><h3>★ 지금 만들 것</h3><span class="c" id="c-next"></span></div>
        <div class="rmcards" id="rm-next"></div>
      </section>
      <section class="rmcol planned">
        <div class="rmcol-h"><span class="bar"></span><h3>예정</h3><span class="c" id="c-planned"></span></div>
        <div class="rmcards" id="rm-planned"></div>
      </section>
      <section class="rmcol later">
        <div class="rmcol-h"><span class="bar"></span><h3>나중 · 백로그</h3><span class="c" id="c-later"></span></div>
        <div class="rmcards" id="rm-later"></div>
        <div class="bl-h">백로그 (가로지르는 항목)</div>
        <div class="rmcards" id="rm-backlog"></div>
      </section>
    </div>
  </div>

  <!-- ===================== MAP VIEW ===================== -->
  <div class="view" id="view-map">
    <div class="legend">
      <span class="lg"><span class="dot live"></span>운영 중</span>
      <span class="lg"><span class="dot partial"></span>부분 · 다음</span>
      <span class="lg"><span class="dot design"></span>설계확정 · 미구현</span>
      <span class="lg"><span class="dot concept"></span>공통 개념</span>
      <span class="lg"><span class="star" style="color:var(--design)">★</span> = 다음 한 수</span>
    </div>
    <div class="layout">
      <main>
        <section class="band">
          <div class="band-h"><span class="band-tag">Data backbone</span><h2>데이터 등뼈 — 공유 자산</h2>
            <span class="note">숫자 = 소비 워크플로우 수</span></div>
          <div class="assets" id="assets"></div>
        </section>
        <div id="clusters"></div>
      </main>
      <aside>
        <div class="detail" id="detail">
          <div class="empty"><b>노드를 선택하세요.</b><br>그 업무가 <b>소비하는 자산(←)</b>·<b>공급/연계(→)</b>·<b>로드맵 항목</b>·앱 페이지가 표시됩니다. 자산 칩을 누르면 그걸 쓰는 업무가 켜집니다.</div>
        </div>
      </aside>
    </div>
  </div>

  <footer>
    <span>source = systemmap.json (KB private)</span><span>·</span><span>갱신 2026-06-15</span><span>·</span><span>임베드 데이터 = 제안 스키마</span>
  </footer>
</div>

<script>
/* ===================== systemmap.json (제안 스키마 — 단일 진실원천) ===================== */
const MAP = __SYSTEMMAP_JSON__;

/* ===================== render ===================== */
const $=s=>document.querySelector(s);
const nodeById=Object.fromEntries(MAP.nodes.map(n=>[n.id,n]));
const assetById=Object.fromEntries(MAP.assets.map(a=>[a.id,a]));
const STAT_KR={live:"운영 중",partial:"부분 · 다음",design:"설계확정 · 미구현",concept:"공통 개념"};
const onnuri={id:"onnuri-order",status:"live",label:"온누리 발주서"}; // backlog ref (제이티)

/* ---- ROADMAP board ---- */
let rmCount=0;
const tiers={next:[],planned:[],later:[]};
MAP.nodes.forEach(n=>(n.roadmap||[]).forEach(r=>{tiers[r.tier].push({...r, node:n.id, status:n.status}); rmCount++;}));
function rmCardHTML(r){
  const node=nodeById[r.node]||onnuri;
  const ph=r.phase?`<span class="ph2">${r.phase}</span>`:"";
  return `<div class="rmcard${r.star?' lift':''}">
    <div class="chips"><button class="chip ${node.status}" data-jump="${r.node}">${r.node}</button>${ph}</div>
    <h4>${r.star?'<span class="st">★ </span>':''}${r.title}</h4><p>${r.detail||""}</p></div>`;
}
$("#rm-next").innerHTML=tiers.next.map(rmCardHTML).join("");
$("#rm-planned").innerHTML=tiers.planned.map(rmCardHTML).join("");
$("#rm-later").innerHTML=tiers.later.map(rmCardHTML).join("");
$("#rm-backlog").innerHTML=MAP.backlog.map(b=>{
  const node=b.node?(nodeById[b.node]||onnuri):null;
  const chip=node?`<button class="chip ${node.status}" data-jump="${b.node}">${b.node}</button>`:`<span class="chip sys">시스템</span>`;
  return `<div class="rmcard"><div class="chips">${chip}</div><h4>${b.title}</h4><p>${b.detail||""}</p></div>`;
}).join("");
$("#c-next").textContent=tiers.next.length;
$("#c-planned").textContent=tiers.planned.length;
$("#c-later").textContent=tiers.later.length+MAP.backlog.length;
$("#rm-count").textContent=rmCount+MAP.backlog.length;

/* ---- MAP: backbone ---- */
$("#assets").innerHTML=MAP.assets.map(a=>
  `<button class="asset dimmable${a.id==='product_master'?' hub':''}" id="asset-${a.id}" data-kind="asset" data-id="${a.id}">
    ${a.label}<span class="ct">${a.consumers.length}</span></button>`).join("");

/* ---- MAP: clusters ---- */
$("#clusters").innerHTML=MAP.clusters.map(cl=>{
  const ns=MAP.nodes.filter(n=>n.cluster===cl.id);
  let inner;
  if(cl.id==="fulfillment"){
    const o=["openmarket-merge","onnuri-order","logistics-order","cheonnyeon-upload","invoice-fill"].map(id=>ns.find(n=>n.id===id));
    inner=o.map((n,i)=>nodeHTML(n)+(i===2||i===3?`<span class="arrow">→</span>`:"")).join("");
  }else if(cl.id==="intelligence"){
    const o=["dashboard","channel-margin-monitor","upload-monitor","intelligence-layer"].map(id=>ns.find(n=>n.id===id));
    inner=o.map((n,i)=>nodeHTML(n)+(i===2?`<span class="arrow">→</span>`:"")).join("");
  }else{
    const c=ns.find(n=>n.id==="product-registration-common"), rest=ns.filter(n=>n.id!=="product-registration-common");
    inner=nodeHTML(c)+`<span class="arrow">→</span>`+rest.map(nodeHTML).join("");
  }
  return `<section class="cluster"><div class="cl-h"><span class="cl-num">${cl.num}</span><h2>${cl.label}</h2><span class="flow">${cl.flow}</span></div><div class="row">${inner}</div></section>`;
}).join("");
function nodeHTML(n){
  const rmN=(n.roadmap||[]).length;
  const rmTag=rmN?`<div class="nrm">${(n.roadmap.some(r=>r.tier==='next'))?'<span class="pip"></span>':''}로드맵 ${rmN}</div>`:"";
  return `<button class="node dimmable s-${n.status}${n.brain?' brain':''}" id="node-${n.id}" data-kind="node" data-id="${n.id}">
    <div class="ntop"><span class="nstat"></span><span class="ncode">${n.id}</span>${n.star?'<span class="star">★</span>':''}</div>
    <div class="nlabel">${n.label}</div><div class="nline">${n.line}</div>${rmTag}</button>`;
}

/* ---- adjacency + highlight ---- */
function relatedOfNode(id){
  const n=nodeById[id]; const nodes=new Set(), assets=new Set();
  (n.consumes||[]).forEach(c=>{ if(assetById[c]) assets.add(c); else if(nodeById[c]) nodes.add(c); });
  MAP.assets.forEach(a=>{ if(a.consumers.includes(id)) assets.add(a.id); });
  MAP.edges.forEach(e=>{ if(e.from===id) nodes.add(e.to); if(e.to===id) nodes.add(e.from); });
  return {nodes,assets};
}
function clearStates(){document.querySelectorAll('.dimmable').forEach(el=>el.classList.remove('dim','rel','sel'));}
function applyHighlight(keepN,keepA,selId,selKind){
  clearStates();
  document.querySelectorAll('.node').forEach(el=>{const id=el.dataset.id;
    if(id===selId&&selKind==='node')el.classList.add('sel'); else if(keepN.has(id))el.classList.add('rel'); else el.classList.add('dim');});
  document.querySelectorAll('.asset').forEach(el=>{const id=el.dataset.id;
    if(id===selId&&selKind==='asset')el.classList.add('sel'); else if(keepA.has(id))el.classList.add('rel'); else el.classList.add('dim');});
}
function rmMiniHTML(n){
  if(!(n.roadmap||[]).length) return "";
  const items=n.roadmap.map(r=>`<div class="rm-i t-${r.tier}"><div class="rm-t">${r.star?'<span style="color:var(--design)">★</span>':''}${r.phase?`<span class="ph">${r.phase}</span>`:''}${r.title}</div><div class="rm-d">${r.detail||""}</div></div>`).join("");
  return `<div class="d-sec"><div class="h">로드맵 (${n.roadmap.length})</div></div><div class="d-rm">${items}</div>`;
}
function showNode(id){
  const n=nodeById[id]; const {nodes,assets}=relatedOfNode(id);
  applyHighlight(nodes,assets,id,'node');
  const consumesList=(n.consumes||[]).map(c=>assetById[c]?assetById[c].label:c);
  const feedsList=(n.feeds||[]).map(f=>nodeById[f]?nodeById[f].id:f);
  let html=`<div><span class="d-code">${n.id}</span><span class="d-pill ${n.status}">${STAT_KR[n.status]}</span></div>
    <div class="d-label">${n.star?'★ ':''}${n.label}</div><div class="d-line">${n.line}</div>`;
  if(consumesList.length) html+=`<div class="d-sec"><div class="h">← 소비 (입력·자산)</div><ul>${consumesList.map(x=>`<li>${x}</li>`).join("")}</ul></div>`;
  if((n.produces||[]).length) html+=`<div class="d-sec"><div class="h">→ 산출</div><ul>${n.produces.map(x=>`<li>${x}</li>`).join("")}</ul></div>`;
  if(feedsList.length) html+=`<div class="d-sec"><div class="h">→ 공급 / 연계</div><ul>${feedsList.map(x=>`<li>${x}</li>`).join("")}</ul></div>`;
  html+=rmMiniHTML(n);
  html+=`<div class="d-meta"><span class="pg">${n.page}</span></div>`;
  $("#detail").innerHTML=html;
}
function showAsset(id){
  const a=assetById[id]; const nodes=new Set(a.consumers);
  applyHighlight(nodes,new Set([id]),id,'asset');
  const consumers=a.consumers.map(c=>nodeById[c]?nodeById[c].label:c);
  $("#detail").innerHTML=`<div><span class="d-code">${a.label}</span><span class="d-pill concept">공유 자산</span></div>
    <div class="d-line" style="margin-top:10px">${a.note}</div>
    <div class="d-sec"><div class="h">이 자산을 쓰는 업무 (${a.consumers.length})</div><ul>${consumers.map(x=>`<li>${x}</li>`).join("")}</ul></div>`;
}

/* ---- view toggle ---- */
function setView(v){
  $("#tab-roadmap").classList.toggle('active',v==='roadmap');
  $("#tab-map").classList.toggle('active',v==='map');
  $("#view-roadmap").classList.toggle('active',v==='roadmap');
  $("#view-map").classList.toggle('active',v==='map');
}
$("#tab-roadmap").onclick=()=>setView('roadmap');
$("#tab-map").onclick=()=>setView('map');

/* ---- clicks ---- */
document.addEventListener('click',e=>{
  const jump=e.target.closest('[data-jump]');
  if(jump){ setView('map'); showNode(jump.dataset.jump);
    $("#view-map").scrollIntoView({behavior:'smooth',block:'start'}); return; }
  const el=e.target.closest('[data-kind]');
  if(!el){
    if(!e.target.closest('.detail')&&$("#view-map").classList.contains('active')){ clearStates();
      $("#detail").innerHTML=`<div class="empty"><b>노드를 선택하세요.</b><br>그 업무가 <b>소비하는 자산(←)</b>·<b>공급/연계(→)</b>·<b>로드맵 항목</b>·앱 페이지가 표시됩니다. 자산 칩을 누르면 그걸 쓰는 업무가 켜집니다.</div>`; }
    return;
  }
  if(el.dataset.kind==='node') showNode(el.dataset.id); else showAsset(el.dataset.id);
});
</script>
</body>
</html>
'''


st.title("🗺️ 시스템 지도 · 로드맵")
st.caption("출처 = systemmap.json (KB). 노드/자산을 누르면 연결이, 로드맵 카드 칩을 누르면 해당 업무가 보입니다.")

_sm, _err = _fetch_systemmap()
if _err:
    st.error(_err)
    st.info("systemmap.json은 KB repo(work-automation-wb)에 있습니다. Streamlit secrets의 PAT가 그 private repo 읽기 권한을 가져야 합니다.")
    st.stop()

components.html(TEMPLATE.replace("__SYSTEMMAP_JSON__", _sm), height=1300, scrolling=True)

st.divider()
st.subheader("바로가기")
_LINKS = [
    ("pages/1_파일처리.py", "파일처리 (오픈마켓·온누리·발주·천년경영)", "📂"),
    ("pages/5_송장처리.py", "송장처리", "🏷️"),
    ("pages/3_대시보드.py", "대시보드", "📊"),
    ("pages/6_채널마진모니터.py", "채널마진모니터", "💹"),
    ("pages/7_업로드감시.py", "업로드감시", "📦"),
    ("pages/3_연동데이터관리/1_상품관리.py", "상품관리 (연동데이터)", "🔗"),
]
_cols = st.columns(2)
for _i, (_p, _label, _icon) in enumerate(_LINKS):
    with _cols[_i % 2]:
        try:
            st.page_link(_p, label=_label, icon=_icon)
        except Exception:
            st.markdown("%s %s" % (_icon, _label))
