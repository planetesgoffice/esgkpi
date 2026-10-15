// 永續 KPI 管理平台｜畫面與流程
const CFG=window.KPI_CONFIG||{};
const T={KPI:'tbl_KPI',REC:'tbl_Records',PPL:'tbl_People',YR:'tbl_Years',CODE:'tbl_Codes',DEPT:'tbl_Depts'};
const COLS={
  [T.KPI]:['系統識別碼','KPI_ID','年度','面向','部門','負責主管','主要執行人員','目標_ID','永續目標','專案_ID','必勝專案','SDGs','KPI類型','績效評核','達標值','單位','年度狀態','排序','最後修改人','最後修改時間'],
  [T.REC]:['紀錄_ID','KPI_ID','年度','月份','量化數據','執行說明','月狀態','備註','填報人','填報帳號','填報時間','前一版紀錄_ID'],
  [T.PPL]:['英文名','M365帳號','部門','角色','啟用'],
  [T.YR]:['年度','建立方式','建立人','建立時間'],
  [T.CODE]:['類別','代碼','排序'],
  [T.DEPT]:['部門代碼','部門名稱','啟用','改為']
};
const DEF={face:['永續治理','永續供應鏈','環境永續','業務行銷','綠色研發','資訊安全','永續人力','社會共融','永續辦公室'],mstat:['有進度','本月無進度（NA）','已完成','暫緩','公司策略性中止'],ktype:['量化型','質化型','里程碑型','完成型','評分型'],ystat:['執行中','已完成','未達標','已中止']};
const MCLS={'有進度':'d-prog','本月無進度（NA）':'d-na','已完成':'d-ok','暫緩':'d-warn','公司策略性中止':'d-bad'};
const MPILL={'有進度':'p-acc','本月無進度（NA）':'p-mute','已完成':'p-ok','暫緩':'p-warn','公司策略性中止':'p-bad'};
let kpis=[],recs=[],staff=[],years=[],depts=[],codes={...DEF},me={mail:'',name:null,role:'',display:''},lastSync=null;
const dirty=new Set();
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const split=s=>String(s??'').split(/[,，、;；\n]+/).map(x=>x.trim()).filter(Boolean);
const uuid=()=>(crypto.randomUUID?crypto.randomUUID():'x'+Date.now().toString(36)+Math.random().toString(36).slice(2));
const nowTW=()=>new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Taipei',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false}).format(new Date()).replace('T',' ');
const TODAY=nowTW();const CUR_Y=+TODAY.slice(0,4),CUR_M=+TODAY.slice(5,7);
const REPORT={y:CUR_M===1?CUR_Y-1:CUR_Y,m:CUR_M===1?12:CUR_M-1}; // 回報上個月
const S={year:REPORT.y,tab:'dash',face:'全部',dept:'',person:'',goal:'',status:'',q:'',fillMonth:REPORT.m,scope:'mine'};
try{const t=sessionStorage.getItem('kpiTab');if(t)S.tab=t;}catch(e){}
const repMonth=y=>y<REPORT.y?12:y>REPORT.y?0:REPORT.m; // 該年度目前應回報到幾月

// ---------- Excel 列 ⇄ 物件 ----------
function mapper(table,data){
  const need=COLS[table],idx={};
  const miss=need.filter(h=>{const i=data.headers.indexOf(h);idx[h]=i;return i<0;});
  if(miss.length)throw new Graph.GraphError('structure',`Excel 的 ${table} 缺少欄位：${miss.join('、')}。請勿更改表格欄位名稱。`);
  return data.rows.map((r,i)=>{const o={_idx:i};for(const h of need)o[h]=r[idx[h]]==null?'':String(r[idx[h]]).trim();return o;});
}
const toKpi=o=>({_idx:o._idx,uid:o['系統識別碼'],id:o['KPI_ID'],year:+o['年度'],face:o['面向'],dept:o['部門'],owner:o['負責主管'],execs:split(o['主要執行人員']),gid:o['目標_ID'],goal:o['永續目標'],pid:o['專案_ID'],proj:o['必勝專案'],sdgs:split(o['SDGs']).map(Number).filter(n=>n>0),type:o['KPI類型'],kpi:o['績效評核'],target:o['達標值'],unit:o['單位'],status:o['年度狀態']||'執行中',order:+o['排序']||0,mby:o['最後修改人'],mat:o['最後修改時間']});
const kpiRow=k=>[k.uid,k.id,String(k.year),k.face,k.dept,k.owner,k.execs.join(', '),k.gid,k.goal,k.pid,k.proj,k.sdgs.join(', '),k.type,k.kpi,k.target||'',k.unit||'',k.status,String(k.order||''),k.mby,k.mat];
const toRec=o=>({_idx:o._idx,rid:o['紀錄_ID'],kid:o['KPI_ID'],year:+o['年度'],m:+o['月份'],q:o['量化數據'],note:o['執行說明'],st:o['月狀態']||'有進度',rem:o['備註'],by:o['填報人'],mail:o['填報帳號'],at:o['填報時間'],prev:o['前一版紀錄_ID']});
const recRow=r=>[r.rid,r.kid,String(r.year),String(r.m),r.q,r.note,r.st,r.rem,r.by,r.mail,r.at,r.prev];
const toPerson=o=>({_idx:o._idx,name:o['英文名'],email:o['M365帳號'].toLowerCase(),dept:o['部門'],role:o['角色']||'一般',active:!/^(否|n|no|false|0)$/i.test(o['啟用'])});
const personRow=p=>[p.name,p.email,p.dept,p.role,p.active?'是':'否'];

async function readKpis(){return mapper(T.KPI,await Graph.readTable(T.KPI)).filter(o=>o['KPI_ID']).map(toKpi);}
async function readRecs(){return mapper(T.REC,await Graph.readTable(T.REC)).filter(o=>o['KPI_ID']&&o['月份']).map(toRec);}
async function readPeople(){return mapper(T.PPL,await Graph.readTable(T.PPL)).filter(o=>o['英文名']).map(toPerson);}
async function readYears(){return mapper(T.YR,await Graph.readTable(T.YR)).map(o=>+o['年度']).filter(Boolean);}
async function readCodes(){const c={face:[],mstat:[],ktype:[],ystat:[]},key={'面向':'face','月狀態':'mstat','KPI類型':'ktype','年度狀態':'ystat'};
  mapper(T.CODE,await Graph.readTable(T.CODE)).sort((a,b)=>(+a['排序'])-(+b['排序'])).forEach(o=>{const k=key[o['類別']];if(k&&o['代碼'])c[k].push(o['代碼']);});
  for(const k in c)if(!c[k].length)c[k]=DEF[k];return c;}
async function readDepts(){return mapper(T.DEPT,await Graph.readTable(T.DEPT)).filter(o=>o['部門代碼']).map(o=>({_idx:o._idx,code:o['部門代碼'],name:o['部門名稱'],active:!/^(否|n|no|false|0)$/i.test(o['啟用']),to:o['改為']}));}
const activeDepts=()=>depts.filter(d=>d.active).map(d=>d.code);
const deptLabel=c=>{const d=depts.find(x=>x.code===c);return d&&d.name?`${c}｜${d.name}`:c;};
// 舊部門依「改為」一路換到目前啟用的部門（例如 QSHE → QA）；找不到就保留原代碼
function mapDept(c){let cur=c;for(let i=0;i<5;i++){const d=depts.find(x=>x.code===cur);if(!d||d.active||!d.to)break;cur=d.to;}return cur;}
const deptInactive=c=>{const d=depts.find(x=>x.code===c);return !!c&&(!d||!d.active);};
function deptSelect(id,name,val){const list=activeDepts(),extra=val&&!list.includes(val)?[val]:[];
  return `<select id="${id}" name="${name}"><option value="">選擇部門</option>${extra.map(v=>`<option value="${esc(v)}" selected>${esc(v)}（已停用，請改選）</option>`).join('')}${list.map(c=>`<option value="${esc(c)}"${c===val?' selected':''}>${esc(deptLabel(c))}</option>`).join('')}</select>`;}
async function loadAll(){
  // 依序讀取（官方建議同一份活頁簿不要平行送出請求）
  codes=await readCodes();depts=await readDepts();staff=await readPeople();years=await readYears();kpis=await readKpis();recs=await readRecs();
  const p=staff.find(x=>x.email&&x.email===me.mail);
  me.name=p&&p.active?p.name:null;me.role=p&&p.active?p.role:'';me.inList=!!p;me.inactive=p&&!p.active;
  lastSync=nowTW();
}

// 寫入一律排隊，同一時間只送一個請求
let chain=Promise.resolve();
const enqueue=fn=>{const p=chain.then(fn);chain=p.catch(()=>{});return p;};
const TRANSIENT=['network','timeout','server','throttle'];

function errText(e){
  const k=e&&e.kind;
  if(k==='conflict'||k==='structure'||k==='validate'||k==='verify')return e.message;
  if(k==='auth')return '登入已過期，資料尚未寫入。請按「重新登入」，你填的內容會先幫你保留。';
  if(k==='forbidden')return '你的帳號沒有這份 Excel 的編輯權限，資料尚未寫入。請聯絡 Kiko 確認你是否在「年度永續指標管理」團隊中。';
  if(k==='notfound')return '找不到設定的 Excel 檔案或表格，資料尚未寫入。檔案可能被移動、改名或刪除，請聯絡 Kiko。';
  if(k==='locked')return 'Excel 檔案目前被鎖定，資料尚未寫入。可能有人用桌面版 Excel 開啟且沒有開啟自動儲存，請稍後再試。';
  if(k==='throttle')return 'Excel 目前忙碌中，資料尚未寫入。請等幾秒後再按一次儲存。';
  if(k==='server')return 'Microsoft 365 暫時無法回應，資料尚未寫入。請稍等一下再按一次儲存。';
  if(k==='timeout'||k==='network')return '無法連線到 Excel，資料尚未寫入。請確認網路後再按一次儲存。';
  return `Excel 回應錯誤（${esc(e&&(e.code||e.status)||'未知')}），資料尚未寫入。請稍後再試；持續發生請截圖給 Kiko。`;
}
const err=(kind,msg)=>new Graph.GraphError(kind,msg);

// ---------- 共用：以識別欄找到列、確認後覆寫 ----------
async function updateRow(table,keyCol,keyVal,fresh,rowValues){
  const target=fresh.find(o=>o[keyCol]===keyVal);
  if(!target)throw err('conflict','在 Excel 中找不到這筆資料，可能已被刪除。請按「重新整理」。');
  const rr=await Graph.rowRange(table,target._idx);
  const col=COLS[table].indexOf(keyCol);
  if(!rr||!rr.values||String(rr.values[0][col]).trim()!==keyVal)throw err('conflict','Excel 的列順序剛被調整過，為避免改錯列，這次沒有寫入。請按「重新整理」後再試。');
  await Graph.patchAddress(rr.address,[rowValues]);
}

// ---------- 年度與篩選 ----------
const yk=()=>kpis.filter(k=>k.year===S.year).sort((a,b)=>(a.order||9999)-(b.order||9999)||a.id.localeCompare(b.id));
const allYears=()=>[...new Set([...years,...kpis.map(k=>k.year)])].sort();
function latest(kid,m,list=recs){let r=null;for(const x of list)if(x.kid===kid&&x.m===m&&(!r||x.at>r.at||(x.at===r.at&&x._idx>r._idx)))r=x;return r;}
function history(kid,m){return recs.filter(x=>x.kid===kid&&x.m===m).sort((a,b)=>a.at<b.at?1:a.at>b.at?-1:b._idx-a._idx);}
function lastUpd(kid){let r=null;for(const x of recs)if(x.kid===kid&&(!r||x.at>r.at))r=x;return r;}
const yearPeople=()=>[...new Set(yk().flatMap(k=>[k.owner,...k.execs]).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
const activeNames=()=>staff.filter(p=>p.active).map(p=>p.name).sort((a,b)=>a.localeCompare(b));
const isMine=k=>!!me.name&&(k.execs.includes(me.name)||k.owner===me.name);
const isAdmin=()=>me.role==='管理者';
const statusPill=s=>`<span class="pill ${s==='已完成'?'p-ok':s==='未達標'?'p-bad':s==='已中止'?'p-mute':'p-warn'}">${esc(s||'執行中')}</span>`;
function fillSelect(el,opts,val,allLabel){el.innerHTML=(allLabel?`<option value="">${allLabel}</option>`:'')+opts.map(o=>{const [v,t]=Array.isArray(o)?o:[o,o];return `<option value="${esc(v)}"${String(v)===String(val)?' selected':''}>${esc(t)}</option>`;}).join('');}

// ---------- 畫面 ----------
function renderStats(){
  const ks=yk(),rm=repMonth(S.year),done=ks.filter(k=>k.status==='已完成').length,filled=rm?ks.filter(k=>latest(k.id,rm)).length:0;
  const cards=[[ks.length,'KPI 總數','acc'],[new Set(ks.map(k=>k.face)).size,'涵蓋面向',''],[ks.length-done,'未結案',''],[done,'已完成','ok'],
    [rm?filled:'—',rm?`${rm} 月已填報`:'尚未開始填報','ok'],[rm?ks.length-filled:'—',rm?`${rm} 月尚未填報`:'—','warn']];
  $('stats').innerHTML=cards.map(([n,l,c])=>`<div class="stat"><div class="n ${c}">${n}</div><div class="l">${l}</div></div>`).join('');
  $('heroTitle').textContent=`${S.year} 年度永續 KPI`;
}
function bars(rows,max){return rows.map(r=>`<div class="bar"><span class="k" title="${esc(r.k)}">${esc(r.k)}</span><span class="t">${r.parts.map(p=>`<i style="width:${max?p.v/max*100:0}%;background:var(${p.c})" title="${esc(p.n)} ${p.v}"></i>`).join('')}</span><span class="v">${r.parts.reduce((a,p)=>a+p.v,0)}</span></div>`).join('');}
function groupBars(keyFn){
  const m=new Map();for(const k of yk())for(const key of [].concat(keyFn(k)).filter(Boolean)){const o=m.get(key)||{done:0,doing:0};k.status==='已完成'?o.done++:o.doing++;m.set(key,o);}
  const rows=[...m].map(([k,o])=>({k,parts:[{n:'已完成',v:o.done,c:'--ok'},{n:'未結案',v:o.doing,c:'--sky'}]})).sort((a,b)=>(b.parts[0].v+b.parts[1].v)-(a.parts[0].v+a.parts[1].v));
  return bars(rows,Math.max(1,...rows.map(r=>r.parts[0].v+r.parts[1].v)));
}
function renderDash(){
  const ks=yk(),p=$('p-dash'),rm=repMonth(S.year);
  if(!ks.length){p.innerHTML=`<div class="card empty" style="margin-top:14px"><p>${S.year} 年度還沒有任何 KPI。</p><p class="small">請各負責人到這裡新增自己的 KPI；也可以請 Kiko 在「歷年資料」複製上一年度的設定。</p>${me.name?'<button class="btn pri" type="button" data-newkpi>＋ 新增 KPI</button>':''}</div>`;return;}
  const lg=`<div class="legend"><span style="--c:var(--ok)">已完成</span><span style="--c:var(--sky)">未結案</span></div>`;
  const ms=new Map(codes.mstat.map(s=>[s,0]));let none=0;if(rm)for(const k of ks){const r=latest(k.id,rm);r?ms.set(r.st,(ms.get(r.st)||0)+1):none++;}
  const col=s=>s==='已完成'?'--ok':s==='有進度'?'--sky':/中止/.test(s)?'--bad':s==='暫緩'?'--warn':'--line2';
  const msRows=[...ms].map(([s,v])=>({k:s,parts:[{n:s,v,c:col(s)}]})).concat([{k:'尚未填報',parts:[{n:'尚未填報',v:none,c:'--warn'}]}]);
  const tm=new Map();ks.forEach(k=>tm.set(k.type||'未設定',(tm.get(k.type||'未設定')||0)+1));
  const months=Array.from({length:12},(_,i)=>i+1).map(m=>({m,r:ks.filter(k=>latest(k.id,m)).length/ks.length,fut:m>rm}));
  p.innerHTML=`<div class="grid2" style="margin-top:14px">
   <div class="card"><h3>各面向 KPI 數</h3>${lg}<div class="bars">${groupBars(k=>k.face)}</div></div>
   <div class="card"><h3>每月填報率</h3><p class="sub">有填報紀錄的 KPI ÷ KPI 總數${rm?`；藕紫色為目前應回報的月份（${rm} 月）`:''}</p>
     <div class="months"><div class="mchart">${months.map(o=>`<div class="mcol"><span class="pct">${o.fut?'—':Math.round(o.r*100)+'%'}</span><span class="b ${o.m===rm?'cur':''}" style="height:${o.fut?0:o.r*120}px"></span></div>`).join('')}</div>
     <div class="mlab">${months.map(o=>`<span>${o.m}月</span>`).join('')}</div></div></div>
   <div class="card"><h3>各部門 KPI 數</h3>${lg}<div class="bars">${groupBars(k=>k.dept)}</div></div>
   <div class="card"><h3>${rm?rm+' 月執行概況':'執行概況'}</h3><p class="sub">${rm?'依各 KPI 該月最新一筆紀錄的月狀態':'這個年度還沒到填報月份'}</p>${rm?`<div class="bars">${bars(msRows,ks.length)}</div>`:''}
     <h3 style="padding-top:4px">KPI 類型</h3><div class="bars">${bars([...tm].sort((a,b)=>b[1]-a[1]).map(([k,v])=>({k,parts:[{n:k,v,c:'--ink3'}]})),ks.length)}</div></div>
   <div class="card" style="grid-column:1/-1"><h3>各人員 KPI 數（主要執行人員）</h3>${lg}<div class="bars" style="grid-template-columns:repeat(auto-fill,minmax(260px,1fr));column-gap:28px">${groupBars(k=>k.execs)}</div></div>
  </div>`;
}
function renderListFilters(){
  const ks=yk(),rm=repMonth(S.year);
  $('faceChips').innerHTML='<label>面向</label>'+['全部',...codes.face].map(f=>`<button class="chip" type="button" aria-pressed="${S.face===f}" data-face="${esc(f)}">${esc(f)}</button>`).join('');
  fillSelect($('fDept'),[...new Set(ks.map(k=>k.dept).filter(Boolean))].sort(),S.dept,'全部部門');
  fillSelect($('fPerson'),yearPeople(),S.person,'全部人員');
  fillSelect($('fGoal'),[...new Map(ks.map(k=>[k.gid,k.goal])).entries()].map(([id,g])=>[id,g.replace(/￭\s*/g,'').replace(/\n/g,'／').slice(0,40)]),S.goal,'全部目標');
  fillSelect($('fStatus'),[...codes.ystat.map(s=>[s,'年度：'+s]),...(rm?[['unfilled',`${rm} 月尚未填報`]]:[])],S.status,'全部狀態');
  $('newKpiBtn').hidden=!me.name;
}
function filtered(){
  const q=S.q.trim().toLowerCase(),rm=repMonth(S.year);
  return yk().filter(k=>(S.face==='全部'||k.face===S.face)&&(!S.dept||k.dept===S.dept)&&(!S.person||k.owner===S.person||k.execs.includes(S.person))&&(!S.goal||k.gid===S.goal)
   &&(!S.status||(S.status==='unfilled'?!latest(k.id,rm):k.status===S.status))
   &&(!q||[k.proj,k.kpi,k.goal,k.dept,k.owner,k.execs.join(' '),k.id,...recs.filter(r=>r.kid===k.id).map(r=>r.note)].join(' ').toLowerCase().includes(q)));
}
const dotsFor=k=>`<span class="dots" aria-label="1 到 12 月填報狀態">${Array.from({length:12},(_,i)=>{const r=latest(k.id,i+1);return `<i class="dot ${r?MCLS[r.st]||'d-prog':''}" title="${i+1}月：${r?esc(r.st):'未填'}"></i>`;}).join('')}</span>`;
function renderList(){
  const ks=filtered();let html='',g=null;
  for(const k of ks){
    if(k.gid!==g){g=k.gid;html+=`<tr class="goal-row"><td colspan="9"><span class="mono">${esc(k.gid)}</span>　永續目標：${esc(k.goal.replace(/\n/g,'　'))}</td></tr>`;}
    const lu=lastUpd(k.id);
    html+=`<tr class="row" data-kid="${esc(k.id)}" tabindex="0"><td><span class="tag">${esc(k.face)}</span></td><td>${esc(k.dept)}</td><td>${esc(k.owner)}</td><td>${k.execs.map(esc).join('<br>')}</td>
     <td><div style="font-weight:500">${esc(k.proj)}</div><div class="muted small clamp">${esc(k.kpi)}</div><div class="mono small muted">${esc(k.id)}</div></td>
     <td><span class="pill p-mute">${esc(k.type||'—')}</span></td><td>${statusPill(k.status)}</td><td>${dotsFor(k)}</td><td class="small muted mono" style="white-space:nowrap">${lu?esc(lu.at.slice(5,16)):'—'}</td></tr>`;
  }
  $('listBody').innerHTML=html||`<tr><td colspan="9" class="empty">沒有符合條件的 KPI。試著清除篩選或改用其他關鍵字。</td></tr>`;
  $('listCount').textContent=`顯示 ${ks.length} ／ ${yk().length} 筆 KPI。月份色塊：綠＝已完成、青＝有進度、灰＝本月無進度（NA）、藕紫＝暫緩、紅＝公司策略性中止、淡色＝未填。`;
}
function cardHtml(k,m){
  const r=latest(k.id,m),prev=m>1?latest(k.id,m-1):null,h=history(k.id,m),can=!!me.name;
  return `<article class="card fcard" data-kid="${esc(k.id)}" data-base="${r?esc(r.rid):''}">
   <div><div class="meta"><span class="tag">${esc(k.face)}</span><span class="tag">${esc(k.dept)}</span><span class="pill p-mute">${esc(k.type||'—')}</span>${r?`<span class="pill ${MPILL[r.st]||'p-acc'}">${m} 月已填</span>`:`<span class="pill p-warn">${m} 月未填</span>`}</div>
    <h4>${esc(k.proj)}</h4><div class="small muted clamp" title="${esc(k.kpi)}">${esc(k.kpi)}</div>
    <div class="small muted" style="margin-top:4px">主要執行：${k.execs.map(esc).join('、')}　<a href="#" data-open="${esc(k.id)}">查看詳情</a></div>
    ${prev?`<div class="prev"><b>${m-1} 月：</b>${esc(prev.note.slice(0,90))}${prev.note.length>90?'…':''}</div>`:''}</div>
   <form class="form" data-kid="${esc(k.id)}" novalidate>
    <label>量化數據（選填，可填 NA）<input id="q-${esc(k.id)}" name="q" value="${esc(r?.q||'')}" placeholder="例：18 筆、NA"${can?'':' disabled'}></label>
    <label>月狀態<select id="s-${esc(k.id)}" name="st"${can?'':' disabled'}>${codes.mstat.map(s=>`<option${(r?r.st:'有進度')===s?' selected':''}>${esc(s)}</option>`).join('')}</select></label>
    <label>備註（選填）<input id="r-${esc(k.id)}" name="rem" value="${esc(r?.rem||'')}"${can?'':' disabled'}></label>
    <label class="full">執行狀況／質化說明<textarea id="n-${esc(k.id)}" name="note" placeholder="本月做了什麼、進度到哪裡"${can?'':' disabled'}>${esc(r?.note||'')}</textarea></label>
    <div class="full fact"><span class="pill p-warn" data-dirty hidden>尚未儲存</span><span class="msg" hidden></span>${h.length?`<span class="small muted">最後修改：${esc(h[0].by)}・${esc(h[0].at.slice(0,16))}${h.length>1?`（共 ${h.length} 版）`:''}</span>`:''}${can?'<button class="btn pri" type="submit">儲存</button>':''}</div>
   </form></article>`;
}
function renderFill(){
  fillSelect($('fillMonth'),Array.from({length:12},(_,i)=>[i+1,`${S.year} 年 ${i+1} 月${S.year===REPORT.y&&i+1===REPORT.m?'（本期）':''}`]),S.fillMonth);
  $('scopeChips').innerHTML='<label>範圍</label>'+[['mine','我負責的'],['all','全部 KPI']].map(([v,t])=>`<button class="chip" type="button" aria-pressed="${S.scope===v}" data-scope="${v}">${t}</button>`).join('');
  const m=S.fillMonth,ks=yk().filter(k=>S.scope==='all'||isMine(k));
  $('fillSummary').textContent=`${ks.length} 項 KPI，已填 ${ks.filter(k=>latest(k.id,m)).length} 項`;
  dirty.clear();
  if(!ks.length){$('fillList').innerHTML=`<div class="card empty">${me.name?`你在 ${S.year} 年沒有負責的 KPI。可切換成「全部 KPI」查看或代填。`:'你的帳號還不在人員名單中，可以查看但不能填報。'}</div>`;return;}
  $('fillList').innerHTML=ks.map(k=>cardHtml(k,m)).join('');
  restoreDrafts();
}
function showMsg(form,cls,html){const msg=form.querySelector('.msg');msg.hidden=false;msg.className='msg '+cls;msg.innerHTML=html;}
async function saveRecord(form){
  const kid=form.dataset.kid,card=form.closest('.fcard'),btn=form.querySelector('button[type=submit]'),m=S.fillMonth,yr=S.year;
  const d=Object.fromEntries(new FormData(form));
  if(!me.name)return showMsg(form,'bad','你的帳號不在人員名單中，無法填報。請聯絡 Kiko。');
  if(d.st!=='本月無進度（NA）'&&!d.note.trim())return showMsg(form,'bad','請填寫執行狀況／質化說明。本月沒有進度時，請把月狀態改成「本月無進度（NA）」。');
  btn.disabled=true;btn.textContent='寫入 Excel 中…';
  const rec={rid:'R-'+uuid(),kid,year:yr,m,q:d.q.trim(),note:d.note.trim(),st:d.st,rem:d.rem.trim(),by:me.name,mail:me.mail,at:'',prev:card.dataset.base};
  try{
    await enqueue(async()=>{
      const fresh=await readRecs();                       // 1. 寫入前重新讀取，確認沒有別人先改
      const cur=latest(kid,m,fresh);
      if((cur?cur.rid:'')!==card.dataset.base){recs=fresh;throw err('conflict',`這個月份在你開啟後已被 ${esc(cur.by)} 於 ${esc(cur.at.slice(0,16))} 更新，這次沒有寫入。對方的內容：「${esc(cur.note.slice(0,80))}」。請複製你的內容後按「重新整理」，再決定是否修改。`);}
      rec.at=nowTW();
      try{await Graph.addRows(T.REC,[recRow(rec)]);}      // 2. 新增一列（不覆蓋舊資料）
      catch(e){
        if(!TRANSIENT.includes(e.kind))throw e;
        const again=await readRecs();                      // 逾時等情況：先確認是否其實已寫入，避免重複
        if(!again.some(r=>r.rid===rec.rid))throw e;
      }
      const after=await readRecs();                        // 3. 重新讀回確認
      if(!after.some(r=>r.rid===rec.rid))throw err('verify','寫入後在 Excel 中找不到這筆紀錄，可能沒有存成功。請按「重新整理」確認，必要時再存一次。');
      recs=after;lastSync=nowTW();
    });
    dirty.delete(kid);clearDraft(kid);
    const k=kpis.find(x=>x.id===kid);card.outerHTML=cardHtml(k,m);
    const nf=document.querySelector(`.fcard[data-kid="${CSS.escape(kid)}"] .form`);
    showMsg(nf,'ok',`已儲存到 Excel（${rec.at.slice(11,16)}）`);
    renderStats();setSync();
  }catch(e){
    btn.disabled=false;btn.textContent='儲存';
    showMsg(form,'bad',errText(e)+(e.kind==='auth'?' <button class="btn" type="button" data-relogin>重新登入</button>':''));
  }
}
// 未儲存內容保護：只在登入過期需要重新登入時，暫存在這個分頁（不是正式資料）
const DKEY='kpiDrafts';
function saveDrafts(){const o={};document.querySelectorAll('.fcard .form').forEach(f=>{if(dirty.has(f.dataset.kid))o[f.dataset.kid+'|'+S.year+'|'+S.fillMonth]=Object.fromEntries(new FormData(f));});try{sessionStorage.setItem(DKEY,JSON.stringify(o));}catch(e){}}
function clearDraft(kid){try{const o=JSON.parse(sessionStorage.getItem(DKEY)||'{}');delete o[kid+'|'+S.year+'|'+S.fillMonth];sessionStorage.setItem(DKEY,JSON.stringify(o));}catch(e){}}
function restoreDrafts(){let o={};try{o=JSON.parse(sessionStorage.getItem(DKEY)||'{}');}catch(e){}
  document.querySelectorAll('.fcard .form').forEach(f=>{const d=o[f.dataset.kid+'|'+S.year+'|'+S.fillMonth];if(!d)return;
    for(const [n,v] of Object.entries(d)){const el=f.elements[n];if(el)el.value=v;}dirty.add(f.dataset.kid);f.querySelector('[data-dirty]').hidden=false;showMsg(f,'info','已還原你上次尚未儲存的內容，請確認後按儲存。');});}

function openDrawer(kid){
  const k=kpis.find(x=>x.id===kid);if(!k)return;const rm=repMonth(k.year);
  const items=Array.from({length:12},(_,i)=>{const m=i+1,h=history(kid,m),r=h[0];
    return `<li><span class="mo">${m} 月</span><span class="ax"><i class="${r?MCLS[r.st]||'':''}"></i></span><div class="body">${r?`<span class="pill ${MPILL[r.st]||'p-acc'}">${esc(r.st)}</span>${r.q?` <span class="mono small">${esc(r.q)}</span>`:''}\n${esc(r.note)}${r.rem?`\n<span class="muted small">備註：${esc(r.rem)}</span>`:''}<div class="hist">${esc(r.by)}・${esc(r.at.slice(0,16))}${h.length>1?`・共 ${h.length} 版`:''}</div>`:`<span class="muted small">${m>rm?'尚未到期':'未填報'}</span>`}</div></li>`;}).join('');
  $('drawer').innerHTML=`<button class="x" type="button" aria-label="關閉" data-close>×</button>
   <div class="meta" style="display:flex;gap:6px;flex-wrap:wrap"><span class="tag">${esc(k.face)}</span><span class="tag">${esc(k.dept)}</span>${statusPill(k.status)}<span class="pill p-mute">${esc(k.type||'—')}</span></div>
   <h3>${esc(k.proj)}</h3><div class="mono small muted">${esc(k.id)}・${esc(k.gid)}・${esc(k.pid)}　${me.name?`<button class="btn" type="button" data-editkpi="${esc(k.id)}" style="font-family:var(--f-body)">編輯 KPI 設定</button>`:''}</div>
   <dl class="kv"><dt>永續目標</dt><dd>${esc(k.goal)}</dd><dt>績效評核</dt><dd>${esc(k.kpi)}</dd>${k.target?`<dt>達標值</dt><dd>${esc(k.target)} ${esc(k.unit)}</dd>`:''}<dt>負責主管</dt><dd>${esc(k.owner)||'—'}</dd><dt>主要執行人員</dt><dd>${k.execs.map(esc).join('、')}</dd><dt>對應 SDGs</dt><dd>${k.sdgs.length?k.sdgs.map(n=>`<span class="tag">SDG ${n}</span>`).join(' '):'—'}</dd><dt>最後修改</dt><dd class="small">${esc(k.mby)}・${esc(k.mat.slice(0,16))}</dd></dl>
   <div style="display:flex;justify-content:space-between;align-items:center;margin:6px 0 10px"><b>每月執行紀錄</b>${me.name&&rm?`<button class="btn" type="button" data-gofill="${esc(k.id)}">填寫 ${rm} 月</button>`:''}</div>
   <ol class="tl">${items}</ol>`;
  $('drawer').hidden=false;$('scrim').hidden=false;$('drawer').querySelector('[data-close]').focus();
}
function closeDrawer(){$('drawer').hidden=true;$('scrim').hidden=true;}

function openKpiForm(kid){
  const k=kid?kpis.find(x=>x.id===kid):null,yr=k?k.year:S.year,ks=kpis.filter(x=>x.year===yr);
  const goals=[...new Map(ks.map(x=>[x.gid,x])).values()];
  const opt=(v,t,sel)=>`<option value="${esc(v)}"${sel?' selected':''}>${esc(t)}</option>`;
  const names=[...new Set([...activeNames(),...(k?[k.owner,...k.execs].filter(Boolean):[])])].sort((a,b)=>a.localeCompare(b));
  $('drawer').innerHTML=`<button class="x" type="button" aria-label="關閉" data-close>×</button>
   <div class="small muted">${yr} 年度・由 KPI 負責人自行${k?'修改':'新增'}</div><h3>${k?'編輯 KPI 設定':'新增 KPI'}</h3>
   <form class="form" id="kpiForm" data-kid="${k?esc(k.id):''}" data-year="${yr}" novalidate style="grid-template-columns:repeat(2,minmax(0,1fr));margin-top:12px">
    <label>面向<select id="kfFace" name="face">${codes.face.map(f=>opt(f,f,k?k.face===f:false)).join('')}</select></label>
    <label>部門${deptSelect('kfDept','dept',k?k.dept:(staff.find(p=>p.name===me.name)?.dept&&!deptInactive(staff.find(p=>p.name===me.name).dept)?staff.find(p=>p.name===me.name).dept:''))}</label>
    <label class="full">永續目標<select id="kfGoal" name="gid"><option value="__new">＋ 新增永續目標</option>${goals.map(g=>opt(g.gid,g.goal.replace(/￭\s*/g,'').replace(/\n/g,'／').slice(0,60),k&&k.gid===g.gid)).join('')}</select></label>
    <label class="full" id="kfGoalNewWrap">新的永續目標<textarea id="kfGoalNew" name="goalNew" style="min-height:48px" placeholder="一個目標格可以寫多項，例如：￭ 深化永續治理文化 ￭ 精進永續資訊揭露"></textarea></label>
    <label class="full">必勝專案（工作項目）<input id="kfProj" name="proj" value="${esc(k?.proj||'')}" list="projList" placeholder="輸入新專案，或選擇同一目標下已有的專案"></label>
    <datalist id="projList">${[...new Set(ks.map(x=>x.proj))].map(p=>`<option value="${esc(p)}">`).join('')}</datalist>
    <label>負責主管<select id="kfOwner" name="owner"><option value="">選擇人員</option>${names.map(n=>opt(n,n,k&&k.owner===n)).join('')}</select></label>
    <label>KPI 類型<select id="kfType" name="type">${codes.ktype.map(t=>opt(t,t,k&&k.type===t)).join('')}</select></label>
    <fieldset class="full" style="border:1px solid var(--line2);border-radius:8px;padding:8px 10px;margin:0"><legend class="small muted">主要執行人員（可複選）</legend>
     <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:2px 8px;font-size:13px">${names.map(n=>`<label style="display:flex;gap:6px;align-items:center;color:var(--ink)"><input type="checkbox" name="execs" value="${esc(n)}"${(k?k.execs.includes(n):n===me.name)?' checked':''} style="width:auto">${esc(n)}</label>`).join('')}</div></fieldset>
    <label class="full">績效評核（KPI）<textarea id="kfKpi" name="kpi" placeholder="例：完成 100 家供應商 ESG 評估">${esc(k?.kpi||'')}</textarea></label>
    <label>達標值（選填）<input id="kfTarget" name="target" value="${esc(k?.target||'')}" placeholder="例：100、定稿、完成"></label>
    <label>單位（選填）<input id="kfUnit" name="unit" value="${esc(k?.unit||'')}" placeholder="例：家、%"></label>
    <label>對應 SDGs（數字，以逗號分隔）<input id="kfSdg" name="sdgs" value="${k?k.sdgs.join(', '):''}" placeholder="例：12, 13"></label>
    <label>年度狀態<select id="kfStatus" name="status">${codes.ystat.map(t=>opt(t,t,(k?.status||'執行中')===t)).join('')}</select></label>
    <div class="full fact"><span class="msg" id="kfMsg" hidden></span><button class="btn" type="button" data-close>取消</button><button class="btn pri" type="submit">${k?'儲存變更':'新增 KPI'}</button></div>
   </form>`;
  const sync=()=>{$('kfGoalNewWrap').hidden=$('kfGoal').value!=='__new';};$('kfGoal').addEventListener('change',sync);sync();
  $('drawer').hidden=false;$('scrim').hidden=false;$('kfDept').focus();
}
async function saveKpi(f){
  const fd=new FormData(f),d=Object.fromEntries(fd),execs=fd.getAll('execs'),yr=+f.dataset.year,m=$('kfMsg'),btn=f.querySelector('button[type=submit]');
  const say=(c,t)=>{m.hidden=false;m.className='msg '+c;m.innerHTML=t;};
  if(d.gid==='__new'&&!d.goalNew.trim())return say('bad','請填寫新的永續目標，或從清單選擇既有目標。');
  if(!d.dept)return say('bad','請選擇部門。');
  if(deptInactive(d.dept))return say('bad',`部門「${esc(d.dept)}」已停用，請改選目前的部門。`);
  if(!d.proj.trim())return say('bad','請填寫必勝專案。');
  if(!execs.length)return say('bad','請至少勾選一位主要執行人員。');
  if(!d.kpi.trim())return say('bad','請填寫績效評核（KPI）內容。');
  btn.disabled=true;btn.textContent='寫入 Excel 中…';
  let savedId=f.dataset.kid;
  try{
    await enqueue(async()=>{
      const freshRows=mapper(T.KPI,await Graph.readTable(T.KPI)),fresh=freshRows.filter(o=>o['KPI_ID']).map(toKpi),ks=fresh.filter(x=>x.year===yr);
      const nextNum=(list,re)=>Math.max(0,...list.map(s=>{const x=re.exec(s);return x?+x[1]:0;}))+1;
      let gid=d.gid,goal;
      if(gid==='__new'){gid=`${yr}-G${String(nextNum(ks.map(x=>x.gid),/G(\d+)$/)).padStart(2,'0')}`;goal=d.goalNew.trim();}
      else{const g=ks.find(x=>x.gid===gid);if(!g)throw err('conflict','選擇的永續目標已不存在，請重新整理。');goal=g.goal;}
      const same=ks.find(x=>x.gid===gid&&x.proj===d.proj.trim());
      const pid=same?same.pid:`${yr}-P${String(nextNum(ks.map(x=>x.pid),/P(\d+)$/)).padStart(2,'0')}`;
      const obj={face:d.face,dept:d.dept.trim(),owner:d.owner,execs,gid,goal,pid,proj:d.proj.trim(),type:d.type,kpi:d.kpi.trim(),target:d.target.trim(),unit:d.unit.trim(),sdgs:(d.sdgs.match(/\d+/g)||[]).map(Number).filter(n=>n>=1&&n<=17),status:d.status,mby:me.name,mat:nowTW()};
      if(savedId){
        const old=fresh.find(x=>x.id===savedId);if(!old)throw err('conflict','在 Excel 中找不到這個 KPI，可能已被刪除。');
        const k={...old,...obj};
        await updateRow(T.KPI,'系統識別碼',old.uid,freshRows,kpiRow(k));
      }else{
        savedId=`${yr}-K${String(nextNum(ks.map(x=>x.id),/K(\d+)$/)).padStart(3,'0')}`;
        const k={uid:uuid(),id:savedId,year:yr,order:ks.length+1,...obj};
        try{await Graph.addRows(T.KPI,[kpiRow(k)]);}
        catch(e){if(!TRANSIENT.includes(e.kind))throw e;const again=await readKpis();if(!again.some(x=>x.uid===k.uid))throw e;}
        let after=await readKpis();
        // 兩人同時新增而拿到同一個編號時，後寫入的一方自動改用下一個編號
        const dup=after.filter(x=>x.id===savedId);
        if(dup.length>1&&dup[0].uid!==k.uid){
          savedId=`${yr}-K${String(nextNum(after.filter(x=>x.year===yr).map(x=>x.id),/K(\d+)$/)).padStart(3,'0')}`;
          await updateRow(T.KPI,'系統識別碼',k.uid,mapper(T.KPI,await Graph.readTable(T.KPI)),kpiRow({...k,id:savedId}));
        }
      }
      kpis=await readKpis();
      if(!kpis.some(x=>x.id===savedId))throw err('verify','寫入後在 Excel 中找不到這個 KPI，請按「重新整理」確認。');
      if(!years.includes(yr)){await Graph.addRows(T.YR,[[String(yr),'新增KPI時建立',me.name,nowTW()]]);years=await readYears();}
      lastSync=nowTW();
    });
    S.year=yr;renderTop();render();openDrawer(savedId);
    $('drawer').querySelector('h3').insertAdjacentHTML('beforebegin',`<p class="msg ok" style="margin:8px 0">KPI ${esc(savedId)} 已${f.dataset.kid?'更新':'新增'}並儲存到 Excel。</p>`);
  }catch(e){btn.disabled=false;btn.textContent=f.dataset.kid?'儲存變更':'新增 KPI';say('bad',errText(e));}
}

function renderHist(){
  const ys=allYears(),next=Math.max(CUR_Y,...ys)+ (ys.includes(Math.max(CUR_Y,...ys))?1:0),prev=Math.max(...ys.filter(y=>y<next));
  const card=y=>{const ks=kpis.filter(k=>k.year===y),ids=new Set(ks.map(k=>k.id));return `<div class="card ycard"><div class="y">${y}</div><p class="small">${ks.length} 項 KPI・${recs.filter(r=>ids.has(r.kid)).length} 筆月紀錄</p><button class="btn" type="button" data-year="${y}">切換到 ${y}</button></div>`;};
  $('p-hist').innerHTML=`<div class="years" style="margin-top:14px">${ys.map(card).join('')}
   <div class="card ycard"><div class="y">${next}</div>${isAdmin()?`<p class="small muted">尚未建立。空白年度：大家從頭新增各自的 KPI。複製：只帶入 ${prev} 年的 KPI 設定並產生新編號，每月紀錄不會複製，${prev} 年資料也不受影響。</p>
     <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="btn pri" type="button" data-mkyear="${next}">建立空白 ${next} 年度</button>${Number.isFinite(prev)?`<button class="btn" type="button" data-cpyear="${next}" data-from="${prev}">複製 ${prev} 主資料建立 ${next}</button>`:''}</div><span class="msg" id="yrMsg" hidden></span>`
     :'<p class="small muted">尚未建立。新年度由 Kiko 或 Jessie 建立。</p>'}</div></div>`;
}
async function makeYear(y,from){
  const m=$('yrMsg'),say=(c,t)=>{m.hidden=false;m.className='msg '+c;m.innerHTML=t;};
  document.querySelectorAll('[data-mkyear],[data-cpyear]').forEach(b=>b.disabled=true);say('info','寫入 Excel 中…');
  try{await enqueue(async()=>{
    const ys=await readYears();if(ys.includes(y))throw err('conflict',`${y} 年度已經建立過了，請重新整理。`);
    if(from){
      const src=(await readKpis()).filter(k=>k.year===from);
      if((await readKpis()).some(k=>k.year===y))throw err('conflict',`${y} 年度已經有 KPI，為避免重複，這次沒有複製。`);
      const t=nowTW(),rows=src.map((k,i)=>kpiRow({...k,dept:mapDept(k.dept),uid:uuid(),id:`${y}-K${String(i+1).padStart(3,'0')}`,year:y,gid:k.gid.replace(/^\d{4}-/,y+'-'),pid:k.pid.replace(/^\d{4}-/,y+'-'),status:'執行中',order:i+1,mby:me.name,mat:t}));
      if(rows.length)await Graph.addRows(T.KPI,rows);
    }
    await Graph.addRows(T.YR,[[String(y),from?`複製${from}`:'空白',me.name,nowTW()]]);
    years=await readYears();kpis=await readKpis();
    if(!years.includes(y))throw err('verify','寫入後在 Excel 中找不到新年度，請重新整理確認。');
  });S.year=y;renderTop();setTab('dash');
  const bad=kpis.filter(k=>k.year===y&&deptInactive(k.dept));
  if(bad.length)alertBar(`${y} 年度已建立。有 ${bad.length} 項 KPI 的部門（${[...new Set(bad.map(k=>k.dept))].map(esc).join('、')}）已停用且沒有設定「改為」，請負責人打開 KPI 改選部門。`);}
  catch(e){document.querySelectorAll('[data-mkyear],[data-cpyear]').forEach(b=>b.disabled=false);say('bad',errText(e));}
}

function renderAdmin(){
  const adm=isAdmin(),tbl=(n,d,cols)=>`<div class="card" style="padding:14px 16px"><b class="mono">${n}</b><p class="small muted" style="margin:2px 0 8px">${d}</p><code>${cols.join('・')}</code></div>`;
  const rows=staff.slice().sort((a,b)=>(b.active-a.active)||a.name.localeCompare(b.name)).map(p=>{const n=kpis.filter(k=>k.year===S.year&&(k.owner===p.name||k.execs.includes(p.name))).length;
    return `<tr><td>${esc(p.name)}</td><td class="mono small">${p.email?esc(p.email):'<span class="pill p-warn">待補</span>'}</td><td>${esc(p.dept)}${deptInactive(p.dept)?' <span class="pill p-warn">部門已停用</span>':''}</td><td>${esc(p.role)}</td><td class="mono">${n}</td><td>${p.active?'<span class="pill p-ok">在職</span>':'<span class="pill p-mute">已停用</span>'}</td><td>${adm?`<button class="btn" type="button" data-toggle="${esc(p.name)}">${p.active?'停用':'重新啟用'}</button>`:''}</td></tr>`;}).join('');
  $('p-admin').innerHTML=`<div class="card" style="padding:16px;margin-top:14px"><b>正式資料來源</b>
   <dl class="kv" style="margin-bottom:0"><dt>檔案</dt><dd>${esc(Graph.file?.name||'')}　${Graph.file?.webUrl?`<a href="${esc(Graph.file.webUrl)}" target="_blank" rel="noopener">在 Excel 中開啟</a>`:''}</dd><dt>讀寫方式</dt><dd>以你的 Microsoft 365 帳號透過 Microsoft Graph 直接讀寫這份 Excel；網站本身不保存資料</dd><dt>最後同步</dt><dd>${esc(lastSync||'—')}</dd></dl></div>
   <h3 style="font-size:14px;margin:22px 0 4px">人員名單</h3>
   <p class="note" style="margin:0 0 10px">由 Kiko 建檔維護，Jessie 為備援管理者。新同仁加入就新增一筆；同仁離職改成「停用」，不要刪除。M365 帳號必須是公司信箱，系統才認得登入的人。</p>
   ${adm?'':`<p class="msg info" style="margin:0 0 10px">人員名單與 KPI 交接由管理者 Kiko 維護（備援：Jessie）。需要異動請聯絡 Kiko 或 Jessie。</p>`}
   <form class="card form" id="addPerson"${adm?'':' hidden'} style="padding:14px 16px;grid-template-columns:repeat(4,minmax(0,1fr)) auto;align-items:end" novalidate>
    <label>英文名<input id="apName" name="name" required></label><label>M365 帳號<input id="apMail" name="email" type="email" placeholder="name@planet.com.tw"></label>
    <label>部門${deptSelect('apDept','dept','')}</label><label>角色<select id="apRole" name="role"><option>一般</option><option>管理者</option></select></label>
    <button class="btn pri" type="submit">新增人員</button><span class="msg full" id="apMsg" hidden></span></form>
   <div class="card tbl-wrap" style="margin-top:10px"><table><thead><tr><th>英文名</th><th>M365 帳號</th><th>部門</th><th>角色</th><th>${S.year} KPI 數</th><th>狀態</th><th></th></tr></thead><tbody>${rows}</tbody></table></div>
   <span class="msg" id="ppMsg" hidden></span>
   ${staff.some(p=>p.active&&deptInactive(p.dept))?`<p class="note">有 ${staff.filter(p=>p.active&&deptInactive(p.dept)).length} 位在職人員的部門已停用，請 Kiko 在 Excel「人員」工作表更新為現在的部門。</p>`:''}
   <h3 style="font-size:14px;margin:22px 0 4px">部門</h3>
   <p class="note" style="margin:0 0 10px">由 Kiko 在 Excel「部門」工作表維護。組織調整時，舊部門改為停用並填寫「改為」；過去年度的 KPI 保留當時的部門名稱。</p>
   <div class="card" style="padding:12px 16px;display:flex;flex-wrap:wrap;gap:6px">${depts.filter(d=>d.active).map(d=>`<span class="tag">${esc(deptLabel(d.code))}</span>`).join('')}${depts.filter(d=>!d.active).map(d=>`<span class="tag" style="opacity:.6">${esc(d.code)}（已停用${d.to?' → '+esc(d.to):''}）</span>`).join('')}</div>
   <div class="card" style="padding:14px 16px;margin-top:10px"${adm?'':' hidden'}><b>交接：轉移負責的 KPI</b><p class="note" style="margin:2px 0 10px">同仁離職或調動時，把他在某一年度負責的 KPI 一次轉給接手的人。過去的填報紀錄不會改動。</p>
    <div class="filters" style="padding:0"><div class="fg"><label for="mvYear">年度</label><select id="mvYear" class="ctl">${allYears().map(y=>`<option${y===S.year?' selected':''}>${y}</option>`).join('')}</select></div>
    <div class="fg"><label for="mvFrom">從</label><select id="mvFrom" class="ctl"><option value="">選擇人員</option>${staff.map(p=>`<option>${esc(p.name)}</option>`).join('')}</select></div>
    <div class="fg"><label for="mvTo">轉給</label><select id="mvTo" class="ctl"><option value="">選擇人員</option>${activeNames().map(n=>`<option>${esc(n)}</option>`).join('')}</select></div>
    <button class="btn pri" type="button" id="moveBtn">轉移</button><span class="msg" id="mvMsg" hidden></span></div></div>
   <h3 style="font-size:14px;margin:22px 0 8px">Excel 內的資料表</h3>
   <div class="schema">${tbl('tbl_KPI','KPI 主資料',COLS[T.KPI])}${tbl('tbl_Records','每月紀錄，只新增不覆蓋',COLS[T.REC])}${tbl('tbl_People','人員名單',COLS[T.PPL])}${tbl('tbl_Depts','部門對照',COLS[T.DEPT])}${tbl('tbl_Years ／ tbl_Codes','年度、下拉選單選項',[...COLS[T.YR],'｜',...COLS[T.CODE]])}</div>`;
}
async function addPerson(f){
  const d=Object.fromEntries(new FormData(f)),m=$('apMsg'),say=(c,t)=>{m.hidden=false;m.className='msg full '+c;m.innerHTML=t;};
  const name=d.name.trim(),mail=d.email.trim().toLowerCase();
  if(!name)return say('bad','請輸入英文名。');
  if(mail&&!/^[^@\s]+@[^@\s]+$/.test(mail))return say('bad','M365 帳號格式不正確，請填公司信箱。');
  try{await enqueue(async()=>{
    const fresh=await readPeople();
    if(fresh.some(p=>p.name.toLowerCase()===name.toLowerCase()))throw err('validate',`名單裡已經有「${esc(name)}」。同名不同人請加上姓氏縮寫，例如 Aaron.W。`);
    if(mail&&fresh.some(p=>p.email===mail))throw err('validate','這個 M365 帳號已經在名單中。');
    await Graph.addRows(T.PPL,[personRow({name,email:mail,dept:d.dept.trim(),role:d.role,active:true})]);
    staff=await readPeople();if(!staff.some(p=>p.name===name))throw err('verify','寫入後在 Excel 中找不到這位人員，請重新整理確認。');
  });renderTop();renderAdmin();}catch(e){say('bad',errText(e));}
}
async function togglePerson(name){
  const m=$('ppMsg');
  try{await enqueue(async()=>{const rows=mapper(T.PPL,await Graph.readTable(T.PPL)),p=rows.map(toPerson).find(x=>x.name===name);if(!p)throw err('conflict','找不到這位人員，請重新整理。');
    await updateRow(T.PPL,'英文名',name,rows,personRow({...p,active:!p.active}));staff=await readPeople();});renderTop();renderAdmin();}
  catch(e){m.hidden=false;m.className='msg bad';m.innerHTML=errText(e);}
}
async function moveKpis(){
  const from=$('mvFrom').value,to=$('mvTo').value,yr=+$('mvYear').value,out=$('mvMsg'),say=(c,t)=>{out.hidden=false;out.className='msg '+c;out.innerHTML=t;};
  if(!from||!to||from===to)return say('bad','請選擇不同的兩位人員。');
  say('info','寫入 Excel 中…');let done=0,total=0;
  try{await enqueue(async()=>{
    const rows=mapper(T.KPI,await Graph.readTable(T.KPI)),targets=rows.map(toKpi).filter(k=>k.year===yr&&(k.owner===from||k.execs.includes(from)));total=targets.length;
    for(const k of targets){const n={...k,owner:k.owner===from?to:k.owner,execs:[...new Set(k.execs.map(x=>x===from?to:x))],mby:me.name,mat:nowTW()};
      await updateRow(T.KPI,'系統識別碼',k.uid,rows,kpiRow(n));done++;}
    kpis=await readKpis();});
    say('ok',`${yr} 年共 ${done} 項 KPI 已由 ${esc(from)} 轉給 ${esc(to)}，並儲存到 Excel。`);}
  catch(e){kpis=await readKpis().catch(()=>kpis);say('bad',`已完成 ${done} ／ ${total} 項。${errText(e)}`);}
}

// ---------- 外框 ----------
function setSync(){$('syncTxt').textContent=lastSync?`已同步 Excel：${lastSync.slice(11,16)}`:'';}
function renderTop(){
  fillSelect($('yearSel'),allYears().map(y=>[y,y+' 年']),S.year);
  $('userChip').textContent=me.name?`${me.name}${isAdmin()?'（管理者）':''}`:(me.display||me.mail);
  const w=$('acctWarn');
  if(!me.name){w.hidden=false;w.innerHTML=me.inactive?`你的帳號（${esc(me.mail)}）在人員名單中已停用，目前只能查看。`:`你的帳號（${esc(me.mail)}）還不在人員名單中，目前只能查看，不能填報或新增 KPI。請聯絡 Kiko 把你加入名單。`;}else w.hidden=true;
  setSync();
}
function setTab(t){S.tab=t;try{sessionStorage.setItem('kpiTab',t);}catch(e){}
  document.querySelectorAll('.tab').forEach(b=>b.setAttribute('aria-selected',b.dataset.tab===t));
  ['dash','list','fill','hist','admin'].forEach(x=>$('p-'+x).hidden=x!==t);render();}
function render(){renderStats();
  if(S.tab==='dash')renderDash();if(S.tab==='list'){renderListFilters();renderList();}if(S.tab==='fill')renderFill();if(S.tab==='hist')renderHist();if(S.tab==='admin')renderAdmin();}
function fatal(html){$('boot').hidden=false;$('app').hidden=true;$('boot').innerHTML=`<div class="card" style="padding:22px;max-width:640px;margin:40px auto">${html}</div>`;}
async function refresh(){
  if(dirty.size&&!confirm('有尚未儲存的填報內容，重新整理會清除。確定要重新整理嗎？'))return;
  $('syncTxt').textContent='讀取 Excel 中…';
  try{await enqueue(async()=>{await Graph.createSession();await loadAll();});renderTop();render();}
  catch(e){$('syncTxt').textContent='讀取失敗';alertBar(errText(e).replace('資料尚未寫入。',''));}
}
function alertBar(t){const w=$('acctWarn');w.hidden=false;w.innerHTML=t;}

document.addEventListener('click',e=>{const t=e.target.closest('button,a,tr.row');if(!t)return;
  if(t.dataset.tab)setTab(t.dataset.tab);
  else if(t.dataset.face){S.face=t.dataset.face;renderListFilters();renderList();}
  else if(t.dataset.scope){if(dirty.size&&!confirm('切換範圍會清除尚未儲存的內容，確定嗎？'))return;S.scope=t.dataset.scope;renderFill();}
  else if(t.dataset.open){e.preventDefault();openDrawer(t.dataset.open);}
  else if(t.matches('tr.row'))openDrawer(t.dataset.kid);
  else if(t.dataset.close!==undefined)closeDrawer();
  else if(t.dataset.gofill){closeDrawer();S.fillMonth=repMonth(S.year);S.scope='all';setTab('fill');document.querySelector(`.fcard[data-kid="${CSS.escape(t.dataset.gofill)}"]`)?.scrollIntoView({block:'center'});}
  else if(t.dataset.year){S.year=+t.dataset.year;renderTop();setTab('list');}
  else if(t.dataset.newkpi!==undefined)openKpiForm(null);
  else if(t.dataset.editkpi)openKpiForm(t.dataset.editkpi);
  else if(t.dataset.toggle&&isAdmin())togglePerson(t.dataset.toggle);
  else if(t.dataset.mkyear&&isAdmin())makeYear(+t.dataset.mkyear,0);
  else if(t.dataset.cpyear&&isAdmin())makeYear(+t.dataset.cpyear,+t.dataset.from);
  else if(t.id==='moveBtn'&&isAdmin())moveKpis();
  else if(t.id==='reloadBtn')refresh();
  else if(t.id==='loginBtn')Graph.login();
  else if(t.id==='logoutBtn'){if(dirty.size&&!confirm('有尚未儲存的內容，確定要登出嗎？'))return;Graph.logout();}
  else if(t.dataset.relogin!==undefined){saveDrafts();Graph.relogin();}
});
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeDrawer();if(e.key==='Enter'&&e.target.matches&&e.target.matches('tr.row'))openDrawer(e.target.dataset.kid);});
document.addEventListener('input',e=>{const f=e.target.closest&&e.target.closest('.fcard .form');if(!f)return;dirty.add(f.dataset.kid);f.querySelector('[data-dirty]').hidden=false;});
document.addEventListener('submit',e=>{e.preventDefault();const f=e.target;
  if(f.matches('.fcard .form'))saveRecord(f);else if(f.id==='kpiForm')saveKpi(f);else if(f.id==='addPerson'&&isAdmin())addPerson(f);});
window.addEventListener('beforeunload',e=>{if(dirty.size){e.preventDefault();e.returnValue='';}});

async function boot(){
  $('scrim').addEventListener('click',closeDrawer);
  $('yearSel').addEventListener('change',e=>{S.year=+e.target.value;S.goal='';S.fillMonth=repMonth(S.year)||1;render();});
  $('fillMonth').addEventListener('change',e=>{if(dirty.size&&!confirm('切換月份會清除尚未儲存的內容，確定嗎？')){e.target.value=S.fillMonth;return;}S.fillMonth=+e.target.value;renderFill();});
  ['fDept','fPerson','fGoal','fStatus'].forEach(id=>$(id).addEventListener('change',e=>{S[{fDept:'dept',fPerson:'person',fGoal:'goal',fStatus:'status'}[id]]=e.target.value;renderList();}));
  $('fQ').addEventListener('input',e=>{S.q=e.target.value;renderList();});
  const miss=['tenantId','clientId','fileShareUrl'].filter(k=>!CFG[k]&&!CFG.devMockBase);
  if(miss.length)return fatal(`<h2 style="margin-top:0">網站尚未完成設定</h2><p>請在 <code>config.js</code> 填入：${miss.join('、')}。詳見設定步驟。</p>`);
  let st;
  try{st=await Graph.init(CFG);}catch(e){return fatal(`<h2 style="margin-top:0">登入元件啟動失敗</h2><p class="small">${esc(e.message)}</p><p>請確認 config.js 的 tenantId 與 clientId 是否正確，以及應用程式註冊的重新導向 URI 是否為：<br><code>${esc(location.origin+location.pathname)}</code></p>`);}
  if(!st.signedIn)return fatal(`<h2 style="margin-top:0;font-family:var(--f-display)">普萊德 永續 KPI 管理平台</h2><p>請使用公司 Microsoft 365 帳號登入。登入後，網站會以你的帳號讀寫「年度永續指標管理」團隊中的 KPI Excel。</p><button class="btn pri" id="loginBtn" type="button">使用 Microsoft 365 登入</button>`);
  $('boot').innerHTML='<p class="empty">讀取 Excel 中…</p>';
  try{
    const u=await Graph.me();me.mail=(u.mail||u.userPrincipalName||'').toLowerCase();me.display=u.displayName||'';
    await Graph.open(CFG);await loadAll();
  }catch(e){
    const k=e.kind;
    return fatal(`<h2 style="margin-top:0">無法讀取 KPI Excel</h2><p>${k==='forbidden'||k==='notfound'?'你的帳號無法開啟設定的 Excel。請確認你是「年度永續指標管理」Teams 團隊的成員，或聯絡 Kiko。':k==='structure'?esc(e.message):k==='auth'?'登入狀態失效或權限尚未同意。':'連線發生問題，請稍後重新整理。'}</p><p class="small muted">技術資訊：${esc(k)} ${esc(e.status||'')} ${esc(e.code||'')} ${esc(e.message)}</p><button class="btn" type="button" onclick="location.reload()">重新整理</button> <button class="btn" type="button" id="logoutBtn">登出</button>`);
  }
  $('boot').hidden=true;$('app').hidden=false;$('openXlsx').href=Graph.file?.webUrl||'#';
  renderTop();setTab(S.tab);
}
boot();
