// 資料層：Microsoft 365 登入（MSAL）＋ Microsoft Graph Excel API
// 所有讀寫都以「登入者本人」的權限執行；網站不保存任何密碼、Client Secret 或資料副本。
(function(){
  const SCOPES=['User.Read','Files.ReadWrite.All'];
  let G='https://graph.microsoft.com/v1.0';
  let app=null,account=null,base=null,sessionId=null,devMode=false,fileInfo=null;

  class GraphError extends Error{
    constructor(kind,message,status,code){super(message);this.kind=kind;this.status=status;this.code=code;}
  }
  function kindOf(status,code){
    if(status===401)return 'auth';
    if(status===403)return 'forbidden';
    if(status===404)return 'notfound';
    if(status===409&&/conflict|EditMode/i.test(code||''))return 'locked';
    if(status===423)return 'locked';
    if(status===429)return 'throttle';
    if(status===503||status===504)return 'server';
    return status>=500?'server':'other';
  }
  const sleep=ms=>new Promise(r=>setTimeout(r,ms));

  async function init(cfg){
    // 本機測試模式：只在 localhost 且設定了 devMockBase 時啟用，正式網址不會進入此模式
    if(cfg.devMockBase&&['localhost','127.0.0.1'].includes(location.hostname)){devMode=true;G=cfg.devMockBase;return {signedIn:true};}
    app=new msal.PublicClientApplication({
      auth:{clientId:cfg.clientId,authority:'https://login.microsoftonline.com/'+cfg.tenantId,redirectUri:location.origin+location.pathname},
      cache:{cacheLocation:'sessionStorage'}
    });
    await app.initialize();
    const r=await app.handleRedirectPromise();
    account=(r&&r.account)||app.getAllAccounts()[0]||null;
    return {signedIn:!!account};
  }
  function login(){return app.loginRedirect({scopes:SCOPES,prompt:'select_account'});}
  function logout(){return app.logoutRedirect({account});}
  function relogin(){return app.acquireTokenRedirect({scopes:SCOPES,account});}
  async function token(force){
    if(devMode)return 'dev';
    try{const r=await app.acquireTokenSilent({scopes:SCOPES,account,forceRefresh:!!force});return r.accessToken;}
    catch(e){throw new GraphError('auth','登入已過期，需要重新登入。');}
  }

  // 單一 HTTP 呼叫。idempotent=true 的請求（讀取、覆寫同一列）遇到忙碌或逾時會自動重試；
  // 新增列（POST rows）不自動重試，交由呼叫端先確認是否已寫入，避免重複新增。
  async function call(method,path,body,opt={}){
    const idem=opt.idempotent??(method==='GET'||method==='PATCH');
    let tries=0,refreshed=false,resessioned=false;
    for(;;){
      tries++;
      const ctl=new AbortController(),timer=setTimeout(()=>ctl.abort(),opt.timeout||30000);
      let res;
      try{
        const h={'Authorization':'Bearer '+await token(refreshed),'Content-Type':'application/json'};
        if(sessionId&&path.startsWith(base||'\u0000'))h['workbook-session-id']=sessionId;
        res=await fetch(G+path,{method,headers:h,body:body?JSON.stringify(body):undefined,signal:ctl.signal,cache:'no-store'});
      }catch(e){
        clearTimeout(timer);
        if(e instanceof GraphError)throw e;
        const kind=e.name==='AbortError'?'timeout':'network';
        if(idem&&tries<3){await sleep(1500*tries);continue;}
        throw new GraphError(kind,kind==='timeout'?'Excel 回應逾時。':'無法連線到 Microsoft 365。');
      }
      clearTimeout(timer);
      if(res.ok){return res.status===204?null:res.json().catch(()=>null);}
      let code='',msg='';try{const j=await res.json();code=j?.error?.innerError?.code||j?.error?.code||'';msg=j?.error?.message||'';}catch(e){}
      if(res.status===401&&!refreshed){refreshed=true;continue;}
      if(res.status===404&&sessionId&&!resessioned&&/session/i.test(code+msg)){resessioned=true;await createSession();continue;}
      const kind=kindOf(res.status,code);
      if(idem&&(kind==='throttle'||kind==='server')&&tries<4){
        const ra=+res.headers.get('Retry-After');await sleep((ra>0?ra:2*tries)*1000);continue;
      }
      throw new GraphError(kind,msg||('HTTP '+res.status),res.status,code);
    }
  }
  function encShare(url){
    const b=btoa(unescape(encodeURIComponent(url))).replace(/=+$/,'').replace(/\//g,'_').replace(/\+/g,'-');
    return 'u!'+b;
  }
  async function createSession(){
    sessionId=null;
    const s=await call('POST',base+'/createSession',{persistChanges:true},{idempotent:true});
    sessionId=s&&s.id;
  }
  async function open(cfg){
    const it=await call('GET','/shares/'+encShare(cfg.fileShareUrl)+'/driveItem?$select=id,name,webUrl,parentReference');
    if(!it||!it.parentReference)throw new GraphError('notfound','找不到設定的 Excel 檔案。');
    fileInfo={name:it.name,webUrl:it.webUrl};
    base='/drives/'+it.parentReference.driveId+'/items/'+it.id+'/workbook';
    await createSession();
    return fileInfo;
  }
  async function me(){
    if(devMode)return call('GET','/me');
    return call('GET','/me?$select=displayName,mail,userPrincipalName');
  }
  // 讀整張表（含表頭）。回傳 {headers, rows:[[...]]}，rows[i] 對應表格第 i 列（0 起算）
  async function readTable(name){
    const r=await call('GET',base+'/tables/'+encodeURIComponent(name)+'/range?$select=values');
    const v=(r&&r.values)||[];
    return {headers:(v[0]||[]).map(x=>String(x).trim()),rows:v.slice(1)};
  }
  async function addRows(name,rows){
    return call('POST',base+'/tables/'+encodeURIComponent(name)+'/rows',{index:null,values:rows},{idempotent:false,timeout:45000});
  }
  async function rowRange(name,idx){
    return call('GET',base+'/tables/'+encodeURIComponent(name)+'/rows/itemAt(index='+idx+')/range?$select=address,values');
  }
  async function patchAddress(address,values){
    const i=address.lastIndexOf('!');
    const sheet=address.slice(0,i).replace(/^'|'$/g,'').replace(/''/g,"'"),cells=address.slice(i+1);
    return call('PATCH',base+'/worksheets/'+encodeURIComponent(sheet)+"/range(address='"+cells+"')",{values});
  }
  window.Graph={GraphError,init,login,logout,relogin,open,me,readTable,addRows,rowRange,patchAddress,createSession,get file(){return fileInfo;},get dev(){return devMode;}};
})();
