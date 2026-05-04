/* ═══════════ CONFIG ═══════════ */
var CLIENT_ID    = '4ae73f12-9d89-467d-9dec-79be71a18224';
var TENANT_ID    = '21827f82-8236-44d6-b5fe-bd2d5d65b3a9';
var REDIRECT     = 'https://shantala-mitchell.github.io/ITServiceDesk/dashboard.html';
var SITE_ID      = 'cd2c6dbd-3685-41c9-9d76-820fb6f2350d';
var LIST_ID      = 'c83d54e5-0c6a-4274-aeb2-b584e48826b8';
var AGENTS_LIST_ID = 'c3b2a174-0cbc-4f8b-9d0f-075c60746aca';
var GRAPH_BASE   = 'https://graph.microsoft.com/v1.0';
var LIST_URL     = GRAPH_BASE+'/sites/'+SITE_ID+'/lists/'+LIST_ID;
var AGENTS_URL   = GRAPH_BASE+'/sites/'+SITE_ID+'/lists/'+AGENTS_LIST_ID;

var DEFAULT_AGENTS = [{email:'shantala.mitchell@ecnz.ac.nz',role:'admin'}];
var FLOW_REPLY  = '';
var FLOW_SUBMIT = '';
var cachedAgents = null;

/* ═══════════ MSAL ═══════════ */
var msalInstance = new msal.PublicClientApplication({
  auth:{clientId:CLIENT_ID,authority:'https://login.microsoftonline.com/'+TENANT_ID,redirectUri:REDIRECT,navigateToLoginRequestUrl:true},
  cache:{cacheLocation:'sessionStorage',storeAuthStateInCookie:true}
});
var loginRequest = {scopes:['User.Read','openid','profile']};
var graphRequest = {scopes:['https://graph.microsoft.com/Sites.Selected']};
var currentUser  = null;
var currentRole  = 'agent';
var msalReady    = false;

msalInstance.initialize().then(function(){
  msalReady=true;
  return msalInstance.handleRedirectPromise();
}).then(function(resp){
  var acct=(resp&&resp.account)?resp.account:null;
  if(!acct){var all=msalInstance.getAllAccounts();if(all.length)acct=all[0];}
  if(acct){msalInstance.setActiveAccount(acct);handleAuthSuccess(acct);}
}).catch(function(e){console.error('MSAL:',e);showAuthErr('Sign in error: '+(e.message||'Please try again.'));});

function getToken(){
  var acct=msalInstance.getActiveAccount();
  if(!acct)return Promise.reject('No account');
  return msalInstance.acquireTokenSilent({scopes:['https://graph.microsoft.com/Sites.Selected'],account:acct})
  .then(function(r){return r.accessToken;})
  .catch(function(e){
    console.warn('Silent token failed, trying popup:',e);
    return msalInstance.acquireTokenPopup({scopes:['https://graph.microsoft.com/Sites.Selected'],account:acct})
    .then(function(r){return r.accessToken;});
  });
}

function graphFetch(url,method,body){
  return getToken().then(function(token){
    var opts={method:method||'GET',headers:{'Authorization':'Bearer '+token,'Content-Type':'application/json'}};
    if(body) opts.body=JSON.stringify(body);
    return fetch(url,opts);
  });
}

function handleAuthSuccess(acct){
  currentUser=acct;
  currentRole='agent'; // role will be confirmed by checkAccess against live SharePoint data
  applyStoredSettings();
  document.getElementById('nav-name').textContent=acct.name||acct.username;
  var rb=document.getElementById('nav-role-badge');
  rb.textContent='Agent';rb.classList.remove('admin');
  document.getElementById('nav-user').style.display='flex';
  document.getElementById('nav-tabs').style.visibility='visible';
  document.getElementById('nb-settings').style.display='none'; // hidden until checkAccess confirms admin
  checkAccess(acct.username);
}

function checkAccess(email){
  document.getElementById('screen-auth').style.display='none';
  loadAgentsFromGraph(true).then(function(agents){
    var allowed=agents.map(function(a){return a.email.toLowerCase();});
    var me=agents.find(function(a){return a.email.toLowerCase()===(email||'').toLowerCase();});
    var meRole=me?(me.role||'agent').toLowerCase().trim().replace('administrator','admin'):'agent';
    if(meRole==='admin'){
      currentRole='admin';
      var rb=document.getElementById('nav-role-badge');
      rb.textContent='Admin';rb.classList.add('admin');
    }
    document.getElementById('nb-settings').style.display=currentRole==='admin'?'':'none';
    if(allowed.indexOf((email||'').toLowerCase())!==-1){
      document.getElementById('screen-denied').style.display='none';
      document.getElementById('page-tickets').classList.add('active');
      loadTickets();
      startSLATicker();
    } else {
      document.getElementById('denied-msg').textContent=(email||'Your account')+' does not have agent access. Contact your administrator.';
      document.getElementById('screen-denied').style.display='flex';
    }
  }).catch(function(e){
    console.error('ITDesk checkAccess error:',e);
    var fallback=DEFAULT_AGENTS.find(function(a){return a.email.toLowerCase()===(email||'').toLowerCase();});
    if(fallback&&fallback.role==='admin'){
      currentRole='admin';
      var rb=document.getElementById('nav-role-badge');
      rb.textContent='Admin';rb.classList.add('admin');
      document.getElementById('nb-settings').style.display='';
      document.getElementById('screen-denied').style.display='none';
      document.getElementById('page-tickets').classList.add('active');
      loadTickets();startSLATicker();
    } else {
      document.getElementById('screen-denied').style.display='flex';
      document.getElementById('denied-msg').textContent='Could not verify access — check your connection and try again.';
    }
  });
}

function signIn(){
  if(!msalReady){showAuthErr('Still loading — please wait a moment.');return;}
  showAuthErr('');
  msalInstance.loginRedirect(loginRequest).catch(function(e){showAuthErr('Sign in failed: '+(e.message||'Please try again.'));});
}
function signOut(){msalInstance.logoutRedirect({postLogoutRedirectUri:REDIRECT}).catch(function(){window.location.reload();});}
function showAuthErr(m){var e=document.getElementById('auth-err');e.textContent=m;e.style.display=m?'block':'none';}

/* ═══════════ PAGES ═══════════ */
function showPage(p){
  document.querySelectorAll('.page').forEach(function(el){el.classList.remove('active');});
  document.querySelectorAll('.nav-btn').forEach(function(b){b.classList.remove('active');});
  document.getElementById('page-'+p).classList.add('active');
  document.getElementById('nb-'+p).classList.add('active');
  if(p==='archive') renderArchive();
  if(p==='reports'){ loadSyskitSaved(); loadDefenderSaved(); loadPhishingSaved(); renderTicketReport(); document.getElementById('report-updated').textContent='Updated '+new Date().toLocaleTimeString('en-NZ',{hour:'2-digit',minute:'2-digit'}); }
  if(p==='settings'){if(currentRole!=='admin'){alert('Settings are restricted to administrators.');return;}loadSettingsUI();}
}

/* ═══════════ UTILS ═══════════ */
var tickets=[],activeFilter='all',openId=null,sortCol='created',sortDir=-1;
function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/'/g,'&#39;').replace(/"/g,'&quot;');}
function relTime(ts){var d=Date.now()-new Date(ts).getTime();if(d<3600000)return Math.floor(d/60000)+'m ago';if(d<86400000)return Math.floor(d/3600000)+'h ago';return Math.floor(d/86400000)+'d ago';}
function fmtDate(ts){return new Date(ts).toLocaleDateString('en-NZ',{day:'numeric',month:'short',year:'numeric'});}
function fmtTime(ts){return new Date(ts).toLocaleString('en-NZ',{day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'});}
function fmtDur(ms){var h=Math.floor(ms/3600000);var m=Math.floor((ms%3600000)/60000);if(h>=24){var d=Math.floor(h/24);return d+'d '+(h%24)+'h';}if(h>0)return h+'h '+m+'m';return m+'m';}
function priBadge(p){return '<span class="pri-badge"><span class="dot dot-'+esc(p)+'"></span>'+esc(p)+'</span>';}
function statusBadge(s){var map={open:'st-open','in-progress':'st-progress','on-hold':'st-hold',resolved:'st-resolved',closed:'st-closed'};var labels={open:'Open','in-progress':'In Progress','on-hold':'On Hold',resolved:'Resolved',closed:'Closed'};return'<span class="st-badge '+(map[s]||'')+'">'+esc(labels[s]||s)+'</span>';}

/* ═══════════ UNSEEN HIGHLIGHTS ═══════════ */
var _viewedTickets=new Set(); // in-memory; cleared on page reload only

function getUnseenStore(){try{return JSON.parse(localStorage.getItem('itdesk_unseen')||'{}');}catch(e){return{};}}
function getSeenState(){try{return JSON.parse(localStorage.getItem('itdesk_seen_state')||'{}');}catch(e){return{};}}

function updateSeenState(tks){
  var seenState=getSeenState();
  var isFirstLoad=Object.keys(seenState).length===0;
  var existingUnseen=getUnseenStore();
  var newState={};
  var newUnseen={};
  tks.forEach(function(t){
    if(t._archived)return;
    var rc=(t.replies||[]).length;
    newState[t.id]=rc;
    if(!isFirstLoad){
      if(rc>(seenState[t.id]||0)&&seenState[t.id]!==undefined){
        // genuine new reply — re-highlight even if previously viewed
        newUnseen[t.id]='reply';
        _viewedTickets.delete(t.id);
      } else if(seenState[t.id]===undefined){
        // brand new ticket
        if(!_viewedTickets.has(t.id))newUnseen[t.id]='new';
      } else if(existingUnseen[t.id]&&!_viewedTickets.has(t.id)){
        // preserve existing unseen entry if not yet viewed
        newUnseen[t.id]=existingUnseen[t.id];
      }
    }
  });
  localStorage.setItem('itdesk_seen_state',JSON.stringify(newState));
  localStorage.setItem('itdesk_unseen',JSON.stringify(newUnseen));
}

function clearUnseen(id){
  _viewedTickets.add(id);
  var unseen=getUnseenStore();
  if(unseen[id]){
    delete unseen[id];
    localStorage.setItem('itdesk_unseen',JSON.stringify(unseen));
  }
  var row=document.querySelector('#ticket-body tr[data-id="'+id+'"]');
  if(row){row.classList.remove('ticket-unseen');var dot=row.querySelector('.unseen-dot');if(dot)dot.remove();}
}

/* ═══════════ SLA ═══════════ */
var _refreshDebounce=null;
function loadTicketsDebounced(){
  if(_refreshDebounce)return;
  loadTickets();
  _refreshDebounce=setTimeout(function(){_refreshDebounce=null;},5000);
}
var DEFAULT_SLA={critical:1,high:4,medium:8,low:24};
function getSLATargets(){var s=getStoredSettings();return{critical:(s.slaCrit||DEFAULT_SLA.critical)*3600000,high:(s.slaHigh||DEFAULT_SLA.high)*3600000,medium:(s.slaMed||DEFAULT_SLA.medium)*3600000,low:(s.slaLow||DEFAULT_SLA.low)*3600000};}
function getSLAInfo(t){
  if(t.status==='resolved'||t.status==='closed')return{state:'met',label:'SLA met',cd:'&#10003; Met'};
  if(t.status==='on-hold')return{state:'hold',label:'SLA paused (on hold)',cd:'&#9646;&#9646; Paused'};
  var targets=getSLATargets();var target=targets[t.priority]||targets.low;
  var holdMs=t.holdDuration||0;
  var elapsed=Date.now()-new Date(t.created).getTime()-holdMs;
  var remaining=target-elapsed;var pct=Math.min(100,(elapsed/target)*100);
  if(remaining<=0)return{state:'breach',label:'Breached by '+fmtDur(Math.abs(remaining)),cd:'+'+fmtDur(Math.abs(remaining))};
  return{state:pct>=80?'warn':'ok',label:'Due in '+fmtDur(remaining),cd:fmtDur(remaining)};
}
function slaBadge(t){var i=getSLAInfo(t);var cls={breach:'sla-breach',warn:'sla-warn',ok:'sla-ok',hold:'sla-hold',met:'sla-met'};return'<span class="sla-badge '+(cls[i.state]||'')+'">'+i.cd+'</span>';}
function updatePanelSLA(t){var i=getSLAInfo(t);var b=document.getElementById('p-sla-block');b.className='sla-block '+i.state;document.getElementById('p-sla-cd').innerHTML=i.cd;document.getElementById('p-sla-lbl').textContent=i.label;}
function startSLATicker(){if(window._slaTick)clearInterval(window._slaTick);window._slaTick=setInterval(function(){var active=tickets.filter(function(t){return !t._archived;});document.getElementById('s-breached').textContent=active.filter(function(t){return getSLAInfo(t).state==='breach';}).length;document.querySelectorAll('#ticket-body tr[data-id]').forEach(function(row){var t=getTicket(row.getAttribute('data-id'));if(!t)return;var sc=row.querySelector('.sla-cell');if(sc)sc.innerHTML=slaBadge(t);var info=getSLAInfo(t);var slaCls=info.state==='breach'?'sla-breach':info.state==='warn'?'sla-warn':'';var hasUnseen=row.classList.contains('ticket-unseen');row.className=slaCls+(hasUnseen?' ticket-unseen':'');});if(openId){var t=getTicket(openId);if(t)updatePanelSLA(t);}},30000);}

/* ═══════════ SHAREPOINT GRAPH API ═══════════ */
function mapItem(i){
  var f=i.fields||{};
  var replies=[];
  try{replies=JSON.parse(f.Replies||'[]');}catch(e){}
  return{
    spId:      i.id,
    id:        f.TicketID||('INC-'+i.id),
    name:      f.RequesterName||'',
    email:     f.RequesterEmail||'',
    dept:      f.Department||'',
    cat:       f.Category||'',
    subject:   f.Title||'',
    desc:      f.Description||'',
    priority:  (f.Priority||'low').toLowerCase(),
    status:    (f.Status||'open').toLowerCase(),
    assignee:  f.Assignee||'',
    notes:     f.AgentNotes||'',
    watchers:  f.Watchers||'',
    replies:   replies,
    holdDuration: f.HoldDuratiom||0,
    holdStart: f.HoldStart||null,
    created:   f.Created||i.createdDateTime||new Date().toISOString(),
    closed:    f.Closed||null,
    _archived: false
  };
}

function fetchAllPages(url, allItems){
  allItems = allItems || [];
  return graphFetch(url)
  .then(function(r){
    if(!r.ok) throw new Error('Graph error '+r.status);
    return r.json();
  })
  .then(function(data){
    allItems = allItems.concat(data.value||[]);
    if(data['@odata.nextLink']){
      return fetchAllPages(data['@odata.nextLink'], allItems);
    }
    return allItems;
  });
}

function loadTickets(){
  document.getElementById('dash-date').textContent=new Date().toLocaleDateString('en-NZ',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
  document.getElementById('ticket-body').innerHTML='<tr class="loading-row"><td colspan="7">Loading&hellip;</td></tr>';
  document.getElementById('empty-state').style.display='none';

  var url=LIST_URL+'/items?expand=fields&$top=500';
  fetchAllPages(url)
  .then(function(items){
    tickets=items.map(mapItem);
    updateSeenState(tickets);
    autoArchive();renderDash();
    if(document.getElementById('page-reports').classList.contains('active')) renderReports();
  })
  .catch(function(e){
    console.error('Load error:',e);
    var s=getStoredSettings();
    tickets=[];
    if(s.demo!==false)tickets=JSON.parse(JSON.stringify(DEMO));
    autoArchive();renderDash();
    if(document.getElementById('page-reports').classList.contains('active')) renderReports();
    document.getElementById('ticket-body').insertAdjacentHTML('afterbegin','<tr class="loading-row"><td colspan="7" style="color:var(--warn);">&#9888; Using demo data — could not connect to SharePoint.</td></tr>');
  });
}

function persistToGraph(t){
  var statusMap={'open':'open','in-progress':'in-progress','on-hold':'on-hold','closed':'closed'};
  var priorityMap={'low':'low','medium':'medium','high':'high','critical':'critical'};
  var fields={
    Status:       statusMap[t.status]||t.status,
    Priority:     priorityMap[t.priority]||t.priority,
    AgentNotes:   t.notes||'',
    Assignee:     t.assignee||'',
    Replies:      JSON.stringify(t.replies||[]),
    HoldDuratiom: t.holdDuration||0,
    HoldStart:    t.holdStart||'',
    Closed:       t.closed||''
  };
  console.log('[Save] persistToGraph called. spId:', t.spId);
  return graphFetch(LIST_URL+'/items/'+t.spId+'/fields','PATCH',fields)
  .then(function(r){
    if(!r.ok){
      return r.text().then(function(txt){
        console.error('PATCH failed '+r.status+':', txt);
        throw new Error('PATCH failed: '+r.status);
      });
    }
  });
}

function createTicketInGraph(ticket){
  var fields={
    Title:          ticket.subject,
    TicketID:       ticket.id,
    RequesterName:  ticket.name,
    RequesterEmail: ticket.email,
    Department:     ticket.dept,
    Category:       ticket.cat,
    Description:    ticket.desc,
    Priority:       ticket.priority,
    Status:         'open',
    Replies:        '[]',
    HoldDuratiom:   0
  };
  return graphFetch(LIST_URL+'/items','POST',{fields:fields})
  .then(function(r){return r.json();})
  .then(function(data){
    // update TicketID with real SP ID
    var realId='INC-'+data.id;
    return graphFetch(LIST_URL+'/items/'+data.id+'/fields','PATCH',{TicketID:realId})
    .then(function(){return realId;});
  });
}

function autoArchive(){
  var s=getStoredSettings();var days=(s.archivedays||7)*86400000;
  tickets.forEach(function(t){
    if(t.status==='closed'&&!t.closed)t.closed=new Date().toISOString();
    if(t.status==='closed'&&t.closed&&(Date.now()-new Date(t.closed).getTime())>days)t._archived=true;
  });
}

/* ═══════════ RENDER ═══════════ */
function renderDash(){
  var active=tickets.filter(function(t){return !t._archived;});
  document.getElementById('s-open').textContent    =active.filter(function(t){return t.status==='open';}).length;
  document.getElementById('s-progress').textContent=active.filter(function(t){return t.status==='in-progress';}).length;
  document.getElementById('s-hold').textContent    =active.filter(function(t){return t.status==='on-hold';}).length;
  document.getElementById('s-resolved').textContent=active.filter(function(t){return t.status==='closed';}).length;
  document.getElementById('s-breached').textContent=active.filter(function(t){return getSLAInfo(t).state==='breach';}).length;

  var me=currentUser?(currentUser.username||'').toLowerCase():'';
  var f=active;
  if(activeFilter==='mine')           f=active.filter(function(t){return(t.assignee||'').toLowerCase()===me;});
  else if(activeFilter==='unassigned')f=active.filter(function(t){return !t.assignee;});
  else if(activeFilter==='breached')  f=active.filter(function(t){return getSLAInfo(t).state==='breach';});
  else if(activeFilter!=='all')       f=active.filter(function(t){return t.status===activeFilter;});

  var priOrder={critical:0,high:1,medium:2,low:3};
  f=f.slice().sort(function(a,b){
    var av,bv;
    if(sortCol==='priority'){av=priOrder[a.priority]||99;bv=priOrder[b.priority]||99;}
    else if(sortCol==='created'){av=new Date(a.created).getTime();bv=new Date(b.created).getTime();}
    else{av=(a[sortCol]||'').toLowerCase();bv=(b[sortCol]||'').toLowerCase();}
    if(av<bv)return-1*sortDir;if(av>bv)return sortDir;return 0;
  });

  var tbody=document.getElementById('ticket-body');tbody.innerHTML='';
  if(!f.length){document.getElementById('empty-state').style.display='block';return;}
  document.getElementById('empty-state').style.display='none';

  var unseen=getUnseenStore();
  f.forEach(function(t){
    var tr=document.createElement('tr');
    var info=getSLAInfo(t);
    var cls=[];
    if(unseen[t.id]&&!_viewedTickets.has(t.id))cls.push('ticket-unseen');
    if(info.state==='breach')cls.push('sla-breach');
    else if(info.state==='warn')cls.push('sla-warn');
    tr.className=cls.join(' ');
    tr.setAttribute('data-id',t.id);
    var asgHtml=t.assignee?'<span class="asg-badge'+(t.assignee.toLowerCase()===me?' me':'')+'">'+(t.assignee.toLowerCase()===me?'Me':esc(t.assignee.split('@')[0]))+'</span>':'<span class="asg-badge">Unassigned</span>';
    var stMap={open:'st-open','in-progress':'st-progress','on-hold':'st-hold',resolved:'st-resolved',closed:'st-closed'};
    var stLabels={open:'Open','in-progress':'In Progress','on-hold':'On Hold',resolved:'Resolved',closed:'Closed'};
    var unreadDot=unseen[t.id]?'<span class="unseen-dot" title="'+(unseen[t.id]==='new'?'New ticket':'New reply')+'"></span>':'';
    tr.innerHTML=
      '<td class="tid">'+esc(t.id)+'</td>'+
      '<td class="ttitle">'+esc(t.subject)+unreadDot+(t.replies&&t.replies.length?'<span style="color:var(--ok);margin-left:6px;font-size:10px;">\u2709 '+t.replies.length+'</span>':'')+'</td>'+
      '<td>'+priBadge(t.priority)+'</td>'+
      '<td><span class="st-badge '+(stMap[t.status]||'')+'">'+esc(stLabels[t.status]||t.status)+'</span></td>'+
      '<td class="sla-cell">'+slaBadge(t)+'</td>'+
      '<td>'+asgHtml+'</td>'+
      '<td class="tmeta">'+relTime(t.created)+'</td>';
    (function(id){tr.onclick=function(){openPanel(id);};})(t.id);
    tbody.appendChild(tr);
  });
}

function setFilter(f,el){activeFilter=f;document.querySelectorAll('.chip').forEach(function(c){c.classList.remove('active');});el.classList.add('active');renderDash();}

function sortBy(col){
  if(sortCol===col)sortDir*=-1;else{sortCol=col;sortDir=1;}
  document.querySelectorAll('#main-table th[data-col]').forEach(function(th){
    th.classList.toggle('sorted',th.getAttribute('data-col')===col);
    var icon=th.querySelector('.sort-icon');
    if(icon)icon.innerHTML=th.getAttribute('data-col')===col?(sortDir===1?'&#8593;':'&#8595;'):'&#8597;';
  });
  renderDash();
}

/* ═══════════ ARCHIVE ═══════════ */
function renderArchive(){
  var archived=tickets.filter(function(t){return t._archived;});
  var tbody=document.getElementById('archive-body');tbody.innerHTML='';
  if(!archived.length){document.getElementById('archive-empty').style.display='block';return;}
  document.getElementById('archive-empty').style.display='none';
  archived.forEach(function(t){
    var tr=document.createElement('tr');
    tr.innerHTML='<td class="tid">'+esc(t.id)+'</td><td class="ttitle">'+esc(t.subject)+'</td><td>'+priBadge(t.priority)+'</td><td class="tmeta">'+(t.closed?fmtDate(t.closed):'—')+'</td><td class="tmeta">'+esc(t.assignee||'Unassigned')+'</td>';
    (function(id){tr.onclick=function(){openPanel(id);};})(t.id);
    tbody.appendChild(tr);
  });
}

/* ═══════════ REPORTS ═══════════ */
var activeReport='tickets';

function showReport(key,chipEl){
  activeReport=key;
  document.querySelectorAll('.report-section').forEach(function(s){s.classList.remove('active');});
  document.getElementById('rs-'+key).classList.add('active');
  document.querySelectorAll('#page-reports .chip').forEach(function(c){c.classList.remove('active');});
  if(chipEl)chipEl.classList.add('active');
  refreshActiveReport();
}

function refreshActiveReport(){
  document.getElementById('report-updated').textContent='Updated '+new Date().toLocaleTimeString('en-NZ',{hour:'2-digit',minute:'2-digit'});
  if(activeReport==='tickets')  renderTicketReport();
  if(activeReport==='defender') loadDefenderAlerts();
  if(activeReport==='phishing') loadPhishingData();
  if(activeReport==='syskit')   renderSyskitManual();
}

function renderReports(){
  renderTicketReport();
  loadDefenderSaved();
  loadPhishingSaved();
  loadSyskitSaved();
  document.getElementById('report-updated').textContent='Updated '+new Date().toLocaleTimeString('en-NZ',{hour:'2-digit',minute:'2-digit'});
}

/* ── Tickets ── */
function renderTicketReport(){
  var days=parseInt(document.getElementById('rpt-range').value)||30;
  var now=Date.now();var cutoff=days===0?0:now-days*86400000;
  var src=(tickets&&tickets.length)?tickets:(typeof DEMO!=='undefined'?DEMO:[]);
  var t=src.filter(function(x){return new Date(x.created).getTime()>=cutoff&&!x._archived;});
  var total=t.length;
  var resolved=t.filter(function(x){return x.status==='resolved'||x.status==='closed';}).length;
  var unassigned=t.filter(function(x){return !x.assignee&&(x.status==='open'||x.status==='in-progress');}).length;
  var s=getStoredSettings();
  var slaMap={critical:(s.slaCrit||1)*3600000,high:(s.slaHigh||4)*3600000,medium:(s.slaMed||8)*3600000,low:(s.slaLow||24)*3600000};
  var metSla=0;
  t.forEach(function(x){if(x.status==='on-hold'||now-new Date(x.created).getTime()<=(slaMap[x.priority]||slaMap.low))metSla++;});
  var slaRate=total?Math.round(metSla/total*100):100;
  document.getElementById('rpt-total').textContent=total;
  document.getElementById('rpt-total-sub').textContent=days===0?'all time':'last '+days+' days';
  document.getElementById('rpt-resolved').textContent=resolved;
  document.getElementById('rpt-resolved-sub').textContent=total?Math.round(resolved/total*100)+'% of total':'';
  document.getElementById('rpt-resolved-bar').style.width=(total?Math.round(resolved/total*100):0)+'%';
  document.getElementById('rpt-sla').textContent=slaRate+'%';
  var slaBar=document.getElementById('rpt-sla-bar');
  slaBar.style.width=slaRate+'%';slaBar.className='rc-bar-fill'+(slaRate<70?' crit':slaRate<85?' warn':'');
  document.getElementById('rpt-unassigned').textContent=unassigned;
  // By category
  var cats={};t.forEach(function(x){cats[x.cat||'Uncategorised']=(cats[x.cat||'Uncategorised']||0)+1;});
  var catArr=Object.entries(cats).sort(function(a,b){return b[1]-a[1];});var maxC=catArr[0]?catArr[0][1]:1;
  document.getElementById('rpt-by-cat').innerHTML=catArr.length?catArr.map(function(e){return rpBarRow(e[0],e[1],Math.round(e[1]/maxC*100),'var(--accent)');}).join(''):'<div style="color:var(--muted);font-size:12px;">No tickets in range</div>';
  // By priority
  var priColors={critical:'#b91c1c',high:'var(--accent)',medium:'var(--warn)',low:'var(--ok)'};
  var priCounts={};t.forEach(function(x){priCounts[x.priority]=(priCounts[x.priority]||0)+1;});
  var maxP=Math.max.apply(null,Object.values(priCounts).concat([1]));
  document.getElementById('rpt-by-pri').innerHTML=['critical','high','medium','low'].map(function(p){return rpBarRow(p.charAt(0).toUpperCase()+p.slice(1),priCounts[p]||0,Math.round((priCounts[p]||0)/maxP*100),priColors[p]);}).join('');
  // By agent
  var agents={};t.forEach(function(x){var a=x.assignee||'Unassigned';agents[a]=(agents[a]||0)+1;});
  var agArr=Object.entries(agents).sort(function(a,b){return b[1]-a[1];});var maxA=agArr[0]?agArr[0][1]:1;
  document.getElementById('rpt-by-agent').innerHTML=agArr.length?agArr.map(function(e){return rpBarRow(e[0]==='Unassigned'?'Unassigned':e[0].split('@')[0],e[1],Math.round(e[1]/maxA*100),'var(--ok)');}).join(''):'<div style="color:var(--muted);font-size:12px;">No tickets in range</div>';
}

function rpBarRow(label,val,pct,color){
  return'<div class="rp-row"><span class="rp-label">'+esc(label)+'</span><div class="rp-bar-wrap"><div class="rp-bar-seg" style="width:'+pct+'%;background:'+color+'"></div></div><span class="rp-val">'+val+'</span></div>';
}

/* ── Defender ── */
function loadDefenderAlerts(){
  var badge=document.getElementById('defender-status-badge');
  badge.textContent='⬤ Checking…';badge.className='data-status';
  loadDefenderSaved();
  if(!msalInstance||!currentUser){document.getElementById('defender-config-note').style.display='block';badge.textContent='⬤ Manual';badge.className='data-status manual';return;}
  msalInstance.acquireTokenSilent({scopes:['https://graph.microsoft.com/SecurityAlert.Read.All'],account:currentUser})
  .then(function(resp){return fetch('https://graph.microsoft.com/v1.0/security/alerts_v2?$top=20&$orderby=createdDateTime desc',{headers:{Authorization:'Bearer '+resp.accessToken}});})
  .then(function(r){return r.json();})
  .then(function(data){
    if(data.error)throw new Error(data.error.message);
    var alerts=data.value||[];
    var h=alerts.filter(function(a){return a.severity==='high';}).length;
    var m=alerts.filter(function(a){return a.severity==='medium';}).length;
    var l=alerts.filter(function(a){return a.severity==='low';}).length;
    var i=alerts.filter(function(a){return a.severity==='informational';}).length;
    applyDefenderData(h,m,l,i,alerts,'live','via Microsoft Graph');
    document.getElementById('defender-manual').style.display='none';
    document.getElementById('defender-config-note').style.display='none';
  })
  .catch(function(){document.getElementById('defender-config-note').style.display='block';badge.textContent='⬤ Manual';badge.className='data-status manual';});
}

function applyDefenderData(h,m,l,i,alerts,mode,src){
  var badge=document.getElementById('defender-status-badge');
  badge.textContent='⬤ '+(mode==='live'?'Live':'Manual');badge.className='data-status '+(mode==='live'?'live':'manual');
  document.getElementById('def-s-high').textContent=h;document.getElementById('def-s-med').textContent=m;
  document.getElementById('def-s-low').textContent=l;document.getElementById('def-s-info').textContent=i;
  document.getElementById('def-alerts-src').textContent=src?'· '+src:'';
  var list=document.getElementById('def-alert-list');
  if(!alerts||!alerts.length){list.innerHTML='<div style="color:var(--muted);font-size:12px;padding:12px 0;">'+(mode==='manual'?'Manual mode — connect via Graph API to see individual alerts.':'No active alerts.')+'</div>';return;}
  list.innerHTML=alerts.slice(0,10).map(function(a){
    var sev=a.severity||'informational';var title=a.title||a.alertDisplayName||'Alert';
    var cat=a.category||a.serviceSource||'';var ts=a.createdDateTime||a.created||'';
    var date=ts?new Date(ts).toLocaleDateString('en-NZ',{day:'numeric',month:'short'}):'';
    return'<div class="alert-row"><span class="alert-sev sev-'+sev+'">'+sev+'</span><div class="alert-title">'+esc(title)+(cat?'<div style="font-size:11px;color:var(--muted);font-weight:400;margin-top:2px;">'+esc(cat)+'</div>':'')+'</div><span class="alert-meta">'+date+'</span></div>';
  }).join('');
}

function renderDefenderManual(){
  document.getElementById('def-s-high').textContent=parseInt(document.getElementById('def-high').value)||0;
  document.getElementById('def-s-med').textContent=parseInt(document.getElementById('def-med').value)||0;
  document.getElementById('def-s-low').textContent=parseInt(document.getElementById('def-low').value)||0;
  document.getElementById('def-s-info').textContent=parseInt(document.getElementById('def-info').value)||0;
}

function saveDefenderManual(){
  var d={high:parseInt(document.getElementById('def-high').value)||0,med:parseInt(document.getElementById('def-med').value)||0,low:parseInt(document.getElementById('def-low').value)||0,info:parseInt(document.getElementById('def-info').value)||0,newThisWeek:parseInt(document.getElementById('def-new').value)||0,resolvedThisWeek:parseInt(document.getElementById('def-resolved').value)||0,saved:new Date().toISOString()};
  localStorage.setItem('itdesk_def_manual',JSON.stringify(d));
  var msg=document.getElementById('def-saved');msg.style.display='inline';setTimeout(function(){msg.style.display='none';},2500);
  applyDefenderData(d.high,d.med,d.low,d.info,[],'manual',null);
}

function loadDefenderSaved(){
  var d=JSON.parse(localStorage.getItem('itdesk_def_manual')||'null');if(!d)return;
  document.getElementById('def-high').value=d.high||0;document.getElementById('def-med').value=d.med||0;
  document.getElementById('def-low').value=d.low||0;document.getElementById('def-info').value=d.info||0;
  document.getElementById('def-new').value=d.newThisWeek||0;document.getElementById('def-resolved').value=d.resolvedThisWeek||0;
  applyDefenderData(d.high||0,d.med||0,d.low||0,d.info||0,[],'manual',null);
}

/* ── Phishing ── */
function loadPhishingData(){
  var s=getStoredSettings();var badge=document.getElementById('phish-status-badge');
  loadPhishingSaved();
  var flowUrl=s.flowPhishing;
  if(!flowUrl){document.getElementById('phishing-flow-note').style.display='block';badge.textContent='⬤ Manual';badge.className='data-status manual';return;}
  badge.textContent='⬤ Fetching…';badge.className='data-status';
  fetch(flowUrl,{method:'GET'}).then(function(r){return r.json();}).then(function(d){
    badge.textContent='⬤ Live';badge.className='data-status live';
    document.getElementById('phishing-flow-note').style.display='none';
    populatePhishingInputs(d);renderPhishingManual();
  }).catch(function(){badge.textContent='⬤ Manual';badge.className='data-status manual';document.getElementById('phishing-flow-note').style.display='block';});
}

function populatePhishingInputs(d){
  if(d.campaign)document.getElementById('ph-campaign').value=d.campaign;
  ['sent','opened','clicked','submitted','reported','trained','atrisk'].forEach(function(k){if(d[k]!==undefined)document.getElementById('ph-'+k).value=d[k];});
}

function renderPhishingManual(){
  var sent=parseInt(document.getElementById('ph-sent').value)||0;
  var clicked=parseInt(document.getElementById('ph-clicked').value)||0;
  var reported=parseInt(document.getElementById('ph-reported').value)||0;
  var trained=parseInt(document.getElementById('ph-trained').value)||0;
  var atrisk=parseInt(document.getElementById('ph-atrisk').value)||0;
  var campaign=document.getElementById('ph-campaign').value;
  var clickRate=sent?Math.round(clicked/sent*100):0;
  var reportRate=sent?Math.round(reported/sent*100):0;
  var trainedPct=sent?Math.round(trained/sent*100):0;
  document.getElementById('ph-s-click').textContent=clickRate+'%';document.getElementById('ph-s-report').textContent=reportRate+'%';
  document.getElementById('ph-s-trained').textContent=trainedPct+'%';document.getElementById('ph-s-atrisk').textContent=atrisk+'%';
  document.getElementById('ph-bar-click').style.width=Math.min(clickRate,100)+'%';
  document.getElementById('ph-bar-report').style.width=Math.min(reportRate,100)+'%';
  document.getElementById('ph-bar-trained').style.width=Math.min(trainedPct,100)+'%';
  document.getElementById('ph-campaign-label').textContent=campaign?'· '+campaign:'';
  if(!sent){document.getElementById('ph-breakdown').innerHTML='<div style="color:var(--muted);font-size:12px;padding:12px 0;">Enter campaign data above.</div>';return;}
  var opened=parseInt(document.getElementById('ph-opened').value)||0;
  var submitted=parseInt(document.getElementById('ph-submitted').value)||0;
  document.getElementById('ph-breakdown').innerHTML=[
    {label:'Emails sent',val:sent,n:sent,color:'var(--ok)'},
    {label:'Opened email',val:opened,n:opened,color:'var(--warn)'},
    {label:'Clicked link',val:clicked+' ('+clickRate+'%)',n:clicked,color:'var(--accent)'},
    {label:'Submitted data',val:submitted,n:submitted,color:'#b91c1c'},
    {label:'Reported phish',val:reported+' ('+reportRate+'%)',n:reported,color:'var(--ok)'},
    {label:'Training completed',val:trained,n:trained,color:'#7c3aed'},
  ].map(function(r){return rpBarRow(r.label,r.val,sent?Math.round(r.n/sent*100):0,r.color);}).join('');
}

function savePhishingManual(){
  var keys=['campaign','sent','opened','clicked','submitted','reported','trained','atrisk'];
  var d={saved:new Date().toISOString()};
  keys.forEach(function(k){d[k]=document.getElementById('ph-'+k).value;});
  localStorage.setItem('itdesk_ph_manual',JSON.stringify(d));
  var msg=document.getElementById('ph-saved');msg.style.display='inline';setTimeout(function(){msg.style.display='none';},2500);
}

function loadPhishingSaved(){
  var d=JSON.parse(localStorage.getItem('itdesk_ph_manual')||'null');if(!d)return;
  populatePhishingInputs(d);renderPhishingManual();
}

/* ── Syskit ── */
function renderSyskitManual(){
  var teams=parseInt(document.getElementById('sk-teams').value)||0;
  var teamsI=parseInt(document.getElementById('sk-teams-inactive').value)||0;
  var groups=parseInt(document.getElementById('sk-groups').value)||0;
  var orphaned=parseInt(document.getElementById('sk-orphaned').value)||0;
  var users=parseInt(document.getElementById('sk-users').value)||0;
  var guests=parseInt(document.getElementById('sk-guests').value)||0;
  var inactive=parseInt(document.getElementById('sk-inactive').value)||0;
  var shared=parseInt(document.getElementById('sk-shared').value)||0;
  var anon=parseInt(document.getElementById('sk-anon').value)||0;
  var ext=parseInt(document.getElementById('sk-ext').value)||0;
  var storage=parseFloat(document.getElementById('sk-storage').value)||0;
  var storageTotal=parseFloat(document.getElementById('sk-storage-total').value)||0;
  var storagePct=storageTotal?Math.round(storage/storageTotal*100):0;
  document.getElementById('sk-s-teams').textContent=teams;
  document.getElementById('sk-s-teams-sub').textContent=teamsI?teamsI+' inactive':'';
  document.getElementById('sk-s-guests').textContent=guests;
  document.getElementById('sk-s-anon').textContent=anon;
  document.getElementById('sk-s-storage').textContent=storage+'GB';
  document.getElementById('sk-s-storage-sub').textContent=storageTotal?'of '+storageTotal+'GB ('+storagePct+'%)':'';
  var bar=document.getElementById('sk-bar-storage');bar.style.width=storagePct+'%';bar.className='rc-bar-fill'+(storagePct>85?' crit':storagePct>70?' warn':'');
  document.getElementById('sk-teams-breakdown').innerHTML=[
    {label:'Active Teams',val:teams,color:'var(--ok)'},{label:'Inactive Teams',val:teamsI,color:'var(--warn)'},
    {label:'M365 Groups',val:groups,color:'var(--muted)'},{label:'Orphaned (no owner)',val:orphaned,color:'#b91c1c'},
  ].map(function(r){return'<div class="rp-row"><span class="rp-label">'+r.label+'</span><span class="rp-val" style="color:'+r.color+'">'+r.val+'</span></div>';}).join('');
  document.getElementById('sk-users-breakdown').innerHTML=[
    {label:'Licensed users',val:users},{label:'Guest users',val:guests,color:'var(--warn)'},
    {label:'Inactive (90+ days)',val:inactive,color:'var(--accent)'},{label:'Shared mailboxes',val:shared},
  ].map(function(r){return'<div class="rp-row"><span class="rp-label">'+r.label+'</span><span class="rp-val"'+(r.color?' style="color:'+r.color+'"':'')+'>'+r.val+'</span></div>';}).join('');
  document.getElementById('sk-sharing-breakdown').innerHTML=[
    {label:'Anonymous sharing links',val:anon,color:anon>0?'#b91c1c':'var(--ok)',note:'Anyone with link'},
    {label:'External sharing links',val:ext,color:ext>0?'var(--warn)':'var(--ok)',note:'Outside organisation'},
    {label:'Storage used',val:storage+'GB / '+(storageTotal||'?')+'GB',color:'var(--ink)'},
    {label:'Orphaned workspaces',val:orphaned,color:orphaned>0?'var(--accent)':'var(--ok)',note:'No active owner'},
  ].map(function(r){return'<div class="rp-row"><span class="rp-label">'+r.label+(r.note?'<span style="color:var(--muted);font-size:10px;margin-left:6px;">'+r.note+'</span>':'')+'</span><span class="rp-val" style="color:'+r.color+'">'+r.val+'</span></div>';}).join('');
}

function saveSyskitManual(){
  var ids=['sk-teams','sk-teams-inactive','sk-groups','sk-orphaned','sk-users','sk-guests','sk-inactive','sk-shared','sk-anon','sk-ext','sk-storage','sk-storage-total'];
  var d={saved:new Date().toISOString()};
  ids.forEach(function(id){d[id]=document.getElementById(id).value;});
  localStorage.setItem('itdesk_sk_manual',JSON.stringify(d));
  var msg=document.getElementById('sk-saved');msg.style.display='inline';setTimeout(function(){msg.style.display='none';},2500);
}

function loadSyskitSaved(){
  var d=JSON.parse(localStorage.getItem('itdesk_sk_manual')||'null');if(!d)return;
  ['sk-teams','sk-teams-inactive','sk-groups','sk-orphaned','sk-users','sk-guests','sk-inactive','sk-shared','sk-anon','sk-ext','sk-storage','sk-storage-total'].forEach(function(id){if(d[id]!==undefined)document.getElementById(id).value=d[id];});
  renderSyskitManual();
}

/* ═══════════ REPORT EXPORT ═══════════ */
// Prefix fields starting with formula chars to prevent CSV injection in Excel/Sheets
function csvSafe(s){s=String(s||'');return /^[=+\-@\t\r]/.test(s)?'\t'+s:s;}

function toggleExportMenu(){var m=document.getElementById('export-menu');m.style.display=m.style.display==='none'?'block':'none';}
document.addEventListener('click',function(e){var w=document.getElementById('export-wrap');if(w&&!w.contains(e.target)){var m=document.getElementById('export-menu');if(m)m.style.display='none';}});

function getReportDate(){return new Date().toLocaleDateString('en-NZ',{day:'numeric',month:'long',year:'numeric'});}

function getDefenderNums(){return{high:parseInt(document.getElementById('def-s-high').textContent)||0,med:parseInt(document.getElementById('def-s-med').textContent)||0,low:parseInt(document.getElementById('def-s-low').textContent)||0,info:parseInt(document.getElementById('def-s-info').textContent)||0};}
function getPhishingNums(){return{campaign:document.getElementById('ph-campaign').value||'—',sent:parseInt(document.getElementById('ph-sent').value)||0,opened:parseInt(document.getElementById('ph-opened').value)||0,clicked:parseInt(document.getElementById('ph-clicked').value)||0,submitted:parseInt(document.getElementById('ph-submitted').value)||0,reported:parseInt(document.getElementById('ph-reported').value)||0,trained:parseInt(document.getElementById('ph-trained').value)||0,atrisk:parseInt(document.getElementById('ph-atrisk').value)||0};}
function getSyskitNums(){return{teams:parseInt(document.getElementById('sk-teams').value)||0,teamsInactive:parseInt(document.getElementById('sk-teams-inactive').value)||0,groups:parseInt(document.getElementById('sk-groups').value)||0,orphaned:parseInt(document.getElementById('sk-orphaned').value)||0,users:parseInt(document.getElementById('sk-users').value)||0,guests:parseInt(document.getElementById('sk-guests').value)||0,inactive:parseInt(document.getElementById('sk-inactive').value)||0,shared:parseInt(document.getElementById('sk-shared').value)||0,anon:parseInt(document.getElementById('sk-anon').value)||0,ext:parseInt(document.getElementById('sk-ext').value)||0,storage:parseFloat(document.getElementById('sk-storage').value)||0,storageTotal:parseFloat(document.getElementById('sk-storage-total').value)||0};}

/* ── PDF ── */
function exportReportPDF(scope){
  document.getElementById('export-menu').style.display='none';
  var sections=scope==='active'?[activeReport]:['tickets','defender','phishing','syskit'];
  var parts=[];sections.forEach(function(s){if(s==='tickets')parts.push(pdfTickets());if(s==='defender')parts.push(pdfDefender());if(s==='phishing')parts.push(pdfPhishing());if(s==='syskit')parts.push(pdfSyskit());});
  var date=getReportDate();var org='Te Rito Maioha';
  var names=sections.map(function(s){return{tickets:'Helpdesk Tickets',defender:'Microsoft Defender',phishing:'Phriendly Phishing',syskit:'Syskit M365 Governance'}[s];});
  var css='*{box-sizing:border-box;margin:0;padding:0;}body{font-family:Georgia,serif;font-size:12pt;color:#111;}@page{size:A4;margin:18mm 16mm;}.cover{padding:28mm 0 18mm;border-bottom:3px solid #c94f2c;margin-bottom:10mm;}.org{font-size:12pt;color:#c94f2c;font-weight:bold;font-family:monospace;text-transform:uppercase;letter-spacing:.08em;margin-bottom:5pt;}h1{font-size:24pt;font-weight:normal;letter-spacing:-.02em;margin-bottom:7pt;}.scope{font-size:11pt;color:#555;line-height:1.8;margin-bottom:8pt;}.ddate{font-size:10pt;color:#999;font-family:monospace;}.section{padding:7mm 0;}.section+.section{border-top:2px solid #eee;margin-top:5mm;padding-top:7mm;}h2{font-size:14pt;font-weight:normal;color:#c94f2c;margin-bottom:4mm;}h3{font-size:9pt;text-transform:uppercase;letter-spacing:.1em;color:#888;margin:4mm 0 2mm;font-family:monospace;}.krow{display:flex;gap:4mm;margin-bottom:4mm;}.kpi{flex:1;background:#f9f7f3;border:1px solid #e0dbd4;border-left:3px solid #c94f2c;padding:3mm 4mm;border-radius:2px;}.kpi.ok{border-left-color:#2a7c5f;}.kpi.warn{border-left-color:#c47d1a;}.kpi.crit{border-left-color:#b91c1c;}.kv{font-size:18pt;font-weight:bold;line-height:1;}.kl{font-size:8pt;text-transform:uppercase;letter-spacing:.06em;color:#888;margin-top:1mm;font-family:monospace;}table{width:100%;border-collapse:collapse;font-size:10pt;margin-bottom:3mm;}th{font-size:8pt;text-transform:uppercase;letter-spacing:.06em;color:#888;text-align:left;padding:1.5mm 2.5mm;border-bottom:1.5px solid #333;font-family:monospace;font-weight:normal;}td{padding:2mm 2.5mm;border-bottom:1px solid #e8e3d8;vertical-align:top;}tr:last-child td{border-bottom:none;}.ok{color:#2a7c5f;}.warn{color:#c47d1a;}.crit{color:#b91c1c;}.bar-bg{background:#e8e3d8;height:5px;border-radius:3px;overflow:hidden;}.bar-fill{height:100%;border-radius:3px;background:#2a7c5f;}.footer{margin-top:7mm;padding-top:3mm;border-top:1px solid #ddd;font-size:8pt;color:#aaa;font-family:monospace;display:flex;justify-content:space-between;}';
  var html='<!DOCTYPE html><html><head><meta charset="UTF-8"/><title>IT Report — '+org+' — '+date+'</title><style>'+css+'</style></head><body>'+
    '<div class="cover"><div class="org">'+org+'</div><h1>IT Report</h1><div class="scope">'+names.join(' &nbsp;·&nbsp; ')+'</div><div class="ddate">Prepared '+date+'</div></div>'+
    parts.join('')+
    '<div class="footer"><span>'+org+' — Confidential</span><span>'+date+'</span></div>'+
    '</body></html>';
  var win=window.open('','_blank');win.document.write(html);win.document.close();win.focus();setTimeout(function(){win.print();},700);
}

function pdfKpi(val,label,cls){return'<div class="kpi '+(cls||'')+'"><div class="kv">'+val+'</div><div class="kl">'+label+'</div></div>';}

function pdfTickets(){
  var days=parseInt(document.getElementById('rpt-range').value)||30;
  var now=Date.now();var cutoff=days===0?0:now-days*86400000;
  var src=(tickets&&tickets.length)?tickets:(typeof DEMO!=='undefined'?DEMO:[]);
  var t=src.filter(function(x){return new Date(x.created).getTime()>=cutoff&&!x._archived;});
  var total=t.length,resolved=t.filter(function(x){return x.status==='resolved'||x.status==='closed';}).length,open=t.filter(function(x){return x.status==='open';}).length;
  var s=getStoredSettings();var slaMap={critical:(s.slaCrit||1)*3600000,high:(s.slaHigh||4)*3600000,medium:(s.slaMed||8)*3600000,low:(s.slaLow||24)*3600000};
  var metSla=0;t.forEach(function(x){if(x.status==='on-hold'||now-new Date(x.created).getTime()<=(slaMap[x.priority]||slaMap.low))metSla++;});
  var slaRate=total?Math.round(metSla/total*100):100;
  var cats={};t.forEach(function(x){cats[x.cat||'Uncategorised']=(cats[x.cat||'Uncategorised']||0)+1;});
  var catRows=Object.entries(cats).sort(function(a,b){return b[1]-a[1];}).map(function(e){var pct=total?Math.round(e[1]/total*100):0;return'<tr><td>'+esc(e[0])+'</td><td style="width:40%;"><div class="bar-bg"><div class="bar-fill" style="width:'+pct+'%"></div></div></td><td>'+e[1]+' <span style="color:#999;">('+pct+'%)</span></td></tr>';}).join('');
  return'<div class="section"><h2>Helpdesk Tickets <span style="font-size:10pt;color:#999;">'+(days===0?'All time':'Last '+days+' days')+'</span></h2>'+
    '<div class="krow">'+pdfKpi(total,'Total')+pdfKpi(resolved,'Resolved','ok')+pdfKpi(slaRate+'%','SLA compliance',slaRate>=85?'ok':slaRate>=70?'warn':'crit')+pdfKpi(open,'Open now')+'</div>'+
    '<h3>By category</h3><table><thead><tr><th>Category</th><th>Volume</th><th>Count</th></tr></thead><tbody>'+catRows+'</tbody></table></div>';
}

function pdfDefender(){
  var d=getDefenderNums();var total=d.high+d.med+d.low+d.info;
  return'<div class="section"><h2>Microsoft Defender</h2>'+
    '<div class="krow">'+pdfKpi(d.high,'High','crit')+pdfKpi(d.med,'Medium','warn')+pdfKpi(d.low,'Low','ok')+pdfKpi(d.info,'Info')+'</div>'+
    '<table><thead><tr><th>Severity</th><th>Alerts</th><th>Status</th></tr></thead><tbody>'+
    '<tr><td>High</td><td>'+d.high+'</td><td class="'+(d.high>0?'crit':'ok')+'">'+(d.high>0?'Action required':'Clear')+'</td></tr>'+
    '<tr><td>Medium</td><td>'+d.med+'</td><td class="'+(d.med>0?'warn':'ok')+'">'+(d.med>0?'Review needed':'Clear')+'</td></tr>'+
    '<tr><td>Low</td><td>'+d.low+'</td><td class="ok">'+(d.low>0?'Monitor':'Clear')+'</td></tr>'+
    '<tr><td>Informational</td><td>'+d.info+'</td><td style="color:#999;">Awareness</td></tr>'+
    '<tr><td><strong>Total</strong></td><td><strong>'+total+'</strong></td><td></td></tr>'+
    '</tbody></table></div>';
}

function pdfPhishing(){
  var d=getPhishingNums();var sent=d.sent;
  var clickRate=sent?Math.round(d.clicked/sent*100):0;var reportRate=sent?Math.round(d.reported/sent*100):0;var trainedPct=sent?Math.round(d.trained/sent*100):0;
  return'<div class="section"><h2>Phriendly Phishing <span style="font-size:10pt;color:#999;">'+d.campaign+'</span></h2>'+
    '<div class="krow">'+pdfKpi(clickRate+'%','Click rate',clickRate<=5?'ok':clickRate<=15?'warn':'crit')+pdfKpi(reportRate+'%','Report rate',reportRate>=20?'ok':'warn')+pdfKpi(trainedPct+'%','Training done',trainedPct>=80?'ok':trainedPct>=50?'warn':'crit')+pdfKpi(d.atrisk+'%','At risk',d.atrisk<=10?'ok':d.atrisk<=25?'warn':'crit')+'</div>'+
    '<table><thead><tr><th>Metric</th><th>Count</th><th>% of sent</th></tr></thead><tbody>'+
    '<tr><td>Emails sent</td><td>'+sent+'</td><td>—</td></tr>'+
    '<tr><td>Opened email</td><td>'+d.opened+'</td><td>'+(sent?Math.round(d.opened/sent*100):0)+'%</td></tr>'+
    '<tr><td>Clicked link</td><td>'+d.clicked+'</td><td class="'+(clickRate<=5?'ok':clickRate<=15?'warn':'crit')+'">'+clickRate+'%</td></tr>'+
    '<tr><td>Submitted data</td><td>'+d.submitted+'</td><td class="'+(d.submitted>0?'crit':'ok')+'">'+(sent?Math.round(d.submitted/sent*100):0)+'%</td></tr>'+
    '<tr><td>Reported phish</td><td>'+d.reported+'</td><td class="ok">'+reportRate+'%</td></tr>'+
    '<tr><td>Training completed</td><td>'+d.trained+'</td><td class="'+(trainedPct>=80?'ok':trainedPct>=50?'warn':'crit')+'">'+trainedPct+'%</td></tr>'+
    '</tbody></table></div>';
}

function pdfSyskit(){
  var d=getSyskitNums();var storagePct=d.storageTotal?Math.round(d.storage/d.storageTotal*100):0;
  return'<div class="section"><h2>Syskit — M365 Governance</h2>'+
    '<div class="krow">'+pdfKpi(d.teams,'Active Teams')+pdfKpi(d.guests,'Guest users',d.guests>0?'warn':'ok')+pdfKpi(d.anon,'Anon links',d.anon>0?'crit':'ok')+pdfKpi(storagePct+'%','Storage used',storagePct>85?'crit':storagePct>70?'warn':'ok')+'</div>'+
    '<table><thead><tr><th>Area</th><th>Metric</th><th>Value</th><th>Status</th></tr></thead><tbody>'+
    '<tr><td>Teams</td><td>Active</td><td>'+d.teams+'</td><td class="ok">Good</td></tr>'+
    '<tr><td></td><td>Inactive</td><td>'+d.teamsInactive+'</td><td class="'+(d.teamsInactive>0?'warn':'ok')+'">'+(d.teamsInactive>0?'Review':'Clear')+'</td></tr>'+
    '<tr><td></td><td>Orphaned</td><td>'+d.orphaned+'</td><td class="'+(d.orphaned>0?'crit':'ok')+'">'+(d.orphaned>0?'Action required':'Clear')+'</td></tr>'+
    '<tr><td>Users</td><td>Licensed</td><td>'+d.users+'</td><td>—</td></tr>'+
    '<tr><td></td><td>Guests</td><td>'+d.guests+'</td><td class="'+(d.guests>0?'warn':'ok')+'">'+(d.guests>0?'Monitor':'Clear')+'</td></tr>'+
    '<tr><td></td><td>Inactive (90+ days)</td><td>'+d.inactive+'</td><td class="'+(d.inactive>0?'warn':'ok')+'">'+(d.inactive>0?'Review':'Clear')+'</td></tr>'+
    '<tr><td>Sharing</td><td>Anonymous links</td><td>'+d.anon+'</td><td class="'+(d.anon>0?'crit':'ok')+'">'+(d.anon>0?'Action required':'Clear')+'</td></tr>'+
    '<tr><td></td><td>External links</td><td>'+d.ext+'</td><td class="'+(d.ext>0?'warn':'ok')+'">'+(d.ext>0?'Monitor':'Clear')+'</td></tr>'+
    '<tr><td>Storage</td><td>Used / Total</td><td>'+d.storage+'GB / '+(d.storageTotal||'?')+'GB</td><td class="'+(storagePct>85?'crit':storagePct>70?'warn':'ok')+'">'+storagePct+'%</td></tr>'+
    '</tbody></table></div>';
}

/* ── CSV ── */
function exportReportCSV(scope){
  document.getElementById('export-menu').style.display='none';
  var sections=scope==='active'?[activeReport]:['tickets','defender','phishing','syskit'];
  var rows=[];var date=getReportDate();
  rows.push(['Report Date',date]);rows.push(['Organisation','Te Rito Maioha']);rows.push([]);
  sections.forEach(function(s){
    if(s==='tickets'){
      var days=parseInt(document.getElementById('rpt-range').value)||30;
      var now=Date.now();var cutoff=days===0?0:now-days*86400000;
      var src=(tickets&&tickets.length)?tickets:(typeof DEMO!=='undefined'?DEMO:[]);
      var t=src.filter(function(x){return new Date(x.created).getTime()>=cutoff&&!x._archived;});
      rows.push(['=== HELPDESK TICKETS (Last '+days+' days) ===']);
      rows.push(['ID','Subject','Requester','Email','Dept','Category','Priority','Status','Assignee','Submitted']);
      t.forEach(function(x){rows.push([csvSafe(x.id),'"'+csvSafe(x.subject||'').replace(/"/g,'""')+'"','"'+csvSafe(x.name||'').replace(/"/g,'""')+'"',csvSafe(x.email||''),csvSafe(x.dept||''),csvSafe(x.cat||''),csvSafe(x.priority),csvSafe(x.status),csvSafe(x.assignee||'Unassigned'),csvSafe(new Date(x.created).toLocaleDateString('en-NZ'))]);});
      rows.push([]);
    }
    if(s==='defender'){
      var defData=getDefenderNums();
      rows.push(['=== MICROSOFT DEFENDER ===']);rows.push(['Severity','Active alerts','Status']);
      rows.push(['High',defData.high,defData.high>0?'Action required':'Clear']);rows.push(['Medium',defData.med,defData.med>0?'Review needed':'Clear']);
      rows.push(['Low',defData.low,defData.low>0?'Monitor':'Clear']);rows.push(['Informational',defData.info,'Awareness']);
      rows.push(['Total',defData.high+defData.med+defData.low+defData.info,'']);rows.push([]);
    }
    if(s==='phishing'){
      var phData=getPhishingNums();var sent=phData.sent;
      rows.push(['=== PHRIENDLY PHISHING — '+phData.campaign+' ===']);rows.push(['Metric','Count','% of sent']);
      rows.push(['Emails sent',sent,'100%']);rows.push(['Opened email',phData.opened,(sent?Math.round(phData.opened/sent*100):0)+'%']);
      rows.push(['Clicked link',phData.clicked,(sent?Math.round(phData.clicked/sent*100):0)+'%']);rows.push(['Submitted data',phData.submitted,(sent?Math.round(phData.submitted/sent*100):0)+'%']);
      rows.push(['Reported phish',phData.reported,(sent?Math.round(phData.reported/sent*100):0)+'%']);rows.push(['Training completed',phData.trained,(sent?Math.round(phData.trained/sent*100):0)+'%']);
      rows.push(['At risk (%)',phData.atrisk+'%','']);rows.push([]);
    }
    if(s==='syskit'){
      var skData=getSyskitNums();
      rows.push(['=== SYSKIT — M365 GOVERNANCE ===']);rows.push(['Metric','Value']);
      rows.push(['Active Teams',skData.teams]);rows.push(['Inactive Teams',skData.teamsInactive]);rows.push(['M365 Groups',skData.groups]);rows.push(['Orphaned workspaces',skData.orphaned]);
      rows.push(['Licensed users',skData.users]);rows.push(['Guest users',skData.guests]);rows.push(['Inactive users (90+ days)',skData.inactive]);rows.push(['Shared mailboxes',skData.shared]);
      rows.push(['Anonymous sharing links',skData.anon]);rows.push(['External sharing links',skData.ext]);rows.push(['Storage used (GB)',skData.storage]);rows.push(['Storage total (GB)',skData.storageTotal]);
      rows.push([]);
    }
  });
  var csv=rows.map(function(r){return r.map(function(c){var sv=String(c===null||c===undefined?'':c);return sv.includes(',')||sv.includes('"')||sv.includes('\n')?'"'+sv.replace(/"/g,'""')+'"':sv;}).join(',');}).join('\r\n');
  var blob=new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8;'});
  var a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download='IT-Report-'+(scope==='active'?activeReport:'all-sections')+'-'+new Date().toISOString().slice(0,10)+'.csv';
  document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(a.href);
}

function exportReport(){exportReportCSV('active');}
function printReport(){exportReportPDF('active');}
function exportCSV(archive){
  var rows=archive?tickets.filter(function(t){return t._archived;}):tickets.filter(function(t){return !t._archived;});
  var cols=['ID','Subject','Requester','Email','Department','Category','Priority','Status','Assignee','Created','Closed','SLA','Notes'];
  var csv=[cols.join(',')];
  rows.forEach(function(t){var info=getSLAInfo(t);csv.push([csvSafe(t.id),'"'+csvSafe(t.subject||'').replace(/"/g,'""')+'"','"'+csvSafe(t.name||'').replace(/"/g,'""')+'"',csvSafe(t.email||''),csvSafe(t.dept||''),csvSafe(t.cat||''),csvSafe(t.priority),csvSafe(t.status),csvSafe(t.assignee||''),csvSafe(t.created||''),csvSafe(t.closed||''),csvSafe(info.label),'"'+csvSafe(t.notes||'').replace(/"/g,'""')+'"'].join(','));});
  var blob=new Blob(['\uFEFF'+csv.join('\n')],{type:'text/csv;charset=utf-8'});
  var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=(archive?'archive':'tickets')+'-'+new Date().toISOString().slice(0,10)+'.csv';
  document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(a.href);
}

/* ═══════════ PANEL ═══════════ */
function getTicket(id){for(var i=0;i<tickets.length;i++){if(tickets[i].id===id)return tickets[i];}return null;}

function openPanel(id){
  openId=id;var t=getTicket(id);if(!t)return;
  clearUnseen(id);
  document.getElementById('p-id').textContent   =t.id;
  document.getElementById('p-title').textContent=t.subject;
  document.getElementById('p-body').textContent =t.desc;
  document.getElementById('p-name').textContent =t.name;
  document.getElementById('p-email').textContent=t.email;
  document.getElementById('p-dept').textContent =t.dept;
  document.getElementById('p-notes').value      =t.notes||'';
  document.getElementById('p-reply').value      ='';
  document.getElementById('reply-status').textContent='';
  document.getElementById('p-priority-sel').value=t.priority||'low';

  // category dropdown — include ticket's current category even if not in list
  var catSel=document.getElementById('p-cat-sel');
  var cats=CATEGORIES.slice();
  if(t.cat&&cats.indexOf(t.cat)===-1)cats.unshift(t.cat);
  catSel.innerHTML=cats.map(function(c){return'<option value="'+esc(c)+'"'+(t.cat===c?' selected':'')+'>'+esc(c)+'</option>';}).join('');

  // watchers
  renderWatcherTags(t);
  document.getElementById('watcher-input').value='';

  var sel=document.getElementById('p-status-sel');
  sel.innerHTML=[['open','Open'],['in-progress','In Progress'],['on-hold','On Hold'],['closed','Closed']].map(function(s){return'<option value="'+s[0]+'"'+(t.status===s[0]?' selected':'')+'>'+s[1]+'</option>';}).join('');

  var agents=getStoredAgents();
  var me=(currentUser?currentUser.username:'').toLowerCase();
  var asel=document.getElementById('p-asg-sel');
  asel.innerHTML='<option value="">Unassigned</option>'+agents.map(function(a){var lbl=a.email.toLowerCase()===me?'Me ('+a.email+')':a.email;return'<option value="'+esc(a.email)+'"'+(t.assignee&&t.assignee.toLowerCase()===a.email.toLowerCase()?' selected':'')+'>'+esc(lbl)+'</option>';}).join('');

  document.getElementById('pickup-btn').textContent=(t.assignee&&t.assignee.toLowerCase()===me)?'\u2713 Assigned to me':'\u261D Pick up';
  document.getElementById('del-btn').style.display=currentRole==='admin'?'':'none';
  document.getElementById('btn-merge').style.display=currentRole==='admin'?'':'none';

  var thread=document.getElementById('reply-thread');thread.innerHTML='';
  var replies=t.replies||[];
  if(!replies.length){thread.innerHTML='<div style="font-size:12px;color:var(--muted)">No replies yet.</div>';}
  else{replies.forEach(function(r){var d=document.createElement('div');d.className='rb'+(r.from==='agent'?' sent':'');d.innerHTML='<div class="rb-meta">'+(r.from==='agent'?'You (IT Support)':'Staff member')+' \u2014 '+fmtTime(r.ts)+'</div>'+esc(r.text);thread.appendChild(d);});}

  updateCycleBtn(t.status);updatePanelSLA(t);
  document.getElementById('panel-ov').classList.add('open');
  document.getElementById('panel').classList.add('open');
}

function closePanel(){document.getElementById('panel-ov').classList.remove('open');document.getElementById('panel').classList.remove('open');openId=null;}
function updateCycleBtn(s){var m={open:'Mark in progress','in-progress':'Put on hold','on-hold':'Resume (in progress)',closed:'Reopen'};document.getElementById('cycle-btn').textContent=m[s]||'Mark in progress';}

function showPersistError(){
  var existing=document.getElementById('persist-err-banner');
  if(existing)return;
  var banner=document.createElement('div');
  banner.id='persist-err-banner';
  banner.style.cssText='position:fixed;bottom:20px;left:50%;transform:translateX(-50%);background:#fee2e2;border:1.5px solid #fca5a5;color:#b91c1c;font-family:"DM Mono",monospace;font-size:12px;padding:10px 18px;border-radius:4px;z-index:200;box-shadow:0 2px 12px rgba(0,0,0,.15);';
  banner.textContent='\u26A0 Changes may not have saved — check your connection and refresh.';
  var close=document.createElement('button');
  close.textContent='\u2715';
  close.style.cssText='margin-left:12px;background:none;border:none;cursor:pointer;color:#b91c1c;font-size:14px;';
  close.onclick=function(){banner.remove();};
  banner.appendChild(close);
  document.body.appendChild(banner);
  setTimeout(function(){if(banner.parentNode)banner.remove();},8000);
}

function persist(t){
  persistToGraph(t).catch(function(e){console.error('Persist error:',e);showPersistError();});
  renderDash();
}

function updatePriority(){var t=getTicket(openId);if(!t)return;t.priority=document.getElementById('p-priority-sel').value;updatePanelSLA(t);persist(t);}

function applyStatusChange(t,newStatus){
  var old=t.status;
  if(old==='on-hold'&&newStatus!=='on-hold'&&t.holdStart){t.holdDuration=(t.holdDuration||0)+(Date.now()-new Date(t.holdStart).getTime());t.holdStart=null;}
  if(newStatus==='on-hold'&&old!=='on-hold')t.holdStart=new Date().toISOString();
  if(newStatus==='closed'&&!t.closed)t.closed=new Date().toISOString();
  t.status=newStatus;
}

function updateStatus(){var t=getTicket(openId);if(!t)return;applyStatusChange(t,document.getElementById('p-status-sel').value);updateCycleBtn(t.status);updatePanelSLA(t);persist(t);}

function cycleStatus(){
  var t=getTicket(openId);if(!t)return;
  var cycle={open:'in-progress','in-progress':'on-hold','on-hold':'in-progress',closed:'open'};
  var next=cycle[t.status]||'in-progress';
  applyStatusChange(t,next);
  document.getElementById('p-status-sel').value=t.status;
  updateCycleBtn(t.status);updatePanelSLA(t);persist(t);
}

function updateAssignee(){var t=getTicket(openId);if(!t)return;t.assignee=document.getElementById('p-asg-sel').value;var me=(currentUser?currentUser.username:'').toLowerCase();document.getElementById('pickup-btn').textContent=(t.assignee&&t.assignee.toLowerCase()===me)?'\u2713 Assigned to me':'\u261D Pick up';persist(t);}

function pickupTicket(){
  var t=getTicket(openId);if(!t||!currentUser)return;
  t.assignee=currentUser.username;
  if(t.status==='open'){applyStatusChange(t,'in-progress');document.getElementById('p-status-sel').value='in-progress';updateCycleBtn('in-progress');updatePanelSLA(t);}
  document.getElementById('p-asg-sel').value=currentUser.username;
  document.getElementById('pickup-btn').textContent='\u2713 Assigned to me';
  persist(t);
}

function saveNote(e){var t=getTicket(openId);if(!t)return;t.notes=document.getElementById('p-notes').value;persist(t);var b=e.currentTarget;b.textContent='\u2713 Saved';setTimeout(function(){b.textContent='Save notes';},1500);}

function deleteTicket(){
  if(currentRole!=='admin'){alert('Only administrators can delete tickets.');return;}
  if(!confirm('Delete this ticket? Cannot be undone.'))return;
  var t=getTicket(openId);
  if(!t)return;
  var btn=document.getElementById('del-btn');
  btn.disabled=true;btn.textContent='Deleting\u2026';
  graphFetch(LIST_URL+'/items/'+t.spId,'DELETE')
  .then(function(r){
    if(!r.ok&&r.status!==204)throw new Error('Delete failed: '+r.status);
    tickets=tickets.filter(function(x){return x.id!==openId;});
    closePanel();renderDash();
  })
  .catch(function(e){
    console.error('Delete error:',e);
    btn.disabled=false;btn.textContent='Delete';
    alert('Could not delete ticket from SharePoint — please try again.');
  });
}

function sendReply(){
  var t=getTicket(openId);if(!t)return;
  var text=document.getElementById('p-reply').value.trim();
  if(!text){document.getElementById('reply-status').textContent='Please type a reply first.';return;}
  var btn=document.getElementById('send-btn');btn.disabled=true;btn.textContent='Sending\u2026';
  var r={from:'agent',text:text,ts:new Date().toISOString()};
  if(FLOW_REPLY){
    fetch(FLOW_REPLY,{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({spId:t.spId,ticketId:t.id,subject:'Re: ['+t.id+'] '+t.subject,toEmail:t.email,toName:t.name,replyText:text,ccEmails:getWatchers(t).join(',')})
    })
    .then(function(res){if(!res.ok)throw new Error('Flow returned '+res.status);t.replies=t.replies||[];t.replies.push(r);persist(t);openPanel(t.id);document.getElementById('reply-status').textContent='\u2713 Reply sent';})
    .catch(function(e){console.error('Reply error:',e);document.getElementById('reply-status').textContent='\u26A0 Failed. Check Flow URL in Settings.';})
    .finally(function(){btn.disabled=false;btn.textContent='\u2709 Send reply';});
  } else {
    t.replies=t.replies||[];t.replies.push(r);persist(t);openPanel(t.id);
    document.getElementById('reply-status').textContent='\u2713 Saved (add Flow 4 URL in Settings to email staff)';
    btn.disabled=false;btn.textContent='\u2709 Send reply';
  }
}

/* ═══════════ SETTINGS ═══════════ */
function getStoredSettings(){try{return JSON.parse(localStorage.getItem('itdesk_settings')||'{}');}catch(e){return{};}}

function applyStoredSettings(){
  var s=getStoredSettings();
  if(s.flowReply) FLOW_REPLY =s.flowReply;
  if(s.flowSubmit)FLOW_SUBMIT=s.flowSubmit;
  if(s.autorefresh&&!window._refreshTimer)window._refreshTimer=setInterval(loadTicketsDebounced,300000);
}

/* ═══════════ AGENT MANAGEMENT — SHAREPOINT ═══════════ */
function loadAgentsFromGraph(forceRefresh){
  // return cached agents unless forced refresh
  if(cachedAgents && !forceRefresh) return Promise.resolve(cachedAgents);
  return fetchAllPages(AGENTS_URL+'/items?expand=fields&$top=200')
  .then(function(items){
    var agents=items.map(function(i){
      return{spId:i.id,email:(i.fields.Title||'').toLowerCase(),role:(i.fields.Role||'agent').toLowerCase().trim().replace('administrator','admin'),displayName:i.fields.DisplayName||''};
    });
    if(!agents.length) agents=DEFAULT_AGENTS.slice();
    cachedAgents=agents;
    return agents;
  })
  .catch(function(e){
    console.warn('Could not load agents from SharePoint, using defaults:',e);
    if(!cachedAgents) cachedAgents=DEFAULT_AGENTS.slice();
    return cachedAgents;
  });
}

function getStoredAgents(){
  return cachedAgents||DEFAULT_AGENTS.slice();
}

function loadSettingsUI(){
  // force refresh agents from SharePoint when opening settings
  loadAgentsFromGraph(true).then(function(agents){
    renderAgentList(agents);
  });
  var isAdmin=currentRole==='admin';
  ['s-sla-crit','s-sla-high','s-sla-med','s-sla-low'].forEach(function(id){document.getElementById(id).disabled=!isAdmin;});
  document.getElementById('agent-add-section').style.display=isAdmin?'block':'none';
  document.getElementById('agent-locked-msg').style.display=isAdmin?'none':'block';
  var s=getStoredSettings();
  if(s.slaCrit)      document.getElementById('s-sla-crit').value=s.slaCrit;
  if(s.slaHigh)      document.getElementById('s-sla-high').value=s.slaHigh;
  if(s.slaMed)       document.getElementById('s-sla-med').value=s.slaMed;
  if(s.slaLow)       document.getElementById('s-sla-low').value=s.slaLow;
  if(typeof s.demo==='boolean')        document.getElementById('s-demo').checked=s.demo;
  if(typeof s.autorefresh==='boolean') document.getElementById('s-autorefresh').checked=s.autorefresh;
  if(s.archivedays)  document.getElementById('s-archivedays').value=s.archivedays;
  if(s.flowReply)    document.getElementById('s-fl-reply').value=s.flowReply;
  if(s.flowSubmit)   document.getElementById('s-fl-submit').value=s.flowSubmit;
  if(s.flowPhishing) document.getElementById('s-fl-phishing').value=s.flowPhishing;
  if(s.flowSyskit)   document.getElementById('s-fl-syskit').value=s.flowSyskit;
}

function renderAgentList(agents){
  var list=document.getElementById('agent-list');list.innerHTML='';
  if(!agents.length){list.innerHTML='<div style="font-size:12px;color:var(--muted);">No agents added yet.</div>';return;}
  var me=(currentUser?currentUser.username:'').toLowerCase();
  var isAdmin=currentRole==='admin';
  agents.forEach(function(a){
    var isMe=a.email.toLowerCase()===me;
    var row=document.createElement('div');row.className='agent-row';

    var emailDiv=document.createElement('div');
    emailDiv.className='agent-email';
    emailDiv.style.flex='1';
    emailDiv.textContent=a.email;
    row.appendChild(emailDiv);

    if(isAdmin&&!isMe){
      var sel=document.createElement('select');
      sel.className='si';
      sel.style.cssText='font-size:11px;padding:4px 8px;';
      var optAgent=document.createElement('option');optAgent.value='agent';optAgent.textContent='Agent';if(a.role!=='admin')optAgent.selected=true;
      var optAdmin=document.createElement('option');optAdmin.value='admin';optAdmin.textContent='Admin';if(a.role==='admin')optAdmin.selected=true;
      sel.appendChild(optAgent);sel.appendChild(optAdmin);
      (function(email,spId){sel.addEventListener('change',function(){changeRole(email,spId,sel.value);});})(a.email,a.spId);
      row.appendChild(sel);

      var rmBtn=document.createElement('button');
      rmBtn.className='rm-btn';rmBtn.innerHTML='&#10005;';rmBtn.setAttribute('aria-label','Remove '+a.email);
      (function(email,spId){rmBtn.addEventListener('click',function(){removeAgent(email,spId);});})(a.email,a.spId);
      row.appendChild(rmBtn);
    } else {
      var aRole=(a.role||'agent').trim();
      var badge=document.createElement('span');
      badge.className='role-badge '+aRole+(isMe?' you':'');
      badge.textContent=(aRole==='admin'?'Admin':'Agent')+(isMe?' (you)':'');
      row.appendChild(badge);
    }

    list.appendChild(row);
  });
}

function changeRole(email,spId,role){
  if(currentRole!=='admin')return;
  graphFetch(AGENTS_URL+'/items/'+spId+'/fields','PATCH',{Role:role})
  .then(function(){
    var a=cachedAgents&&cachedAgents.find(function(x){return x.email===email;});
    if(a)a.role=role;
    renderAgentList(cachedAgents||[]);
  })
  .catch(function(e){alert('Could not update role: '+e.message);});
}

function addAgent(){
  if(currentRole!=='admin'){alert('Only administrators can add agents.');return;}
  var emailEl=document.getElementById('new-agent-email');
  var roleEl=document.getElementById('new-agent-role');
  var email=emailEl.value.trim().toLowerCase();
  var role=roleEl.value||'agent';
  var err=document.getElementById('agent-err');err.style.display='none';
  if(!email||!email.includes('@')){err.textContent='Please enter a valid email address.';err.style.display='block';return;}
  var agents=cachedAgents||[];
  if(agents.find(function(a){return a.email.toLowerCase()===email;})){err.textContent='That agent is already in the list.';err.style.display='block';return;}
  // write to SharePoint
  graphFetch(AGENTS_URL+'/items','POST',{fields:{Title:email,Role:role,DisplayName:email.split('@')[0]}})
  .then(function(r){if(!r.ok)throw new Error();return r.json();})
  .then(function(data){
    agents.push({spId:data.id,email:email,role:role,displayName:email.split('@')[0]});
    cachedAgents=agents;
    renderAgentList(agents);
    emailEl.value='';
  })
  .catch(function(e){err.textContent='Could not add agent: '+(e.message||'Please try again.');err.style.display='block';});
}

function removeAgent(email,spId){
  if(currentRole!=='admin'){alert('Only administrators can remove agents.');return;}
  if(!confirm('Remove '+email+' from agent access?'))return;
  graphFetch(AGENTS_URL+'/items/'+spId,'DELETE')
  .then(function(){
    cachedAgents=(cachedAgents||[]).filter(function(a){return a.email!==email;});
    renderAgentList(cachedAgents);
  })
  .catch(function(e){alert('Could not remove agent: '+(e.message||'Please try again.'));});
}

function saveSettings(){
  var s=getStoredSettings();
  var err=document.getElementById('settings-err');err.style.display='none';err.textContent='';
  if(currentRole==='admin'){s.slaCrit=parseInt(document.getElementById('s-sla-crit').value)||1;s.slaHigh=parseInt(document.getElementById('s-sla-high').value)||4;s.slaMed=parseInt(document.getElementById('s-sla-med').value)||8;s.slaLow=parseInt(document.getElementById('s-sla-low').value)||24;}
  s.demo       =document.getElementById('s-demo').checked;
  s.autorefresh=document.getElementById('s-autorefresh').checked;
  s.archivedays=parseInt(document.getElementById('s-archivedays').value)||7;
  var replyUrl=document.getElementById('s-fl-reply').value.trim();
  var submitUrl=document.getElementById('s-fl-submit').value.trim();
  if(replyUrl&&!replyUrl.startsWith('https://')){err.textContent='Reply Flow URL must start with https://';err.style.display='block';return;}
  if(submitUrl&&!submitUrl.startsWith('https://')){err.textContent='Submit Flow URL must start with https://';err.style.display='block';return;}
  s.flowReply  =replyUrl;
  s.flowSubmit =submitUrl;
  s.flowPhishing=(document.getElementById('s-fl-phishing').value||'').trim();
  s.flowSyskit  =(document.getElementById('s-fl-syskit').value||'').trim();
  localStorage.setItem('itdesk_settings',JSON.stringify(s));
  if(s.flowReply) FLOW_REPLY =s.flowReply;
  if(s.flowSubmit)FLOW_SUBMIT=s.flowSubmit;
  if(s.autorefresh){if(!window._refreshTimer)window._refreshTimer=setInterval(loadTicketsDebounced,300000);}
  else{clearInterval(window._refreshTimer);window._refreshTimer=null;}
  var msg=document.getElementById('saved-msg');msg.style.display='inline';setTimeout(function(){msg.style.display='none';},2500);
}

/* ═══════════ DEMO DATA ═══════════ */
var DEMO=[
  {id:'INC-0001',spId:'1',name:'Aroha Tuohoe', email:'aroha@example.com',  dept:'Academic',       cat:'Email / Outlook',       subject:'Cannot send emails — relay error', desc:'Error 5.7.57 externally. Started after password change.',  priority:'high',    status:'open',       assignee:'',notes:'',replies:[],holdDuration:0,holdStart:null,created:new Date(Date.now()-172800000).toISOString(),closed:null,_archived:false},
  {id:'INC-0002',spId:'2',name:'James Piripi', email:'james@example.com',  dept:'Administration', cat:'Device / Hardware',     subject:'Laptop will not power on',        desc:'Unresponsive. Was working Friday. Adapter fine.',         priority:'critical',status:'in-progress',assignee:'agent@example.com',notes:'Battery ordered.',replies:[{from:'agent',text:'Battery ordered, arriving Thursday.',ts:new Date(Date.now()-86400000).toISOString()}],holdDuration:0,holdStart:null,created:new Date(Date.now()-90000000).toISOString(),closed:null,_archived:false},
  {id:'INC-0003',spId:'3',name:'Mere Whaanga', email:'mere@example.com',   dept:'Finance',        cat:'Account / Access',      subject:'Locked out of SharePoint',        desc:'Access denied on Finance intranet. Had access last week.', priority:'medium',  status:'on-hold',    assignee:'',notes:'Waiting on manager approval.',replies:[],holdDuration:0,holdStart:new Date(Date.now()-3600000).toISOString(),created:new Date(Date.now()-36000000).toISOString(),closed:null,_archived:false},
  {id:'INC-0004',spId:'4',name:'Hemi Tuohoe',  email:'hemi@example.com',   dept:'HR',             cat:'Microsoft 365 / Teams', subject:'Teams calls drop after 2 mins',   desc:'Video calls disconnect. Audio calls fine.',               priority:'medium',  status:'closed',     assignee:'agent@example.com',notes:'Cleared Teams cache.',replies:[],holdDuration:0,holdStart:null,created:new Date(Date.now()-345600000).toISOString(),closed:new Date(Date.now()-86400000).toISOString(),_archived:false},
  {id:'INC-0005',spId:'5',name:'Ngahuia Rei',  email:'ngahuia@example.com',dept:'Operations',     cat:'Printer / Peripheral',  subject:'Printer on Level 2 offline',      desc:'HP LaserJet Pro offline for everyone on L2.',             priority:'low',     status:'open',       assignee:'',notes:'',replies:[],holdDuration:0,holdStart:null,created:new Date(Date.now()-7200000).toISOString(),closed:null,_archived:false}
];

/* ═══════════ EVENT BINDINGS ═══════════ */
function bindEvents(){
  // Nav
  document.getElementById('nb-tickets').addEventListener('click',function(){showPage('tickets');});
  document.getElementById('nb-archive').addEventListener('click',function(){showPage('archive');});
  document.getElementById('nb-reports').addEventListener('click',function(){showPage('reports');});
  document.getElementById('nb-settings').addEventListener('click',function(){showPage('settings');});
  document.getElementById('btn-signout').addEventListener('click',signOut);
  document.getElementById('btn-signin').addEventListener('click',signIn);
  document.getElementById('btn-signout-denied').addEventListener('click',signOut);

  // Tickets
  document.getElementById('btn-refresh').addEventListener('click',loadTicketsDebounced);
  document.getElementById('btn-export-csv').addEventListener('click',function(){exportCSV(false);});

  // Filter chips — use data-filter attribute
  document.querySelectorAll('.filter-bar .chip[data-filter]').forEach(function(btn){
    btn.addEventListener('click',function(){setFilter(this.dataset.filter,this);});
  });

  // Sort headers — use data-col attribute
  document.querySelectorAll('#main-table thead th[data-col]').forEach(function(th){
    th.addEventListener('click',function(){sortBy(this.dataset.col);});
  });

  // Archive
  document.getElementById('btn-export-archive').addEventListener('click',function(){exportCSV(true);});

  // Reports toolbar
  document.getElementById('btn-report-refresh').addEventListener('click',refreshActiveReport);
  document.getElementById('btn-export-toggle').addEventListener('click',toggleExportMenu);
  document.getElementById('btn-export-pdf-active').addEventListener('click',function(){exportReportPDF('active');});
  document.getElementById('btn-export-pdf-all').addEventListener('click',function(){exportReportPDF('all');});
  document.getElementById('btn-export-csv-active').addEventListener('click',function(){exportReportCSV('active');});
  document.getElementById('btn-export-csv-all').addEventListener('click',function(){exportReportCSV('all');});

  // Report sub-nav
  document.getElementById('rc-tickets').addEventListener('click',function(){showReport('tickets',this);});
  document.getElementById('rc-defender').addEventListener('click',function(){showReport('defender',this);});
  document.getElementById('rc-phishing').addEventListener('click',function(){showReport('phishing',this);});
  document.getElementById('rc-syskit').addEventListener('click',function(){showReport('syskit',this);});

  // Report range select
  document.getElementById('rpt-range').addEventListener('change',renderTicketReport);

  // Defender manual inputs
  ['def-high','def-med','def-low','def-info','def-new','def-resolved'].forEach(function(id){
    document.getElementById(id).addEventListener('input',renderDefenderManual);
  });
  document.getElementById('btn-defender-fetch').addEventListener('click',loadDefenderAlerts);
  document.getElementById('btn-defender-save').addEventListener('click',saveDefenderManual);

  // Phishing manual inputs
  ['ph-sent','ph-opened','ph-clicked','ph-submitted','ph-reported','ph-trained','ph-atrisk'].forEach(function(id){
    document.getElementById(id).addEventListener('input',renderPhishingManual);
  });
  document.getElementById('btn-phishing-fetch').addEventListener('click',loadPhishingData);
  document.getElementById('btn-phishing-save').addEventListener('click',savePhishingManual);

  // Syskit manual inputs
  ['sk-teams','sk-teams-inactive','sk-groups','sk-orphaned','sk-users','sk-guests',
   'sk-inactive','sk-shared','sk-anon','sk-ext','sk-storage','sk-storage-total'].forEach(function(id){
    document.getElementById(id).addEventListener('input',renderSyskitManual);
  });
  document.getElementById('btn-syskit-save').addEventListener('click',saveSyskitManual);

  // Settings
  document.getElementById('btn-settings-save').addEventListener('click',saveSettings);
  document.getElementById('btn-add-agent').addEventListener('click',addAgent);
  document.getElementById('new-agent-email').addEventListener('keydown',function(e){if(e.key==='Enter')addAgent();});

  // Panel
  document.getElementById('panel-ov').addEventListener('click',closePanel);
  document.getElementById('btn-panel-close').addEventListener('click',closePanel);
  document.getElementById('p-priority-sel').addEventListener('change',updatePriority);
  document.getElementById('p-status-sel').addEventListener('change',updateStatus);
  document.getElementById('p-asg-sel').addEventListener('change',updateAssignee);
  document.getElementById('pickup-btn').addEventListener('click',pickupTicket);
  document.getElementById('send-btn').addEventListener('click',sendReply);
  document.getElementById('btn-save-notes').addEventListener('click',function(e){saveNote(e);});
  document.getElementById('cycle-btn').addEventListener('click',cycleStatus);
  document.getElementById('del-btn').addEventListener('click',deleteTicket);

  // Phase 3 — category, watchers, workload, merge
  document.getElementById('p-cat-sel').addEventListener('change',updateCategory);
  document.getElementById('btn-add-watcher').addEventListener('click',addWatcher);
  document.getElementById('watcher-input').addEventListener('keydown',function(e){if(e.key==='Enter')addWatcher();});
  document.getElementById('btn-workload').addEventListener('click',showWorkload);
  document.getElementById('btn-workload-close').addEventListener('click',closeWorkload);
  document.getElementById('workload-modal-ov').addEventListener('click',function(e){if(e.target===this)closeWorkload();});
  document.getElementById('btn-merge').addEventListener('click',showMergeModal);
  document.getElementById('btn-merge-close').addEventListener('click',closeMergeModal);
  document.getElementById('btn-merge-cancel').addEventListener('click',closeMergeModal);
  document.getElementById('btn-merge-confirm').addEventListener('click',confirmMerge);
  document.getElementById('merge-modal-ov').addEventListener('click',function(e){if(e.target===this)closeMergeModal();});
  document.getElementById('merge-search').addEventListener('input',function(){renderMergeList(this.value);});
}

document.addEventListener('DOMContentLoaded', bindEvents);

/* ═══════════ PHASE 3: CATEGORIES ═══════════ */
var CATEGORIES = [
  'Account / Access',
  'Audio / Visual',
  'Device / Hardware',
  'Email / Outlook',
  'Microsoft 365 / Teams',
  'Network / Connectivity',
  'Onboarding / Offboarding',
  'Printer / Peripheral',
  'Security / Virus',
  'Software / Application',
  'Training / How-to',
  'Website / Intranet',
  'Other'
];

function updateCategory(){
  var t=getTicket(openId);if(!t)return;
  t.cat=document.getElementById('p-cat-sel').value;
  graphFetch(LIST_URL+'/items/'+t.spId+'/fields','PATCH',{Category:t.cat})
  .catch(function(e){console.error('Category update failed:',e);showPersistError();});
  renderDash();
}

/* ═══════════ PHASE 3: WATCHERS ═══════════ */
function getWatchers(t){
  if(!t.watchers||!t.watchers.trim())return[];
  return t.watchers.split(',').map(function(w){return w.trim();}).filter(Boolean);
}

function saveWatchers(t){
  t.watchers=(getWatchers(t)).join(',');
  graphFetch(LIST_URL+'/items/'+t.spId+'/fields','PATCH',{Watchers:t.watchers})
  .catch(function(e){console.error('Watcher save failed:',e);showPersistError();});
}

function renderWatcherTags(t){
  var container=document.getElementById('watcher-tags');
  container.innerHTML='';
  var ws=getWatchers(t);
  if(!ws.length){
    var empty=document.createElement('span');
    empty.style.cssText='font-size:11px;color:var(--muted);';
    empty.textContent='No watchers added';
    container.appendChild(empty);
    return;
  }
  ws.forEach(function(w){
    var tag=document.createElement('span');tag.className='watcher-tag';
    tag.innerHTML=esc(w)+'<button title="Remove" data-w="'+esc(w)+'">\u00d7</button>';
    tag.querySelector('button').addEventListener('click',function(){
      var t2=getTicket(openId);if(!t2)return;
      var cur=getWatchers(t2).filter(function(x){return x!==w;});
      t2.watchers=cur.join(',');
      saveWatchers(t2);
      renderWatcherTags(t2);
    });
    container.appendChild(tag);
  });
}

function addWatcher(){
  var t=getTicket(openId);if(!t)return;
  var inp=document.getElementById('watcher-input');
  var email=(inp.value||'').trim().toLowerCase();
  if(!email||!/^[^@]+@[^@]+\.[^@]+$/.test(email)){inp.style.borderColor='var(--accent)';setTimeout(function(){inp.style.borderColor='';},1500);return;}
  var cur=getWatchers(t);
  if(cur.indexOf(email)===-1){cur.push(email);t.watchers=cur.join(',');saveWatchers(t);}
  inp.value='';
  renderWatcherTags(t);
}

/* ═══════════ PHASE 3: WORKLOAD VIEW ═══════════ */
function showWorkload(){
  var active=tickets.filter(function(t){return !t._archived&&t.status!=='closed';});
  var agents=getStoredAgents();
  var body=document.getElementById('workload-body');

  // tally per agent
  var tally={};
  var unassigned=0;
  agents.forEach(function(a){tally[a.email.toLowerCase()]={email:a.email,crit:0,high:0,med:0,low:0,total:0};});
  active.forEach(function(t){
    var key=(t.assignee||'').toLowerCase();
    if(!key){unassigned++;return;}
    if(!tally[key])tally[key]={email:t.assignee,crit:0,high:0,med:0,low:0,total:0};
    tally[key][t.priority==='critical'?'crit':t.priority==='high'?'high':t.priority==='medium'?'med':'low']++;
    tally[key].total++;
  });

  var rows=Object.values(tally).sort(function(a,b){return b.total-a.total;});
  var max=rows.reduce(function(m,r){return Math.max(m,r.total);},1);

  var html='';
  if(!rows.length){html='<p style="font-size:12px;color:var(--muted);">No agents found.</p>';}
  rows.forEach(function(r){
    var pct=Math.round((r.total/max)*100);
    var barCls=r.total===0?'':r.total>=8?'crit':r.total>=5?'warn':'';
    var breakdown='';
    if(r.crit)breakdown+='<span style="color:#b91c1c;">'+r.crit+' crit</span> ';
    if(r.high)breakdown+='<span style="color:var(--warn);">'+r.high+' high</span> ';
    if(r.med) breakdown+='<span style="color:var(--ink);">'+r.med+' med</span> ';
    if(r.low) breakdown+='<span style="color:var(--muted);">'+r.low+' low</span>';
    html+='<div class="wl-row">'+
      '<div class="wl-name">'+esc(r.email.split('@')[0])+'<div style="font-size:10px;color:var(--muted);">'+esc(r.email)+'</div></div>'+
      '<div class="wl-bar-wrap"><div class="wl-bar '+barCls+'" style="width:'+pct+'%"></div></div>'+
      '<div class="wl-count">'+r.total+' open</div>'+
      '<div class="wl-breakdown">'+breakdown+'</div>'+
      '</div>';
  });
  if(unassigned>0){
    html+='<div class="wl-row" style="opacity:.7;">'+
      '<div class="wl-name" style="color:var(--muted);">Unassigned</div>'+
      '<div class="wl-bar-wrap"><div class="wl-bar crit" style="width:'+Math.round((unassigned/max)*100)+'%"></div></div>'+
      '<div class="wl-count">'+unassigned+' open</div>'+
      '<div class="wl-breakdown"></div></div>';
  }

  body.innerHTML=html||'<p style="font-size:12px;color:var(--muted);">No open tickets.</p>';
  document.getElementById('workload-modal-ov').classList.add('open');
}

function closeWorkload(){document.getElementById('workload-modal-ov').classList.remove('open');}

/* ═══════════ PHASE 3: MERGE TICKETS ═══════════ */
var _mergeSelected=null;

function showMergeModal(){
  _mergeSelected=null;
  var current=getTicket(openId);if(!current)return;
  document.getElementById('merge-search').value='';
  renderMergeList('');
  document.getElementById('merge-modal-ov').classList.add('open');
}

function closeMergeModal(){document.getElementById('merge-modal-ov').classList.remove('open');_mergeSelected=null;}

function renderMergeList(query){
  var q=(query||'').toLowerCase();
  var current=getTicket(openId);
  var candidates=tickets.filter(function(t){
    if(t.id===openId||t._archived)return false;
    if(q)return t.id.toLowerCase().includes(q)||t.subject.toLowerCase().includes(q)||t.name.toLowerCase().includes(q);
    return true;
  });
  var list=document.getElementById('merge-list');
  if(!candidates.length){list.innerHTML='<div style="font-size:12px;color:var(--muted);padding:12px 0;">No matching tickets found.</div>';return;}
  list.innerHTML='';
  candidates.slice(0,20).forEach(function(t){
    var div=document.createElement('div');
    div.className='merge-item'+(_mergeSelected===t.id?' selected':'');
    div.innerHTML='<div class="mi-id">'+esc(t.id)+'</div>'+
      '<div class="mi-title">'+esc(t.subject)+'</div>'+
      '<div class="mi-meta">'+esc(t.name)+' &middot; '+esc(t.status)+' &middot; '+priBadge(t.priority)+'</div>';
    div.addEventListener('click',function(){
      _mergeSelected=t.id;
      document.querySelectorAll('.merge-item').forEach(function(el){el.classList.remove('selected');});
      div.classList.add('selected');
    });
    list.appendChild(div);
  });
}

function confirmMerge(){
  if(!_mergeSelected){alert('Please select a ticket to merge.');return;}
  var primary=getTicket(openId);
  var secondary=getTicket(_mergeSelected);
  if(!primary||!secondary)return;
  if(!confirm('Merge '+secondary.id+' into '+primary.id+'? '+secondary.id+' will be closed and its replies moved to '+primary.id+'.'))return;

  // combine replies, sorted by timestamp
  var combined=(primary.replies||[]).concat(secondary.replies||[]);
  combined.sort(function(a,b){return new Date(a.ts)-new Date(b.ts);});

  // add merge notice
  combined.push({from:'agent',text:'[Merged from '+secondary.id+': '+secondary.subject+']',ts:new Date().toISOString()});

  // update primary
  primary.replies=combined;
  if(primary.notes){primary.notes+='\n[Merged from '+secondary.id+']';}
  else{primary.notes='[Merged from '+secondary.id+']';}

  // close secondary
  applyStatusChange(secondary,'closed');
  secondary.notes=(secondary.notes?secondary.notes+'\n':'')+'[Merged into '+primary.id+']';

  // persist both
  var btn=document.getElementById('btn-merge-confirm');
  btn.disabled=true;btn.textContent='Merging\u2026';

  Promise.all([persistToGraph(primary),persistToGraph(secondary)])
  .then(function(){
    // remove secondary from local array so it disappears immediately without waiting for SP refresh
    tickets=tickets.filter(function(t){return t.id!==_mergeSelected;});
    closeMergeModal();
    openPanel(primary.id);
    renderDash();
    btn.disabled=false;btn.textContent='\u8644 Merge tickets';
  })
  .catch(function(e){
    console.error('Merge error:',e);
    showPersistError();
    btn.disabled=false;btn.textContent='\u8644 Merge tickets';
  });
}
