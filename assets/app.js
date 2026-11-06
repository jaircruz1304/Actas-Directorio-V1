/* FIAS Directorio · Front-End estático para GitHub Pages. Ningún dato del Excel se incrusta en la distribución. */
const DEPLOY_CONFIG=window.FIAS_APP_CONFIG||{};
const SETTINGS={tenantId:'5e23e4af-237d-4d97-bf6e-dca808015787',clientId:'c8a828c1-3a20-4876-96af-b4f30ce4abeb',matrixUrl:'https://fiasec-my.sharepoint.com/:x:/g/personal/dvillacis_fias_org_ec/IQCho4wt1qyXT6DnsDAU-r33AavyNjkIhd9xhXy5QgQvLkk',folderUrl:'https://fiasec-my.sharepoint.com/:f:/g/personal/dvillacis_fias_org_ec/IgAPa6xunze7Rp5Sbb3eoG4bAbnTiJ3TP2iPh5nhaefQklY?e=WH2jtK',scopes:['User.Read','Files.Read.All','Sites.Read.All'],indexShareUrl:DEPLOY_CONFIG.indexShareUrl||'',autoPublishNewActas:DEPLOY_CONFIG.autoPublishNewActas===true};
const $=id=>document.getElementById(id),state={rows:[],acts:0,warnings:[],account:null,client:null,origin:'',lastSync:'',search:'',year:'',fund:'',session:'',visible:12,page:'inicio',semantic:null,vectors:null,semanticBusy:false,pdfChunks:[],pdfVectors:null,pdfCancel:false,repository:[],changeLog:[],baselineStats:null,scanRunning:false,autoTimer:null,privateIndexLoaded:false,baselineDocuments:[],ocrWorker:null,lastMatrixSignature:null,privateIndexUrl:'',remoteIndexMeta:null,remoteIndexDirty:false,remoteIndexBusy:false,remoteIndexError:'',remoteIndexFromSharePoint:false};const num=new Intl.NumberFormat('es-EC');const clean=x=>String(x??'').trim();const fold=x=>clean(x).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLocaleLowerCase('es');const esc=x=>clean(x).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));const snippet=(s,n=250)=>clean(s).slice(0,n)+(clean(s).length>n?'…':'');const funds=['FAP','FEIG','REM','PASF','PASNAP','CONSERVA AVES','BIOECONOMÍA','SOCIO BOSQUE','GALÁPAGOS','KFW'];const toast=(message)=>{const e=$('toast');e.textContent=message;e.classList.add('show');clearTimeout(toast.timer);toast.timer=setTimeout(()=>e.classList.remove('show'),5600)};
function safeUrl(value){try{const url=new URL(value);return url.protocol==='https:'&&(url.hostname==='fiasec-my.sharepoint.com'||url.hostname.endsWith('.sharepoint.com'))?url.href:''}catch{return ''}}
function normLink(value){const s=clean(value);if(s.startsWith('https://'))return safeUrl(s);const m=s.match(/\/:([bfx]):\/g\/personal\/[^\s"<>]+/i);return m?safeUrl('https://fiasec-my.sharepoint.com'+m[0]):''}
function rowTags(text){const s=fold(text);return funds.filter(f=>new RegExp(`(^|[^\\p{L}])${fold(f).replace(/ /g,'\\s+')}($|[^\\p{L}])`,'iu').test(s))}
function dateString(x,XLSX){if(x instanceof Date&&!Number.isNaN(x.getTime()))return x.toISOString().slice(0,10);if(typeof x==='number'&&x>20000&&x<70000){const d=XLSX.SSF.parse_date_code(x);if(d)return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`}return clean(x)}
function actYear(d,sheet){return (clean(d).match(/20\d{2}/)||clean(sheet).match(/20\d{2}/)||[])[0]||''}
function parseExcel(arrayBuffer){const X=window.XLSX;if(!X)throw Error('No está disponible el lector Excel.');const wb=X.read(new Uint8Array(arrayBuffer),{type:'array',cellDates:false,cellFormula:true});const result=[],warnings=[];
 for(const sheet of wb.SheetNames){const ws=wb.Sheets[sheet];const matrix=X.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});const header=matrix.findIndex(r=>fold(r.map(clean).join(' ')).includes('puntos a tratar')&&fold(r.map(clean).join(' ')).includes('resolucion'));if(header<0){warnings.push(`Hoja ${sheet}: no se identificó encabezado de asuntos y resoluciones.`);continue}let current=null,n=0;
 for(let i=header+1;i<matrix.length;i++){const a=matrix[i]||[];const [d,t,p,r,f]=[0,1,2,3,4].map(j=>a[j]);const date=clean(d),title=clean(t),subject=clean(p),resolution=clean(r),filename=clean(f);if(!date&&!title&&!subject&&!resolution&&!filename)continue;
 if(date||title){n++;current={date:dateString(d,X),title,actId:`${sheet}-${n}`,documentUrl:'',documentName:'',members:[]}}if(!current)continue;const cell=ws[`E${i+1}`];const link=normLink(clean(cell?.l?.Target)||clean(cell?.v));if(link)current.documentUrl=link;if(filename)current.documentName=filename;for(const prev of current.members){if(current.documentUrl)prev.documentUrl=current.documentUrl;if(current.documentName)prev.documentName=current.documentName}if(!subject&&!resolution)continue;const year=actYear(current.date,sheet);const program=rowTags([subject,resolution,current.title].join(' '));const amount=([...`${subject} ${resolution}`.matchAll(/(?:USD\s*|US\$\s*|\$\s*)([\d.,]+)/gi)]).map(x=>x[0]).slice(0,3).join(', ');const rec={id:`${sheet}-${i+1}`,year,date:current.date,act:current.title,actId:current.actId,session:/extraordinaria/i.test(current.title)?'Extraordinaria':/ordinaria/i.test(current.title)?'Ordinaria':'',resolution,subject,program,amount,documentUrl:current.documentUrl,documentName:current.documentName,sheet,rowNumber:i+1};result.push(rec);current.members.push(rec)} }
 return{rows:result,warnings,acts:new Set(result.map(r=>r.actId)).size}}
function loadScript(url,name){if(window[name])return Promise.resolve();return new Promise((resolve,reject)=>{const script=document.createElement('script');script.src=url;script.onload=resolve;script.onerror=()=>reject(Error('No se pudo cargar '+name+'. Revisa la conexión o las políticas de red.'));document.head.appendChild(script)})}
async function excelReader(){await loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js','XLSX')}
async function msalReader(){if(window.msal)return window.msal;try{await loadScript('https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.13.0/lib/msal-browser.min.js','msal')}catch{ /* Alternate ESM build for hosts that do not carry the UMD build. */ const m=await import('https://esm.sh/@azure/msal-browser@4.13.0');return m}return window.msal}
async function client(){if(state.client)return state.client;const m=await msalReader();const uri=location.origin+location.pathname;const c=new m.PublicClientApplication({auth:{clientId:SETTINGS.clientId,authority:`https://login.microsoftonline.com/${SETTINGS.tenantId}`,redirectUri:uri},cache:{cacheLocation:'sessionStorage'}});await c.initialize();await c.handleRedirectPromise();state.client=c;state.account=c.getAllAccounts()[0]||null;authUi();return c}
// Caché privado por cuenta en IndexedDB. No se publica ni se envía a GitHub.
const LOCAL_DB='fias-directorio-v08';
function localDB(){return new Promise((resolve,reject)=>{if(!('indexedDB' in window))return reject(Error('Este navegador no permite IndexedDB.'));const r=indexedDB.open(LOCAL_DB,1);r.onupgradeneeded=()=>{if(!r.result.objectStoreNames.contains('indices'))r.result.createObjectStore('indices')};r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error||Error('Error de almacenamiento local'))})}
function privateKey(){return state.account?.homeAccountId||''}
async function cacheOp(operation,value,overrideKey=''){const key=overrideKey||privateKey();if(!key)return null;const db=await localDB();try{return await new Promise((resolve,reject)=>{const tx=db.transaction('indices',operation==='get'?'readonly':'readwrite');const store=tx.objectStore('indices');let req=operation==='get'?store.get(key):operation==='put'?store.put(value,key):store.delete(key);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error||Error('Error de caché'));tx.onerror=()=>reject(tx.error||Error('Error de transacción'))})}finally{db.close()}}
async function savePrivateCache(){if(!state.account||!state.privateIndexLoaded)return;await cacheOp('put',{version:8,pdfChunks:state.pdfChunks,baselineDocuments:state.baselineDocuments,baselineStats:state.baselineStats,repository:state.repository,changeLog:state.changeLog,rows:state.rows,acts:state.acts,updatedAt:new Date().toISOString(),remoteIndexMeta:state.remoteIndexMeta,remoteIndexDirty:state.remoteIndexDirty,remoteIndexFromSharePoint:state.remoteIndexFromSharePoint,indexShareUrl:state.privateIndexUrl})}
async function restorePrivateCache(){let data;try{data=await cacheOp('get')}catch(e){console.warn('No se pudo leer el caché privado:',e.message);return false}if(!data||data.version!==8||!Array.isArray(data.pdfChunks))return false;if(data.indexShareUrl&&data.indexShareUrl!==configuredPrivateIndexUrl())return false;state.pdfChunks=data.pdfChunks;state.baselineDocuments=data.baselineDocuments||[];state.baselineStats=data.baselineStats||null;state.repository=data.repository||[];state.changeLog=data.changeLog||[];state.privateIndexLoaded=true;state.remoteIndexMeta=data.indexShareUrl===configuredPrivateIndexUrl()?(data.remoteIndexMeta||null):null;state.remoteIndexDirty=!!data.remoteIndexDirty;state.remoteIndexFromSharePoint=!!data.remoteIndexFromSharePoint;state.pdfVectors=null;if(!state.rows.length&&Array.isArray(data.rows))putData({rows:data.rows,acts:data.acts||new Set(data.rows.map(x=>x.actId)).size,warnings:[]},'Caché privado institucional');$('pdfIndexProgress').textContent=`Índice restaurado: ${state.pdfChunks.length} pasajes PDF. Consultando cambios en SharePoint…`;reportRender();renderPrivateIndexStatus();return true}
async function login(){if(!/^https?:$/.test(location.protocol))throw Error('Abre la página desde HTTPS o localhost para autenticarte.');const c=await client();const response=await c.loginPopup({scopes:SETTINGS.scopes,prompt:'select_account'});state.account=response.account;authUi();toast('Sesión institucional iniciada. Consultando la matriz…');await restorePrivateCache();await restorePrivateIndexUrl();await sync()}
async function logout(){const c=await client();const a=state.account;state.pdfCancel=true;if(state.ocrWorker){try{await state.ocrWorker.terminate()}catch{}state.ocrWorker=null}try{await cacheOp('delete')}catch(e){console.warn('No se eliminó el caché local:',e.message)}try{localStorage.removeItem(userTrackerKey())}catch{}if(a)await c.logoutPopup({account:a,mainWindowRedirectUri:location.origin+location.pathname});state.account=null;state.rows=[];state.acts=0;state.origin='';state.vectors=null;state.pdfChunks=[];state.pdfVectors=null;state.repository=[];state.changeLog=[];state.baselineStats=null;state.privateIndexLoaded=false;state.baselineDocuments=[];state.privateIndexUrl=configuredPrivateIndexUrl();state.remoteIndexMeta=null;state.remoteIndexDirty=false;state.remoteIndexFromSharePoint=false;state.remoteIndexError='';authUi();renderPrivateIndexStatus();updateDashboard();renderResults();reportRender();toast('Sesión finalizada. Se eliminó el índice privado de este navegador.') }
async function accessToken(scopes=SETTINGS.scopes){const c=await client();const account=state.account||c.getAllAccounts()[0];if(!account)throw Error('Primero inicia sesión con tu cuenta Microsoft 365.');try{return(await c.acquireTokenSilent({account,scopes})).accessToken}catch{return(await c.acquireTokenPopup({account,scopes})).accessToken}}
const encodeShare=url=>'u!'+btoa(unescape(encodeURIComponent(url))).replace(/=/g,'').replace(/\+/g,'-').replace(/\//g,'_');
async function graph(path){const token=await accessToken();let res=await fetch('https://graph.microsoft.com/v1.0'+path,{headers:{Authorization:'Bearer '+token,Accept:'application/json'}});if([429,503].includes(res.status)){await new Promise(r=>setTimeout(r,Math.min(5000,Number(res.headers.get('Retry-After')||2)*1000)));res=await fetch('https://graph.microsoft.com/v1.0'+path,{headers:{Authorization:'Bearer '+token}})}if(!res.ok){const detail=await res.text();throw Error(`Microsoft Graph ${res.status}: ${detail.slice(0,180)}`)}return res.json()}

// v1.0: localizador central definido durante el despliegue (config.js) y cargado
// automáticamente DESPUÉS de autenticar al usuario con Microsoft 365.
// La configuración JS es PÚBLICA en GitHub Pages; el enlace NO es una credencial.
// Protege el recurso en SharePoint mediante permisos nominativos, nunca "Cualquiera con vínculo".
function checkedPrivateUrl(input){
 const url=safeUrl(clean(input));
 if(!url||!/^https:\/\/fiasec-my\.sharepoint\.com\//i.test(url))throw Error('El índice debe apuntar a una URL HTTPS de SharePoint FIAS.');
 if(!/\/:[ux]:\//i.test(url))throw Error('La URL configurada no identifica un archivo JSON de SharePoint.');
 return url;
}
function configuredPrivateIndexUrl(){
 try{return checkedPrivateUrl(SETTINGS.indexShareUrl)}catch(e){console.warn('FIAS: falta configuración de localizador del índice.',e.message);return ''}
}
function renderPrivateIndexStatus(){
 const el=$('privateIndexStatus');if(!el)return;
 const url=state.privateIndexUrl,m=state.remoteIndexMeta;
 if($('privateIndexOpen')){$('privateIndexOpen').href=url||'#';$('privateIndexOpen').classList.toggle('hide',!url)}
 let cls='pi-wait',title='Índice no configurado',detail='Falta indexShareUrl en config.js del despliegue.';
 if(url){cls='pi-pending';title='Conexión automática configurada';detail=state.account?'Consultando archivo institucional con Microsoft Graph…':'El índice se consultará al iniciar sesión Microsoft 365.'}
 if(state.privateIndexLoaded){cls='pi-ready';title=state.remoteIndexFromSharePoint?'Índice central sincronizado':'Respaldo disponible localmente';detail=`${num.format(state.pdfChunks.length)} pasajes PDF disponibles`+(m?.lastModifiedDateTime?' · SharePoint: '+new Date(m.lastModifiedDateTime).toLocaleString('es-EC'):'');}
 if(state.remoteIndexDirty){cls='pi-dirty';title='Novedades pendientes de publicar';detail=`${num.format(state.pdfChunks.length)} pasajes. Se intentará publicar automáticamente si la cuenta dispone de permiso de escritura.`}
 if(state.remoteIndexError){cls='pi-error';title='No se pudo consultar o actualizar el índice';detail=state.remoteIndexError}
 el.className='pi-status '+cls;el.innerHTML=`<strong>${esc(title)}</strong><small>${esc(detail)}</small>`;
 $('btnIndexPush').disabled=!(state.account&&state.privateIndexLoaded&&state.privateIndexUrl&&state.remoteIndexMeta&&state.remoteIndexDirty)||state.remoteIndexBusy;
 $('btnRemoteRefresh').disabled=!(state.account&&url)||state.remoteIndexBusy;
}
async function restorePrivateIndexUrl(){
 // No depende de IndexedDB ni exige introducir enlaces por cada usuario/dispositivo.
 const current=configuredPrivateIndexUrl();
 if(state.privateIndexUrl&&current!==state.privateIndexUrl){
  // Cambio de origen: no utilizar el estado ETag de un archivo distinto.
  state.remoteIndexMeta=null;state.remoteIndexFromSharePoint=false;
 }
 state.privateIndexUrl=current;renderPrivateIndexStatus();
}
async function loadRemotePrivateIndex({force=false}={}){
 if(!state.account)throw Error('Conecta tu cuenta Microsoft 365.');
 if(!state.privateIndexUrl)return false;
 if(state.remoteIndexBusy)return false;
 state.remoteIndexBusy=true;state.remoteIndexError='';renderPrivateIndexStatus();
 try{
  const share=encodeShare(state.privateIndexUrl);
  const meta=await graph(`/shares/${share}/driveItem`);
  const real=meta.remoteItem||meta;
  if(!real.file&&!meta.file)throw Error('El enlace no corresponde a un archivo descargable.');
  const driveId=real.parentReference?.driveId||meta.parentReference?.driveId;
  const id=real.id||meta.id;
  if(!driveId||!id)throw Error('Microsoft Graph no entregó driveId/itemId para consultar el índice.');
  const currentEtag=real.eTag||meta.eTag||'';
  if(!force&&state.privateIndexLoaded&&state.remoteIndexMeta?.etag&&currentEtag&&currentEtag===state.remoteIndexMeta.etag){
   state.remoteIndexFromSharePoint=true;renderPrivateIndexStatus();return false;
  }
  if(state.remoteIndexDirty){throw Error('Hay nuevas actas procesadas localmente. Guarda los cambios en SharePoint o exporta el índice antes de reemplazarlo.');}
  let temporary=meta['@microsoft.graph.downloadUrl']||real['@microsoft.graph.downloadUrl'];
  if(!temporary){const resolved=await graph(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(id)}`);temporary=resolved['@microsoft.graph.downloadUrl'];}
  if(!temporary)throw Error('Graph no devolvió una URL temporal para descargar el índice.');
  const response=await fetch(temporary,{cache:'no-store'});
  if(!response.ok)throw Error(`No se pudo descargar el índice de SharePoint (${response.status}).`);
  const contentType=response.headers.get('content-type')||'';
  if(response.headers.get('content-length')&&Number(response.headers.get('content-length'))>35*1024*1024)throw Error('El archivo supera 35 MB; requiere otra estrategia de almacenamiento.');
  const content=await response.text();
  if(content.length>40*1024*1024)throw Error('El índice supera el límite configurado de lectura.');
  let data;try{data=JSON.parse(content.replace(/^\uFEFF/,''))}catch{throw Error('El archivo compartido no contiene un JSON válido. Comprueba que sea INDICE_PRIVADO_IMPORTABLE.json.');}
  applyPrivateIndexData(data,'SharePoint privado');
  state.remoteIndexMeta={driveId,id,etag:currentEtag,webUrl:safeUrl(real.webUrl||meta.webUrl||''),lastModifiedDateTime:real.lastModifiedDateTime||meta.lastModifiedDateTime||''};
  state.remoteIndexFromSharePoint=true;state.remoteIndexDirty=false;
  try{await savePrivateCache()}catch(e){console.warn('Caché no disponible:',e.message)}
  toast(`Índice central conectado: ${num.format(state.pdfChunks.length)} fragmentos de actas.`);
  return true;
 }catch(e){state.remoteIndexError=e.message;throw e}
 finally{state.remoteIndexBusy=false;renderPrivateIndexStatus();}
}
function currentPrivateIndexPayload(){
 const documents=state.baselineDocuments.map(x=>({...x}));
 const existing=new Set(documents.map(x=>`${x.year}|${x.fileName}`));
 for(const f of state.repository){const key=`${f.year}|${f.name}`;if(!existing.has(key)&&hasPdfChunks(f)){
   const pcs=state.pdfChunks.filter(c=>c.fileId===f.id||(c.file===f.name&&String(c.year)===String(f.year)));
   documents.push({fileName:f.name,year:f.year,path:`${f.year}/${f.name}`,shareUrl:f.url,bytes:f.size,graphId:f.id,graphEtag:f.etag,
    pages:Math.max(...pcs.map(c=>c.lastPage||c.page||0),0),readablePages:0,ocrPages:0,indexableChunks:pcs.length,processedAt:new Date().toISOString()});existing.add(key);
  }}
 const rows=state.rows.map(r=>({id:r.id,year:r.year,date:r.date,actId:r.actId,actTitle:r.act,subject:r.subject,resolution:r.resolution,
  fileName:r.documentName,shareUrl:r.documentUrl,sheet:r.sheet,rowNumber:r.rowNumber}));
 const stats={...(state.baselineStats||{}),matrixRecords:rows.length,pdfDocuments:documents.length,pdfPages:documents.reduce((a,b)=>a+(Number(b.pages)||0),0),
   pagesWithText:documents.reduce((a,b)=>a+(Number(b.readablePages)||0),0),indexChunks:state.pdfChunks.length,generatedAt:new Date().toISOString()};
 return {schema:'fias-directorio-private-v07',generatedAt:new Date().toISOString(),stats,rows,
  pdfChunks:state.pdfChunks.map(r=>({id:r.id,fileId:r.fileId,fileName:r.file,year:r.year,page:r.page,lastPage:r.lastPage,
   text:r.text,shareUrl:r.documentUrl})),documents};
}
async function publishRemoteIndex({automatic=false}={}){
 if(!state.account||!state.privateIndexUrl||!state.privateIndexLoaded||!state.remoteIndexMeta)throw Error('Primero vincula y carga el índice de SharePoint.');
 if(!state.remoteIndexDirty){toast('El índice de SharePoint ya está actualizado en este navegador.');return}
 if(state.remoteIndexBusy)throw Error('Ya existe una operación del índice en curso.');
 state.remoteIndexBusy=true;state.remoteIndexError='';renderPrivateIndexStatus();
 try{
  const {driveId,id,etag}=state.remoteIndexMeta;
  if(!etag)throw Error('No existe un ETag de referencia. Refresca primero el índice de SharePoint.');
  // Comprobación optimista: evita sobreescribir cambios efectuados por otra persona.
  // La actualización silenciosa nunca abre ventanas de consentimiento desde tareas automáticas.
  const token=automatic ? (await (await client()).acquireTokenSilent({account:state.account,scopes:['Files.ReadWrite']})).accessToken : await accessToken(['Files.ReadWrite']);
  const check=await graph(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(id)}`);
  if(check.eTag!==etag)throw Error('Otra persona actualizó el índice en SharePoint. No se sobrescribió. Descarga una copia y revisa los cambios.');
  const json=JSON.stringify(currentPrivateIndexPayload());
  const response=await fetch(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(id)}/content`,{
    method:'PUT',headers:{Authorization:'Bearer '+token,'Content-Type':'application/json','If-Match':etag},body:json});
  if(!response.ok){const details=(await response.text()).slice(0,220);throw Error(`No se pudo actualizar el índice (${response.status}). ${details}`)}
  const saved=await response.json();
  state.remoteIndexMeta={driveId,id,etag:saved.eTag||'',lastModifiedDateTime:saved.lastModifiedDateTime||new Date().toISOString(),webUrl:safeUrl(saved.webUrl||'')};
  state.remoteIndexDirty=false;state.remoteIndexFromSharePoint=true;
  state.baselineStats=currentPrivateIndexPayload().stats;
  await savePrivateCache();toast(automatic?'Nuevas actas indexadas y publicadas automáticamente en SharePoint.':'Índice privado actualizado en SharePoint. Los nuevos fragmentos ya están disponibles para el próximo inicio de sesión.');
 }catch(e){state.remoteIndexError=e.message;throw e}
 finally{state.remoteIndexBusy=false;renderPrivateIndexStatus();}
}
async function sync(){if(!state.account)throw Error('Inicia sesión para sincronizar la matriz desde SharePoint.');$('btnSync').disabled=true;connection('Sincronizando…',false);try{if(state.privateIndexUrl){try{await loadRemotePrivateIndex()}catch(error){console.warn('Índice privado no disponible:',error.message);toast('Índice privado: '+error.message)}}await excelReader();const meta=await graph(`/shares/${encodeShare(SETTINGS.matrixUrl)}/driveItem`);let url=meta['@microsoft.graph.downloadUrl'];if(!url&&meta.parentReference?.driveId){const more=await graph(`/drives/${encodeURIComponent(meta.parentReference.driveId)}/items/${encodeURIComponent(meta.id)}`);url=more['@microsoft.graph.downloadUrl']}if(!url)throw Error('Microsoft Graph no devolvió una URL temporal de descarga.');const file=await fetch(url,{cache:'no-store'});if(!file.ok)throw Error('Error al descargar matriz: '+file.status);const arr=await file.arrayBuffer();const bytes=new Uint8Array(arr);if(bytes[0]!==80||bytes[1]!==75)throw Error('La respuesta no parece un archivo XLSX.');const parsed=parseExcel(arr);const oldIds=new Set(state.rows.map(r=>r.id));const initial=state.rows.length===0;const added=parsed.rows.filter(r=>!oldIds.has(r.id)).length;putData(parsed,'Microsoft SharePoint');$('lastSync').textContent='Actualizado '+new Date().toLocaleString('es-EC',{dateStyle:'short',timeStyle:'short'});connection('SharePoint conectado',true);toast(`Matriz actualizada: ${num.format(parsed.rows.length)} registros${!initial?` · ${added} filas nuevas`:''}.`);if(state.privateIndexLoaded){await detectNewActas({automatic:true});await savePrivateCache()}else{$('pdfIndexProgress').textContent='Matriz conectada. No se pudo recuperar el índice institucional preconfigurado; comprueba los permisos de SharePoint y vuelve a sincronizar.';reportRender()}}finally{$('btnSync').disabled=false}}
function connection(label,ok){$('connection').innerHTML=`<i></i> ${esc(label)}`;$('connection').classList.toggle('ok',!!ok)}
function authUi(){const logged=!!state.account;$('btnLogin').classList.toggle('hide',logged);$('btnLogout').classList.toggle('hide',!logged);if(logged){$('btnLogout').title='Cerrar sesión de '+(state.account.name||'Microsoft 365')} }
function putData(parsed,origin){state.rows=parsed.rows;state.acts=parsed.acts;state.warnings=parsed.warnings;state.origin=origin;state.visible=12;state.vectors=null;state.pdfVectors=null;updateDashboard();renderResults();renderAnalytics();if(state.semantic)$('semanticStatus').textContent='Modelo listo · reindexar';if(parsed.warnings.length)toast(parsed.warnings.join(' · ').slice(0,350))}
function setPage(page){state.page=page;document.querySelectorAll('.page').forEach(x=>x.classList.toggle('active',x.id===`page-${page}`));document.querySelectorAll('[data-page]').forEach(x=>x.classList.toggle('active',x.dataset.page===page));$('crumbPage').textContent={inicio:'Panorama ejecutivo',buscar:'Explorador de actas',inteligencia:'Asistente inteligente',analitica:'Analítica histórica'}[page];$('sidebar').classList.remove('open');window.scrollTo({top:0,behavior:'smooth'});if(page==='buscar')renderResults();if(page==='analitica')renderAnalytics()}
function updateDashboard(){const rows=state.rows;const years=[...new Set(rows.map(x=>x.year).filter(Boolean))].sort();$('kpiRows').textContent=rows.length?num.format(rows.length):'—';$('kpiActs').textContent=rows.length?num.format(state.acts):'—';$('kpiYears').textContent=rows.length?num.format(years.length):'—';$('kpiLinks').textContent=rows.length?num.format(rows.filter(x=>!!x.documentUrl).length):'—';$('healthTitle').textContent=rows.length?'Catálogo documental disponible':'Esperando datos';$('healthDesc').textContent=rows.length?`${num.format(rows.length)} registros consultables · ${state.warnings.length} alertas`: 'Conecta Microsoft 365 o importa la matriz local.';$('healthDot').classList.toggle('good',!!rows.length);$('originText').textContent=state.origin||'Sin conectar';const select=$('filterYear'),current=select.value;select.innerHTML='<option value="">Todos los años</option>'+years.map(y=>`<option value="${esc(y)}">${esc(y)}</option>`).join('');select.value=current;const fundSet=[...new Set(rows.flatMap(x=>x.program))].sort(),fund=$('filterFund'),prev=fund.value;fund.innerHTML='<option value="">Todos los fondos</option>'+fundSet.map(f=>`<option value="${esc(f)}">${esc(f)}</option>`).join('');fund.value=prev;const recent=rows.slice().sort((a,b)=>b.year.localeCompare(a.year)||b.rowNumber-a.rowNumber).slice(0,5);$('recentRows').innerHTML=recent.length?recent.map(r=>`<div class="recent-row" data-id="${esc(r.id)}"><span class="recent-year">${esc(r.year||'—')}</span><div class="recent-name"><strong>${esc(snippet(r.subject||r.resolution,125))}</strong><small>${esc(r.act||r.sheet)} · ${esc(r.date||r.year)}</small></div><span class="recent-arrow">↗</span></div>`).join(''):'<div class="empty-line">Conecta la matriz para visualizar el historial documental.</div>'}
function matches(r){if(state.year&&r.year!==state.year)return false;if(state.fund&&!r.program.includes(state.fund))return false;if(state.session&&r.session!==state.session)return false;return true}
const wordTokens=q=>fold(q).match(/[a-z0-9ñ]{2,}/g)||[];const STOP=new Set('de la el los las en por para con del al que una un sus su se fue sobre esta este desde como mas entre ha y o a e lo ante cuales cual quienes quien que son es estan fueron cuales cuales encontrar buscar busca dame mostrar muestra informacion antecedentes relacionados relacionadas relacionado relacionada acerca respecto decisiones tomadas existe existen cual cuales directorio actas acta'.split(' '));const expansions={'aprobacion':['aprobar','aprobado','aprobada','aprobacion','autoriza','autorizar'],'presupuesto':['presupuestaria','presupuestario','presupuesto','financiero'],'convenio':['acuerdo','alianza','convenio','cooperacion'],'donacion':['donaciones','donacion','donar','donante'],'contrato':['contratacion','contrataciones','contrato','servicio'],'fondo':['fondos','fondo','fap','feig','rem'],'gestion':['administracion','manejo','gestion'],'proyecto':['programa','iniciativa','proyecto']};
function lexicalRank(row,q){
 if(!q.trim())return 1;
 const terms=[...new Set(wordTokens(q).filter(t=>t.length>2&&!STOP.has(t)))];
 if(!terms.length)return 0;
 const names=[row.subject||'',row.resolution||'',`${row.act} ${row.program.join(' ')} ${row.year}`];
 const bags=names.map(n=>new Set(wordTokens(n)));
 const canon=t=>t.length>=6&&t.endsWith('es')?t.slice(0,-2):t.length>=5&&t.endsWith('s')?t.slice(0,-1):t;
 let score=0,hits=0,anchorMatches=0;
 const anchors=terms.filter(t=>/^(fap|feig|rem|pasf|pasnap|kfw|fias)$/.test(t));
 for(const term of terms){
   const lookup=anchors.includes(term)?null:Object.entries(expansions).find(([k,v])=>k===term||v.includes(term));
   const candidates=new Set([term,...(lookup?[lookup[0],...lookup[1]]:[])]);
   let best=0;
   for(const bagIndex of [1,0,2]){
     const bag=bags[bagIndex];
     for(const cand of candidates){
       if(bag.has(cand)||[...bag].some(w=>canon(w)===canon(cand)))
         best=Math.max(best,(bagIndex===1?5:bagIndex===0?4.7:2.5)*(term===cand?1:0.65));
     }
   }
   if(best){hits++;score+=best;if(anchors.includes(term))anchorMatches++}
 }
 if(anchors.length&&anchorMatches<anchors.length)return 0;
 if(hits<Math.max(1,Math.ceil(terms.length*.42)))return 0;
 return score/(terms.length*5);
}
function ranked(q,filtered=true){const rr=state.rows.filter(r=>!filtered||matches(r));return rr.map(r=>({r,score:lexicalRank(r,q)})).filter(x=>x.score>0).sort((a,b)=>b.score-a.score||b.r.year.localeCompare(a.r.year)||b.r.rowNumber-a.r.rowNumber)}
function renderResults(){state.year=$('filterYear').value;state.fund=$('filterFund').value;state.session=$('filterSession').value;state.search=$('searchQuery').value;const all=ranked(state.search),show=all.slice(0,state.visible);$('searchCount').textContent=all.length?`${num.format(all.length)} coincidencia${all.length===1?'':'s'} · ${num.format(state.rows.length)} entradas disponibles`:(state.rows.length?'Sin coincidencias':'Conecta SharePoint o importa un Excel para comenzar');$('results').innerHTML=show.length?show.map(({r,score})=>`<article class="result-card"><div class="result-top"><span class="result-id">REGISTRO ${esc(r.id)} · ${esc(r.session||'Sesión')}</span><span class="result-year">◷ ${esc(r.date||r.year||'Sin fecha')}</span></div><h3>${esc(snippet(r.subject||r.resolution||'Sin asunto',220))}</h3><p>${esc(snippet(r.resolution||'No se consigna texto resolutivo en este registro. Consulta el acta para verificar el acuerdo.',360))}</p><div class="result-bottom"><span class="chip">${esc(r.year||r.sheet)}</span>${r.program.slice(0,3).map(p=>`<span class="chip">${esc(p)}</span>`).join('')}<div class="result-actions"><button class="link-btn" data-id="${esc(r.id)}">Ver detalle →</button>${r.documentUrl?`<a href="${esc(r.documentUrl)}" target="_blank" rel="noopener noreferrer" class="link-btn">Abrir acta ↗</a>`:''}</div></div></article>`).join(''):'<div class="empty-state"><div class="empty-icon">⌕</div><h3>'+ (state.rows.length?'Sin resultados para esta búsqueda':'Tu conocimiento empieza aquí')+'</h3><p>'+(state.rows.length?'Ajusta los filtros o prueba con otros términos.':'Conecta Microsoft 365 para explorar el historial del Directorio.')+'</p></div>';$('loadMore').classList.toggle('hide',all.length<=state.visible)}
function showRecord(id){const r=state.rows.find(x=>x.id===id);if(!r)return;$('drawerBody').innerHTML=`<div class="drawer-tag">FICHA DOCUMENTAL · ${esc(r.id)}</div><h2 class="drawer-title">${esc(r.subject||'Registro del Directorio')}</h2><div class="drawer-field"><label>IDENTIFICACIÓN</label><p>${esc(r.act||'Acta sin título')}<br>${esc(r.date||r.year||'Fecha no registrada')} · ${esc(r.session||'Tipo no registrado')}</p></div><div class="drawer-field"><label>ASUNTO TRATADO</label><p>${esc(r.subject||'No disponible')}</p></div><div class="drawer-field"><label>RESOLUCIÓN REGISTRADA</label><p>${esc(r.resolution||'La matriz no registra texto resolutivo en esta entrada.')}</p></div><div class="drawer-field"><label>FONDO / INICIATIVA · MENCIONES TEXTUALES</label><p>${esc(r.program.join(', ')||'Sin etiqueta detectada')}</p></div><div class="drawer-field"><label>TRAZABILIDAD EN LA MATRIZ</label><p>Hoja: ${esc(r.sheet)}<br>Fila de origen: ${r.rowNumber}<br>Identificador de acta: ${esc(r.actId)}<br>Documento: ${esc(r.documentName||'No especificado')}</p></div><div class="drawer-source">${r.documentUrl?`<a class="btn btn-primary" target="_blank" rel="noopener noreferrer" href="${esc(r.documentUrl)}">↗ Abrir documento original en SharePoint</a>`:'<div class="api-status">Documento fuente sin enlace validado en la matriz</div>'}</div><div class="drawer-warning">La transcripción del índice no sustituye la lectura del acta firmada. Los permisos de SharePoint se aplican al abrir el origen.</div>`;$('drawerBackdrop').classList.remove('hide');$('recordDrawer').classList.add('open')}
function renderAnalytics(){const rows=state.rows;const years=[...new Set(rows.map(r=>r.year).filter(Boolean))].sort();const max=Math.max(1,...years.map(y=>rows.filter(r=>r.year===y).length));$('yearChart').innerHTML=years.length?years.map(y=>{const n=rows.filter(r=>r.year===y).length;return `<div class="bar-column" title="${esc(y)}: ${n} entradas"><span class="bar-value">${n}</span><div class="bar-shape" style="height:${Math.max(2,Math.round(n/max*159))}px"></div><small>${esc(y)}</small></div>`}).join(''):'<div class="empty-line">Sin registros disponibles.</div>';const counts=funds.map(f=>({name:f,count:rows.filter(r=>r.program.includes(f)).length})).filter(x=>x.count>0).sort((a,b)=>b.count-a.count).slice(0,8),m=Math.max(1,...counts.map(x=>x.count));$('fundChart').innerHTML=counts.length?counts.map(x=>`<div class="fund-row"><strong>${esc(x.name)}</strong><div class="fund-track"><div class="fund-fill" style="width:${x.count/m*100}%"></div></div><span>${x.count}</span></div>`).join(''):'<div class="empty-line">Sin etiquetas detectadas.</div>'}
// Resumen extractivo: siempre construido únicamente de textos concretos indexados y con identificadores comprobables.
function extractiveAnswer(question,hits,mode='contextual'){if(!hits.length)return '<p>No se encontraron antecedentes suficientemente relacionados con la consulta en la matriz disponible. Prueba con otro término o verifica el PDF original.</p>';const sources=hits.slice(0,5);const heading=`<p class="answer-par">Se localizaron <strong>${hits.length}</strong> registros relacionados en el índice. A continuación se presentan los extractos más pertinentes, sin agregar afirmaciones ajenas a la matriz.</p>`;return heading+sources.map(({r},i)=>`<div class="answer-citation"><b>[${i+1}] ${esc(r.year||'Sin año')} · ${esc(snippet(r.act||r.sheet,110))}</b><div><strong>Asunto:</strong> ${esc(snippet(r.subject||'No disponible',260))}</div><div><strong>Texto resolutivo:</strong> ${esc(snippet(r.resolution||'No consta en la matriz',390))}</div><div>${r.documentUrl?`<a href="${esc(r.documentUrl)}" target="_blank" rel="noopener noreferrer">Ver fuente original ↗</a>`:'<span class="muted-source">Sin vínculo de acta validado</span>'} · <button class="link-btn" data-id="${esc(r.id)}">Ficha ${esc(r.id)}</button></div></div>`).join('')+`<div class="answer-footer">Método: ${mode==='semantic'?'similitud neuronal multilingüe + clasificación contextual':'coincidencia contextual y expansión terminológica'}. Este es un extracto documentado, no una conclusión legal ni un resumen generativo. No se analizó el texto completo de los PDF.</div>`}
async function activateSemantic(){if(state.semanticBusy)return;state.semanticBusy=true;$('semanticBtn').disabled=true;$('aiProgress').textContent='Descargando modelo multilingüe para ejecutarlo localmente. La primera activación puede requerir varios minutos…';try{const {pipeline,env}=await import('https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.7.6/+esm');env.allowRemoteModels=true;state.semantic=await pipeline('feature-extraction','Xenova/paraphrase-multilingual-MiniLM-L12-v2',{dtype:'q8'});$('semanticStatus').textContent='Modelo neuronal activo';$('semanticBtn').textContent='✧ Actualizar índice semántico';$('aiMode').textContent='✧ IA semántica activa';toast('Modelo de embeddings multilingüe cargado en el navegador.');await indexSemantic();if(state.pdfChunks.length)await indexPdfVectors()}catch(e){$('aiProgress').textContent='No fue posible cargar el modelo neuronal: '+e.message+'. Puedes seguir usando la búsqueda contextual.';toast('Modelo no disponible en este navegador o red.')}finally{state.semanticBusy=false;$('semanticBtn').disabled=false}}
const embedding=async s=>{const output=await state.semantic(snippet(s,850),{pooling:'mean',normalize:true});return Array.from(output.data)};
async function indexSemantic(){if(!state.semantic){await activateSemantic();return}if(!state.rows.length){$('aiProgress').textContent='Modelo cargado. Conecta primero la matriz para generar vectores.';return}state.semanticBusy=true;$('semanticBtn').disabled=true;try{const vec=[];for(let i=0;i<state.rows.length;i++){const r=state.rows[i];vec.push(await embedding(`${r.subject}. ${r.resolution}. ${r.program.join(' ')}.`));if(i%8===0||i===state.rows.length-1){$('aiProgress').textContent=`Indexación semántica local: ${i+1}/${state.rows.length} registros procesados…`;await new Promise(resolve=>setTimeout(resolve,0))}}state.vectors=vec;$('aiProgress').textContent=`Índice semántico listo: ${vec.length} vectores. Los registros no se enviaron a un servicio de IA externo.`;toast('Índice de similitud semántica generado correctamente.')}catch(e){$('aiProgress').textContent='Error durante indexación: '+e.message;state.vectors=null}finally{state.semanticBusy=false;$('semanticBtn').disabled=false}}
const similarity=(a,b)=>{let n=0;for(let i=0;i<a.length;i++)n+=a[i]*b[i];return n};
async function retrieve(question){let list=ranked(question,false);if(state.semantic&&state.vectors?.length===state.rows.length){const query=await embedding(question);const scored=state.rows.map((r,i)=>({r,score:similarity(query,state.vectors[i])*.8+lexicalRank(r,question)*.2}));list=scored.filter(x=>x.score>.13).sort((a,b)=>b.score-a.score);return{list,mode:'semantic'}}return{list,mode:'contextual'}}
function sendChat(text,html,role){const node=document.createElement('div');node.className='chat-message '+(role==='user'?'chat-user':'chat-assistant');if(role==='user')node.textContent=text;else node.innerHTML='<div class="assistant-title">✳ &nbsp;FIAS · INVESTIGACIÓN DOCUMENTADA</div>'+html;$('chatThread').append(node);$('chatThread').scrollTop=$('chatThread').scrollHeight;return node}
// Indexación de documentos PDF originales desde SharePoint con permisos delegados.
// El navegador mantiene los pasajes exclusivamente en RAM, y no incluye PDFs en el despliegue público.
async function graphPage(url){const token=await accessToken();const res=await fetch(url.startsWith('https://graph.microsoft.com/')?url:'https://graph.microsoft.com/v1.0'+url,{headers:{Authorization:'Bearer '+token,Accept:'application/json'}});if(!res.ok)throw Error('Graph '+res.status+' al recorrer la biblioteca');return res.json()}
async function sharepointPdfs(){
 const root=await graph(`/shares/${encodeShare(SETTINGS.folderUrl)}/driveItem`);
 const driveId=root.parentReference?.driveId;
 if(!driveId)throw Error('No se pudo resolver la biblioteca compartida.');
 const folders=[{id:root.id,year:''}],files=[];
 let visits=0;
 while(folders.length){
  if(state.pdfCancel)break;
  if(++visits>300)throw Error('Se alcanzó el límite de subcarpetas.');
  const folder=folders.shift();
  let next=`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folder.id)}/children?$top=200`;
  while(next){
   const payload=await graphPage(next);
   for(const item of payload.value||[]){
    if(item.folder)folders.push({id:item.id,year:/^20\d{2}$/.test(item.name)?item.name:folder.year});
    else if(/\.pdf$/i.test(item.name||''))files.push({id:item.id,driveId,name:item.name,size:item.size||0,
     etag:item.eTag||item.cTag||'',modified:item.lastModifiedDateTime||'',year:folder.year||(item.name.match(/20\d{2}/)||[])[0]||'',url:safeUrl(item.webUrl||'')});
   }
   next=payload['@odata.nextLink']||null;
  }
 }
 return files.sort((a,b)=>b.year.localeCompare(a.year)||a.name.localeCompare(b.name));
}
async function pdfReader(){if(window.__fiasPdfjs)return window.__fiasPdfjs;const api=await import('https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.mjs');api.GlobalWorkerOptions.workerSrc='https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.mjs';window.__fiasPdfjs=api;return api}
function pdfChunksFromPages(pages,file){const chunks=[];let accumulator='',start=1;for(let i=0;i<pages.length;i++){const text=pages[i].trim();if(!text)continue;if(!accumulator)start=i+1;accumulator+=' '+text;if(accumulator.length>=1400||i===pages.length-1){chunks.push({id:`pdf-${file.id}-${start}`,file:file.name,fileId:file.id,year:file.year||(file.name.match(/20\d{2}/)||[])[0]||'',page:start,lastPage:i+1,text:accumulator.trim().slice(0,5500),documentUrl:file.url});accumulator=''}}if(accumulator.trim())chunks.push({id:`pdf-${file.id}-${start}`,file:file.name,fileId:file.id,year:file.year||'',page:start,lastPage:pages.length,text:accumulator.trim().slice(0,5500),documentUrl:file.url});return chunks}
async function getOcrWorker(){if(state.ocrWorker)return state.ocrWorker;await loadScript('https://cdn.jsdelivr.net/npm/tesseract.js@6.0.1/dist/tesseract.min.js','Tesseract');if(!window.Tesseract?.createWorker)throw Error('No se pudo iniciar OCR en el navegador.');$('pdfIndexProgress').textContent='Cargando OCR español en el navegador (primera vez)…';state.ocrWorker=await window.Tesseract.createWorker('spa');return state.ocrWorker}
async function parsePdfFile(file,api){const details=await graph(`/drives/${encodeURIComponent(file.driveId)}/items/${encodeURIComponent(file.id)}`);const url=details['@microsoft.graph.downloadUrl'];if(!url)throw Error('Sin URL temporal de descarga');const response=await fetch(url,{cache:'no-store'});if(!response.ok)throw Error('Descarga '+response.status);const bytes=await response.arrayBuffer();const doc=await api.getDocument({data:new Uint8Array(bytes),useSystemFonts:true}).promise;const pages=[],total=doc.numPages;let ocrPages=0,unreadablePages=0;try{for(let i=1;i<=total;i++){if(state.pdfCancel)break;const page=await doc.getPage(i);let text='';try{const content=await page.getTextContent();text=content.items.map(it=>it.str||'').join(' ').trim();if(text.length<65){$('pdfIndexProgress').textContent=`${file.name} · OCR página ${i}/${total}…`;const worker=await getOcrWorker();const view=page.getViewport({scale:Math.min(2,2600/Math.max(page.getViewport({scale:1}).width,page.getViewport({scale:1}).height))});const canvas=document.createElement('canvas');canvas.width=Math.ceil(view.width);canvas.height=Math.ceil(view.height);try{const ctx=canvas.getContext('2d',{willReadFrequently:true});if(!ctx)throw Error('Canvas OCR no disponible');await page.render({canvasContext:ctx,viewport:view,canvas}).promise;const result=await worker.recognize(canvas);if((result.data?.text||'').trim().length>text.length)text=result.data.text.trim();ocrPages++}finally{canvas.width=0;canvas.height=0}}if(text.length<20)unreadablePages++;pages.push(text)}finally{page.cleanup()}if(i%3===0)await new Promise(resolve=>setTimeout(resolve,0))}}finally{await doc.destroy()}if(state.pdfCancel)throw Error('Procesamiento cancelado');const chunks=pdfChunksFromPages(pages,file);return {chunks,ocrPages,unreadablePages,pages:total}}
async function indexSharePointPdfs(){return detectNewActas({})}
async function indexPdfVectors(){if(!state.semantic||!state.pdfChunks.length)return;const vectors=[];for(let i=0;i<state.pdfChunks.length;i++){if(state.pdfCancel)break;vectors.push(await embedding(state.pdfChunks[i].text.slice(0,1200)));if(i%10===0){$('pdfIndexProgress').textContent=`Embeddings PDF: ${i+1}/${state.pdfChunks.length} pasajes`;await new Promise(r=>setTimeout(r,0))}}state.pdfVectors=vectors.length===state.pdfChunks.length?vectors:null}
function pdfLexical(question){
 if(!state.pdfChunks.length)return [];
 const tokens=[...new Set(wordTokens(question).filter(t=>t.length>2&&!STOP.has(t)))];
 if(!tokens.length)return [];
 const anchors=tokens.filter(t=>/^(fap|feig|rem|pasf|pasnap|kfw|fias)$/.test(t));
 return state.pdfChunks.map(c=>{
    const bag=new Set(wordTokens(c.text||''));
    if(anchors.length&&anchors.some(t=>!bag.has(t)))return {r:c,score:0};
    const hit=tokens.reduce((sum,t)=>sum+(bag.has(t)?1:0),0);
    return {r:c,score:hit/tokens.length};
 }).filter(x=>x.score>0).sort((a,b)=>b.score-a.score).slice(0,5);
}
async function pdfEvidence(question){let hits=pdfLexical(question);if(state.pdfVectors?.length===state.pdfChunks.length&&state.semantic){const v=await embedding(question);hits=state.pdfChunks.map((r,i)=>({r,score:similarity(v,state.pdfVectors[i])})).filter(x=>x.score>.25).sort((a,b)=>b.score-a.score).slice(0,5)}return hits}
function renderPdfEvidence(hits){if(!hits.length)return '';return `<div class="pdf-evidence-title"><strong>Pasajes de documentos PDF originales (${hits.length})</strong><span>Extracción automática · verificar contra el acta firmada</span></div>`+hits.map(({r},i)=>`<div class="answer-citation"><b>[PDF ${i+1}] ${esc(r.file||r.fileName)} · página ${r.page}${r.lastPage>r.page?'–'+r.lastPage:''}</b><div>${esc(snippet(r.text,520))}</div>${r.documentUrl?`<a target="_blank" rel="noopener noreferrer" href="${esc(r.documentUrl)}">Consultar acta en SharePoint ↗</a>`:'<span class="muted-source">Documento sin enlace confirmado</span>'}</div>`).join('')}
// Índice institucional descargado mediante Graph con permisos delegados; no se publica en GitHub.
function userTrackerKey(){return 'fias-actas-seen-v07-'+(state.account?.homeAccountId||'anónimo')}
function readTracker(){try{const v=JSON.parse(localStorage.getItem(userTrackerKey())||'{}');return v && typeof v==='object'?v:{}}catch{return {}}}
function storedTracker(items,previous={}){const out={...previous};for(const x of items){out[x.id]={etag:x.etag,size:x.size,modified:x.modified}}localStorage.setItem(userTrackerKey(),JSON.stringify(out))}
function stampFile(f){return JSON.stringify([f.etag,f.size,f.modified])}
function reportStatus(line){if(!line)return;const box=$('repoStatus');if(box)box.textContent=line}
function reportRender(){
 const docs=state.repository||[];
 const total=docs.length||state.baselineDocuments.length,news=state.changeLog.filter(x=>x.status==='nuevo').length,updates=state.changeLog.filter(x=>x.status==='modificado').length;
 reportStatus(`${total} ${docs.length?'PDF revisados en SharePoint':'PDF del índice local'} · ${news} nuevos · ${updates} actualizados · ${state.pdfChunks.length} fragmentos en memoria`);
 if($('repoReport'))$('repoReport').innerHTML=state.changeLog.length?
   state.changeLog.slice(0,16).map(x=>`<div class="report-entry"><span class="chip">${esc(x.status)}</span><strong>${esc(x.name)}</strong><small>${esc(x.year||'')}</small>${safeUrl(x.url)?`<a href="${esc(safeUrl(x.url))}" target="_blank" rel="noopener noreferrer">Fuente ↗</a>`:''}<span>${esc(x.note||'')}</span></div>`).join(''):
   '<p class="mini-copy">No se han identificado altas posteriores a la primera inspección de esta cuenta. La primera revisión establece la línea base.</p>';
}
function historicalDoc(file){return state.baselineDocuments.find(d=>String(d.year)===String(file.year)&&d.fileName===file.name)}
function hasPdfChunks(file){return state.pdfChunks.some(c=>(c.fileId===file.id)||(c.file===file.name&&String(c.year)===String(file.year)))}
async function detectNewActas({automatic=false}={}){
 if(!state.account||state.scanRunning)return;
 if(!state.privateIndexLoaded){$('pdfIndexProgress').textContent='No se dispone del índice histórico. Conéctate a Microsoft 365 y sincroniza SharePoint; revisa los permisos del archivo si persiste el problema.';if(!automatic)toast('El índice histórico no está disponible. Revisa la conexión de SharePoint.');return}
 state.scanRunning=true;state.pdfCancel=false;$('pdfCancelBtn').classList.remove('hide');
 $('pdfIndexProgress').textContent='Comparando los archivos de SharePoint con el índice histórico…';
 try{
   const files=await sharepointPdfs(),previous=readTracker();
   state.repository=files;
   const changes=[];
   for(const f of files){const old=previous[f.id];const indexed=hasPdfChunks(f)||!!historicalDoc(f);const changed=!!old&&(old.etag!==f.etag||old.size!==f.size);
     // Un archivo nuevo se procesa incluso en la primera visita, siempre que NO figure en el índice histórico.
     if(!indexed||changed)changes.push({file:f,status:old?(changed?'modificado':'pendiente'):'nuevo'});

   }
   const handled=[],success=[];let done=0;
   if(changes.length){const api=await pdfReader();for(const item of changes){if(state.pdfCancel)break;const f=item.file;
     const entry={id:f.id,name:f.name,url:f.url,year:f.year,status:item.status,note:''};
     try{if(f.size>45*1024*1024)throw Error('Supera 45 MB: requiere procesamiento fuera del navegador.');
       const parsed=await parsePdfFile(f,api);
       // Si ninguna página obtuvo texto, sigue pendiente: no marcar como indexada.
       if(!parsed.chunks.length)throw Error('El OCR no obtuvo texto suficiente. Revisar este PDF.');
       state.pdfChunks=state.pdfChunks.filter(c=>c.fileId!==f.id && !(c.file===f.name&&String(c.year)===String(f.year)));
       state.pdfChunks.push(...parsed.chunks);state.pdfVectors=null;
       state.baselineDocuments=state.baselineDocuments.filter(d=>!(String(d.year)===String(f.year)&&d.fileName===f.name));
       state.baselineDocuments.push({fileName:f.name,year:f.year,shareUrl:f.url,bytes:f.size,path:`${f.year}/${f.name}`,pages:parsed.pages,readablePages:parsed.pages-parsed.unreadablePages,ocrPages:parsed.ocrPages,pagesNeedingOCR:parsed.unreadablePages,indexableChunks:parsed.chunks.length,graphId:f.id,graphEtag:f.etag,processedAt:new Date().toISOString()});
       state.remoteIndexDirty=true;
       entry.note=`Indexado: ${parsed.chunks.length} pasajes · OCR ${parsed.ocrPages}/${parsed.pages} páginas${parsed.unreadablePages?` · ${parsed.unreadablePages} pendientes`:''}`;
       success.push(f)
     }catch(err){entry.status='incidencia';entry.note='Pendiente de reintento: '+err.message}
     handled.push(entry);done++;$('pdfIndexProgress').textContent=`Novedades ${done}/${changes.length} · ${f.name}`;
     await new Promise(resolve=>setTimeout(resolve,0))
   }}
   // No registrar como procesadas las actas fallidas ni interrumpidas.
   const trackerSave=files.filter(f=>(!!historicalDoc(f)||hasPdfChunks(f))&&!handled.some(h=>h.id===f.id&&h.status==='incidencia'));
   storedTracker(trackerSave,previous);
   state.changeLog=[...handled,...state.changeLog].slice(0,120);
   if(state.semantic&&state.pdfChunks.length&&success.length)await indexPdfVectors();
   const label=`${files.length} PDF consultados · ${changes.length} pendientes/nuevos · ${success.length} indexados${handled.filter(h=>h.status==='incidencia').length?' · hay incidencias que requieren revisión':''}.`;
   $('pdfIndexProgress').textContent=label+(success.length?' · Publica los cambios en el índice SharePoint para compartirlos.':'');reportRender();renderPrivateIndexStatus();
   try{await savePrivateCache()}catch(e){$('pdfIndexProgress').textContent+=` No se pudo guardar caché local: ${e.message}. Exporta tu índice privado.`}
   if(state.remoteIndexDirty&&SETTINGS.autoPublishNewActas){
     try{await publishRemoteIndex({automatic:true});$('pdfIndexProgress').textContent+=' Índice central actualizado automáticamente.'}
     catch(e){console.warn('La actualización automática requiere permiso de edición/consentimiento M365:',e.message);$('pdfIndexProgress').textContent+=' Índice central pendiente de publicación (puedes pulsar Publicar novedades).';renderPrivateIndexStatus();}
   }
   if(!automatic||changes.length)toast(label)
 }catch(e){$('pdfIndexProgress').textContent='No se completó el control de novedades: '+e.message;toast(e.message)}
 finally{state.scanRunning=false;$('pdfCancelBtn').classList.add('hide')}
}

// Importar un catálogo JSON privado. Nunca se conserva en LocalStorage ni se transmite a GitHub.

function applyPrivateIndexData(data,origin='Índice privado'){
 if(data?.schema!=='fias-directorio-private-v07'||!Array.isArray(data.rows)||!Array.isArray(data.pdfChunks)||!Array.isArray(data.documents)){
  throw Error('El archivo no corresponde al índice FIAS v07 o está incompleto.');
 }
 const rows=data.rows.map((r,i)=>({id:r.id||`privado-${i}`,year:String(r.year||''),date:r.date||'',
  act:r.actTitle||'',actId:r.actId||'',session:/extraordinaria/i.test(r.actTitle||'')?'Extraordinaria':/ordinaria/i.test(r.actTitle||'')?'Ordinaria':'',
  subject:r.subject||'',resolution:r.resolution||'',program:rowTags(`${r.subject} ${r.resolution} ${r.actTitle}`),
  amount:'',documentUrl:safeUrl(r.shareUrl||''),documentName:r.fileName||'',sheet:r.sheet||'',rowNumber:r.rowNumber||0}));
 if(!state.rows.length)putData({rows,acts:new Set(rows.map(x=>x.actId)).size,warnings:[]},origin);
 state.pdfChunks=data.pdfChunks.map((r,i)=>({id:r.id||`pdf-${i}`,file:r.fileName||r.file||'Documento',fileId:r.fileId||'',year:String(r.year||''),
  page:r.page||0,lastPage:r.lastPage||r.page||0,text:r.text||r.content||'',documentUrl:safeUrl(r.shareUrl||r.documentUrl||'')}));
 state.privateIndexLoaded=true;state.baselineStats=data.stats||{};state.baselineDocuments=data.documents||[];
 state.pdfVectors=null;reportRender();renderPrivateIndexStatus();
 $('pdfIndexProgress').textContent=`Índice disponible: ${rows.length} registros de matriz y ${state.pdfChunks.length} pasajes PDF; ${(data.stats?.pagesNeedingOCR||0)} páginas pendientes de OCR.`;
}
async function loadPrivateIndex(file){
 if(file.size>35*1024*1024)throw Error('Índice superior a 35 MB.');
 let data;try{data=JSON.parse((await file.text()).replace(/^\uFEFF/,''))}catch{throw Error('El archivo seleccionado no contiene JSON válido.')}
 applyPrivateIndexData(data,'Respaldo privado local');
 state.remoteIndexDirty=!!state.privateIndexUrl;state.remoteIndexFromSharePoint=false;
 try{await savePrivateCache()}catch(e){toast('No se pudo guardar el caché local: '+e.message)}
 renderPrivateIndexStatus();toast('Respaldo privado importado en este navegador.');
 if(state.account)await detectNewActas({automatic:true});
}
function downloadFile(name,mime,text){const u=URL.createObjectURL(new Blob([text],{type:mime}));const a=document.createElement('a');a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),3000)}
function exportChangesReport(){
 const head=['Fecha de control','Año','Archivo','Tamaño bytes','Estado en esta revisión','Resultado de procesamiento','Enlace SharePoint','Capacidad de lectura'];
 const now=new Date().toISOString(),statusMap=new Map(state.changeLog.map(e=>[e.id||'',e]));
 const baseline=new Map((state.baselineDocuments||[]).map(x=>[`${x.year}|${x.fileName}`,x]));
 const inventory=state.repository.length?state.repository:(state.baselineDocuments||[]).map(d=>({id:d.path,name:d.fileName,year:d.year,size:d.bytes||'',url:d.shareUrl||''}));
 const rows=inventory.map(f=>{
   const last=statusMap.get(f.id)||{},historical=baseline.get(`${f.year}|${f.name}`),chunks=state.pdfChunks.filter(x=>x.file===f.name&&String(x.year)===String(f.year)).length;
   const lectura=historical?`${historical.readablePages}/${historical.pages} páginas con texto`:(chunks?`${chunks} pasajes indexados`:'Por comprobar / PDF escaneado');
   return [now,f.year,f.name,f.size,last.status||'Existente (sin cambios)',last.note||'Inventariado en SharePoint',f.url||historical?.shareUrl||'',lectura];
 });
 downloadFile('FIAS_Reporte_integral_actas.csv','text/csv;charset=utf-8','\ufeff'+[head,...rows].map(r=>r.map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(';')).join('\r\n'));
 toast('Reporte integral exportado: incluye el histórico visible y las actas nuevas o actualizadas.');
}

function exportCurrentIndex(){
 if(!state.account){toast('Conecta Microsoft 365 antes de exportar un índice institucional.');return}
 downloadFile('FIAS_INDICE_PRIVADO_ACTUALIZADO.json','application/json;charset=utf-8',JSON.stringify(currentPrivateIndexPayload()));
 toast('Copia de respaldo del índice descargada. No sustituye la publicación en SharePoint.');
}
async function ask(question){
 question=String(question||'').trim();if(!question)return;
 if(!state.rows.length){toast('Sincroniza la matriz o importa el índice privado para consultar.');return}
 if(state.semanticBusy){toast('Indexación semántica en curso.');return}
 if($('chatThread').querySelector('.chat-empty'))$('chatThread').innerHTML='';
 sendChat(question,'','user');$('aiQuery').value='';
 const node=sendChat('','<p class="rag-loading">Recuperando antecedentes y fuentes documentales locales…</p>','assistant');
 try{
  const {list,mode}=await retrieve(question),pdfHits=await pdfEvidence(question);
  node.innerHTML='<div class="assistant-title">◈ &nbsp; SÍNTESIS EXTRACTIVA CON REFERENCIAS</div>'+extractiveAnswer(question,list,mode)+renderPdfEvidence(pdfHits)+
   '<div class="answer-footer">Sin servicios Azure AI/OpenAI. La síntesis extrae texto de fuentes, no genera afirmaciones no verificadas. Las páginas escaneadas sin texto siguen pendientes de OCR.</div>';
 }catch(e){node.innerHTML='<div class="assistant-title">Consulta incompleta</div><p class="rag-error">'+esc(e.message)+'</p>'}
 node.scrollIntoView({block:'nearest',behavior:'smooth'});
}
async function runHybridSearch(){
 const q=$('searchQuery').value.trim();if(!q){toast('Escribe tu consulta.');return}
 const box=$('ragRemoteResults');box.innerHTML='<div class="rag-loading">◈ Evaluando índice documental local…</div>';
 try{
  const {list,mode}=await retrieve(q),pdfHits=await pdfEvidence(q);
  box.innerHTML=`<div class="rag-results-summary"><b>Consulta local · ${esc(mode)}</b><span>${list.length} referencias</span><button id="clearHybrid" class="text-btn">Cerrar</button></div>`+extractiveAnswer(q,list,mode)+renderPdfEvidence(pdfHits);
  $('clearHybrid').onclick=()=>box.innerHTML='';
 }catch(e){box.innerHTML='<div class="rag-error">'+esc(e.message)+'</div>'}
}
function downloadResults(){const entries=ranked(state.search).map(({r})=>r);if(!entries.length){toast('No existen resultados para exportar.');return}const cols=['id','year','date','act','subject','resolution','program','documentUrl','sheet','rowNumber'];const csv=[cols.join(';'),...entries.map(r=>cols.map(c=>`"${clean(c==='program'?r[c].join(', '):r[c]).replace(/"/g,'""')}"`).join(';'))].join('\r\n');const url=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download='FIAS_Directorio_consulta.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000)}
function events(){document.querySelectorAll('[data-page]').forEach(b=>b.addEventListener('click',()=>setPage(b.dataset.page)));document.querySelectorAll('[data-go]').forEach(b=>b.addEventListener('click',()=>setPage(b.dataset.go)));$('menu').onclick=()=>$('sidebar').classList.toggle('open');$('btnLogin').onclick=()=>login().catch(e=>toast(e.message));$('btnLogout').onclick=()=>logout().catch(e=>toast(e.message));$('btnSync').onclick=()=>sync().catch(e=>{connection('Error de conexión',false);toast(e.message)});$('importBtn').onclick=()=>$('fileInput').click();$('fileInput').onchange=async(e)=>{const file=e.target.files?.[0];if(!file)return;try{await excelReader();const parsed=parseExcel(await file.arrayBuffer());putData(parsed,'Importación local');connection('Excel local',false);toast(`Archivo procesado: ${num.format(parsed.rows.length)} registros.`)}catch(err){toast('Error al importar: '+err.message)}e.target.value=''};$('quickForm').onsubmit=e=>{e.preventDefault();setPage('buscar');$('searchQuery').value=$('quickQuery').value;state.visible=12;renderResults()};document.querySelectorAll('[data-suggest]').forEach(b=>b.onclick=()=>{$('quickQuery').value=b.dataset.suggest;$('quickForm').requestSubmit()});$('searchGo').onclick=()=>{state.visible=12;renderResults()};$('searchQuery').oninput=()=>{state.visible=12;renderResults()};['filterYear','filterFund','filterSession'].forEach(id=>$(id).onchange=()=>{state.visible=12;renderResults()});$('clearFilters').onclick=()=>{$('filterYear').value='';$('filterFund').value='';$('filterSession').value='';$('searchQuery').value='';state.visible=12;renderResults()};$('loadMore').onclick=()=>{state.visible+=12;renderResults()};$('results').addEventListener('click',e=>{const b=e.target.closest('[data-id]');if(b)showRecord(b.dataset.id)});$('recentRows').addEventListener('click',e=>{const b=e.target.closest('[data-id]');if(b)showRecord(b.dataset.id)});$('chatThread').addEventListener('click',e=>{const b=e.target.closest('[data-id]');if(b)showRecord(b.dataset.id)});$('closeDrawer').onclick=closeDrawer;$('drawerBackdrop').onclick=closeDrawer;document.addEventListener('keydown',e=>{if(e.key==='Escape')closeDrawer();if(e.key==='Enter'&&e.ctrlKey&&document.activeElement===$('aiQuery'))$('aiForm').requestSubmit()});$('aiForm').onsubmit=e=>{e.preventDefault();ask($('aiQuery').value)};document.querySelectorAll('[data-prompt]').forEach(b=>b.onclick=()=>{$('aiQuery').value=b.dataset.prompt;$('aiForm').requestSubmit()});$('semanticBtn').onclick=()=>state.semantic?indexSemantic().then(()=>indexPdfVectors()):activateSemantic();$('pdfIndexBtn').onclick=()=>detectNewActas({});$('pdfCancelBtn').onclick=()=>{state.pdfCancel=true;$('pdfIndexProgress').textContent='Solicitando cancelación…'};$('hybridSearchBtn').onclick=runHybridSearch;$('exportBtn').onclick=downloadResults;$('themeBtn').onclick=()=>{const dark=document.documentElement.classList.toggle('dark');localStorage.setItem('fias_theme',dark?'dark':'light')};$('btnChanges').onclick=()=>sync().catch(e=>toast(e.message));$('btnReport').onclick=exportChangesReport;$('btnIndexImport').onclick=()=>$('indexInput').click();$('indexInput').onchange=async e=>{const f=e.target.files?.[0];if(!f)return;try{await loadPrivateIndex(f)}catch(err){toast(err.message)}e.target.value=''};$('btnIndexExport').onclick=exportCurrentIndex;$('btnRemoteRefresh').onclick=()=>loadRemotePrivateIndex({force:true}).catch(err=>toast(err.message));$('btnIndexPush').onclick=()=>publishRemoteIndex().catch(err=>toast(err.message))}
function closeDrawer(){$('recordDrawer').classList.remove('open');$('drawerBackdrop').classList.add('hide')}
function init(){
 $('matrixLink').href=SETTINGS.matrixUrl;$('folderLink').href=SETTINGS.folderUrl;
 if(localStorage.getItem('fias_theme')==='dark')document.documentElement.classList.add('dark');
 state.privateIndexUrl=configuredPrivateIndexUrl();events();updateDashboard();renderAnalytics();reportRender();renderPrivateIndexStatus();
 if(/^https?:$/.test(location.protocol))client().then(async()=>{
   if(state.account){try{await restorePrivateCache();await restorePrivateIndexUrl();await sync()}catch(e){console.warn('La sincronización M365 requiere interacción:',e.message);toast('No se pudo conectar SharePoint: '+e.message)}}
 }).catch(e=>console.warn('FIAS: No se pudo restaurar la sesión M365.',e));
  // Sin vigilancia. La matriz y el índice privado se comprueban al iniciar sesión o pulsar Sincronizar.
}
init();
