import { PublicClientApplication, InteractionRequiredAuthError, type AccountInfo } from '@azure/msal-browser';
export const SETTINGS={
  tenantId:import.meta.env.VITE_ENTRA_TENANT_ID || '5e23e4af-237d-4d97-bf6e-dca808015787',
  clientId:import.meta.env.VITE_ENTRA_CLIENT_ID || 'c8a828c1-3a20-4876-96af-b4f30ce4abeb',
  matrixUrl:'https://fiasec-my.sharepoint.com/:x:/g/personal/dvillacis_fias_org_ec/IQCho4wt1qyXT6DnsDAU-r33AavyNjkIhd9xhXy5QgQvLkk',
  folderUrl:'https://fiasec-my.sharepoint.com/:f:/g/personal/dvillacis_fias_org_ec/IgAPa6xunze7Rp5Sbb3eoG4bAbnTiJ3TP2iPh5nhaefQklY?e=WH2jtK',
  scopes:['User.Read','Files.Read.All','Sites.Read.All']
};
let client:PublicClientApplication|undefined;
let current:AccountInfo|null=null;
const redirectUri=()=>window.location.origin+window.location.pathname;
async function app(){if(!client){client=new PublicClientApplication({auth:{clientId:SETTINGS.clientId,authority:`https://login.microsoftonline.com/${SETTINGS.tenantId}`,redirectUri:redirectUri()},cache:{cacheLocation:'sessionStorage'}});await client.initialize();await client.handleRedirectPromise();current=client.getAllAccounts()[0]||null;}return client;}
export async function restoreAccount(){const c=await app();current=c.getAllAccounts()[0]||null;return current;}
export async function signIn(){const c=await app();const res=await c.loginPopup({scopes:SETTINGS.scopes,prompt:'select_account'});current=res.account;return current;}
export async function signOut(){const c=await app();const a=current;current=null;if(a)await c.logoutPopup({account:a,mainWindowRedirectUri:redirectUri()});}
async function token(){const c=await app();const a=current||c.getAllAccounts()[0];if(!a)throw new Error('Inicie sesión con su cuenta institucional Microsoft 365.');try{return (await c.acquireTokenSilent({account:a,scopes:SETTINGS.scopes})).accessToken;}catch(e){if(e instanceof InteractionRequiredAuthError)return (await c.acquireTokenPopup({account:a,scopes:SETTINGS.scopes})).accessToken;throw e;}}
const shareId=(url:string)=>'u!'+btoa(unescape(encodeURIComponent(url))).replace(/=/g,'').replace(/\+/g,'-').replace(/\//g,'_');
async function graph(path:string){const access=await token();let response=await fetch(`https://graph.microsoft.com/v1.0${path}`,{headers:{Authorization:`Bearer ${access}`,'Accept':'application/json'}});if(response.status===429||response.status===503){const delay=Math.min(5000,Number(response.headers.get('Retry-After')||2)*1000);await new Promise(r=>setTimeout(r,delay));response=await fetch(`https://graph.microsoft.com/v1.0${path}`,{headers:{Authorization:`Bearer ${access}`}});}if(!response.ok){let msg=await response.text();try{msg=JSON.parse(msg)?.error?.message||msg;}catch{}throw new Error(`Microsoft Graph (${response.status}): ${msg.slice(0,220)}`);}return response.json();}
export type SourceMeta={name:string;modified:string;size:number;webUrl:string;id:string;driveId:string};
export async function fetchMatrix():Promise<{file:File;metadata:SourceMeta}>{
  const item=await graph(`/shares/${shareId(SETTINGS.matrixUrl)}/driveItem?$select=id,name,size,lastModifiedDateTime,webUrl,parentReference,@microsoft.graph.downloadUrl`);
  if(!item?.id)throw new Error('No se pudo resolver la matriz compartida. Compruebe acceso y permisos.');
  const metadata={name:item.name||'Matriz del Directorio.xlsx',modified:item.lastModifiedDateTime||'',size:item.size||0,webUrl:item.webUrl||SETTINGS.matrixUrl,id:item.id,driveId:item.parentReference?.driveId||''};
  // La URL preautenticada devuelta por Graph es temporal; nunca persistirla.
  let url=item['@microsoft.graph.downloadUrl'];
  if(!url && metadata.driveId){const fresh=await graph(`/drives/${encodeURIComponent(metadata.driveId)}/items/${encodeURIComponent(item.id)}?$select=id,@microsoft.graph.downloadUrl`);url=fresh['@microsoft.graph.downloadUrl'];}
  if(!url)throw new Error('Microsoft Graph no proporcionó una URL de descarga. Revise permisos Files.Read.All y políticas del tenant.');
  const response=await fetch(url,{cache:'no-store'});
  if(!response.ok)throw new Error(`Error al descargar Excel desde SharePoint (${response.status}).`);
  const array=await response.arrayBuffer();const bytes=new Uint8Array(array);
  if(bytes.length<4||bytes[0]!==80||bytes[1]!==75)throw new Error('SharePoint devolvió un archivo que no es un XLSX válido.');
  return {file:new File([array],metadata.name,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),metadata};
}
