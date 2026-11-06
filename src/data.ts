import * as XLSX from 'xlsx';
export type RecordRow={id:string;year:string;date:string;act:string;actId:string;session:string;resolution:string;subject:string;program:string;amount:string;members:string;keywords:string;documentUrl:string;documentName:string;sheet:string;rowNumber:number;raw:Record<string,string>};
const clean=(x:unknown)=>String(x??'').trim();
export function fold(v:string){return String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLocaleLowerCase('es');}
export function safeUrl(v:string){try{const u=new URL(v);return u.protocol==='https:'&&(u.hostname==='fiasec-my.sharepoint.com'||u.hostname.endsWith('.sharepoint.com'))?u.toString():'';}catch{return '';}}
function normalizeLink(value:string){if(value.startsWith('https://'))return safeUrl(value);const match=value.match(/\/:([bfx]):\/g\/personal\/[^\s"<>]+/);return match?safeUrl('https://fiasec-my.sharepoint.com'+match[0]):'';}
function dateText(value:unknown){if(typeof value==='number'&&value>20000&&value<70000){const d=XLSX.SSF.parse_date_code(value);if(d)return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;}return clean(value);}
const funds=['FEIG','REM','PASF','FAP','PASNAP','CONSERVA AVES','BIOECONOMÍA','SOCIO BOSQUE','GALÁPAGOS','KFW'];
function tags(text:string){return funds.filter(t=>new RegExp(`(^|[^\\p{L}])${t}($|[^\\p{L}])`,'iu').test(text)).join(', ');}
function hlink(ws:XLSX.WorkSheet,addr:string){const cell=ws[addr];return normalizeLink(clean(cell?.l?.Target)||clean(cell?.v));}
export async function parseWorkbook(file:File):Promise<{rows:RecordRow[];warnings:string[];acts:number}>{
 const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false,cellFormula:true});const rows:RecordRow[]=[];const warnings:string[]=[];
 for(const sheet of wb.SheetNames){const ws=wb.Sheets[sheet];const mat=XLSX.utils.sheet_to_json<unknown[]>(ws,{header:1,defval:'',raw:true}) as unknown[][];
 const header=mat.findIndex(r=>fold(r.map(clean).join(' ')).includes('puntos a tratar')&&fold(r.map(clean).join(' ')).includes('resolucion'));
 if(header<0){warnings.push(`Hoja ${sheet}: no se encontró la cabecera de puntos y resoluciones.`);continue;}
 let current:{date:string;title:string;actId:string;documentUrl:string;documentName:string;members:RecordRow[]}|null=null;let count=0;
 for(let i=header+1;i<mat.length;i++){
  const [d,t,p,r,f]=[0,1,2,3,4].map(j=>mat[i]?.[j]);const date=clean(d),title=clean(t),subject=clean(p),resolution=clean(r),filename=clean(f);
  if(!date&&!title&&!subject&&!resolution&&!filename)continue;
  if(date||title){count++;current={date:dateText(d),title,actId:`${sheet}-${count}`,documentUrl:'',documentName:'',members:[]};}
  if(!current)continue;
  const url=hlink(ws,`E${i+1}`);if(url)current.documentUrl=url;if(filename)current.documentName=filename;
  // Los enlaces del acta suelen aparecer en una fila posterior a los asuntos: propagar a todas sus resoluciones.
  for(const prev of current.members){if(current.documentUrl)prev.documentUrl=current.documentUrl;if(current.documentName)prev.documentName=current.documentName;}
  if(!subject&&!resolution)continue;
  const year=(current.date.match(/20\d{2}/)||sheet.match(/20\d{2}/))?.[0]||'';
  const tag=tags([subject,resolution,current.title].join(' '));
  const item:RecordRow={id:`${sheet}-${i+1}`,year,date:current.date,act:current.title,actId:current.actId,session:/extraordinaria/i.test(current.title)?'Extraordinaria':/ordinaria/i.test(current.title)?'Ordinaria':'',resolution,subject,program:tag,amount:'',members:'',keywords:'',documentUrl:current.documentUrl,documentName:current.documentName,sheet,rowNumber:i+1,raw:{FECHA:current.date,'TÍTULO / TIPO DE SESIÓN':current.title,'PUNTOS A TRATAR':subject,'RESOLUCIÓN':resolution,'LINK DE ACCESO':current.documentName}};
  rows.push(item);current.members.push(item);
 }
 }
 return {rows,warnings,acts:new Set(rows.map(r=>r.actId)).size};
}
