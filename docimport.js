/* docimport.js — extraction de texte depuis un document d'exercice ou de sujet, 100 % locale (RGPD).
   Formats : .docx .odt .pdf .xlsx .ods .pptx .odp .pages/.numbers/.key (aperçu) .rtf .html .txt .md .csv .json
   Dépend de JSZip (ZIP) et pdf.js (PDF) chargés avant ce fichier.
   API : DocImport.extract(file, onProgress) → Promise<{text, format, details}> */
(function(root){
'use strict';
var MAX_BYTES=25*1024*1024;

var NAMED={lt:'<',gt:'>',quot:'"',apos:"'",amp:'&',nbsp:'\u00A0'};
function unxml(s){return s.replace(/&(lt|gt|quot|apos|amp|nbsp|#\d+|#x[0-9a-fA-F]+);/g,function(_,k){
 if(NAMED[k]!==undefined)return NAMED[k];
 try{return String.fromCodePoint(k[1]==='x'?parseInt(k.slice(2),16):+k.slice(1))}catch(e){return ''}})}
function clean(t){return t.replace(/\r\n?/g,'\n').replace(/\u00A0/g,' ').replace(/[ \t]+\n/g,'\n').replace(/\n{3,}/g,'\n\n').replace(/\u0000/g,'').trim()}
function fail(msg){return new Error(msg)}

/* ---------- décodage texte (UTF-8, repli Windows-1252) ---------- */
function decodeText(buf){
 try{return new TextDecoder('utf-8',{fatal:true}).decode(buf)}
 catch(e){return new TextDecoder('windows-1252').decode(buf)}}

/* ---------- Word (.docx) ---------- */
function docxText(xml){
 var out='',re=/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<m:t(?:\s[^>]*)?>([\s\S]*?)<\/m:t>|<w:tab\b[^>]*\/>|<w:(?:br|cr)\b[^>]*\/>|<\/w:p>|<\/w:tc>|<\/w:tr>/g,m;
 while((m=re.exec(xml))){
  var t=m[0];
  if(m[1]!==undefined)out+=unxml(m[1]);
  else if(m[2]!==undefined)out+=unxml(m[2]);
  else if(t.indexOf('<w:tab')===0)out+='\t';
  else if(t.indexOf('<w:br')===0||t.indexOf('<w:cr')===0)out+='\n';
  else if(t==='</w:p>')out+='\n';
  else if(t==='</w:tc>'){if(out.charAt(out.length-1)==='\n')out=out.slice(0,-1);out+='\t'}
  else if(t==='</w:tr>'){if(out.charAt(out.length-1)==='\t')out=out.slice(0,-1);out+='\n'}}
 return out}

/* ---------- OpenDocument (.odt .ods .odp) ---------- */
function odfText(xml){
 xml=xml.replace(/<office:(?:automatic-styles|font-face-decls|styles)\b[\s\S]*?<\/office:(?:automatic-styles|font-face-decls|styles)>/g,'')
  .replace(/<text:tracked-changes\b[\s\S]*?<\/text:tracked-changes>/g,'');
 return unxml(xml
  .replace(/<text:tab\s*\/>/g,'\t')
  .replace(/<text:line-break\s*\/>/g,'\n')
  .replace(/<text:s(?:\s+text:c="(\d+)")?\s*\/>/g,function(_,n){return ' '.repeat(+n||1)})
  .replace(/<\/text:(?:p|h)>/g,'\n')
  .replace(/<\/table:table-cell>/g,'\t')
  .replace(/<\/table:table-row>/g,'\n')
  .replace(/<\/draw:page>/g,'\n\n')
  .replace(/<[^>]+>/g,''))}

/* ---------- Excel (.xlsx) ---------- */
function attr(tag,n){var m=new RegExp('(?:^|\\s)'+n+'\\s*=\\s*(?:"([^"]*)"|\'([^\']*)\')').exec(tag);return m?unxml(m[1]!==undefined?m[1]:m[2]):null}
function sstrs(b){var o='',re=/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g,m;b=b.replace(/<rPh\b[\s\S]*?<\/rPh>/g,'');while((m=re.exec(b)))o+=unxml(m[1]);return o}
function colIdx(ref){var l=/^[A-Z]+/.exec(ref);if(!l)return -1;var n=0;for(var i=0;i<l[0].length;i++)n=n*26+l[0].charCodeAt(i)-64;return n-1}
async function xlsxText(zip,details){
 var rd=async function(n){var f=zip.file(n);return f?await f.async('string'):null};
 var wb=await rd('xl/workbook.xml'),rel=await rd('xl/_rels/workbook.xml.rels');
 if(!wb||!rel)throw fail('Classeur Excel illisible.');
 var targets={},m,re=/<Relationship\b[^>]*>/g;
 while((m=re.exec(rel))){var id=attr(m[0],'Id'),t=attr(m[0],'Target');if(id&&t)targets[id]=t.charAt(0)==='/'?t.slice(1):'xl/'+t}
 var ss=[],sx=await rd('xl/sharedStrings.xml');
 if(sx){var rs=/<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g;while((m=rs.exec(sx)))ss.push(sstrs(m[1]||''))}
 var out=[],rs2=/<sheet\b[^>]*>/g,nb=0;
 while((m=rs2.exec(wb))){
  var name=attr(m[0],'name'),path=targets[attr(m[0],'r:id')],x=path&&await rd(path);if(!x)continue;
  nb++;var lines=[],rows=0,rr=/<row\b[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g,rm;
  while((rm=rr.exec(x))){
   if(rows++>=1500){lines.push('[… feuille tronquée à 1500 lignes]');break}
   var cells=[],cr=/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g,cm;
   while((cm=cr.exec(rm[1]||''))){
    var ref=attr(cm[1],'r'),ty=attr(cm[1],'t'),b=cm[2]||'',ci=ref?colIdx(ref):cells.length,v='';
    if(ty==='inlineStr')v=sstrs(b);
    else{var vm=/<v>([\s\S]*?)<\/v>/.exec(b);if(vm){v=unxml(vm[1]);if(ty==='s')v=ss[+vm[1]]||'';else if(ty==='b')v=vm[1]==='1'?'VRAI':'FAUX'}}
    if(ci>=0)cells[ci]=v}
   var line=Array.from(cells,function(c){return c==null?'':String(c)}).join('\t').replace(/\s+$/,'');
   if(line)lines.push(line)}
  if(lines.length)out.push('### Feuille « '+name+' »\n'+lines.join('\n'))}
 details.sheets=nb;
 return out.join('\n\n')}

/* ---------- PowerPoint (.pptx) ---------- */
async function pptxText(zip,details){
 var names=Object.keys(zip.files).filter(function(n){return /^ppt\/slides\/slide\d+\.xml$/.test(n)})
  .sort(function(a,b){return +/(\d+)\.xml$/.exec(a)[1]-+/(\d+)\.xml$/.exec(b)[1]});
 var out=[];
 for(var i=0;i<names.length;i++){
  var x=await zip.file(names[i]).async('string'),s='',re=/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>|<\/a:p>/g,m;
  while((m=re.exec(x)))s+=m[1]!==undefined?unxml(m[1]):'\n';
  s=clean(s);if(s)out.push('### Diapositive '+(i+1)+'\n'+s)}
 details.slides=names.length;
 return out.join('\n\n')}

/* ---------- PDF ---------- */
async function pdfText(buf,onProgress,details){
 var lib=root.pdfjsLib;if(!lib)throw fail('Module PDF indisponible (fichier lib/pdf.min.js manquant).');
 lib.GlobalWorkerOptions.workerSrc=root.DOC_PDF_WORKER||'lib/pdf.worker.min.js';
 var pdf;
 try{pdf=await lib.getDocument({data:new Uint8Array(buf)}).promise}
 catch(e){
  if(e&&e.name==='PasswordException')throw fail('Ce PDF est protégé par un mot de passe : retirez la protection puis réessayez.');
  throw fail('PDF illisible ou corrompu.')}
 var pages=[];
 for(var i=1;i<=pdf.numPages;i++){
  var page=await pdf.getPage(i),tc=await page.getTextContent(),s='',lastY=null;
  for(var j=0;j<tc.items.length;j++){
   var it=tc.items[j];if(typeof it.str!=='string')continue;
   var y=it.transform?it.transform[5]:null;
   if(lastY!==null&&y!==null&&Math.abs(y-lastY)>2&&s&&s.charAt(s.length-1)!=='\n')s+='\n';
   s+=it.str;if(it.hasEOL)s+='\n';lastY=y}
  pages.push(s.trim());
  if(onProgress)onProgress(i,pdf.numPages)}
 details.pages=pdf.numPages;
 var text=pages.join('\n\n');
 if(!text.replace(/\s/g,''))throw fail("Aucun texte trouvé dans ce PDF : il s'agit probablement d'un scan (image). Utilisez un PDF « texte » ou copiez le texte à la main dans la zone ci-dessous.");
 return text}

/* ---------- RTF (extraction simple) ---------- */
function rtfText(s){
 var out='',skip=[false],i=0,n=s.length,SKIP=/^\\(\*|fonttbl|colortbl|stylesheet|info|pict|header|footer)/;
 while(i<n){
  var c=s.charAt(i),top=skip[skip.length-1];
  if(c==='{'){skip.push(top);i++;if(s.charAt(i)==='\\'&&SKIP.test(s.substr(i,14)))skip[skip.length-1]=true}
  else if(c==='}'){if(skip.length>1)skip.pop();i++}
  else if(c==='\\'){
   var m=/^\\([a-zA-Z]+)(-?\d+)? ?/.exec(s.substr(i,40));
   if(m){var w=m[1];
    if(!top){
     if(w==='par'||w==='line'||w==='sect')out+='\n';
     else if(w==='tab')out+='\t';
     else if(w==='u'&&m[2]!==undefined){var code=+m[2];out+=String.fromCharCode(code<0?code+65536:code)}}
    i+=m[0].length;if(w==='u')i++}
   else{var nx=s.charAt(i+1);
    if(nx==="'"){if(!top)out+=String.fromCharCode(parseInt(s.substr(i+2,2),16));i+=4}
    else if(nx==='\\'||nx==='{'||nx==='}'){if(!top)out+=nx;i+=2}
    else i+=2}}
  else{if(c!=='\n'&&c!=='\r'&&!top)out+=c;i++}}
 return out}

/* ---------- HTML ---------- */
function htmlText(s){
 return unxml(s.replace(/<(script|style|head)\b[\s\S]*?<\/\1>/gi,'')
  .replace(/<br\s*\/?>/gi,'\n').replace(/<\/(p|div|li|tr|h[1-6]|table|section|article|ul|ol)>/gi,'\n')
  .replace(/<\/t[dh]>/gi,'\t').replace(/<[^>]+>/g,''))}

/* ---------- Apple iWork (.pages .numbers .key) : aperçu PDF intégré ---------- */
async function iworkText(zip,onProgress,details,ext){
 var key=Object.keys(zip.files).filter(function(n){return /(^|\/)preview\.pdf$/i.test(n)})[0];
 if(!key)throw fail('Ce fichier Apple ('+ext+') ne contient pas d\'aperçu lisible. Dans l\'application : Fichier ▸ Exporter vers ▸ PDF (ou Word), puis importez le fichier exporté.');
 var buf=await zip.file(key).async('arraybuffer');
 details.note='Fichier Apple : seul l\'aperçu intégré est lu, il peut être incomplet. Pour un résultat fiable, exportez le document en PDF ou Word depuis Pages.';
 return pdfText(buf,onProgress,details)}

/* ---------- détection du type ---------- */
function extOf(name){var m=/\.([a-z0-9]+)$/i.exec(name||'');return m?m[1].toLowerCase():''}
function sniff(u8){
 if(u8.length>=4&&u8[0]===0x25&&u8[1]===0x50&&u8[2]===0x44&&u8[3]===0x46)return 'pdf';
 if(u8.length>=2&&u8[0]===0x50&&u8[1]===0x4B)return 'zip';
 if(u8.length>=8&&u8[0]===0xD0&&u8[1]===0xCF&&u8[2]===0x11&&u8[3]===0xE0)return 'ole';
 return ''}

var LABEL={docx:'Word (.docx)',odt:'OpenDocument texte (.odt)',pdf:'PDF',xlsx:'Excel (.xlsx)',ods:'OpenDocument tableur (.ods)',pptx:'PowerPoint (.pptx)',odp:'OpenDocument présentation (.odp)',
 pages:'Apple Pages',numbers:'Apple Numbers',key:'Apple Keynote',rtf:'RTF',html:'HTML',txt:'Texte',md:'Markdown',csv:'CSV',json:'JSON'};

async function extract(file,onProgress){
 if(!file)throw fail('Aucun fichier.');
 if(file.size>MAX_BYTES)throw fail('Fichier trop volumineux (maximum 25 Mo).');
 var buf=await file.arrayBuffer(),u8=new Uint8Array(buf),ext=extOf(file.name),kind=sniff(u8),details={},text='',fmt=ext;

 if(kind==='ole'||ext==='doc'||ext==='xls'||ext==='ppt'){
  if(kind==='ole'||ext==='doc'||ext==='xls'||ext==='ppt')
   throw fail("Ancien format Microsoft Office (.doc, .xls, .ppt) non pris en charge. Ouvrez le fichier puis enregistrez-le en .docx / .xlsx / .pptx ou en PDF.")}

 if(kind==='pdf'||ext==='pdf'){fmt='pdf';text=await pdfText(buf,onProgress,details)}
 else if(kind==='zip'||['docx','odt','ods','odp','xlsx','pptx','pages','numbers','key'].indexOf(ext)>=0){
  if(!root.JSZip)throw fail('Module ZIP indisponible (fichier lib/jszip.min.js manquant).');
  var zip;try{zip=await root.JSZip.loadAsync(buf)}catch(e){throw fail('Archive illisible ou corrompue.')}
  var has=function(n){return !!zip.file(n)};
  if(has('word/document.xml')){fmt='docx';text=docxText(await zip.file('word/document.xml').async('string'))}
  else if(has('xl/workbook.xml')){fmt='xlsx';text=await xlsxText(zip,details)}
  else if(has('ppt/presentation.xml')){fmt='pptx';text=await pptxText(zip,details)}
  else if(has('content.xml')&&has('mimetype')){
   var mt=(await zip.file('mimetype').async('string')).trim();
   fmt=/spreadsheet/.test(mt)?'ods':/presentation/.test(mt)?'odp':'odt';
   text=odfText(await zip.file('content.xml').async('string'))}
  else if(['pages','numbers','key'].indexOf(ext)>=0||Object.keys(zip.files).some(function(n){return /^(index\.xml|Index\/|QuickLook\/)/i.test(n)})){
   if(['pages','numbers','key'].indexOf(ext)<0)ext='pages';fmt=ext;text=await iworkText(zip,onProgress,details,'.'+ext)}
  else throw fail("Type d'archive non reconnu. Formats acceptés : Word, PDF, Excel, PowerPoint, OpenDocument, Pages, texte.")}
 else if(ext==='rtf'||/^\{\\rtf/.test(decodeText(u8.subarray(0,10)))){fmt='rtf';text=rtfText(decodeText(u8))}
 else if(ext==='html'||ext==='htm'){fmt='html';text=htmlText(decodeText(u8))}
 else if(['txt','md','markdown','csv','tsv','json','text'].indexOf(ext)>=0||!ext){fmt=ext==='markdown'?'md':(ext||'txt');text=decodeText(u8)}
 else throw fail('Format « .'+ext+' » non pris en charge. Formats acceptés : .docx .pdf .xlsx .pptx .odt .ods .odp .pages .rtf .html .txt .md .csv');

 text=clean(text);
 if(!text)throw fail('Aucun texte exploitable trouvé dans ce document.');
 return{text:text,format:LABEL[fmt]||fmt,details:details}}

root.DocImport={extract:extract,clean:clean,_docxText:docxText,_odfText:odfText,_rtfText:rtfText,_htmlText:htmlText,ACCEPT:'.docx,.pdf,.xlsx,.pptx,.odt,.ods,.odp,.pages,.numbers,.key,.rtf,.html,.htm,.txt,.md,.csv,.tsv,.json'};
})(typeof window!=='undefined'?window:globalThis);
