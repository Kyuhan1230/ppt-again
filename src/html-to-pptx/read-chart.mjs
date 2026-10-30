import path from 'node:path';
import {strFromU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
import {diagnostic} from './model.mjs';
import {validateChart} from './charts.mjs';
import {readChartWorkbook} from './chart-workbook.mjs';

const C='http://schemas.openxmlformats.org/drawingml/2006/chart',A='http://schemas.openxmlformats.org/drawingml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',REL='http://schemas.openxmlformats.org/package/2006/relationships';
const CHART='application/vnd.openxmlformats-officedocument.drawingml.chart+xml',XLSX='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const attrs=(n,k,fallback=null)=>n?.hasAttribute(k)?n.getAttribute(k):fallback;
const reject=(message,code='UNSUPPORTED_OBJECT')=>{throw Object.assign(new Error(message),{code});};
const requireValue=(condition,message,code='UNSUPPORTED_OBJECT')=>{if(!condition)reject(message,code);};
const children=(n,k,ns=C)=>elements(n).filter(c=>c.localName===k&&c.namespaceURI===ns);
function one(n,k,required=false,ns=C){const nodes=children(n,k,ns);requireValue(nodes.length<2&&(!required||nodes.length===1),'차트 필수 항목이 없거나 중복됩니다.','INPUT_INVALID');return nodes[0]??null;}
function allowed(n,names,ns=C){requireValue(elements(n).every(c=>c.namespaceURI===ns&&names.includes(c.localName)),'지원 범위 밖의 차트 구조·서식이 있습니다.');}
function numeric(value){requireValue(typeof value==='string'&&/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(value.trim())&&Number.isFinite(Number(value)),'빈값·유한 수치가 아닌 차트 값은 읽을 수 없습니다.','CHART_DATA_MISMATCH');return Number(value);}
const number=(n,k,fallback=null)=>numeric(attrs(n,k,fallback===null?null:String(fallback)));
function index(n,k='val',max=1000){const value=attrs(n,k);requireValue(/^\d+$/.test(value??'')&&Number(value)<=max,'차트 인덱스가 잘못되었습니다.','CHART_DATA_MISMATCH');return Number(value);}
function flag(n,fallback=false){const value=attrs(n,'val',fallback?'1':'0');requireValue(['0','1','false','true'].includes(value),'차트 표시 값이 잘못되었습니다.','INPUT_INVALID');return ['1','true'].includes(value);}
const exact=(n,k,value,fallback=value)=>requireValue(attrs(n,k,fallback)===value,'변경된 차트 배치·서식은 아직 지원하지 않습니다.');
function extensions(n){
  if(!n)return;allowed(n,['ext']);
  for(const ext of elements(n)){
    const uri=attrs(ext,'uri'),items=elements(ext);
    requireValue(uri==='{C3380CC4-5D6E-409C-BE32-E72D297353CC}'&&items.length===1&&items[0].namespaceURI==='http://schemas.microsoft.com/office/drawing/2014/chart'&&items[0].localName==='uniqueId'&&!elements(items[0]).length||uri==='{CE6537A1-D6FC-4f65-9D91-7224C49458BB}'&&!items.length,'확장 차트 데이터·서식은 아직 지원하지 않습니다.');
  }
}
function rgb(fill,alpha=100000){
  requireValue(!!fill,'명시적인 차트 색이 없습니다.');allowed(fill,['srgbClr'],A);const c=one(fill,'srgbClr',true,A),hex=attrs(c,'val');
  requireValue(/^[\da-f]{6}$/i.test(hex??''),'이 차트 색은 명시적인 RGB가 아닙니다.');allowed(c,['alpha'],A);
  requireValue(number(one(c,'alpha',false,A),'val',100000)===alpha,'반투명·색 변환 차트 서식은 아직 지원하지 않습니다.');return '#'+hex.toLowerCase();
}
function area(n){
  if(!n)return;allowed(n,['solidFill','noFill','ln','effectLst'],A);
  const fill=one(n,'solidFill',false,A);if(fill)rgb(fill,0);
  const ln=one(n,'ln',false,A);if(ln){allowed(ln,['noFill'],A);requireValue(!!one(ln,'noFill',false,A),'차트 영역의 테두리는 아직 지원하지 않습니다.');}
  requireValue(!elements(one(n,'effectLst',false,A)).length,'차트 효과는 아직 지원하지 않습니다.');
}
function line(n,color,width){
  requireValue(!!n,'차트 선 서식이 없습니다.');allowed(n,['solidFill','prstDash','round','miter','bevel','headEnd','tailEnd'],A);
  requireValue(rgb(one(n,'solidFill',true,A))===color&&Math.abs(number(n,'w')-width)<=1,'변경된 차트 선 색·굵기는 아직 지원하지 않습니다.');
  exact(n,'cap','flat');exact(n,'cmpd','sng');exact(n,'algn','ctr');
  exact(one(n,'prstDash',false,A),'val','solid');for(const k of ['headEnd','tailEnd'])exact(one(n,k,false,A),'type','none');
}
function noLayout(n){if(n)requireValue(!elements(n).length,'수동으로 배치한 차트 요소는 아직 지원하지 않습니다.');}

// This reader retains one common visible typeface/size. Per-label overrides
// that merely replace inactive Office defaults are resolved before comparison.
function textReader(){
  const styles=[];
  const style=(pr,def=null,color='#223845')=>{
    const get=k=>attrs(pr,k,attrs(def,k)),part=k=>one(pr,k,false,A)??one(def,k,false,A);
    for(const p of [pr,def].filter(Boolean))allowed(p,['solidFill','latin','ea','cs'],A);
    requireValue([null,'0','false'].includes(get('b'))&&[null,'0','false'].includes(get('i'))&&[null,'none'].includes(get('u'))&&[null,'noStrike'].includes(get('strike'))&&[null,'none'].includes(get('cap'))&&Number(get('baseline')??0)===0&&Number(get('spc')??0)===0,'개별 차트 글자 서식은 아직 지원하지 않습니다.');
    const family=attrs(part('latin'),'typeface'),size=numeric(get('sz'))/100;
    requireValue(typeof family==='string'&&family.trim()&&family.length<=256&&family.isWellFormed()&&!/[\u0000-\u001f]/.test(family)&&!family.startsWith('+')&&size>0&&rgb(part('solidFill'))===color,'차트의 명시적인 공통 글꼴·색이 필요합니다.');
    for(const k of ['ea','cs'])requireValue(!attrs(part(k),'typeface')||attrs(part(k),'typeface')===family,'서로 다른 문자권 글꼴은 아직 지원하지 않습니다.');
    styles.push({family,size});
  };
  const body=n=>{allowed(n,['bodyPr','lstStyle','p'],A);const b=one(n,'bodyPr',true,A);requireValue(number(b,'rot',0)===0&&attrs(b,'vert','horz')==='horz'&&!elements(b).length&&!elements(one(n,'lstStyle',false,A)).length,'회전·상속 차트 글자 서식은 아직 지원하지 않습니다.');};
  const txPr=(n,color='#223845')=>{
    requireValue(!!n,'명시적인 차트 글자 서식이 없습니다.');body(n);const paragraphs=children(n,'p',A);requireValue(paragraphs.length===1,'복잡한 차트 문단은 아직 지원하지 않습니다.');
    const p=paragraphs[0];allowed(p,['pPr','endParaRPr'],A);const pp=one(p,'pPr',true,A);allowed(pp,['defRPr'],A);style(one(pp,'defRPr',true,A),null,color);
  };
  const title=n=>{
    if(!n)return '';allowed(n,['tx','layout','overlay','spPr','txPr','extLst']);noLayout(one(n,'layout'));requireValue(!flag(one(n,'overlay')),'겹쳐 배치한 차트 제목은 지원하지 않습니다.');area(one(n,'spPr'));extensions(one(n,'extLst'));
    const tx=one(n,'tx',true);allowed(tx,['rich']);const rich=one(tx,'rich',true);body(rich);const ps=children(rich,'p',A);requireValue(ps.length===1,'여러 문단 차트 제목은 지원하지 않습니다.');
    const p=ps[0],pp=one(p,'pPr',false,A),def=one(pp,'defRPr',false,A);allowed(p,['pPr','r','endParaRPr'],A);if(pp)allowed(pp,['defRPr'],A);
    let text='';for(const r of children(p,'r',A)){allowed(r,['rPr','t'],A);style(one(r,'rPr',false,A),def);text+=one(r,'t',true,A).textContent;}
    requireValue(text.length>0,'빈 차트 제목은 지원하지 않습니다.');return text;
  };
  return {style,txPr,title,finish(){requireValue(styles.length>0&&styles.every(s=>s.family===styles[0].family&&s.size===styles[0].size),'차트 글자의 공통 글꼴·크기가 서로 다릅니다.');return {fontFamily:styles[0].family,fontSize:styles[0].size*4/3,fontWeight:400,fontStyle:'normal',color:'#223845',opacity:1};}};
}

function range(formula,workbook){
  requireValue(typeof formula==='string'&&formula.length<=512&&!/[\[\]\\]/.test(formula),'외부·복잡한 차트 셀 참조는 지원하지 않습니다.');
  const match=formula.match(/^(?:'((?:[^']|'')+)'|([^'!:]+))!\$?([A-Z]{1,3})\$?([1-9]\d{0,6})(?::\$?([A-Z]{1,3})\$?([1-9]\d{0,6}))?$/);
  requireValue(!!match,'차트 참조는 내부 단일 시트 A1 범위여야 합니다.');
  const sheet=(match[1]?.replaceAll("''","'")??match[2]),cells=workbook.cells.get(sheet);requireValue(cells instanceof Map,'차트가 참조하는 내부 시트가 없습니다.','CHART_DATA_MISMATCH');
  const column=s=>[...s].reduce((v,c)=>v*26+c.charCodeAt(0)-64,0),name=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
  const x1=column(match[3]),y1=Number(match[4]),x2=column(match[5]??match[3]),y2=Number(match[6]??match[4]);
  requireValue(x2>=x1&&y2>=y1&&x2<=16384&&y2<=1048576&&(x1===x2||y1===y2)&&(x2-x1+1)*(y2-y1+1)<=1000,'차트 셀 범위가 너무 크거나 단일 행·열이 아닙니다.');
  const values=[];for(let y=y1;y<=y2;y++)for(let x=x1;x<=x2;x++){const key=`${name(x)}${y}`;requireValue(cells.has(key),'차트의 현재 셀이 비어 있거나 없습니다.','CHART_DATA_MISMATCH');values.push(cells.get(key));}return values;
}
function referenced(parent,kind,workbook){
  const numericKind=kind==='num',multi=kind==='str'&&parent?.localName==='cat'&&!!one(parent,'multiLvlStrRef'),refName=multi?'multiLvlStrRef':kind+'Ref',cacheName=multi?'multiLvlStrCache':kind+'Cache';allowed(parent,[refName]);const ref=one(parent,refName,true);allowed(ref,['f',cacheName]);
  const values=range(one(ref,'f',true).textContent,workbook),cache=one(ref,cacheName,true);allowed(cache,multi?['ptCount','lvl','extLst']:['formatCode','ptCount','pt','extLst']);extensions(one(cache,'extLst'));
  // PptxGenJS wraps even a single category level; true multilevel labels remain unsupported.
  const points=multi?one(cache,'lvl',true):cache;if(multi)allowed(points,['pt']);
  if(numericKind)requireValue((one(cache,'formatCode')?.textContent??'General')==='General','서식이 지정된 차트 수치는 아직 지원하지 않습니다.');
  requireValue(index(one(cache,'ptCount',true))===values.length&&children(points,'pt').length===values.length,'현재 셀 수와 차트 캐시 수가 다릅니다.','CHART_DATA_MISMATCH');
  const seen=new Set();for(const point of children(points,'pt')){
    allowed(point,['v']);const i=index(point,'idx');requireValue(i<values.length&&!seen.has(i),'차트 캐시 인덱스가 중복되거나 범위를 벗어났습니다.','CHART_DATA_MISMATCH');seen.add(i);
    const cached=one(point,'v',true).textContent,value=numericKind?numeric(cached):cached;
    requireValue(typeof values[i]===(numericKind?'number':'string')&&value===values[i],'차트 캐시와 현재 내부 Excel 값이 다릅니다.','CHART_DATA_MISMATCH');
  }return values;
}

export function readChartPackage(parts,chartPath,{budget,contentTypeFor,workbooks=new Map()}={}){
  const diagnostics=[];let workbookPath=null;
  try{
    requireValue(typeof contentTypeFor==='function'&&contentTypeFor(chartPath)===CHART,'지원하는 차트 ContentType이 아닙니다.');
    const xml=file=>{requireValue(parts[file] instanceof Uint8Array&&parts[file].length>0&&parts[file].length<=16*1024*1024,'차트 부분이 없거나 크기 제한을 초과했습니다.','INPUT_INVALID');const value=strFromU8(parts[file]);requireValue(!/<!DOCTYPE|<!ENTITY/i.test(value),'DTD·엔터티는 지원하지 않습니다.');return new DOMParser({onError:()=>reject('차트 XML이 잘못되었습니다.','INPUT_INVALID')}).parseFromString(value,'application/xml');};
    const root=xml(chartPath).documentElement;requireValue(root.namespaceURI===C&&root.localName==='chartSpace','차트 루트가 잘못되었습니다.','INPUT_INVALID');
    allowed(root,['date1904','lang','roundedCorners','style','chart','spPr','txPr','externalData','extLst']);extensions(one(root,'extLst'));
    requireValue(!flag(one(root,'date1904'))&&!flag(one(root,'roundedCorners')),'이 차트 날짜·영역 서식은 지원하지 않습니다.');exact(one(root,'style'),'val','2');area(one(root,'spPr'));
    const ext=one(root,'externalData',true);allowed(ext,['autoUpdate']);const rid=ext.getAttributeNS(R,'id');requireValue(!!rid,'차트 내부 데이터 관계가 없습니다.','INPUT_INVALID');
    const relPath=path.posix.join(path.posix.dirname(chartPath),'_rels',path.posix.basename(chartPath)+'.rels'),rels=xml(relPath).documentElement;
    requireValue(rels.namespaceURI===REL&&rels.localName==='Relationships','차트 관계 루트가 잘못되었습니다.','INPUT_INVALID');allowed(rels,['Relationship'],REL);
    const ids=new Set();let target;
    for(const r of elements(rels)){
      const id=attrs(r,'Id'),to=attrs(r,'Target','');requireValue(id&&!ids.has(id),'차트 관계 ID가 중복되거나 없습니다.','INPUT_INVALID');ids.add(id);
      requireValue(['Internal',null].includes(attrs(r,'TargetMode'))&&attrs(r,'Type')===`${R}/package`,'차트의 외부·추가 관계는 지원하지 않습니다.');
      requireValue(id===rid&&!target&&!/[:\\%?#]/.test(to)&&to,'차트 데이터 관계가 잘못되었습니다.','INPUT_INVALID');target=to;
    }
    requireValue(!!target,'차트 내부 데이터가 없습니다.','INPUT_INVALID');workbookPath=path.posix.normalize(target.startsWith('/')?target.slice(1):path.posix.join(path.posix.dirname(chartPath),target));
    requireValue(/^ppt\/embeddings\/[^/]+\.xlsx$/.test(workbookPath)&&parts[workbookPath] instanceof Uint8Array&&contentTypeFor(workbookPath)===XLSX,'차트 데이터는 검증된 내부 XLSX여야 합니다.');
    if(!workbooks.has(workbookPath))workbooks.set(workbookPath,readChartWorkbook(parts[workbookPath],{budget}));
    const workbook=workbooks.get(workbookPath);if(workbook.diagnostics.some(d=>d.severity==='error'))return {content:null,style:null,workbookPath:null,diagnostics:[...workbook.diagnostics]};
    const chart=one(root,'chart',true);allowed(chart,['title','autoTitleDeleted','plotArea','legend','plotVisOnly','dispBlanksAs','showDLblsOverMax','extLst']);extensions(one(chart,'extLst'));
    requireValue(flag(one(chart,'plotVisOnly'),true),'숨은 데이터 표시는 아직 지원하지 않습니다.');flag(one(chart,'showDLblsOverMax')); // Every value must remain inside the explicit axis range below.
    exact(one(chart,'dispBlanksAs'),'val','gap');
    const plot=one(chart,'plotArea',true);allowed(plot,['layout','barChart','lineChart','pieChart','catAx','valAx','spPr','extLst']);noLayout(one(plot,'layout'));area(one(plot,'spPr'));extensions(one(plot,'extLst'));
    const types=elements(plot).filter(n=>['barChart','lineChart','pieChart'].includes(n.localName));requireValue(types.length===1,'혼합·3D 차트는 지원하지 않습니다.');const type=types[0],kind=type.localName.slice(0,-5);
    allowed(type,['varyColors','ser','dLbls','extLst',...(kind==='bar'?['barDir','grouping','gapWidth','overlap','axId']:kind==='line'?['grouping','marker','smooth','axId']:['firstSliceAng'])]);extensions(one(type,'extLst'));
    if(kind==='bar'){exact(one(type,'barDir',true),'val','col');exact(one(type,'grouping',true),'val','clustered');requireValue(number(one(type,'gapWidth'),'val',33)===33&&number(one(type,'overlap'),'val',0)===0,'변경된 막대 간격은 지원하지 않습니다.');}
    if(kind==='line'){exact(one(type,'grouping'),'val','standard');requireValue(!flag(one(type,'smooth'))&&flag(one(type,'marker'),true),'곡선·표시가 없는 선 차트는 지원하지 않습니다.');}
    if(kind==='pie')requireValue(number(one(type,'firstSliceAng'),'val',0)===0,'회전한 원형 차트는 지원하지 않습니다.');
    requireValue(flag(one(type,'varyColors'))===(kind==='pie'),'개별 색이 지정된 차트는 지원하지 않습니다.');
    const texts=textReader(),title=texts.title(one(chart,'title'));requireValue(!title||!flag(one(chart,'autoTitleDeleted')),'제목 표시 정보가 서로 다릅니다.');
    const legend=one(chart,'legend');if(legend){allowed(legend,['legendPos','layout','overlay','txPr','spPr','extLst']);exact(one(legend,'legendPos',true),'val','b');requireValue(!flag(one(legend,'overlay')),'겹친 차트 범례는 지원하지 않습니다.');noLayout(one(legend,'layout'));area(one(legend,'spPr'));extensions(one(legend,'extLst'));texts.txPr(one(legend,'txPr',true));}
    const axes={},axisNodes=[...children(plot,'catAx'),...children(plot,'valAx')];
    if(kind==='pie')requireValue(axisNodes.length===0,'원형 차트에는 축을 사용할 수 없습니다.');
    else{
      const cat=one(plot,'catAx',true),val=one(plot,'valAx',true);requireValue(axisNodes.length===2,'보조축은 아직 지원하지 않습니다.');
      for(const [axis,name,other,titleKey] of [[cat,'b',val,'xTitle'],[val,'l',cat,'yTitle']]){
        allowed(axis,['axId','scaling','delete','axPos','majorGridlines','minorGridlines','title','numFmt','majorTickMark','minorTickMark','tickLblPos','spPr','txPr','crossAx','crosses','crossBetween','auto','lblAlgn','lblOffset','noMultiLvlLbl','majorUnit','extLst']);extensions(one(axis,'extLst'));
        requireValue(!flag(one(axis,'delete'))&&!one(axis,'minorGridlines'),'숨긴 축·보조 눈금선은 지원하지 않습니다.');exact(one(axis,'axPos',true),'val',name);exact(one(axis,'majorTickMark'),'val','none');exact(one(axis,'minorTickMark'),'val','none');exact(one(axis,'tickLblPos'),'val','low');exact(one(axis,'crosses'),'val','autoZero');
        requireValue(index(one(axis,'crossAx',true),'val',4294967295)===index(one(other,'axId',true),'val',4294967295),'차트 축 연결이 다릅니다.','INPUT_INVALID');
        const scaling=one(axis,'scaling',true);allowed(scaling,axis===val?['orientation','min','max']:['orientation']);exact(one(scaling,'orientation'),'val','minMax');
        const format=one(axis,'numFmt');exact(format,'formatCode','General');texts.txPr(one(axis,'txPr',true));const label=texts.title(one(axis,'title'));if(label)axes[titleKey]=label;
        if(axis===val){requireValue(!!one(axis,'majorGridlines'),'눈금선을 제거한 차트는 아직 지원하지 않습니다.');axes.min=number(one(scaling,'min',true),'val');axes.max=number(one(scaling,'max',true),'val');axes.majorUnit=number(one(axis,'majorUnit',true),'val');exact(one(axis,'crossBetween'),'val','between');}
        else{requireValue(!one(axis,'majorGridlines')&&!one(axis,'majorUnit')&&flag(one(axis,'auto'),true)&&flag(one(axis,'noMultiLvlLbl'),true),'다단·수치 범주 축은 지원하지 않습니다.');exact(one(axis,'lblAlgn'),'val','ctr');requireValue(number(one(axis,'lblOffset'),'val',100)===100,'변경된 축 레이블 간격은 지원하지 않습니다.');}
      }
      const required=[cat,val].map(n=>attrs(one(n,'axId',true),'val')),references=children(type,'axId').map(n=>attrs(n,'val'));
      // PptxGenJS 4 also writes one unused series-axis ID; Office removes it.
      requireValue(new Set(required).size===2&&new Set(references).size===references.length&&references.length>=2&&references.length<=3&&required.every(id=>references.includes(id)),'차트 축 참조가 잘못되었습니다.','INPUT_INVALID');
    }
    const rawSeries=children(type,'ser');requireValue(rawSeries.length>0&&rawSeries.length<=10,'차트 계열 수 제한을 초과했습니다.');const order=new Set(),indices=new Set();
    for(const s of rawSeries){const i=index(one(s,'idx',true)),o=index(one(s,'order',true));requireValue(!indices.has(i)&&!order.has(o),'차트 계열 순서가 중복됩니다.','INPUT_INVALID');indices.add(i);order.add(o);}
    rawSeries.sort((a,b)=>index(one(a,'order',true))-index(one(b,'order',true)));let categories;
    const labels=(series,values)=>{
      const global=one(type,'dLbls'),own=one(series,'dLbls'),points=children(own,'dLbl'),byIndex=new Map();
      requireValue(!children(global,'dLbl').length,'차트 전체의 개별 레이블 지정은 아직 지원하지 않습니다.');
      for(const group of [global,own].filter(Boolean)){allowed(group,['dLbl','numFmt','spPr','txPr','dLblPos','showLegendKey','showVal','showCatName','showSerName','showPercent','showBubbleSize','showLeaderLines','extLst']);extensions(one(group,'extLst'));}
      for(const p of points){allowed(p,['idx','numFmt','spPr','txPr','dLblPos','showLegendKey','showVal','showCatName','showSerName','showPercent','showBubbleSize','showLeaderLines','extLst']);extensions(one(p,'extLst'));const i=index(one(p,'idx',true));requireValue(i<values.length&&!byIndex.has(i),'차트 레이블 인덱스가 잘못되었습니다.');byIndex.set(i,p);}
      for(let i=0;i<values.length;i++){
        const p=byIndex.get(i),pick=k=>one(p,k)??one(own,k)??one(global,k);
        for(const k of ['showLegendKey','showVal','showCatName','showSerName','showPercent','showBubbleSize','showLeaderLines'])requireValue(flag(pick(k))===(kind==='pie'&&k==='showCatName'),'현재 계약에 없는 차트 레이블은 지원하지 않습니다.');
        if(kind==='pie'){exact(pick('dLblPos'),'val','ctr');texts.txPr(pick('txPr'),'#ffffff');area(pick('spPr'));exact(pick('numFmt'),'formatCode','General');}
      }
    };
    const series=rawSeries.map(s=>{
      allowed(s,['idx','order','tx','spPr','invertIfNegative','marker','dPt','dLbls','cat','val','smooth','explosion','extLst']);extensions(one(s,'extLst'));
      requireValue(!flag(one(s,'invertIfNegative'))&&!flag(one(s,'smooth'))&&number(one(s,'explosion'),'val',0)===0,'변경된 계열 효과는 지원하지 않습니다.');
      const names=referenced(one(s,'tx',true),'str',workbook),cats=referenced(one(s,'cat',true),'str',workbook),values=referenced(one(s,'val',true),'num',workbook);
      requireValue(names.length===1&&cats.length===values.length&&(!categories||JSON.stringify(categories)===JSON.stringify(cats)),'차트 계열의 항목·값 수가 서로 다릅니다.','CHART_DATA_MISMATCH');categories=cats;labels(s,values);
      const props=one(s,'spPr',true);allowed(props,['solidFill','ln','effectLst'],A);requireValue(!elements(one(props,'effectLst',false,A)).length,'계열 효과는 지원하지 않습니다.');
      const points=children(s,'dPt');let color;
      if(kind!=='pie'){requireValue(!points.length,'개별 데이터 색은 아직 지원하지 않습니다.');color=rgb(one(kind==='line'?one(props,'ln',true,A):props,'solidFill',true,A));}
      else{
        const pointMap=new Map();for(const p of points){allowed(p,['idx','bubble3D','spPr','extLst']);extensions(one(p,'extLst'));requireValue(!flag(one(p,'bubble3D')),'3D 데이터는 지원하지 않습니다.');const i=index(one(p,'idx',true));requireValue(i<values.length&&!pointMap.has(i),'데이터 색 인덱스가 잘못되었습니다.');pointMap.set(i,p);}
        for(let i=0;i<values.length;i++){const props=one(pointMap.get(i),'spPr')??one(s,'spPr',true),fill=rgb(one(props,'solidFill',true,A));requireValue(!color||color===fill,'각 조각의 다른 색은 아직 지원하지 않습니다.');color=fill;}
      }return {name:names[0],values,color};
    });
    const content={type:kind,categories,series,title,legend:!!legend,axes},style=texts.finish(),unit=style.fontSize*3/4/12*12700;
    for(const axis of axisNodes){const props=one(axis,'spPr',true);allowed(props,['ln','effectLst'],A);line(one(props,'ln',true,A),'#475569',unit);requireValue(!elements(one(props,'effectLst',false,A)).length,'축 효과는 지원하지 않습니다.');const grid=one(axis,'majorGridlines');if(grid){allowed(grid,['spPr']);const gp=one(grid,'spPr',true);allowed(gp,['ln'],A);line(one(gp,'ln',true,A),'#cbd5e1',unit);}}
    for(const [i,s] of rawSeries.entries()){
      const props=one(s,'spPr',true);
      if(kind==='line'){
        line(one(props,'ln',true,A),series[i].color,2*unit);const marker=one(s,'marker',true);allowed(marker,['symbol','size','spPr']);exact(one(marker,'symbol',true),'val','circle');requireValue(number(one(marker,'size',true),'val')===Math.round(6*unit/12700),'변경된 마커 크기는 지원하지 않습니다.');
        const mp=one(marker,'spPr',true);allowed(mp,['solidFill','ln','effectLst'],A);requireValue(rgb(one(mp,'solidFill',true,A))===series[i].color&&!elements(one(mp,'effectLst',false,A)).length,'개별 마커 서식은 지원하지 않습니다.');line(one(mp,'ln',true,A),series[i].color,unit);
      }else if(kind==='pie'){
        const pointProps=new Map(children(s,'dPt').map(p=>[index(one(p,'idx',true)),one(p,'spPr',true)]));
        for(let j=0;j<categories.length;j++){const p=pointProps.get(j)??props;allowed(p,['solidFill','ln','effectLst'],A);requireValue(!elements(one(p,'effectLst',false,A)).length,'원형 차트 효과는 지원하지 않습니다.');line(one(p,'ln',true,A),'#ffffff',unit);}
      }
      else requireValue(!one(props,'ln',false,A),'변경된 막대 테두리는 아직 지원하지 않습니다.');
    }
    const validation=validateChart(content);requireValue(validation.valid,'현재 차트 값·축이 지원 계약에 맞지 않습니다.','INPUT_INVALID');
    return {content,style,workbookPath,diagnostics};
  }catch(e){diagnostics.push(diagnostic(['INPUT_INVALID','UNSUPPORTED_OBJECT','CHART_DATA_MISMATCH','LIMIT_EXCEEDED'].includes(e.code)?e.code:'INPUT_INVALID',e.message));return {content:null,style:null,workbookPath:null,diagnostics};}
}
