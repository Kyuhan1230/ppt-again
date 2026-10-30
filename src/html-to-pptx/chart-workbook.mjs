import path from 'node:path';
import {DOMParser} from '@xmldom/xmldom';
import {unzipBudgeted} from './package-budget.mjs';

const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',P='http://schemas.openxmlformats.org/package/2006/relationships',C='http://schemas.openxmlformats.org/package/2006/content-types';
const types={workbook:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',worksheet:'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',sharedStrings:'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',styles:'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',theme:'application/vnd.openxmlformats-officedocument.theme+xml',table:'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',core:'application/vnd.openxmlformats-package.core-properties+xml',app:'application/vnd.openxmlformats-officedocument.extended-properties+xml'};
const elements=node=>[...(node?.childNodes??[])].filter(n=>n.nodeType===1),all=(node,name)=>[...node.getElementsByTagNameNS('*',name)],attribute=(node,name)=>node.hasAttribute(name)?node.getAttribute(name):null;
const reject=(message,code='INPUT_INVALID')=>{throw Object.assign(new Error(message),{code});};
const check=(condition,message,code)=>{if(!condition)reject(message,code);};
const integer=value=>typeof value==='string'&&/^(?:0|[1-9]\d*)$/.test(value)&&Number.isSafeInteger(Number(value));
const address=value=>{const m=/^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(value??'');if(!m)return null;const col=[...m[1]].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0),row=Number(m[2]);return col<=16384&&row<=1048576?{row,col}:null;};

// Read values only. Neither OOXML formulas nor workbook code are evaluated.
// A caller caches this result by the validated XLSX part path; every first read
// must receive the same budget already charged for the outer PPTX package.
export function readChartWorkbook(bytes,{budget}={}){
  try{
    const parts=unzipBudgeted(bytes,{budget}),documents=new Map(),used=new Set(['[Content_Types].xml']);
    const xml=file=>{
      if(documents.has(file))return documents.get(file);
      check(parts[file],'필수 차트 워크북 부분이 없습니다.');
      let text;try{text=new TextDecoder('utf-8',{fatal:true}).decode(parts[file]);}catch{reject('워크북 XML 문자 인코딩이 올바르지 않습니다.');}
      check(!/<!DOCTYPE|<!ENTITY/i.test(text),'워크북 DTD·엔터티는 지원하지 않습니다.','UNSUPPORTED_OBJECT');
      const doc=new DOMParser({onError:()=>reject('잘못된 워크북 XML입니다.')}).parseFromString(text,'application/xml');documents.set(file,doc);return doc;
    };
    for(const name of Object.keys(parts)){
      if(name.endsWith('/'))continue;
      check(!/vbaProject|activeX|embeddings\/|externalLinks\/|\.xls[mbx]$|\.bin$/i.test(name),'매크로·외부 데이터·재귀 내장 문서는 지원하지 않습니다.','UNSUPPORTED_OBJECT');
      check(name.endsWith('.xml')||name.endsWith('.rels'),'워크북에 지원하지 않는 데이터가 있습니다.','UNSUPPORTED_OBJECT');
      xml(name);
    }
    const typeRoot=xml('[Content_Types].xml').documentElement,defaults=new Map(),overrides=new Map();
    check(typeRoot.namespaceURI===C&&typeRoot.localName==='Types','워크북 ContentType 네임스페이스가 올바르지 않습니다.');
    for(const node of elements(typeRoot)){
      check(node.namespaceURI===C&&['Default','Override'].includes(node.localName),'잘못된 워크북 ContentType입니다.');
      const override=node.localName==='Override',key=attribute(node,override?'PartName':'Extension'),map=override?overrides:defaults,type=attribute(node,'ContentType');
      check(key&&type&&!map.has(key),'워크북 ContentType이 중복되거나 없습니다.');
      if(override)check(key.startsWith('/')&&parts[key.slice(1)],'워크북 ContentType 대상이 없습니다.');
      map.set(key,type);
    }
    const contentTypeFor=file=>overrides.get('/'+file)??defaults.get(file.slice(file.lastIndexOf('.')+1));
    const requireType=(file,type)=>{check(parts[file]&&contentTypeFor(file)===type,'워크북 부분의 ContentType이 지원 형식과 다릅니다.','UNSUPPORTED_OBJECT');used.add(file);};
    const relationships=file=>{
      const relFile=file?path.posix.join(path.posix.dirname(file),'_rels',path.posix.basename(file)+'.rels'):'_rels/.rels';
      check(parts[relFile],'필수 워크북 관계가 없습니다.');used.add(relFile);
      check(contentTypeFor(relFile)==='application/vnd.openxmlformats-package.relationships+xml','워크북 관계 ContentType이 올바르지 않습니다.');
      const root=xml(relFile).documentElement,map=new Map();check(root.namespaceURI===P&&root.localName==='Relationships','워크북 관계 네임스페이스가 올바르지 않습니다.');
      for(const node of elements(root)){
        check(node.namespaceURI===P&&node.localName==='Relationship','잘못된 워크북 관계입니다.');
        const id=attribute(node,'Id'),target=attribute(node,'Target'),type=attribute(node,'Type'),mode=attribute(node,'TargetMode');
        check(id&&!map.has(id)&&target&&type,'워크북 관계가 중복되거나 불완전합니다.');
        check(mode===null||mode==='Internal','외부 워크북 관계는 지원하지 않습니다.','UNSUPPORTED_OBJECT');
        check(!/[\\\u0000-\u001f:%?#]/.test(target),'워크북 관계의 외부 또는 인코딩 경로는 지원하지 않습니다.','UNSUPPORTED_OBJECT');
        const resolved=target.startsWith('/')?target.slice(1):path.posix.normalize(path.posix.join(file?path.posix.dirname(file):'',target));
        check(resolved&&!resolved.startsWith('../')&&!resolved.startsWith('/')&&parts[resolved],'워크북 관계 경로가 패키지 밖이거나 없습니다.');
        map.set(id,{type,path:resolved});
      }
      return map;
    };
    const rootRels=relationships(''),workbooks=[];
    for(const item of rootRels.values()){
      if(item.type===R+'/officeDocument'){requireType(item.path,types.workbook);workbooks.push(item.path);}
      else if(item.type===P+'/metadata/core-properties')requireType(item.path,types.core);
      else if(item.type===R+'/extended-properties')requireType(item.path,types.app);
      else reject('지원하지 않는 워크북 최상위 관계입니다.','UNSUPPORTED_OBJECT');
    }
    check(workbooks.length===1,'차트 워크북 본문은 한 개여야 합니다.');
    const workbookPath=workbooks[0],workbook=xml(workbookPath),root=workbook.documentElement,links=relationships(workbookPath),shared=[];
    check(root.namespaceURI===S&&root.localName==='workbook','워크북 네임스페이스가 올바르지 않습니다.');
    for(const forbidden of ['externalReferences','externalReference','definedName','oleObjects','oleObject','controls','control','connections','queryTable','f','formula','calculatedColumnFormula','totalsRowFormula','hyperlinks','legacyDrawing','legacyDrawingHF','drawing','tableParts','autoFilter']){
      for(const doc of documents.values())check(!all(doc,forbidden).length,'수식·외부 연결·활성 워크북 요소는 지원하지 않습니다.','UNSUPPORTED_OBJECT');
    }
    const sharedPaths=[],sheetPaths=new Set();
    for(const item of links.values()){
      const kind=item.type.startsWith(R+'/')?item.type.slice(R.length+1):'';
      check(['worksheet','sharedStrings','styles','theme'].includes(kind),'지원하지 않는 워크북 내부 관계입니다.','UNSUPPORTED_OBJECT');requireType(item.path,types[kind]);
      if(kind==='sharedStrings')sharedPaths.push(item.path);
      if(kind==='worksheet'){check(!sheetPaths.has(item.path),'시트 관계가 중복됩니다.');sheetPaths.add(item.path);}
    }
    check(sharedPaths.length<=1,'공유 문자열 부분이 중복됩니다.');
    let stringUnits=0,cellCount=0;
    const textValue=node=>{
      check(node.namespaceURI===S&&['si','is'].includes(node.localName),'잘못된 워크북 문자열입니다.');let value='';
      for(const child of elements(node)){
        check(child.namespaceURI===S,'워크북 문자열 네임스페이스가 올바르지 않습니다.');
        if(child.localName==='t'){check(!elements(child).length,'문자열 안에 요소를 넣을 수 없습니다.');value+=child.textContent;}
        else if(child.localName==='r'){
          const items=elements(child),texts=items.filter(n=>n.localName==='t');
          check(texts.length===1&&items.every(n=>n.namespaceURI===S&&['rPr','t'].includes(n.localName))&&!elements(texts[0]).length,'지원하지 않는 공유 문자열 조각입니다.');value+=texts[0].textContent;
        }else if(child.localName==='phoneticPr')check(!elements(child).length&&!child.textContent.trim(),'발음 설정에 추가 콘텐츠가 있습니다.');
        else reject('지원하지 않는 공유 문자열 표현입니다.','UNSUPPORTED_OBJECT');
      }
      check(value.length<=100000&&(stringUnits+=value.length)<=2000000,'워크북 문구가 글자 수 한도를 초과했습니다.','LIMIT_EXCEEDED');return value;
    };
    if(sharedPaths.length){
      const strings=xml(sharedPaths[0]).documentElement;check(strings.namespaceURI===S&&strings.localName==='sst','공유 문자열 구조가 올바르지 않습니다.');
      for(const item of elements(strings)){check(item.localName==='si','잘못된 공유 문자열 항목입니다.');shared.push(textValue(item));}
    }
    const lists=elements(root).filter(n=>n.namespaceURI===S&&n.localName==='sheets');check(lists.length===1,'워크북 시트 목록이 하나여야 합니다.');
    const sheets=[],cells=new Map(),ids=new Set(),names=new Set(),selectedPaths=new Set();
    for(const node of elements(lists[0])){
      const id=attribute(node,'sheetId'),name=attribute(node,'name'),link=links.get(node.getAttributeNS(R,'id'));
      check(node.namespaceURI===S&&node.localName==='sheet'&&integer(id)&&Number(id)>0&&!ids.has(id)&&name&&name.length<=31&&!/[\u0000-\u001f\\/?*\[\]:]/.test(name)&&!names.has(name.toLowerCase())&&link?.type===R+'/worksheet'&&!selectedPaths.has(link.path),'시트 이름·ID·관계가 잘못되었거나 중복됩니다.');
      ids.add(id);names.add(name.toLowerCase());selectedPaths.add(link.path);sheets.push({id,name,path:link.path});
      const sheet=xml(link.path).documentElement;check(sheet.namespaceURI===S&&sheet.localName==='worksheet','시트 네임스페이스가 올바르지 않습니다.');
      for(const element of [...all(sheet,'row'),...all(sheet,'col'),...all(sheet,'sheetFormatPr'),...all(sheet,'sheetPr')]){
        check(['hidden','collapsed','zeroHeight','filterMode'].every(key=>[null,'0','false'].includes(attribute(element,key))),'숨긴 행·열과 필터 상태는 차트로 읽을 수 없습니다.','UNSUPPORTED_OBJECT');
        for(const key of ['ht','width','defaultRowHeight','defaultColWidth'])if(element.hasAttribute(key))check(Number.isFinite(Number(attribute(element,key)))&&Number(attribute(element,key))>0,'크기가 0인 행·열은 차트로 읽을 수 없습니다.','UNSUPPORTED_OBJECT');
      }
      const sheetRels=path.posix.join(path.posix.dirname(link.path),'_rels',path.posix.basename(link.path)+'.rels');
      if(parts[sheetRels])for(const tableLink of relationships(link.path).values()){
        // PptxGenJS emits an unused table style part, including a trailing quote
        // in its ref. Accept only this inert shape, never a second data source.
        check(tableLink.type===R+'/table','지원하지 않는 시트 관계입니다.','UNSUPPORTED_OBJECT');requireType(tableLink.path,types.table);
        const table=xml(tableLink.path).documentElement,children=elements(table),columns=children.filter(n=>n.localName==='tableColumns'),style=children.filter(n=>n.localName==='tableStyleInfo');
        const ref=/^([A-Z]{1,3}[1-9]\d{0,6}):([A-Z]{1,3}[1-9]\d{0,6})'?$/.exec(attribute(table,'ref')??''),first=ref&&address(ref[1]),last=ref&&address(ref[2]);
        check(table.namespaceURI===S&&table.localName==='table'&&children.every(n=>n.namespaceURI===S&&['tableColumns','tableStyleInfo'].includes(n.localName))&&columns.length===1&&style.length<=1&&(!style.length||!elements(style[0]).length)&&first&&last&&first.row<=last.row&&first.col<=last.col&&[null,'0'].includes(attribute(table,'totalsRowShown')),'계산·필터를 포함한 표 메타데이터는 지원하지 않습니다.','UNSUPPORTED_OBJECT');
        check([...table.attributes].every(a=>['xmlns','id','name','displayName','ref','totalsRowShown'].includes(a.name)),'지원하지 않는 표 메타데이터 속성입니다.','UNSUPPORTED_OBJECT');
        const items=elements(columns[0]);check(integer(attribute(columns[0],'count'))&&items.length===Number(attribute(columns[0],'count'))&&items.length===last.col-first.col+1,'표 메타데이터의 열 수가 다릅니다.');
        const columnIds=new Set();for(const column of items){const columnId=attribute(column,'id');check(column.namespaceURI===S&&column.localName==='tableColumn'&&!elements(column).length&&integer(columnId)&&Number(columnId)>0&&!columnIds.has(columnId)&&[...column.attributes].every(a=>['id','name'].includes(a.name)),'수식·계산 열은 지원하지 않습니다.','UNSUPPORTED_OBJECT');columnIds.add(columnId);}
      }
      const data=elements(sheet).filter(n=>n.namespaceURI===S&&n.localName==='sheetData');check(data.length===1,'시트 데이터가 하나여야 합니다.');
      const values=new Map(),rows=new Set();
      for(const row of elements(data[0])){
        const r=attribute(row,'r');check(row.namespaceURI===S&&row.localName==='row'&&integer(r)&&Number(r)>0&&Number(r)<=1048576&&!rows.has(r),'잘못되거나 중복된 시트 행입니다.');rows.add(r);
        for(const cell of elements(row)){
          check(++cellCount<=100000,'워크북 셀 수 한도를 초과했습니다.','LIMIT_EXCEEDED');
          const ref=attribute(cell,'r'),where=address(ref),kind=attribute(cell,'t')??'n',children=elements(cell);
          check(cell.namespaceURI===S&&cell.localName==='c'&&where?.row===Number(r)&&!values.has(ref),'잘못되거나 중복된 셀 주소입니다.');
          check(['n','s','inlineStr'].includes(kind),'수치·문자열 이외의 셀은 지원하지 않습니다.','UNSUPPORTED_OBJECT');
          check(children.length===1&&children[0].namespaceURI===S&&children[0].localName===(kind==='inlineStr'?'is':'v'),'셀 값이 없거나 중복·수식 형태입니다.');
          let value;
          if(kind==='inlineStr')value=textValue(children[0]);
          else{
            check(!elements(children[0]).length,'셀 값 안에 요소를 넣을 수 없습니다.');const raw=children[0].textContent.trim();
            if(kind==='s'){check(integer(raw)&&Number(raw)<shared.length,'공유 문자열 참조가 범위를 벗어났습니다.');value=shared[Number(raw)];}
            else{check(/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(raw)&&Number.isFinite(Number(raw)),'빈 값 또는 유한하지 않은 수치 셀입니다.');value=Number(raw);}
          }
          values.set(ref,value);
        }
      }
      cells.set(name,values);
    }
    check(sheets.length>0&&selectedPaths.size===sheetPaths.size,'시트 목록과 실제 워크시트 관계가 다릅니다.');
    for(const name of Object.keys(parts))if(!name.endsWith('/'))check(used.has(name),'차트에 필요한 범위를 벗어난 워크북 부분이 있습니다.','UNSUPPORTED_OBJECT');
    return {sheets,cells,diagnostics:[]};
  }catch(error){return {sheets:[],cells:new Map(),diagnostics:[{code:error.code??'INPUT_INVALID',severity:'error',slideId:null,nodeId:null,sourceRef:null,message:error.message}]};}
}
