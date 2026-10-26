import path from 'node:path';
import {unzipSync,strFromU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
import {diagnostic,validateModel} from './model.mjs';
import {resolveAsset,sha256} from './assets.mjs';
import {EMU_PER_CSS_PX} from './canvas.mjs';

const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const child=(n,name)=>elements(n).find(c=>c.localName===name);
const all=(n,name)=>n?[...n.getElementsByTagNameNS('*',name)]:[];
const attr=(n,k,fallback=null)=>n?.hasAttribute(k)?n.getAttribute(k):fallback;
const num=(n,k,fallback=0)=>Number(attr(n,k,fallback));
const px=n=>n/EMU_PER_CSS_PX;
const mimeByExt={'.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.svg':'image/svg+xml'};

// Strict subset reader. No Office process, document code, network or disk extraction.
export async function readPptx(bytes){
  const diagnostics=[],model={version:1,sourceDigest:sha256(bytes??''),slides:[],assets:[]};
  const fail=(code,message,slideId=null,nodeId=null)=>diagnostics.push(diagnostic(code,message,{slideId,nodeId}));
  const done=()=>{const failed=diagnostics.some(d=>d.severity==='error');return {status:failed?'failed':'parsed',model:failed?null:model,diagnostics};};
  try{
    if(!(bytes instanceof Uint8Array)||bytes.length<22||bytes.length>64*1024*1024)throw new Error('PPTX 바이트는 64MiB 이내여야 합니다.');
    let count=0,total=0;const names=new Set();
    const parts=unzipSync(bytes,{filter:f=>{
      if(f.compression===0&&f.size!==f.originalSize)throw new Error('저장된 ZIP 항목의 크기 정보가 일치하지 않습니다.');
      if(![0,8].includes(f.compression)||++count>10000||Math.max(f.size,f.originalSize)>16*1024*1024||(total+=Math.max(f.size,f.originalSize))>128*1024*1024)throw new Error('PPTX 압축 해제 한도를 초과했습니다.');
      if(!f.name||f.name.startsWith('/')||f.name.includes('\\')||f.name.split('/').some(p=>p==='..')||names.has(f.name))throw new Error('PPTX 내부 경로가 잘못되었거나 중복됩니다.');
      names.add(f.name);return true;
    }});
    const xml=file=>{
      if(!parts[file])throw new Error('PPTX 필수 부분이 없습니다.');
      const s=strFromU8(parts[file]);if(/<!DOCTYPE|<!ENTITY/i.test(s))throw new Error('DTD·엔터티는 지원하지 않습니다.');
      return new DOMParser({onError:()=>{throw new Error('잘못된 PPTX XML입니다.');}}).parseFromString(s,'application/xml');
    };
    const relationships=file=>{
      const rfile=path.posix.join(path.posix.dirname(file),'_rels',path.posix.basename(file)+'.rels');
      const map=new Map();if(!parts[rfile])return map;
      for(const r of all(xml(rfile),'Relationship')){
        if(attr(r,'TargetMode')==='External')throw new Error('외부 참조가 있는 PPTX는 지원하지 않습니다.');
        const target=attr(r,'Target','');
        if(/^[a-z][a-z\d+.-]*:|\\/i.test(target))throw new Error('외부 PPTX 경로입니다.');
        const resolved=target.startsWith('/')?target.slice(1):path.posix.normalize(path.posix.join(path.posix.dirname(file),target));
        if(resolved.startsWith('../'))throw new Error('PPTX 경로가 범위를 벗어났습니다.');
        map.set(attr(r,'Id'),resolved);
      }
      return map;
    };
    // Inspect every relationship, including unused external/OLE/script packages.
    for(const name of Object.keys(parts)){
      if(!name.endsWith('/')&&/vbaProject|activeX|embeddings\//i.test(name))throw new Error('매크로·ActiveX·내장 문서는 지원하지 않습니다.');
      if(name.endsWith('.rels'))for(const r of all(xml(name),'Relationship'))if(attr(r,'TargetMode')==='External')throw new Error('외부 참조가 있는 PPTX는 지원하지 않습니다.');
    }
    const presentation=xml('ppt/presentation.xml'),size=all(presentation,'sldSz')[0];
    model.pptxSizeEmu={width:num(size,'cx'),height:num(size,'cy')};
    const slideRefs=all(presentation,'sldId');if(!slideRefs.length||slideRefs.length>200)throw new Error('장표는 1~200장이어야 합니다.');
    const presRels=relationships('ppt/presentation.xml');
    for(const [index,ref] of slideRefs.entries()){
      const file=presRels.get(attr(ref,'r:id')),doc=xml(file),rels=relationships(file),tree=all(doc,'spTree')[0];
      const slide={id:`slide-${index+1}`,widthCssPx:px(model.pptxSizeEmu.width),heightCssPx:px(model.pptxSizeEmu.height),background:{color:'#ffffff',image:'none'},nodes:[]};model.slides.push(slide);
      const issue=(message,nodeId=null)=>fail('UNSUPPORTED_OBJECT',message,slide.id,nodeId);
      const inherited=new Set([...rels.values()].filter(p=>p.startsWith('ppt/slideLayouts/')));
      for(const layout of [...inherited])for(const target of relationships(layout).values())if(target.startsWith('ppt/slideMasters/'))inherited.add(target);
      for(const part of inherited){
        const inheritedDoc=xml(part),inheritedTree=all(inheritedDoc,'spTree')[0];
        if(elements(inheritedTree).some(n=>['sp','pic','cxnSp','grpSp','graphicFrame'].includes(n.localName)&&!all(n,'ph').length))issue('마스터·레이아웃의 배경 객체는 아직 지원하지 않습니다.');
      }
      if(all(doc,'timing').length||all(doc,'transition').length)issue('타이밍·전환 효과는 정적 HTML로 보존할 수 없습니다.');
      const color=(fill,fallback='transparent')=>{
        if(!fill)return fallback;
        const c=child(fill,'srgbClr');if(!c){issue('테마·시스템 색은 명시적 RGB로 바꿔 주세요.');return fallback;}
        const hex=attr(c,'val','');if(!/^[a-f\d]{6}$/i.test(hex))throw new Error('잘못된 RGB 색입니다.');
        if(elements(c).some(n=>n.localName!=='alpha'))issue('색 변환 효과는 지원하지 않습니다.');
        const alpha=num(child(c,'alpha'),'val',100000)/100000;
        return alpha===1?`#${hex}`:`rgba(${parseInt(hex.slice(0,2),16)},${parseInt(hex.slice(2,4),16)},${parseInt(hex.slice(4,6),16)},${alpha})`;
      };
      const bg=all(doc,'bgPr')[0]??[...inherited].map(part=>all(xml(part),'bgPr')[0]).find(Boolean);if(bg)slide.background.color=color(child(bg,'solidFill'),'#ffffff');
      if(bg&&elements(bg).some(n=>!['solidFill','effectLst'].includes(n.localName)))issue('단색 외의 PPT 배경은 지원하지 않습니다.');
      if(all(doc,'bgRef').length||!bg&&[...inherited].some(part=>all(xml(part),'bgRef').length))issue('테마 배경은 아직 지원하지 않습니다.');
      const seen=new Set();let objectCount=0;
      const visit=(parent,parentId=null,map={sx:1,sy:1,tx:0,ty:0})=>{
        for(const object of elements(parent)){
          if(['nvGrpSpPr','grpSpPr','extLst'].includes(object.localName))continue;
          if(++objectCount>10000)throw new Error('장표 객체 한도를 초과했습니다.');
          const nv=all(object,'cNvPr')[0],base=attr(nv,'name')||`object-${attr(nv,'id',String(objectCount))}`;
          let id=base,suffix=1;while(seen.has(id))id=`${base}-${suffix++}`;seen.add(id);
          if(!['sp','cxnSp','pic','grpSp'].includes(object.localName)){issue('표·차트·SmartArt 등 이 객체는 아직 지원하지 않습니다.',id);continue;}
          const group=object.localName==='grpSp',pr=child(object,group?'grpSpPr':'spPr'),xfrm=child(pr,'xfrm');
          if(!xfrm){issue('마스터에 상속된 객체 위치는 아직 지원하지 않습니다.',id);continue;}
          if(num(xfrm,'rot')!==0||(!['cxnSp'].includes(object.localName)&&!['line'].includes(attr(child(pr,'prstGeom'),'prst'))&&(num(xfrm,'flipH')||num(xfrm,'flipV')))){issue('회전·뒤집힌 객체는 아직 지원하지 않습니다.',id);continue;}
          const off=child(xfrm,'off'),ext=child(xfrm,'ext');
          const raw={x:num(off,'x'),y:num(off,'y'),w:num(ext,'cx'),h:num(ext,'cy')};
          const box={x:px(raw.x*map.sx+map.tx),y:px(raw.y*map.sy+map.ty),w:px(raw.w*map.sx),h:px(raw.h*map.sy)};
          const node={id,kind:'shape',boxCssPx:box,zOrder:slide.nodes.length,content:{},style:{opacity:1},sourceRef:{slideId:slide.id,elementId:id},parentId};
          if(group){
            if(parentId){issue('중첩 그룹은 아직 지원하지 않습니다.',id);continue;}
            node.kind='group';slide.nodes.push(node);
            const co=child(xfrm,'chOff'),ce=child(xfrm,'chExt'),sx=raw.w/num(ce,'cx'),sy=raw.h/num(ce,'cy');
            if(!Number.isFinite(sx)||sx<=0||Math.abs(sx-sy)>1e-7){issue('비균일 그룹 배율은 지원하지 않습니다.',id);continue;}
            if(Math.abs(sx-1)>1e-7){issue('배율이 적용된 그룹의 글자·선 굵기는 아직 보존할 수 없습니다.',id);continue;}
            visit(object,id,{sx,sy,tx:raw.x-num(co,'x')*sx,ty:raw.y-num(co,'y')*sy});continue;
          }
          const unsupported=['custGeom','gradFill','pattFill','blipFill','effectDag','scene3d','sp3d'];
          if(elements(pr).some(n=>unsupported.includes(n.localName))||elements(child(pr,'effectLst')).length)issue('복잡한 도형·채움·효과는 아직 지원하지 않습니다.',id);
          const objectStyle=child(object,'style');
          if((child(objectStyle,'fillRef')&&!elements(pr).some(n=>['solidFill','noFill','blipFill','gradFill','pattFill','grpFill'].includes(n.localName)))||
             (child(objectStyle,'lnRef')&&!child(pr,'ln'))||num(child(objectStyle,'effectRef'),'idx')!==0)issue('상속된 도형 채움·선·효과는 아직 지원하지 않습니다.',id);
          if(object.localName==='pic'){
            node.kind='image';const blip=all(object,'blip')[0],target=rels.get(attr(blip,'r:embed'));
            if(all(object,'srcRect').some(n=>[...n.attributes].some(a=>Number(a.value)!==0)))issue('잘라낸 그림은 아직 지원하지 않습니다.',id);
            if(!target||!parts[target]){issue('내장 그림 자산이 없습니다.',id);continue;}
            const mime=mimeByExt[path.posix.extname(target).toLowerCase()];if(!mime){issue('이 그림 형식은 지원하지 않습니다.',id);continue;}
            const asset=resolveAsset('image'+path.posix.extname(target),[{path:'image'+path.posix.extname(target),mime,bytes:parts[target]}]);
            if(!model.assets.some(a=>a.id===asset.id))model.assets.push(asset);
            node.content={assetId:asset.id};node.style.opacity=num(child(blip,'alphaModFix'),'amt',100000)/100000;
          }else{
            const shape=attr(child(pr,'prstGeom'),'prst'),line=child(pr,'ln'),tx=child(object,'txBody');
            Object.assign(node.style,{svgGeometry:true,shape:shape==='ellipse'?'ellipse':'rect',fill:color(child(pr,'solidFill')),stroke:color(child(line,'solidFill')),strokeWidth:child(line,'noFill')?0:px(num(line,'w')),radius:shape==='roundRect'?Math.min(box.w,box.h)*0.16667:0});
            if(child(line,'prstDash')&&attr(child(line,'prstDash'),'val')!=='solid')issue('점선·화살표 선은 아직 지원하지 않습니다.',id);
            if(elements(line).some(n=>['headEnd','tailEnd'].includes(n.localName)&&attr(n,'type')!=='none'))issue('화살표 선은 아직 지원하지 않습니다.',id);
            if(shape==='line'||object.localName==='cxnSp'){
              node.kind='line';const flipH=num(xfrm,'flipH'),flipV=num(xfrm,'flipV');
              node.style.endpoints={start:{x:box.x+(flipH?box.w:0),y:box.y+(flipV?box.h:0)},end:{x:box.x+(flipH?0:box.w),y:box.y+(flipV?0:box.h)}};
            }else if(tx&&all(tx,'t').length){
              node.kind='text';const body=child(tx,'bodyPr'),paras=elements(tx).filter(n=>n.localName==='p'),runs=[];
              if(num(body,'rot')!==0||attr(body,'vert','horz')!=='horz'||num(body,'numCol',1)!==1||num(body,'rtlCol')!==0||num(body,'fromWordArt')!==0||child(body,'prstTxWarp')&&attr(child(body,'prstTxWarp'),'prst')!=='textNoShape')issue('회전·세로쓰기·다단·워드아트 텍스트는 아직 지원하지 않습니다.',id);
              const runStyle=(rPr,pPr)=>{
                const def=child(pPr,'defRPr'),get=k=>attr(rPr,k,attr(def,k)),latin=child(rPr,'latin')??child(def,'latin');
                if(![null,'noStrike'].includes(get('strike'))||![null,'none'].includes(get('cap'))||Number(get('baseline')??0)!==0||![null,'none','sng'].includes(get('u'))||elements(rPr).some(n=>['highlight','effectLst','effectDag','gradFill','blipFill','uFill','uLn','hlinkClick','hlinkMouseOver'].includes(n.localName)))issue('취소선·대소문자 변환·첨자·효과 등 이 글자 서식은 아직 지원하지 않습니다.',id);
                const outline=child(rPr,'ln')??child(def,'ln');let textOutline;
                if(outline&&!child(outline,'noFill')){
                  if(!child(outline,'solidFill')||elements(outline).some(n=>n.localName!=='solidFill')||all(outline,'alpha').some(n=>num(n,'val')!==100000))issue('불투명 단색 외의 글자 외곽선은 지원하지 않습니다.',id);
                  textOutline={color:color(child(outline,'solidFill')),width:px(num(outline,'w'))};
                }
                const font=attr(latin,'typeface'),sizePt=Number(get('sz'))/100;
                if(!font||font.startsWith('+')||!Number.isFinite(sizePt)||sizePt<=0)issue('명시적 글꼴·크기가 없는 상속 텍스트는 아직 지원하지 않습니다.',id);
                return {fontFamily:font??'',fontSize:sizePt*4/3,fontWeight:get('b')==='1'?700:400,fontStyle:get('i')==='1'?'italic':'normal',color:color(child(rPr,'solidFill')??child(def,'solidFill'),'#000000'),letterSpacing:Number(get('spc')??0)/100*4/3,textDecoration:get('u')&&get('u')!=='none'?'underline':'none',...(textOutline?{outline:textOutline}:{})};
              };
              for(const [pi,p] of paras.entries()){
                const pPr=child(p,'pPr');if(elements(pPr).some(n=>/^bu/.test(n.localName)&&n.localName!=='buNone'))issue('글머리표는 아직 지원하지 않습니다.',id);
                if(num(pPr,'marL')||num(pPr,'marR')||num(pPr,'indent')||child(pPr,'tabLst')||['spcBef','spcAft'].some(k=>elements(child(pPr,k)).some(n=>num(n,'val')!==0)))issue('문단 들여쓰기·탭·문단 앞뒤 간격은 아직 지원하지 않습니다.',id);
                if(pi&&attr(pPr,'algn','l')!==attr(child(paras[0],'pPr'),'algn','l'))issue('서로 다른 문단 정렬은 아직 지원하지 않습니다.',id);
                if(pi)runs.push({text:'\n',style:runs.at(-1)?.style??runStyle(child(p,'endParaRPr'),pPr)});
                for(const r of elements(p))if(['r','fld','br'].includes(r.localName))runs.push({text:r.localName==='br'?'\n':all(r,'t').map(t=>t.textContent).join(''),style:runStyle(child(r,'rPr'),pPr)});
              }
              node.content={text:runs.map(r=>r.text).join(''),runs};
              const pPr=child(paras[0],'pPr'),linePt=child(child(pPr,'lnSpc'),'spcPts');
              Object.assign(node.style,runs[0]?.style,{padding:{left:px(num(body,'lIns',91440)),right:px(num(body,'rIns',91440)),top:px(num(body,'tIns',45720)),bottom:px(num(body,'bIns',45720))},align:({l:'left',ctr:'center',r:'right',just:'justify'})[attr(pPr,'algn','l')],valign:({t:'top',ctr:'mid',b:'bottom'})[attr(body,'anchor','t')],lineHeight:linePt?`${num(linePt,'val')/100*4/3}px`:'normal',whiteSpace:'pre-wrap'});
              if(node.style.fill!=='transparent'||node.style.strokeWidth)issue('텍스트와 채움이 결합된 도형은 아직 지원하지 않습니다.',id);
              if(shape&&!['rect','roundRect'].includes(shape))issue('이 텍스트 도형의 모양은 아직 지원하지 않습니다.',id);
            }else if(!['rect','roundRect','ellipse'].includes(shape))issue('지원 도형은 사각형·둥근 사각형·타원입니다.',id);
          }
          slide.nodes.push(node);
        }
      };
      visit(tree);
    }
    for(const e of validateModel(model).errors)fail('INPUT_INVALID',`${e.path}: ${e.message}`);
  }catch(e){fail('INPUT_INVALID',e.message);}
  return done();
}
