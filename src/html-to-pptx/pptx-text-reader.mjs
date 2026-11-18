import {diagnostic} from './model.mjs';
import {EMU_PER_CSS_PX} from './canvas.mjs';

const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const child=(n,name)=>elements(n).find(c=>c.localName===name);
const all=(n,name)=>n?[...n.getElementsByTagNameNS('*',name)]:[];
const attr=(n,k,fallback=null)=>n?.hasAttribute(k)?n.getAttribute(k):fallback;
const num=(n,k,fallback=0)=>Number(attr(n,k,fallback));
const px=n=>n/EMU_PER_CSS_PX;
const A='http://schemas.openxmlformats.org/drawingml/2006/main';

// Shared slide state carries warnings, inherited styling and font decisions.
export function readSlideText(tx,id,{issue,slide,diagnostics,textApproximations,approximatedParagraphs,selectedFont,allowPartial,color,tableStylesDoc,masterDoc,presentation,themeDoc,placeholderContext}){

      const body=child(tx,'bodyPr'),paras=elements(tx).filter(n=>n.localName==='p'),runs=[];
      const approximate=(reasonKey,message)=>{const key=`${id}:${reasonKey}`;if(!textApproximations.has(key)){textApproximations.add(key);diagnostics.push({...diagnostic('PARAGRAPH_SPACING_APPROXIMATED',message,{slideId:slide.id,nodeId:id},'warning'),reasonKey});}};
      // Text-box defaults are inherited per paragraph level; placeholders can
      // also inherit the corresponding layout/master placeholder's text style.
      // Ambiguous placeholder matches are not guessed.
      const owner=tx?.parentNode?.localName==='sp'?tx.parentNode:null;
      const {ph,matchedLayout,phType,masterShape}=placeholderContext(owner);
      const bodyProperty=(name,fallback)=>attr(body,name)??attr(child(child(matchedLayout,'txBody'),'bodyPr'),name)??attr(child(child(masterShape,'txBody'),'bodyPr'),name)??fallback;
      const paragraphAttribute=(pPr,name,fallback)=>{
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return attr(pPr,name,fallback);
        const fromList=list=>attr(child(list,`lvl${level+1}pPr`),name)??attr(child(list,'defPPr'),name);
        const fromPlaceholder=sp=>{
          const text=child(sp,'txBody');if(!text)return null;
          return fromList(child(text,'lstStyle'))??attr(child(elements(text).find(n=>n.localName==='p'),'pPr'),name);
        };
        const styleName=ph?(['title','ctrTitle'].includes(phType)?'titleStyle':['body','subTitle','obj'].includes(phType)?'bodyStyle':'otherStyle'):'otherStyle';
        return attr(pPr,name)??fromList(child(tx,'lstStyle'))??fromPlaceholder(matchedLayout)??fromPlaceholder(masterShape)??fromList(child(child(masterDoc?.documentElement,'txStyles'),styleName))??fromList(child(presentation.documentElement,'defaultTextStyle'))??fallback;
      };
      const paragraphAlignment=pPr=>paragraphAttribute(pPr,'algn','l');
      const paragraphChoice=(pPr,names)=>{
        const select=node=>elements(node).find(n=>names.includes(n.localName));
        const direct=select(pPr);if(direct)return direct;
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return null;
        const fromList=list=>select(child(list,`lvl${level+1}pPr`))??select(child(list,'defPPr'));
        const fromPlaceholder=sp=>{
          const text=child(sp,'txBody');if(!text)return null;
          return fromList(child(text,'lstStyle'))??select(child(elements(text).find(n=>n.localName==='p'),'pPr'));
        };
        const own=fromList(child(tx,'lstStyle'));if(own||!owner)return own;
        const styleName=ph?(['title','ctrTitle'].includes(phType)?'titleStyle':['body','subTitle','obj'].includes(phType)?'bodyStyle':'otherStyle'):'otherStyle';
        return fromPlaceholder(matchedLayout)??fromPlaceholder(masterShape)??fromList(child(child(masterDoc?.documentElement,'txStyles'),styleName))??fromList(child(presentation.documentElement,'defaultTextStyle'));
      };
      const paragraphSetting=(pPr,name)=>paragraphChoice(pPr,[name]);
      const paragraphBullet=pPr=>paragraphChoice(pPr,['buNone','buChar','buAutoNum','buBlip']);
      const paragraphOwnSetting=(pPr,name)=>{
        const direct=child(pPr,name);if(direct)return direct;
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return null;
        const list=child(tx,'lstStyle');
        return child(child(list,`lvl${level+1}pPr`),name)??child(child(list,'defPPr'),name);
      };
      const inheritedRunAttribute=(pPr,name)=>{
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return null;
        const fromList=list=>attr(child(child(list,`lvl${level+1}pPr`),'defRPr'),name)??attr(child(child(list,'defPPr'),'defRPr'),name);
        const fromPlaceholder=sp=>{
          const text=child(sp,'txBody');if(!text)return null;
          const first=elements(text).find(n=>n.localName==='p');
          return fromList(child(text,'lstStyle'))??attr(child(child(first,'pPr'),'defRPr'),name);
        };
        const own=fromList(child(tx,'lstStyle'));
        if(own!==null)return own;
        if(!owner)return null;
        if(ph){
          const styleName=['title','ctrTitle'].includes(phType)?'titleStyle':['body','subTitle','obj'].includes(phType)?'bodyStyle':'otherStyle';
          const inherited=fromPlaceholder(matchedLayout)??fromPlaceholder(masterShape)??fromList(child(child(masterDoc?.documentElement,'txStyles'),styleName));
          if(inherited!==null)return inherited;
        }
        return fromList(child(presentation.documentElement,'defaultTextStyle'));
      };
      const inheritedSize=(pPr)=>{
        if(!owner)return null;
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return null;
        const fromList=list=>attr(child(child(list,`lvl${level+1}pPr`),'defRPr'),'sz',attr(child(child(list,'defPPr'),'defRPr'),'sz'));
        const placeholderSize=sp=>{
          const text=child(sp,'txBody');if(!text)return null;
          const listed=fromList(child(text,'lstStyle'));if(listed!==null)return listed;
          const values=[...all(text,'defRPr'),...all(text,'rPr'),...all(text,'endParaRPr')].map(n=>attr(n,'sz')).filter(v=>v!==null);
          return new Set(values).size===1?values[0]:null;
        };
        const styleName=ph?(['title','ctrTitle'].includes(phType)?'titleStyle':['body','subTitle','obj'].includes(phType)?'bodyStyle':'otherStyle'):'otherStyle';
        const styles=child(child(masterDoc?.documentElement,'txStyles'),styleName);
        return fromList(child(tx,'lstStyle'))??placeholderSize(matchedLayout)??placeholderSize(masterShape)??fromList(styles)??fromList(child(presentation.documentElement,'defaultTextStyle'));
      };
      const inheritedFill=(pPr)=>{
        if(!owner)return null;
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return null;
        const fromList=list=>child(child(child(list,`lvl${level+1}pPr`),'defRPr'),'solidFill')??child(child(child(list,'defPPr'),'defRPr'),'solidFill');
        const fromPlaceholder=sp=>{
          const text=child(sp,'txBody');if(!text)return null;
          const listed=fromList(child(text,'lstStyle'));if(listed)return listed;
          const fills=[...all(text,'defRPr'),...all(text,'rPr'),...all(text,'endParaRPr')].map(n=>child(n,'solidFill')).filter(Boolean);
          return new Set(fills.map(n=>n.toString())).size===1?fills[0]:null;
        };
        const styleName=ph?(['title','ctrTitle'].includes(phType)?'titleStyle':['body','subTitle','obj'].includes(phType)?'bodyStyle':'otherStyle'):'otherStyle';
        const fontReference=child(child(owner,'style'),'fontRef');
        return fromList(child(tx,'lstStyle'))??(!ph&&elements(fontReference).length?fontReference:null)??fromPlaceholder(matchedLayout)??fromPlaceholder(masterShape)??fromList(child(child(masterDoc?.documentElement,'txStyles'),styleName))??fromList(child(presentation.documentElement,'defaultTextStyle'));
      };
      const inheritedTypeface=(pPr,kind)=>{
        if(!owner)return null;
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return null;
        const fromList=list=>attr(child(child(child(list,`lvl${level+1}pPr`),'defRPr'),kind),'typeface')??attr(child(child(child(list,'defPPr'),'defRPr'),kind),'typeface');
        const fromPlaceholder=sp=>{
          const text=child(sp,'txBody');if(!text)return null;
          const listed=fromList(child(text,'lstStyle'));if(listed)return listed;
          const values=[...all(text,'defRPr'),...all(text,'rPr'),...all(text,'endParaRPr')].map(n=>attr(child(n,kind),'typeface')).filter(Boolean);
          return new Set(values).size===1?values[0]:null;
        };
        const styleName=ph?(['title','ctrTitle'].includes(phType)?'titleStyle':['body','subTitle','obj'].includes(phType)?'bodyStyle':'otherStyle'):'otherStyle';
        const referenced=attr(child(child(owner,'style'),'fontRef'),'idx');
        const suffix=kind==='ea'?'ea':'lt';
        return fromList(child(tx,'lstStyle'))??fromPlaceholder(matchedLayout)??fromPlaceholder(masterShape)??fromList(child(child(masterDoc?.documentElement,'txStyles'),styleName))??fromList(child(presentation.documentElement,'defaultTextStyle'))??(referenced==='major'?`+mj-${suffix}`:referenced==='minor'?`+mn-${suffix}`:null);
      };
      const resolveTypeface=(value,script)=>{
        if(!value)return null;
        const alias=/^\+(mj|mn)-(lt|ea|cs)$/.exec(value);
        if(!alias)return value.startsWith('+')?null:value;
        const group=child(all(themeDoc,'fontScheme')[0],alias[1]==='mj'?'majorFont':'minorFont');
        const kind=({lt:'latin',ea:'ea',cs:'cs'})[alias[2]];
        const direct=attr(child(group,kind),'typeface');if(direct)return direct;
        if(kind==='ea'){
          const languageScript=script==='hangul'?'Hang':script==='japanese'?'Jpan':script==='han'?'Hans':null;
          const match=elements(group).find(n=>n.localName==='font'&&attr(n,'script')===languageScript);
          const specified=attr(match,'typeface');if(specified)return specified;
        }
        return null;
      };
      const scriptOf=character=>/\p{Script=Hangul}/u.test(character)?'hangul':/[\p{Script=Hiragana}\p{Script=Katakana}]/u.test(character)?'japanese':/\p{Script=Han}/u.test(character)?'han':'latin';
      const splitScripts=text=>{
        const pieces=[];for(const char of text){const script=scriptOf(char),last=pieces.at(-1);if(last&&last.script===script)last.text+=char;else pieces.push({text:char,script});}return pieces;
      };
      const spacingValue=node=>{
        const raw=attr(node,'val','').trim();
        if(node?.localName==='spcPct'&&/^\+?(?:\d+(?:\.\d+)?|\.\d+)%$/.test(raw))return Number(raw.slice(0,-1))*1000;
        return /^\+?\d+$/.test(raw)&&Number.isSafeInteger(Number(raw))?Number(raw):NaN;
      };
      const checkedSpacing=spacing=>{
        const parts=elements(spacing),item=parts[0];
        if(parts.length!==1||item.namespaceURI!==A||!['spcPct','spcPts'].includes(item.localName)||elements(item).length)throw new Error('문단의 간격 구조가 올바르지 않습니다.');
        const value=spacingValue(item),max=item.localName==='spcPct'?13200000:158400;
        if(!Number.isFinite(value)||value<0||value>max)throw new Error('문단의 간격 값이 올바르지 않습니다.');
        return {kind:item.localName,value};
      };
      const spacingPx=(pPr,name,paragraph)=>{
        const spacing=paragraphSetting(pPr,name);if(!spacing)return 0;
        const {kind,value}=checkedSpacing(spacing);
        if(kind==='spcPct'&&value===0)return 0;
        if(kind==='spcPts')return value/100*4/3;
        const fallback=attr(child(pPr,'defRPr'),'sz')??inheritedSize(pPr);
        const sizes=elements(paragraph).filter(r=>['r','fld'].includes(r.localName)).map(r=>Number(attr(child(r,'rPr'),'sz',fallback)));
        if(!sizes.length)sizes.push(Number(attr(child(paragraph,'endParaRPr'),'sz',fallback)));
        const largest=Math.max(...sizes);return Number.isFinite(largest)&&largest>0?largest/100*4/3*1.2*value/100000:null;
      };
      const paragraphLineHeight=pPr=>{
        const spacing=paragraphSetting(pPr,'lnSpc');if(!spacing)return '1.2';
        const {kind,value}=checkedSpacing(spacing);if(value<=0)return null;
        return kind==='spcPts'?`${value/100*4/3}px`:`${Number((value*1.2/100000).toFixed(6))}`;
      };
      // Keep individual paragraph layout only where its spacing can be read
      // without guessing. Empty paragraphs still occupy a line in PowerPoint.
      const paragraphLayout=paras.map(p=>{
        const pPr=child(p,'pPr'),tableCell=tx?.parentNode?.localName==='tc';
        for(const key of ['marL','marR','indent'])if(pPr?.hasAttribute(key)){
          const raw=pPr.getAttribute(key).trim(),value=Number(raw);
          if(!/^[+-]?\d+$/.test(raw)||!Number.isSafeInteger(value)||key!=='indent'&&value<0)throw new Error('문단의 들여쓰기 값이 올바르지 않습니다.');
        }
        for(const name of ['lnSpc','spcBef','spcAft']){const spacing=child(pPr,name);if(spacing)checkedSpacing(spacing);}
        // Office defaults hanging punctuation to true when it is unspecified.
        const hanging=tableCell?paragraphAttribute(pPr,'hangingPunct','1'):null;
        if(tableCell&&!['0','1','false','true'].includes(hanging))throw new Error('문단의 문장부호 배치 값이 올바르지 않습니다.');
        const marker=paragraphBullet(pPr),level=Number(attr(pPr,'lvl','0'));
        const bullet=marker?.localName==='buChar'?{type:'char',char:attr(marker,'char',''),level}:marker?.localName==='buAutoNum'?{type:'number',format:attr(marker,'type'),start:Number(attr(marker,'startAt','1')),restart:attr(marker,'startAt')!==null,level}:null;
        if(bullet){if(paragraphChoice(pPr,['buClr','buClrTx'])?.localName==='buClrTx')bullet.followTextColor=true;if(paragraphChoice(pPr,['buFont','buFontTx'])?.localName==='buFontTx')bullet.followTextFont=true;}
        const bulletSize=bullet&&paragraphChoice(pPr,['buSzTx','buSzPct','buSzPts']);
        if(bulletSize){const kind=bulletSize.localName,value=num(bulletSize,'val',NaN);if(kind==='buSzTx'||Number.isInteger(value)&&value>0&&value<=400000)bullet.sourceSize=kind==='buSzTx'?{kind}:{kind,value};}
        const sourceSpacing={};for(const name of ['lnSpc','spcBef','spcAft']){const spacing=paragraphSetting(pPr,name);if(spacing)sourceSpacing[name]=checkedSpacing(spacing);else if(name==='lnSpc')sourceSpacing[name]={kind:'default'};}
        const end=child(p,'endParaRPr'),attrs={};for(const key of ['spc','sz','kern','b','i']){const value=attr(end,key);if(value!==null)attrs[key]=value;}
        const sourceEnd=end?{mode:'explicit',attrs}:{mode:'implicit'};
        return {beforePx:spacingPx(pPr,'spcBef',p),afterPx:spacingPx(pPr,'spcAft',p),leftPx:px(Number(paragraphAttribute(pPr,'marL','0'))),rightPx:px(Number(paragraphAttribute(pPr,'marR','0'))),indentPx:px(Number(paragraphAttribute(pPr,'indent','0'))),lineHeight:paragraphLineHeight(pPr),align:({l:'left',ctr:'center',r:'right',just:'justify'})[paragraphAlignment(pPr)],...(bullet?{bullet}:{}),...(Object.keys(sourceSpacing).length?{sourceSpacing}:{}),sourceEnd,...(tableCell?{hangingPunct:['1','true'].includes(hanging)}:{})};
      });
      const useParagraphLayout=Boolean((owner||tx?.parentNode?.localName==='tc')&&paras.length>0&&paras.every((p,i)=>{
        const layout=paragraphLayout[i];return layout.beforePx!==null&&layout.afterPx!==null&&layout.lineHeight&&layout.align&&[layout.leftPx,layout.rightPx].every(n=>Number.isFinite(n)&&n>=0&&n<=5000)&&Number.isFinite(layout.indentPx)&&Math.abs(layout.indentPx)<=5000;
      })&&(paragraphLayout.some(p=>p.bullet||p.leftPx||p.rightPx||p.indentPx)||
        paras.some(p=>{
          const end=child(p,'endParaRPr');
          return ['spc','b','i'].some(key=>end?.hasAttribute(key))||!end&&elements(p).some(run=>['r','fld'].includes(run.localName)&&['spc','b','i'].some(key=>child(run,'rPr')?.hasAttribute(key)));
        })||
        paras.some(p=>{
          const properties=child(p,'pPr'),line=paragraphOwnSetting(properties,'lnSpc');
          return ['spcBef','spcAft'].some(k=>paragraphOwnSetting(properties,k))||line&&(paras.length>1||child(line,'spcPct'));
        })||
        paras.length>1&&paragraphLayout.some(p=>p.align!==paragraphLayout[0].align)||
        ['1','true'].includes(attr(body,'spcFirstLastPara','0'))&&paras.some(p=>['spcBef','spcAft'].some(k=>paragraphOwnSetting(child(p,'pPr'),k)))));
      const paragraphRuns=[];
      let inheritedTextColorEntry=null,inheritedTextColorAmbiguous=false;
      const cell=tx?.parentNode,row=cell?.parentNode,table=row?.parentNode,tableProps=child(table,'tblPr');
      if(cell?.localName==='tc'&&row?.localName==='tr'&&table?.localName==='tbl'&&tableStylesDoc){
        const styleId=child(tableProps,'tableStyleId')?.textContent,style=elements(tableStylesDoc.documentElement).find(n=>n.localName==='tblStyle'&&attr(n,'styleId')===styleId);
        const isFirstRow=elements(table).filter(n=>n.localName==='tr')[0]===row&&['1','true'].includes(attr(tableProps,'firstRow','0'));
        const rows=elements(table).filter(n=>n.localName==='tr'),cells=elements(row).filter(n=>n.localName==='tc'),col=cells.indexOf(cell);
        const enabled=key=>['1','true'].includes(attr(tableProps,key,'0'));
        const mixedCell=col<0||enabled('firstCol')&&col===0||enabled('lastCol')&&col===cells.length-1||enabled('lastRow')&&rows.length===1;
        if(style&&isFirstRow){
          const textStyle=child(child(style,'firstRow'),'tcTxStyle')??child(child(style,'wholeTbl'),'tcTxStyle');
          const entry=elements(textStyle).find(n=>['schemeClr','srgbClr','sysClr'].includes(n.localName));
          if(entry){inheritedTextColorEntry=entry;inheritedTextColorAmbiguous=mixedCell;}
        }
      }
      if(num(body,'rot')!==0||attr(body,'vert','horz')!=='horz'||num(body,'numCol',1)!==1||num(body,'rtlCol')!==0||num(body,'fromWordArt')!==0||child(body,'prstTxWarp')&&attr(child(body,'prstTxWarp'),'prst')!=='textNoShape')issue('회전·세로쓰기·다단·워드아트 텍스트는 아직 지원하지 않습니다.',id);
      const runStyle=(rPr,pPr,script='latin')=>{
        const def=child(pPr,'defRPr'),get=k=>attr(rPr,k,attr(def,k,['b','i','spc','kern'].includes(k)?inheritedRunAttribute(pPr,k):null)),kind=script==='latin'?'latin':'ea';
        const baseline=Number(get('baseline')??0);
        if(!Number.isInteger(baseline)||Math.abs(baseline)>100000)issue('지원 범위를 벗어난 위첨자·아래첨자 위치입니다.',id);
        const highlighted=child(rPr,'highlight')??child(def,'highlight');
        let highlight=highlighted?color(highlighted):null;
        if(highlight&&!/^#[a-f\d]{6}$/i.test(highlight)){
          highlight=null;
          if(allowPartial)approximate('PPTX_TEXT_HIGHLIGHT_APPROXIMATED','반투명 형광펜 배경 없이 글자를 표시했습니다.');else issue('반투명 글자 형광펜 표시는 아직 지원하지 않습니다.',id);
        }
        for(const [property,allowed,message] of [
          ['strike',[null,'noStrike'],'글자 취소선 서식은 아직 지원하지 않습니다.'],
          ['cap',[null,'none'],'글자 대소문자 변환은 아직 지원하지 않습니다.'],
          ['u',[null,'none','sng'],'이 글자 밑줄 서식은 아직 지원하지 않습니다.']
        ]){const value=get(property);if(!allowed.includes(value))issue(message,id,'PPTX_TEXT_STYLE_UNSUPPORTED',{property,value});}
        for(const effect of elements(rPr)){
          if(effect.localName==='effectLst')for(const item of elements(effect))issue('글자 효과는 아직 지원하지 않습니다.',id,'PPTX_TEXT_STYLE_UNSUPPORTED',{property:'effect',value:item.localName});
          else if(['effectDag','gradFill','blipFill','uFill','uLn','hlinkClick','hlinkMouseOver'].includes(effect.localName))issue('글자 효과 또는 연결은 아직 지원하지 않습니다.',id,'PPTX_TEXT_STYLE_UNSUPPORTED',{property:'effect',value:effect.localName});
        }
        const outline=child(rPr,'ln')??child(def,'ln');let textOutline;
        const outlinePaint=child(outline,'solidFill'),paintColors=elements(outlinePaint),paintEffects=elements(paintColors[0]);
        const invisibleOutline=outline&&(outline.hasAttribute('w')&&num(outline,'w')===0||paintColors.length===1&&paintEffects.length===1&&paintEffects[0].localName==='alpha'&&num(paintEffects[0],'val',NaN)===0);
        if(outline&&!child(outline,'noFill')&&!invisibleOutline){
          if(!child(outline,'solidFill')||elements(outline).some(n=>n.localName!=='solidFill')||all(outline,'alpha').some(n=>num(n,'val')!==100000))issue('불투명 단색 외의 글자 외곽선은 지원하지 않습니다.',id);
          textOutline={color:color(child(outline,'solidFill')),width:px(num(outline,'w'))};
        }
        const declaredFont=attr(child(rPr,kind),'typeface')??attr(child(def,kind),'typeface')??inheritedTypeface(pPr,kind);
        const font=resolveTypeface(declaredFont,script)??(kind==='ea'?resolveTypeface(attr(child(rPr,'latin'),'typeface')??attr(child(def,'latin'),'typeface')??inheritedTypeface(pPr,'latin'),'latin'):null);
        const declaredSize=get('sz'),inherited=declaredSize===null?inheritedSize(pPr):null;
        let sizePt=Number(declaredSize??inherited)/100;
        if(declaredSize===null&&inherited===null&&allowPartial&&(selectedFont||font&&!font.startsWith('+'))){sizePt=18;approximate('PPTX_INHERITED_SIZE_APPROXIMATED','상속된 글자 크기를 확인하지 못해 18pt로 표시했습니다.');}
        if((!selectedFont&&(!font||font.startsWith('+')))||!Number.isFinite(sizePt)||sizePt<=0)issue('명시적 글꼴·크기가 없는 상속 텍스트는 아직 지원하지 않습니다.',id);
        const kerningValue=get('kern'),kerning=kerningValue===null?null:Number(kerningValue.trim());
        const validKerning=kerningValue===null||/^\+?\d+$/.test(kerningValue.trim())&&Number.isInteger(kerning)&&kerning<=400000;
        if(!validKerning)issue('글자 커닝 기준 크기가 올바르지 않습니다.',id);
        const disableKerning=kerning!==null&&validKerning&&(kerning===0||sizePt*100<kerning);
        const textFill=child(rPr,'solidFill')??child(def,'solidFill')??inheritedFill(pPr);
        if(!textFill&&inheritedTextColorEntry&&inheritedTextColorAmbiguous){
          if(allowPartial)approximate('PPTX_TABLE_TEXT_COLOR_APPROXIMATED','표 교차 셀의 상속 글자색은 원본과 다를 수 있습니다.');
          else issue('표 교차 셀의 상속 글자색을 확정할 수 없습니다.',id);
        }
        const fontReference=child(child(owner,'style'),'fontRef');
        const textColor=textFill?color(textFill,'#000000'):inheritedTextColorEntry?color({childNodes:[inheritedTextColorEntry]},'#000000'):fontReference&&elements(fontReference).length?color(fontReference,'#000000'):'#000000';
        return {fontFamily:selectedFont??font??'',fontSize:sizePt*4/3,fontWeight:['1','true'].includes(get('b'))?700:400,fontStyle:['1','true'].includes(get('i'))?'italic':'normal',color:textColor,letterSpacing:Number(get('spc')??0)/100*4/3,textDecoration:get('u')&&get('u')!=='none'?'underline':'none',...(disableKerning?{disableKerning:true}:{}),...(baseline&&Number.isInteger(baseline)&&Math.abs(baseline)<=100000?{baseline}:{}),...(textOutline?{outline:textOutline}:{}),...(highlight?{highlight}:{})};
      };
      const bulletStyle=(p,pPr,characterBullet=false)=>{
        const firstRun=elements(p).find(n=>['r','fld'].includes(n.localName));
        const style=runStyle(child(firstRun,'rPr')??child(p,'endParaRPr'),pPr);
        delete style.highlight;delete style.outline;delete style.baseline;style.textDecoration='none';
        // PowerPoint renders character bullets with their own font face's
        // default weight, independently of the first text run's bold setting.
        if(characterBullet)style.fontWeight=400;
        const font=paragraphChoice(pPr,['buFont','buFontTx']);
        if(!selectedFont&&font?.localName==='buFont'){
          const family=resolveTypeface(attr(font,'typeface'),'latin');
          if(family)style.fontFamily=family;else approximate('PPTX_PARAGRAPH_INDENT_APPROXIMATED','글머리표 글꼴을 확인하지 못해 본문 글꼴로 표시했습니다.');
        }
        const size=paragraphChoice(pPr,['buSzTx','buSzPct','buSzPts']);
        if(size&&size.localName!=='buSzTx'){
          const value=num(size,'val',NaN),scaled=size.localName==='buSzPct'?style.fontSize*value/100000:value/100*4/3;
          if(Number.isFinite(value)&&value>0&&Number.isFinite(scaled)&&scaled>0&&scaled<=5000)style.fontSize=scaled;
          else approximate('PPTX_PARAGRAPH_INDENT_APPROXIMATED','글머리표 크기를 확인하지 못해 본문 크기로 표시했습니다.');
        }
        const bulletColor=paragraphChoice(pPr,['buClr','buClrTx']);if(bulletColor?.localName==='buClr')style.color=color(bulletColor,style.color);
        return style;
      };
      const previousNumbers=[];
      for(const [pi,p] of paras.entries()){
        const pPr=child(p,'pPr'),choice=paragraphBullet(pPr),bullet=choice?.localName==='buChar'?choice:null,bulletChar=attr(bullet,'char',''),auto=choice?.localName==='buAutoNum'?choice:null;
        let numberLabel=null;
        if(auto&&attr(auto,'type')==='arabicPeriod'){
          const start=attr(auto,'startAt'),level=Number(attr(pPr,'lvl','0')),parsed=start===null?null:Number(start);
          if(Number.isInteger(level)&&level>=0&&level<=8&&(parsed===null||Number.isInteger(parsed)&&parsed>=1&&parsed<=32767)){
            const groupStart=parsed??1,previous=previousNumbers[level],restart=!previous||previous.groupStart!==groupStart;
            const count=restart?groupStart:previous.count+1;
            if(count<=32767){numberLabel=`${count}.`;previousNumbers[level]={groupStart,count};previousNumbers.length=level+1;if(paragraphLayout[pi].bullet){paragraphLayout[pi].bullet.start=count;paragraphLayout[pi].bullet.restart=restart;}}
          }
        }
        if(!auto){const level=Number(attr(pPr,'lvl','0'));previousNumbers.length=Number.isInteger(level)&&level>=0&&level<=8?Math.min(previousNumbers.length,level):0;}
        const hasInheritedBulletProperties=elements(pPr).some(n=>/^bu/.test(n.localName))&&!bullet&&!child(pPr,'buNone');
        const unsupportedBullet=auto&&numberLabel===null||choice?.localName==='buBlip'||hasInheritedBulletProperties&&!auto||(bullet&&(!bulletChar||/[\u0000-\u001f]/.test(bulletChar)||[...bulletChar].length!==1));
        if(unsupportedBullet)issue('이 글머리표는 아직 지원하지 않습니다.',id);
        const usesTab=all(p,'t').some(t=>t.textContent.includes('\t'))||elements(p).some(n=>n.localName==='tab');
        if(usesTab){if(allowPartial)approximate('PPTX_TEXT_TAB_APPROXIMATED','기본 탭 간격으로 표시했으며 개별 탭 위치는 원본과 다를 수 있습니다.');else issue('문단 탭은 아직 지원하지 않습니다.',id);}
        if(((bullet||auto)&&!unsupportedBullet||num(pPr,'marL')||num(pPr,'marR')||num(pPr,'indent'))&&!approximatedParagraphs.has(id)){
          diagnostics.push({...diagnostic('PARAGRAPH_SPACING_APPROXIMATED','문단 들여쓰기와 글머리 배치는 근사했습니다.',{slideId:slide.id,nodeId:id},'warning'),reasonKey:'PPTX_PARAGRAPH_INDENT_APPROXIMATED'});approximatedParagraphs.add(id);
        }
        if(!useParagraphLayout&&['spcBef','spcAft'].some(k=>{const spacing=paragraphOwnSetting(pPr,k);return spacing&&checkedSpacing(spacing).value!==0;}))approximate('PPTX_PARAGRAPH_SPACING_APPROXIMATED','문단 앞뒤 간격은 그대로 옮기지 못했습니다.');
        if(useParagraphLayout&&['spcBef','spcAft'].some(k=>{const spacing=paragraphSetting(pPr,k),value=spacing&&checkedSpacing(spacing);return value?.kind==='spcPct'&&value.value!==0;}))approximate('PPTX_PARAGRAPH_SPACING_APPROXIMATED','비율로 지정한 문단 간격의 원래 값은 보존했으며 HTML 표시는 근사했습니다.');
        if(!paragraphLineHeight(pPr)&&paragraphSetting(pPr,'lnSpc')){
          if(allowPartial)approximate('PPTX_LINE_SPACING_APPROXIMATED','줄 간격 0은 기본 간격으로 표시했습니다.');else issue('줄 간격 0은 아직 지원하지 않습니다.',id);
        }
        if(!useParagraphLayout&&pi&&paragraphAlignment(pPr)!==paragraphAlignment(child(paras[0],'pPr'))){
          if(allowPartial&&['l','ctr','r','just'].includes(paragraphAlignment(pPr)))approximate('PPTX_PARAGRAPH_ALIGNMENT_APPROXIMATED','서로 다른 문단 정렬은 첫 문단의 정렬로 표시했습니다.');
          else issue('서로 다른 문단 정렬은 아직 지원하지 않습니다.',id);
        }
        if(pi)runs.push({text:'\n',style:runs.at(-1)?.style??runStyle(child(p,'endParaRPr'),pPr)});
        const paragraphStart=runs.length;
        if(bullet&&!unsupportedBullet&&all(p,'t').some(t=>t.textContent.trim()))runs.push({text:`${bulletChar} `,style:bulletStyle(p,pPr,true),marker:true});
        if(numberLabel&&all(p,'t').some(t=>t.textContent.trim()))runs.push({text:`${numberLabel} `,style:bulletStyle(p,pPr),marker:true});
        for(const r of elements(p))if(['r','fld','br','tab'].includes(r.localName)){
          const text=r.localName==='br'?'\n':r.localName==='tab'?'\t':all(r,'t').map(t=>t.textContent).join('');
          if(r.localName==='tab'||r.localName==='br'){runs.push({text,style:runs.at(-1)?.style??runStyle(child(p,'endParaRPr'),pPr)});continue;}
          const start=runs.length;
          for(const part of selectedFont?[{text,script:'latin'}]:splitScripts(text)){
            const style=runStyle(child(r,'rPr'),pPr,part.script),last=runs.at(-1);
            if(runs.length>start&&JSON.stringify(last.style)===JSON.stringify(style))last.text+=part.text;
            else runs.push({text:part.text,style});
          }
        }
        if(useParagraphLayout){
          const current=runs.slice(paragraphStart);
          if(!current.length){
            const firstRun=elements(p).find(n=>['r','fld'].includes(n.localName));
            const properties=child(firstRun,'rPr')??child(p,'endParaRPr');
            current.push({text:'',style:properties?runStyle(properties,pPr):runs.at(-1)?.style??runStyle(null,pPr)});
          }
          paragraphRuns.push(current);
        }
      }
      const content={text:runs.map(r=>r.text).join(''),runs};
      if(useParagraphLayout){
        const edgeSpacing=['1','true'].includes(attr(body,'spcFirstLastPara','0'));
        content.paragraphs=paragraphLayout.map((layout,index)=>({...layout,beforePx:index===0&&!edgeSpacing?0:layout.beforePx,afterPx:index===paras.length-1&&!edgeSpacing?0:layout.afterPx,runs:paragraphRuns[index]}));
      }
      const pPr=child(paras[0],'pPr');
      // PowerPoint's percentage is relative to its single-line advance,
      // which is 1.2 times the point size in the supported text bodies.
      const lineHeight=paragraphLineHeight(pPr)??'normal';
      const textRun=runs.find(run=>!run.marker&&run.text.trim());
      const style=Object.assign({},textRun?.style??runs[0]?.style??runStyle(child(paras[0],'endParaRPr'),pPr),{padding:{left:px(Number(bodyProperty('lIns',91440))),right:px(Number(bodyProperty('rIns',91440))),top:px(Number(bodyProperty('tIns',45720))),bottom:px(Number(bodyProperty('bIns',45720)))},align:({l:'left',ctr:'center',r:'right',just:'justify'})[paragraphAlignment(pPr)],valign:({t:'top',ctr:'mid',b:'bottom'})[bodyProperty('anchor','t')],lineHeight,whiteSpace:bodyProperty('wrap','square')==='none'?'pre':'pre-wrap'});
      if(runs.some(run=>run.text.includes('\t'))){
        const raw=attr(pPr,'defTabSz'),tabSize=raw===null?96:px(Number(raw));
        if(Number.isFinite(tabSize)&&tabSize>0&&tabSize<=5000)style.tabSize=tabSize;
      }
      // Auto-sized Office text can extend slightly into the right inset before
      // breaking a line. Preserve the actual inset separately for re-export.
      if(owner&&child(body,'spAutoFit')&&style.whiteSpace==='pre-wrap')style.wrapAllowance=Math.min(style.padding.right/2,style.fontSize/4);
        return {content,style};

}
