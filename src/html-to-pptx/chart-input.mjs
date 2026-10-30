import {parseFragment} from 'parse5';
import {createHash} from 'node:crypto';
import {validateChart} from './charts.mjs';
const attribute=(node,name)=>node.attrs?.find(attr=>attr.name===name)?.value;
const hash=value=>createHash('sha256').update(value).digest('hex');
const canonicalJson=value=>JSON.stringify(value,(_key,item)=>item&&typeof item==='object'&&!Array.isArray(item)?Object.fromEntries(Object.keys(item).sort().map(key=>[key,item[key]])):item);
function canonicalSvg(svg){
  const nodes=typeof svg==='string'?parseFragment(svg).childNodes.filter(node=>node.nodeName!=='#text'||node.value.trim()):[svg];
  if(nodes.length!==1||nodes[0]?.tagName!=='svg')throw new Error('Expected one SVG');
  const encode=node=>[node.nodeName,node.namespaceURI??null,node.value??node.data??null,(node.attrs??[]).map(a=>[a.namespace??'',a.prefix??'',a.name,a.value]).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b),'en')),(node.childNodes??[]).map(encode)];
  return JSON.stringify(encode(nodes[0]));
}
export function chartHashes(content,svg){return {json:hash(canonicalJson(content)),svg:hash(canonicalSvg(svg))};}

export function readChartDeclarations(tree){
  const nodes=[],stack=[tree];while(stack.length){const node=stack.pop();if(!node)continue;nodes.push(node);stack.push(...(node.childNodes??[]).slice().reverse());}
  const figures=nodes.filter(n=>n.tagName==='figure'&&attribute(n,'data-chart-id')!==undefined),scripts=nodes.filter(n=>n.tagName==='script'&&attribute(n,'data-chart-for')!==undefined),charts=[],diagnostics=[];
  const issue=(code,message,node)=>diagnostics.push({code,severity:'error',slideId:null,nodeId:null,sourceRef:{path:`${node?.tagName??'chart'}@${node?.sourceCodeLocation?.startOffset??0}`},message});
  let points=0;
  for(const figure of figures){
    const id=attribute(figure,'data-chart-id'),matches=scripts.filter(n=>attribute(n,'data-chart-for')===id);
    if(!id.trim()||figures.filter(n=>attribute(n,'data-chart-id')===id).length!==1||matches.length!==1){issue('CHART_DATA_REQUIRED','차트 영역과 JSON 데이터가 하나씩 연결되어야 합니다.',figure);continue;}
    const script=matches[0];
    if(attribute(script,'type')?.trim().toLowerCase()!=='application/json'||attribute(script,'src')!==undefined){issue('CHART_DATA_REQUIRED','차트 데이터는 파일 안의 application/json 선언이어야 합니다.',script);continue;}
    let content;
    try{content=JSON.parse((script.childNodes??[]).map(n=>n.value??'').join(''));}catch{issue('CHART_DATA_REQUIRED','차트 JSON 데이터를 읽을 수 없습니다.',script);continue;}
    if(!validateChart(content).valid){issue('INPUT_INVALID','차트 종류·항목·값·축의 지원 범위를 확인해 주세요.',script);continue;}
    const children=(figure.childNodes??[]).filter(n=>n.nodeName!=='#comment'&&(n.nodeName!=='#text'||n.value.trim())&&n!==script),svg=children[0];
    if(children.length>1||children.length===1&&svg.tagName!=='svg'){issue('UNSUPPORTED_OBJECT','차트 영역 안의 추가 설명이나 객체는 별도 영역에 넣어 주세요.',figure);continue;}
    const jsonHash=hash(canonicalJson(content)),beforeJson=attribute(figure,'data-chart-json-sha256'),beforeSvg=attribute(figure,'data-chart-svg-sha256');
    if(svg||beforeJson!==undefined||beforeSvg!==undefined){
      if(!svg||!beforeJson||!beforeSvg||!/^[a-f\d]{64}$/.test(beforeJson)||!/^[a-f\d]{64}$/.test(beforeSvg)||hash(canonicalSvg(svg))!==beforeSvg){issue('CHART_DATA_MISMATCH','차트 표시가 변경되었습니다. JSON 데이터에서 표시를 다시 생성해 주세요.',figure);continue;}
    }
    points+=content.categories.length*content.series.length;
    charts.push({id,content,sourceRef:{path:`figure@${figure.sourceCodeLocation?.startOffset??0}`},jsonHash,dataChanged:beforeJson!==undefined&&beforeJson!==jsonHash});
  }
  for(const script of scripts)if(!figures.some(n=>attribute(n,'data-chart-id')===attribute(script,'data-chart-for')))issue('CHART_DATA_REQUIRED','JSON 데이터에 연결된 차트 영역이 없습니다.',script);
  if(points>10000)issue('LIMIT_EXCEEDED','문서의 차트 데이터는 총 10,000점 이하여야 합니다.');
  return {charts,diagnostics};
}
