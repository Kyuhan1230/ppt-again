import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
import path from 'node:path';
import {addNativeChart} from '../src/html-to-pptx/write-chart.mjs';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {renderPowerPointChartImages} from '../src/html-to-pptx/powerpoint-chart-images.mjs';

const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGOQTDkBAAHfAUZ76qvRAAAAAElFTkSuQmCC','base64');
let fixturePromise;
function fixture(){
 fixturePromise??=(async()=>{
  const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';let count=0;
  const cats=['A','B','C','D'],rows=[[15,24,18,28],[12,19,23,17],[22,16,14,26]],colors=['2878B7','E6803D','35A184'];
  const series=()=>rows.map((values,i)=>({name:`Series ${i+1}`,values,color:`#${colors[i]}`}));
  const native=(slide,id,type)=>{count++;addNativeChart(slide,{id,kind:'chart',boxCssPx:{x:60,y:135,w:790,h:390},style:{fontFamily:'Arial',fontSize:15},content:{type,categories:cats,series:series(),title:id,legend:true,axes:{}}},{pointScale:.75,resolveFont:()=>({fontFace:'Arial'})});return count;};
  const raw=(slide,type,id,data,extra={})=>{count++;const opts={x:.65,y:1.45,w:8.35,h:4.3,objectName:id,showTitle:true,title:id,titleFontFace:'Arial',titleFontSize:15,fontFace:'Arial',showLegend:true,legendPos:'b',chartColors:colors,...extra};if(Array.isArray(type))slide.addChart(type,opts);else slide.addChart(type,data,opts);return count;};
  const slide=()=>{const s=ppt.addSlide();s.background={color:'FFFFFF'};s.addText('Synthetic data',{x:.5,y:.3,w:6,h:.5,fontFace:'Arial',fontSize:18});return s;};
  native(slide(),'matrix-1','bar');
  {const s=slide();native(s,'matrix-2a','bar');native(s,'matrix-2b','line');}
  const gantt=slide(),starts=['2026-01-05','2026-01-12','2026-01-20','2026-02-01'].map(x=>(Date.parse(`${x}T00:00:00Z`)-Date.UTC(1904,0,1))/86400000);
  const ganttIndex=raw(gantt,'bar','gantt-1904',[{name:'Start',labels:cats,values:starts},{name:'Duration',labels:cats,values:[8,11,9,12]}],{barDir:'bar',barGrouping:'stacked',barOverlapPct:100,valAxisLabelFormatCode:'yyyy-mm-dd'});
  gantt.addTable([['Task','Start'],['A','2026-01-05']],{x:9,y:2,w:3,h:1,fontFace:'Arial',fontSize:14,objectName:'gantt-table'});
  const data=()=>rows.map((values,i)=>({name:`Series ${i+1}`,labels:cats,values}));
  raw(slide(),'bar','stacked',data(),{barDir:'bar',barGrouping:'stacked',barOverlapPct:100});
  raw(slide(),'bar','percent-stacked',data(),{barDir:'col',barGrouping:'percentStacked',barOverlapPct:100});
  raw(slide(),'scatter','scatter',[{name:'X',values:[1,2,3,4]},{name:'Y1',values:[2,4,3,7]},{name:'Y2',values:[3,2,6,5]}]);
  raw(slide(),'bubble','bubble',[{name:'X',values:[1,2,3,4]},{name:'Y1',values:[2,4,3,7],sizes:[4,7,5,9]},{name:'Y2',values:[3,2,6,5],sizes:[6,4,8,5]}]);
  raw(slide(),[{type:'bar',data:[{name:'Actual',labels:cats,values:rows[0]}]},{type:'line',data:[{name:'Plan',labels:cats,values:rows[1]}]}],'combo',[]);
  raw(slide(),'doughnut','doughnut',[{name:'Share',labels:cats,values:rows[0]}]);
  raw(slide(),'radar','radar',data().slice(0,2));raw(slide(),'area','area',data());
  native(slide(),'native-line','line');native(slide(),'native-bar','bar');
  const rounded=native(slide(),'rounded-inventory','bar'),date=native(slide(),'date-shortage','bar');
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
  const patch=(key,pattern,replacement)=>{const before=strFromU8(parts[key]),after=before.replace(pattern,replacement);assert.notEqual(after,before,key);parts[key]=strToU8(after);};
  patch(`ppt/charts/chart${rounded}.xml`,/<c:roundedCorners(?:\s+val="[^"]*")?\s*\/>/,'<c:roundedCorners val="1"/>');
  for(const n of [ganttIndex,date]){
   patch(`ppt/charts/chart${n}.xml`,/<c:date1904(?:\s+val="[^"]*")?\s*\/>/,'<c:date1904 val="1"/>');
   const rel=`ppt/charts/_rels/chart${n}.xml.rels`,doc=new DOMParser().parseFromString(strFromU8(parts[rel]),'application/xml'),link=[...doc.getElementsByTagNameNS('*','Relationship')].find(x=>x.getAttribute('Type').endsWith('/package'));
   const target=path.posix.normalize(path.posix.join('ppt/charts',link.getAttribute('Target'))),wb=unzipSync(parts[target]),key='xl/workbook.xml',before=strFromU8(wb[key]);
   const after=/<workbookPr\b/.test(before)?before.replace(/<workbookPr([^>]*)\/>/,(_m,a)=>`<workbookPr${a.replace(/\sdate1904="[^"]*"/,'')} date1904="1"/>`):before.replace(/(<workbook\b[^>]*>)/,'$1<workbookPr date1904="1"/>');assert.notEqual(after,before);wb[key]=strToU8(after);parts[target]=zipSync(wb);
  }
  // PPTXGenJS may repeat shape IDs when charts and tables share a slide.
  for(const key of Object.keys(parts).filter(x=>/^ppt\/slides\/slide\d+\.xml$/.test(x))){const source=strFromU8(parts[key]),ids=[...source.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map(m=>Number(m[1])),seen=new Set();let next=Math.max(0,...ids)+1;parts[key]=strToU8(source.replace(/(<p:cNvPr\b[^>]*\bid=")(\d+)(")/g,(_m,a,id,b)=>{if(!seen.has(Number(id))){seen.add(Number(id));return a+id+b;}return a+(next++)+b;}));}
  return zipSync(parts);
 })();return fixturePromise;
}
function assertDiverseResult(result){
 assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));assert.equal(result.model.slides.length,15);
 const nodes=result.model.slides.flatMap(s=>s.nodes),warnings=result.diagnostics.filter(d=>d.reasonKey==='PPTX_CHART_IMAGE_FALLBACK');
 assert.equal(nodes.filter(n=>n.kind==='chart').length,5);assert.equal(warnings.length,11);
 for(const id of ['gantt-1904','scatter','bubble','combo','doughnut','radar','area','stacked','percent-stacked','rounded-inventory','date-shortage'])assert.ok(warnings.some(w=>w.objectName===id),id);
 assert.equal(result.model.slides[1].nodes.filter(n=>n.kind==='chart').length,2);
 assert.ok(result.model.slides[2].nodes.some(n=>n.kind==='table'&&n.id==='gantt-table'));
 assert.equal(exportHtml(result.model).status,'generated');
}
test('diverse validated charts preserve native charts and label image substitutions',async()=>{
 const bytes=await fixture(),result=await readPptx(bytes,{allowPartial:true},{renderChartImages:async (_bytes,charts)=>charts.map(chart=>({...chart,bytes:png}))});assertDiverseResult(result);
});
test('installed PowerPoint renders a diverse 15-slide chart deck', {skip:process.platform!=='win32'||process.env.PPT_AGAIN_OFFICE_TEST!=='1'},async()=>{
 const result=await readPptx(await fixture(),{allowPartial:true},{renderChartImages:renderPowerPointChartImages});assertDiverseResult(result);
 for(const slide of result.model.slides)for(const n of slide.nodes.filter(n=>n.kind==='image'))assert.ok(result.model.assets.find(a=>a.id===n.content.assetId)?.bytes.length>1000);
});
