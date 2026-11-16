import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {addNativeChart} from '../src/html-to-pptx/write-chart.mjs';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGOQTDkBAAHfAUZ76qvRAAAAAElFTkSuQmCC','base64');
let originalPromise;
async function fixture(){
 originalPromise??=(async()=>{
 const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';const slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('Visible neighboring text',{x:1,y:.2,w:7,h:.5,fontFace:'Arial',fontSize:20,objectName:'safe-title'});
 for(const [id,x] of [['fallback-chart',96],['editable-chart',550]])addNativeChart(slide,{id,kind:'chart',boxCssPx:{x,y:100,w:400,h:300},style:{fontFamily:'Arial',fontSize:16},content:{type:'bar',categories:['A','B'],series:[{name:'Series',values:[10,20],color:'#1976d2'}],title:id,legend:true,axes:{}}},{pointScale:.75,resolveFont:()=>({fontFace:'Arial'})});
 slide.addShape(ppt.ShapeType.rect,{x:1,y:5,w:1,h:.5,fill:{color:'225588'},objectName:'neighbor-shape'});
 slide.addTable([['Header','Value'],['Row','42']],{x:3,y:5,w:4,h:1,fontFace:'Arial',fontSize:16,objectName:'neighbor-table'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
 const key='ppt/charts/chart1.xml',xml=strFromU8(parts[key]),changed=xml.replace(/<c:roundedCorners(?:\s+val="[^"]*")?\s*\/>/,'<c:roundedCorners val="1"/>');assert.notEqual(changed,xml);parts[key]=strToU8(changed);return parts;
 })();return {...await originalPromise};
}
const render=async (_bytes,charts)=>charts.map(chart=>({slideNumber:chart.slideNumber,shapeId:chart.shapeId,bytes:png}));
test('unsupported chart is replaced in place while supported chart and text stay editable',async()=>{
 const parts=await fixture(),result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:render});
 assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const slide=result.model.slides[0],image=slide.nodes.find(n=>n.id==='fallback-chart'),native=slide.nodes.find(n=>n.id==='editable-chart');
 assert.equal(image?.kind,'image');assert.equal(native?.kind,'chart');assert.ok(slide.nodes.some(n=>n.id==='safe-title'));
 assert.equal(slide.nodes.find(n=>n.id==='neighbor-shape')?.kind,'shape');assert.equal(slide.nodes.find(n=>n.id==='neighbor-table')?.kind,'table');
 const warning=result.diagnostics.find(d=>d.reasonKey==='PPTX_CHART_IMAGE_FALLBACK'&&d.objectName==='fallback-chart'&&d.slideNumber===1);assert.ok(warning);
 assert.match(publicDiagnostic(warning).message,/그림.*편집할 수 없습니다/);
 assert.match(exportHtml(result.model).html,/data:image\/png;base64,/);
});
test('supported charts remain editable and never invoke PowerPoint',async()=>{
 const parts=await fixture(),key='ppt/charts/chart1.xml';parts[key]=strToU8(strFromU8(parts[key]).replace('<c:roundedCorners val="1"/>','<c:roundedCorners val="0"/>'));let calls=0;
 const result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:async()=>{calls++;return [];}});
 assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));assert.equal(calls,0);assert.equal(result.model.slides[0].nodes.filter(n=>n.kind==='chart').length,2);
});
test('failed image extraction reports the chart and blocks an apparently successful HTML',async()=>{
 const result=await readPptx(zipSync(await fixture()),{allowPartial:true},{renderChartImages:async()=>{throw Object.assign(new Error('unavailable'),{code:'POWERPOINT_UNAVAILABLE'});}});
 assert.equal(result.status,'failed');assert.ok(result.diagnostics.some(d=>d.reasonKey==='PPTX_CHART_IMAGE_UNAVAILABLE'&&d.slideNumber===1&&d.objectName==='fallback-chart'&&d.blocking));
 const diagnostic=result.diagnostics.find(d=>d.reasonKey==='PPTX_CHART_IMAGE_UNAVAILABLE');assert.match(publicDiagnostic(diagnostic).reason,/PowerPoint/);
});
test('a timeout or invalid image is reported instead of producing HTML',async()=>{
 for(const [renderer,reasonKey] of [[async()=>{throw Object.assign(new Error('timeout'),{code:'POWERPOINT_TIMEOUT'});},'PPTX_CHART_IMAGE_TIMEOUT'],[async (_bytes,charts)=>charts.map(c=>({...c,bytes:Buffer.from('not an image')})),'PPTX_CHART_IMAGE_FAILED']]){
  const result=await readPptx(zipSync(await fixture()),{allowPartial:true},{renderChartImages:renderer});assert.equal(result.status,'failed');
  const d=result.diagnostics.find(item=>item.reasonKey===reasonKey);assert.ok(d&&d.slideNumber===1&&d.objectName==='fallback-chart'&&d.blocking);assert.ok(publicDiagnostic(d).reason);
 }
});
test('PowerPoint open and security failures report specific safe reasons',async()=>{
 for(const [code,reasonKey] of [['POWERPOINT_OPEN_FAILED','PPTX_CHART_IMAGE_OPEN_FAILED'],['POWERPOINT_SECURITY_FAILED','PPTX_CHART_IMAGE_SECURITY_FAILED']]){
  const result=await readPptx(zipSync(await fixture()),{allowPartial:true},{renderChartImages:async()=>{throw Object.assign(new Error('private secret'),{code});}});
  const d=result.diagnostics.find(item=>item.reasonKey===reasonKey);assert.equal(result.status,'failed');assert.ok(d&&d.objectName==='fallback-chart'&&d.slideNumber===1);
  assert.doesNotMatch(JSON.stringify(publicDiagnostic(d)),/private secret/);
 }
});
test('unsupported presentation cannot hide a data mismatch or invoke PowerPoint',async()=>{
 const parts=await fixture(),key='ppt/charts/chart1.xml';parts[key]=strToU8(strFromU8(parts[key]).replace('<c:v>10</c:v>','<c:v>999</c:v>'));let calls=0;
 const result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:async()=>{calls++;return [];}});
 assert.equal(result.status,'failed');assert.equal(calls,0);assert.ok(result.diagnostics.some(d=>d.code==='CHART_DATA_MISMATCH'&&d.objectName==='fallback-chart'));
});
test('unsupported appearance cannot hide malformed chart attributes',async()=>{
 for(const mutate of [x=>x.replace('<c:gapWidth val="33"/>','<c:gapWidth val="999"/>'),x=>x.replace('<c:roundedCorners val="1"/>','<c:roundedCorners val="invalid"/>')]){
  const parts=await fixture(),key='ppt/charts/chart1.xml',before=strFromU8(parts[key]),changed=mutate(before);assert.notEqual(changed,before);parts[key]=strToU8(changed);let calls=0;
  const result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:async()=>{calls++;return [];}});
  assert.equal(result.status,'failed');assert.equal(calls,0);assert.ok(result.diagnostics.some(d=>d.code==='INPUT_INVALID'&&d.objectName==='fallback-chart'));
 }
});
test('an external relationship is blocked before image extraction',async()=>{
 const parts=await fixture(),key='ppt/charts/_rels/chart1.xml.rels',before=strFromU8(parts[key]);parts[key]=strToU8(before.replace('</Relationships>','<Relationship Id="outside" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="https://example.invalid/private.xlsx" TargetMode="External"/></Relationships>'));let calls=0;
 const result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:async()=>{calls++;return [];}});
 assert.equal(result.status,'failed');assert.equal(calls,0);assert.ok(result.diagnostics.some(d=>d.severity==='error'));
});
test('an internal relationship cannot disguise a network or encoded traversal path',async()=>{
 for(const target of ['//example.invalid/share/file.png','%2e%2e/%2e%2e/external.png']){
  const parts=await fixture(),key='_rels/.rels',before=strFromU8(parts[key]);parts[key]=strToU8(before.replace('</Relationships>',`<Relationship Id="unsafe-path" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/></Relationships>`));let calls=0;
  const result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:async()=>{calls++;return [];}});
  assert.equal(result.status,'failed',target);assert.equal(calls,0,target);assert.ok(result.diagnostics.some(d=>d.reasonKey==='PPTX_EXTERNAL_REFERENCE'),target);
 }
});
test('a duplicate PowerPoint object ID reports ambiguity before image extraction',async()=>{
 const parts=await fixture(),key='ppt/slides/slide1.xml',before=strFromU8(parts[key]),id=/<p:cNvPr id="(\d+)" name="fallback-chart"/.exec(before)?.[1];assert.ok(id);
 const changed=before.replace(/(<p:cNvPr id=")\d+(" name="neighbor-shape")/,`$1${id}$2`);assert.notEqual(changed,before);parts[key]=strToU8(changed);let calls=0;
 const result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:async()=>{calls++;return [];}});
 assert.equal(result.status,'failed');assert.equal(calls,0);const d=result.diagnostics.find(d=>d.reasonKey==='PPTX_CHART_IMAGE_ID_AMBIGUOUS'&&d.objectName==='fallback-chart');assert.ok(d);assert.match(publicDiagnostic(d).reason,/중복/);
});
test('1904 date flag is passed to the renderer unchanged',async()=>{
 const parts=await fixture(),key='ppt/charts/chart1.xml',before=strFromU8(parts[key]);parts[key]=strToU8(before.replace(/<c:date1904(?:\s+val="[^"]*")?\s*\/>/,'<c:date1904 val="1"/>'));assert.notEqual(strFromU8(parts[key]),before);
 const bytes=zipSync(parts);let passed;
 const result=await readPptx(bytes,{allowPartial:true},{renderChartImages:async (source,charts)=>{passed=source;return render(source,charts);}});
 assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));assert.deepEqual(passed,bytes);assert.match(strFromU8(unzipSync(passed)[key]),/<c:date1904 val="1"\/>/);
});
test('another chart data failure prevents any Office extraction',async()=>{
 const parts=await fixture(),key='ppt/charts/chart2.xml',before=strFromU8(parts[key]),changed=before.replace('<c:v>10</c:v>','<c:v>999</c:v>');assert.notEqual(changed,before);parts[key]=strToU8(changed);let calls=0;
 const result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:async()=>{calls++;return [];}});
 assert.equal(result.status,'failed');assert.equal(calls,0);assert.ok(result.diagnostics.some(d=>d.code==='CHART_DATA_MISMATCH'&&d.objectName==='editable-chart'));
});
test('one failed chart reports its own location without mislabeling the other',async()=>{
 const parts=await fixture(),key='ppt/charts/chart2.xml',before=strFromU8(parts[key]);parts[key]=strToU8(before.replace(/<c:roundedCorners(?:\s+val="[^"]*")?\s*\/>/,'<c:roundedCorners val="1"/>'));
 const result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:async (_bytes,charts)=>charts.map((chart,i)=>i?{...chart,errorCode:'POWERPOINT_CHART_NOT_FOUND'}:{...chart,bytes:png})});
 assert.equal(result.status,'failed');assert.ok(result.diagnostics.some(d=>d.reasonKey==='PPTX_CHART_IMAGE_FALLBACK'&&d.objectName==='fallback-chart'));
 assert.ok(result.diagnostics.some(d=>d.reasonKey==='PPTX_CHART_IMAGE_NOT_FOUND'&&d.objectName==='editable-chart'));
 assert.ok(!result.diagnostics.some(d=>d.reasonKey==='PPTX_CHART_IMAGE_NOT_FOUND'&&d.objectName==='fallback-chart'));
});
