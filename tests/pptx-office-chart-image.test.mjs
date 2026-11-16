import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {addNativeChart} from '../src/html-to-pptx/write-chart.mjs';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {renderPowerPointChartImages} from '../src/html-to-pptx/powerpoint-chart-images.mjs';

test('installed PowerPoint exports unsupported charts on slides 14 and 15', {skip:process.platform!=='win32'||process.env.PPT_AGAIN_OFFICE_TEST!=='1'},async()=>{
 const ppt=new PptxGenJS();for(let i=0;i<13;i++)ppt.addSlide();const slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('Other content',{x:1,y:.3,w:5,h:.5,fontFace:'Arial',fontSize:20,objectName:'neighbor'});
 addNativeChart(slide,{id:'office-chart',kind:'chart',boxCssPx:{x:100,y:100,w:500,h:300},style:{fontFamily:'Arial',fontSize:16},content:{type:'bar',categories:['A','B'],series:[{name:'Series',values:[10,20],color:'#1976d2'}],title:'Office chart',legend:true,axes:{}}},{pointScale:.75,resolveFont:()=>({fontFace:'Arial'})});
 addNativeChart(slide,{id:'native-chart',kind:'chart',boxCssPx:{x:620,y:100,w:200,h:200},style:{fontFamily:'Arial',fontSize:16},content:{type:'bar',categories:['A','B'],series:[{name:'Series',values:[10,20],color:'#1976d2'}],title:'Native chart',legend:true,axes:{}}},{pointScale:.75,resolveFont:()=>({fontFace:'Arial'})});
 const last=ppt.addSlide();last.background={color:'FFFFFF'};
 addNativeChart(last,{id:'office-date-chart',kind:'chart',boxCssPx:{x:100,y:100,w:500,h:300},style:{fontFamily:'Arial',fontSize:16},content:{type:'bar',categories:['A','B'],series:[{name:'Series',values:[10,20],color:'#1976d2'}],title:'Date chart',legend:true,axes:{}}},{pointScale:.75,resolveFont:()=>({fontFace:'Arial'})});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/charts/chart1.xml',xml=strFromU8(parts[key]);
 const changed=xml.replace(/<c:roundedCorners(?:\s+val="[^"]*")?\s*\/>/,'<c:roundedCorners val="1"/>');assert.notEqual(changed,xml);parts[key]=strToU8(changed);
 const dateKey='ppt/charts/chart3.xml',dateXml=strFromU8(parts[dateKey]),dateChanged=dateXml.replace(/<c:date1904(?:\s+val="[^"]*")?\s*\/>/,'<c:date1904 val="1"/>');assert.notEqual(dateChanged,dateXml);parts[dateKey]=strToU8(dateChanged);
 const result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:renderPowerPointChartImages});
 assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));const chart=result.model.slides[13].nodes.find(n=>n.id==='office-chart');assert.equal(chart?.kind,'image');
 const asset=result.model.assets.find(a=>a.id===chart.content.assetId);assert.equal(asset?.mime,'image/png');assert.ok(asset.bytes.length>1000);
 const png=Buffer.from(asset.bytes),ratio=png.readUInt32BE(16)/png.readUInt32BE(20);assert.ok(Math.abs(ratio-500/300)<.04,`chart image aspect ratio ${ratio} differs from the source`);
 assert.equal(result.model.slides[13].nodes.find(n=>n.id==='native-chart')?.kind,'chart');assert.ok(result.model.slides[13].nodes.some(n=>n.id==='neighbor'));
 assert.equal(result.model.slides[14].nodes.find(n=>n.id==='office-date-chart')?.kind,'image');assert.deepEqual(result.diagnostics.filter(d=>d.reasonKey==='PPTX_CHART_IMAGE_FALLBACK').map(d=>d.slideNumber),[14,15]);
});
