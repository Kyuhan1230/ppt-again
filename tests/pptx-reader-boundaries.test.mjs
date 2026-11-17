import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {addNativeChart} from '../src/html-to-pptx/write-chart.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

const image='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGOQTDkBAAHfAUZ76qvRAAAAAElFTkSuQmCC';
const update=(parts,path,edit)=>{const before=strFromU8(parts[path]),after=edit(before);assert.notEqual(after,before,`fixture update: ${path}`);parts[path]=strToU8(after);};
const shape=(name,color,x)=>`<p:sp><p:nvSpPr><p:cNvPr id="810" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`;
let fixturePromise;
async function mixedFixture(){
  fixturePromise??=(async()=>{
    const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';const slide=ppt.addSlide();slide.background={color:'FFFFFF'};
    slide.addText('Editable title',{x:1,y:1,w:3,h:.6,fontFace:'Arial',fontSize:18,shape:'rect',fill:{color:'CCEEFF'},line:{color:'112233',width:1},isTextBox:false,objectName:'title'});
    slide.addShape(ppt.ShapeType.line,{x:1,y:2,w:2,h:0,line:{color:'112233',width:2},objectName:'connector'});
    slide.addImage({data:image,x:4,y:1,w:1,h:1,objectName:'picture'});
    slide.addShape(ppt.ShapeType.rect,{x:1,y:3,w:1,h:1,fill:{color:'335577'},line:{transparency:100},objectName:'group-child'});
    slide.addTable([['Label','Value'],['A','42']],{x:3,y:3,w:3,h:1,fontFace:'Arial',fontSize:14,objectName:'native-table'});
    addNativeChart(slide,{id:'native-chart',kind:'chart',boxCssPx:{x:700,y:100,w:360,h:250},style:{fontFamily:'Arial',fontSize:16},content:{type:'bar',categories:['A','B'],series:[{name:'Series',values:[10,20],color:'#1976d2'}],title:'Native chart',legend:true,axes:{}}},{pointScale:.75,resolveFont:()=>({fontFace:'Arial'})});
    const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
    update(parts,'ppt/slideMasters/slideMaster1.xml',x=>x.replace('</p:spTree>',shape('master-layer','FF0000',0)+'</p:spTree>'));
    update(parts,'ppt/slideLayouts/slideLayout1.xml',x=>x.replace('</p:spTree>',shape('layout-layer','00FF00',914400)+'</p:spTree>'));
    update(parts,'ppt/slides/slide1.xml',xml=>{
      const child=xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?name="group-child"[\s\S]*?<\/p:sp>/)?.[0];assert.ok(child);
      const group=`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="900" name="group-parent"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="914400" y="2743200"/><a:ext cx="914400" cy="914400"/><a:chOff x="914400" y="2743200"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>${child}</p:grpSp>`;
      return xml.replace(child,group);
    });
    return parts;
  })();
  return {...await fixturePromise};
}

test('one PPT retains inherited layer order and mixed object identity across reader boundaries',async()=>{
  const parts=await mixedFixture(),result=await readPptx(zipSync(parts));
  assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const slide=result.model.slides[0],nodes=slide.nodes,byId=id=>nodes.find(n=>n.id===id);
  const master=byId('master:master-layer'),layout=byId('layout:layout-layer'),title=byId('title');
  assert.ok(master&&layout&&title,JSON.stringify(nodes.map(n=>[n.id,n.kind])));
  assert.ok(master.zOrder<layout.zOrder&&layout.zOrder<title.zOrder);
  assert.equal(master.style.fill.toLowerCase(),'#ff0000');assert.equal(layout.style.fill.toLowerCase(),'#00ff00');
  assert.equal(title.kind,'text');assert.match(title.content.text,/Editable title/);
  assert.equal(title.pptxOrigin?.objectId,'title');assert.equal(title.pptxOrigin?.slideId,slide.id);assert.equal(title.pptxOrigin?.parentId,null);
  assert.equal(byId('connector')?.kind,'line');assert.equal(byId('picture')?.kind,'image');
  assert.ok(result.model.assets.some(a=>a.id===byId('picture').content.assetId));
  assert.equal(byId('native-table')?.kind,'table');assert.match(JSON.stringify(byId('native-table').content),/42/);
  assert.equal(byId('native-chart')?.kind,'chart');assert.deepEqual(byId('native-chart').content.series[0].values,[10,20]);
  assert.equal(byId('group-child')?.parentId,'group-parent');assert.deepEqual(byId('group-child').boxCssPx,{x:96,y:288,w:96,h:96});
  assert.ok(nodes.every((node,i)=>node.zOrder===i),'object ordering remains stable through layer and kind dispatch');
});

test('unused workbook stays fatal and chart image fallback never runs before whole-package validation',async()=>{
  const parts=await mixedFixture();
  update(parts,'ppt/charts/chart1.xml',x=>x.replace(/<c:roundedCorners(?:\s+val="[^"]*")?\s*\/>/,'<c:roundedCorners val="1"/>'));
  parts['ppt/embeddings/unrelated.xlsx']=zipSync({'[Content_Types].xml':strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')});
  let calls=0;const result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:async()=>{calls++;return [];}});
  assert.equal(result.status,'failed');assert.equal(result.model,null);assert.equal(calls,0);
  const fatal=result.diagnostics.find(d=>d.reasonKey==='PPTX_UNUSED_EMBEDDED_PACKAGE');assert.ok(fatal,JSON.stringify(result.diagnostics));
  const shown=publicDiagnostic(fatal);assert.equal(shown.blocking,true);assert.equal(shown.slideNumber,null);assert.ok(!shown.objectName);assert.doesNotMatch(JSON.stringify(shown),/unrelated\.xlsx/);
});

test('a verified unsupported chart is replaced in place after the surrounding objects are read',async()=>{
  const parts=await mixedFixture();
  update(parts,'ppt/charts/chart1.xml',x=>x.replace(/<c:roundedCorners(?:\s+val="[^"]*")?\s*\/>/,'<c:roundedCorners val="1"/>'));
  let calls=0,passed=[];
  const result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:async (_bytes,charts)=>{
    calls++;passed=charts;return charts.map(chart=>({...chart,bytes:Buffer.from(image.split(',')[1],'base64')}));
  }});
  assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));assert.equal(calls,1);assert.equal(passed.length,1);
  const nodes=result.model.slides[0].nodes,chart=nodes.find(n=>n.id==='native-chart');
  assert.equal(chart?.kind,'image');assert.ok(result.model.assets.some(a=>a.id===chart.content.assetId));
  assert.ok(nodes.find(n=>n.id==='title').zOrder<chart.zOrder);
  assert.ok(nodes.find(n=>n.id==='native-table').zOrder<chart.zOrder);
  assert.ok(result.diagnostics.some(d=>d.reasonKey==='PPTX_CHART_IMAGE_FALLBACK'&&d.objectName==='native-chart'));
});

test('unused external relationship is rejected before any chart fallback even in partial mode',async()=>{
  const parts=await mixedFixture();
  update(parts,'ppt/charts/chart1.xml',x=>x.replace(/<c:roundedCorners(?:\s+val="[^"]*")?\s*\/>/,'<c:roundedCorners val="1"/>'));
  update(parts,'ppt/slides/_rels/slide1.xml.rels',x=>x.replace('</Relationships>','<Relationship Id="unused-outside" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.invalid/private.png" TargetMode="External"/></Relationships>'));
  let calls=0;const result=await readPptx(zipSync(parts),{allowPartial:true},{renderChartImages:async()=>{calls++;return [];}});
  assert.equal(result.status,'failed');assert.equal(result.model,null);assert.equal(calls,0);
  const fatal=result.diagnostics.find(d=>d.reasonKey==='PPTX_EXTERNAL_REFERENCE');assert.ok(fatal,JSON.stringify(result.diagnostics));
  assert.equal(publicDiagnostic(fatal).blocking,true);assert.doesNotMatch(JSON.stringify(publicDiagnostic(fatal)),/example\.invalid|private\.png/);
});
