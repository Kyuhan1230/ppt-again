import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import {unzipSync,strFromU8} from 'fflate';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';

const cli=path.resolve('src/html-to-pptx/cli.mjs');
const run=(...args)=>spawnSync(process.execPath,[cli,...args],{encoding:'utf8',timeout:180000});
const html='<section style="width:400px;height:300px;background:white"><p>Selected slide</p></section><p>Outside</p>';
async function fixture(t){
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'ppt-region-cli-'));t.after(()=>fs.rm(dir,{recursive:true,force:true}));
 const input=path.join(dir,'input.html'),out=path.join(dir,'out.pptx'),inspection=path.join(dir,'regions.json');
 await fs.writeFile(input,html);return {dir,input,out,inspection};
}

test('inspect reports candidates in JSON and never overwrites its output',async t=>{
 const {input,inspection}=await fixture(t);
 let r=run('inspect',input,'--viewport','1280x720','--output',inspection,'--json');
 assert.equal(r.status,0,r.stderr||r.stdout);
 const data=JSON.parse(await fs.readFile(inspection,'utf8'));
 assert.ok(data.snapshotId);assert.ok(data.choices.length>0);assert.doesNotMatch(JSON.stringify(data),/Selected slide|Outside/);
 const before=await fs.readFile(inspection);
 r=run('inspect',input,'--output',inspection,'--json');assert.equal(r.status,1);assert.equal(JSON.parse(r.stdout).code,'OUTPUT_EXISTS');assert.deepEqual(await fs.readFile(inspection),before);
});

test('inspect text explains candidate evidence and reports region positions',async t=>{
 const {input}=await fixture(t);const r=run('inspect',input);
 assert.equal(r.status,0,r.stdout||r.stderr);assert.match(r.stdout,/장표 후보/);assert.match(r.stdout,/범위 밖 내용/);assert.match(r.stdout,/화면 안/);
 assert.doesNotMatch(r.stdout,/Selected slide|Outside/);
});

test('option conflicts and malformed viewport are usage errors',async t=>{
 const {input,out,inspection}=await fixture(t);
 for(const args of [
  ['inspect',input,'--viewport','0x720'],
  ['convert',input,'--output',out,'--region','c1'],
  ['convert',input,'--output',out,'--inspection',inspection],
  ['convert',input,'--output',out,'--scope','document'],
  ['convert',input,'--output',out,'--scope','document','--page-mode','single','--selector','section'],
  ['convert',input,'--output',out,'--selector','section','--inspection',inspection,'--region','c1']
 ]){const r=run(...args,'--json');assert.equal(r.status,2,`${args.join(' ')}: ${r.stdout||r.stderr}`);assert.equal(JSON.parse(r.stdout).code,'USAGE');}
});

test('unmarked conversion reports selection and safe candidate list without output',async t=>{
 const {input,out}=await fixture(t);const r=run('convert',input,'--output',out,'--json');
 assert.equal(r.status,1,r.stdout||r.stderr);const data=JSON.parse(r.stdout);assert.equal(data.code,'SELECTION_REQUIRED');
 assert.ok(data.inspection?.choices?.length);assert.doesNotMatch(JSON.stringify(data.inspection),/Selected slide|Outside/);await assert.rejects(fs.access(out));
});

test('saved selection creates PPT and stale selections never create output',async t=>{
 const {input,out,inspection}=await fixture(t);let r=run('inspect',input,'--viewport','1280x720','--output',inspection,'--json');assert.equal(r.status,0,r.stdout||r.stderr);
 const list=JSON.parse(await fs.readFile(inspection,'utf8'));const choice=list.choices.find(c=>c.scope==='elements'&&c.regionIds.length===1&&c.convertible);
 assert.ok(choice,'a supported single candidate should be offered');
 // Saved coordinates, selector and member lists are not trusted to choose a
 // different part of the source: only the snapshot and selected ID are used.
 choice.selector='body';choice.regionIds=['forged'];choice.bounds={x:0,y:0,width:1,height:1};await fs.writeFile(inspection,JSON.stringify(list));
 r=run('convert',input,'--output',out,'--viewport','1280x720','--inspection',inspection,'--region',choice.id,'--json');assert.equal(r.status,0,r.stdout||r.stderr);
 const xml=strFromU8(unzipSync(await fs.readFile(out))['ppt/slides/slide1.xml']);assert.match(xml,/Selected slide/);assert.doesNotMatch(xml,/>Outside</);
 await fs.unlink(out);r=run('convert',input,'--output',out,'--viewport','1200x720','--inspection',inspection,'--region',choice.id,'--json');assert.equal(r.status,1);assert.equal(JSON.parse(r.stdout).code,'STALE_INSPECTION');await assert.rejects(fs.access(out));
 await fs.writeFile(input,html.replace('Selected slide','Changed slide'));
 r=run('convert',input,'--output',out,'--viewport','1280x720','--inspection',inspection,'--region',choice.id,'--json');assert.equal(r.status,1);assert.equal(JSON.parse(r.stdout).code,'STALE_INSPECTION');await assert.rejects(fs.access(out));
});

test('explicit single document scope refuses unsupported size instead of creating output',async t=>{
 const {input,out}=await fixture(t);const r=run('convert',input,'--output',out,'--scope','document','--page-mode','single','--json');
 if(r.status===0){const xml=strFromU8(unzipSync(await fs.readFile(out))['ppt/slides/slide1.xml']);assert.match(xml,/Selected slide/);assert.match(xml,/Outside/);}
 else {assert.equal(r.status,1,r.stdout||r.stderr);assert.ok(JSON.parse(r.stdout).code);await assert.rejects(fs.access(out));}
});

test('explicit single document scope converts a fixed-size whole document',async t=>{
 const {input,out}=await fixture(t);
 await fs.writeFile(input,'<!doctype html><html><body style="margin:0;width:1280px;height:960px"><p style="margin:0">Document content</p></body></html>');
 const r=run('convert',input,'--output',out,'--scope','document','--page-mode','single','--json');
 assert.equal(r.status,0,r.stdout||r.stderr);
 const xml=strFromU8(unzipSync(await fs.readFile(out))['ppt/slides/slide1.xml']);assert.match(xml,/Document content/);
});

test('a supported repeated group creates separate PPT slides in order',async t=>{
 const {input,out,inspection}=await fixture(t);
 await fs.writeFile(input,'<main><section style="width:400px;height:300px"><p>First slide</p></section><section style="width:400px;height:300px"><p>Second slide</p></section></main>');
 let r=run('inspect',input,'--output',inspection,'--json');assert.equal(r.status,0,r.stdout||r.stderr);
 const list=JSON.parse(await fs.readFile(inspection,'utf8'));
 const group=list.choices.find(choice=>choice.regionIds.length===2&&choice.evidence.includes('repeated-sibling'));
 assert.ok(group?.convertible,'two same-sized nonoverlapping candidates should be convertible');
 r=run('convert',input,'--output',out,'--inspection',inspection,'--region',group.id,'--json');assert.equal(r.status,0,r.stdout||r.stderr);
 const zip=unzipSync(await fs.readFile(out));assert.match(strFromU8(zip['ppt/slides/slide1.xml']),/First slide/);assert.match(strFromU8(zip['ppt/slides/slide2.xml']),/Second slide/);
});

test('UI and CLI inspect the same image asset and produce the same selected PPT content',async t=>{
 const {dir,input,out,inspection}=await fixture(t),assetsDir=path.join(dir,'assets');await fs.mkdir(assetsDir);
 const image=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=','base64');
 const source='<section style="position:relative;width:400px;height:300px;background:white"><p>Shared slide</p><img src="image.png" style="width:20px;height:20px"></section><p>Excluded text</p>';
 await fs.writeFile(input,source);await fs.writeFile(path.join(assetsDir,'image.png'),image);
 let r=run('inspect',input,'--assets',assetsDir,'--viewport','1280x720','--output',inspection,'--json');assert.equal(r.status,0,r.stdout||r.stderr);
 const cliInspection=JSON.parse(await fs.readFile(inspection,'utf8'));
 const app=await createHtmlServer({port:0});t.after(()=>app.close());
 const session=await fetch(app.url+'/api/session'),cookie=session.headers.get('set-cookie').split(';')[0],csrf=(await session.json()).csrfToken;
 const post=async(route,body)=>fetch(app.url+route,{method:'POST',headers:{cookie,'content-type':'application/json','x-csrf-token':csrf},body:JSON.stringify(body)});
 let response=await post('/api/inspect',{html:source,fileName:'input.html',assets:[{path:'image.png',mime:'image/png',base64:image.toString('base64')}],viewport:{width:1280,height:720}});
 let job=await response.json();assert.equal(response.status,200,JSON.stringify(job));
 assert.equal(job.inspection.snapshotId,cliInspection.snapshotId);
 const choice=cliInspection.choices.find(c=>c.convertible&&c.regionIds.length===1&&cliInspection.regions.some(region=>region.id===c.regionIds[0]&&region.bounds?.width===400&&region.bounds?.height===300));
 assert.ok(choice,'same fixed-size slide candidate should be convertible');assert.ok(job.inspection.choices.some(c=>c.id===choice.id));
 response=await post(`/api/jobs/${job.id}/select`,{snapshotId:job.inspection.snapshotId,choiceId:choice.id});job=await response.json();assert.equal(response.status,200,JSON.stringify(job));
 response=await post('/api/export',{jobId:job.id});job=await response.json();assert.equal(response.status,200,JSON.stringify(job));
 response=await fetch(app.url+job.downloadUrl,{headers:{cookie}});assert.equal(response.status,200);const uiZip=unzipSync(new Uint8Array(await response.arrayBuffer()));
 r=run('convert',input,'--assets',assetsDir,'--viewport','1280x720','--inspection',inspection,'--region',choice.id,'--output',out,'--json');assert.equal(r.status,0,r.stdout||r.stderr);
 const cliZip=unzipSync(await fs.readFile(out)),uiSlide=strFromU8(uiZip['ppt/slides/slide1.xml']),cliSlide=strFromU8(cliZip['ppt/slides/slide1.xml']);
 for(const slide of [uiSlide,cliSlide]){assert.match(slide,/Shared slide/);assert.doesNotMatch(slide,/Excluded text/);assert.match(slide,/<p:pic>/);}
 assert.equal(uiSlide,cliSlide);
});


test('inspect text explains unsupported composition and unverified screen in Korean',async t=>{
 const {input}=await fixture(t);
 await fs.writeFile(input,'<main style="position:relative;width:400px;height:300px;background:red"><section style="width:400px;height:300px">UNDER</section><aside style="position:absolute;left:0;top:0;width:400px;height:300px;background:red;z-index:2"></aside></main>');
 let r=run('inspect',input);assert.equal(r.status,0,r.stdout||r.stderr);
 assert.match(r.stdout,/부모 배경이나 효과를 보존할 수 없음/);assert.match(r.stdout,/겹친 외부 요소나 효과를 보존할 수 없음/);
 assert.doesNotMatch(r.stdout,/ancestor-appearance-not-supported|external-overlay-not-supported/);
 await fs.writeFile(input,'<style>@keyframes move{from{opacity:1}to{opacity:.5}}section{animation:move 2s infinite}</style><section style="width:400px;height:300px">Animated</section>');
 r=run('inspect',input);assert.equal(r.status,0,r.stdout||r.stderr);assert.match(r.stdout,/화면을 충분히 확인하지 못함/);assert.doesNotMatch(r.stdout,/inspection-incomplete/);
});
