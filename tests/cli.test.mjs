import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import {unzipSync,strFromU8} from 'fflate';
const cli=path.resolve('src/html-to-pptx/cli.mjs');
const run=(...args)=>spawnSync(process.execPath,[cli,...args],{encoding:'utf8',timeout:120000});
const source='<section data-slide style="width:800px;height:600px;position:relative;background:white"><p style="position:absolute;left:30px;top:30px;width:400px;height:60px;margin:0;font:20px Arial">Current text</p></section>';
test('converts both directions and refuses to overwrite',async t=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'slidehandoff-cli-'));t.after(()=>fs.rm(dir,{recursive:true,force:true}));
 const input=path.join(dir,'input.html'),ppt=path.join(dir,'output.pptx'),html=path.join(dir,'output.html');await fs.writeFile(input,source);
 let result=run('convert',input,'--output',ppt,'--json');assert.equal(result.status,0,result.stderr || result.stdout);assert.equal(JSON.parse(result.stdout).status,'generated');
 const bytes=await fs.readFile(ppt);assert.match(strFromU8(unzipSync(bytes)['ppt/slides/slide1.xml']),/Current text/);assert.equal(await fs.readFile(input,'utf8'),source);
 result=run('convert',input,'--output',ppt,'--json');assert.equal(result.status,1);assert.equal(JSON.parse(result.stdout).status,'failed');assert.deepEqual(await fs.readFile(ppt),bytes);
 result=run('convert',ppt,'--output',html,'--json');assert.equal(result.status,0,result.stderr || result.stdout);assert.match(await fs.readFile(html,'utf8'),/Current text/);
});
test('invalid invocation produces JSON and usage exit code',()=>{const result=run('convert','--json');assert.equal(result.status,2);assert.equal(JSON.parse(result.stdout).status,'failed');});
test('asset symlinks are refused without creating output',async t=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'slidehandoff-cli-'));t.after(()=>fs.rm(dir,{recursive:true,force:true}));
 const input=path.join(dir,'input.html'),out=path.join(dir,'out.pptx'),assets=path.join(dir,'assets');await fs.mkdir(assets);await fs.writeFile(input,source);await fs.mkdir(path.join(dir,'outside'));await fs.writeFile(path.join(dir,'outside','outside.css'),'p{}');
 try{await fs.symlink(path.join(dir,'outside'),path.join(assets,'link'),process.platform==='win32'?'junction':'dir');}catch(e){if(e.code==='EPERM'){t.skip('symlink privileges unavailable');return;}throw e;}
 const result=run('convert',input,'--output',out,'--assets',assets,'--json');assert.equal(result.status,1);assert.equal(JSON.parse(result.stdout).status,'failed');await assert.rejects(fs.access(out));
});
test('assets are loaded from the specified directory and missing assets fail',async t=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'slidehandoff-cli-'));t.after(()=>fs.rm(dir,{recursive:true,force:true}));
 const assets=path.join(dir,'assets'),input=path.join(dir,'input.html'),output=path.join(dir,'out.pptx');await fs.mkdir(assets);
 const image=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=','base64');
 await fs.writeFile(path.join(assets,'image.png'),image);await fs.writeFile(input,source.replace('</section>','<img src="image.png" style="position:absolute;left:500px;top:100px;width:50px;height:50px"></section>'));
 let result=run('convert',input,'--output',output,'--assets',assets,'--json');assert.equal(result.status,0,result.stdout||result.stderr);assert.ok(Object.keys(unzipSync(await fs.readFile(output))).some(name=>name.startsWith('ppt/media/')));
 await fs.unlink(output);await fs.unlink(path.join(assets,'image.png'));result=run('convert',input,'--output',output,'--assets',assets,'--json');assert.equal(result.status,1);assert.equal(JSON.parse(result.stdout).status,'failed');await assert.rejects(fs.access(output));
});
test('asset directories with too many entries are rejected before conversion',async t=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'slidehandoff-cli-'));t.after(()=>fs.rm(dir,{recursive:true,force:true}));
 const input=path.join(dir,'input.html'),output=path.join(dir,'out.pptx'),assets=path.join(dir,'assets');await fs.mkdir(assets);await fs.writeFile(input,source);
 for(let i=0;i<1025;i++)await fs.writeFile(path.join(assets,`${i}.txt`),'');
 const result=run('convert',input,'--output',output,'--assets',assets,'--json');assert.equal(result.status,1);assert.equal(JSON.parse(result.stdout).code,'ASSET_LIMIT');await assert.rejects(fs.access(output));
});
