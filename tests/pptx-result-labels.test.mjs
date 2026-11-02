import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import vm from 'node:vm';
const js=await fs.readFile(new URL('../web/html-to-pptx/app.js',import.meta.url),'utf8');
const html=await fs.readFile(new URL('../web/html-to-pptx/index.html',import.meta.url),'utf8');
const functionSource=js.slice(js.indexOf('function resultMessages(job){'),js.indexOf('function showResult(job){'));
const messages=vm.runInNewContext(`(${functionSource.trim()})`);
const job={sourceType:'pptx',status:'failed',slides:[],exportAvailable:false,partialExportAvailable:false,fileName:'synthetic.pptx',slideCount:0};
test('PPTX settings omit redundant direction headings and retain accessible controls',()=>{
 assert.doesNotMatch(html,/conversion-direction|PPTX → HTML 설정/);
 assert.doesNotMatch(js,/conversion-direction/);
 assert.match(html,/<fieldset id="pptx-options" aria-label="PPTX 가져오기 옵션" hidden>/);
 for(const id of ['pptx-font-policy','pptx-allow-partial'])assert.ok(html.includes(`id="${id}"`));
});
test('failed input with no model-derived slides says no HTML or preview was generated',()=>{
 const result=messages(job);assert.match(result.status,/입력 분석 중단: HTML과 미리보기 생성 안 됨/);assert.match(result.empty,/입력 분석 중단/);
 assert.doesNotMatch(result.empty,/다운로드할 수/);
});
test('available HTML without preview is described as a rendering failure, not input failure',()=>{
 const result=messages({...job,status:'analyzed',slides:[{previewUrl:null}],slideCount:1,exportAvailable:true});
 assert.match(result.status,/미리보기 렌더링만 실패/);assert.match(result.empty,/HTML은 다운로드할 수/);assert.doesNotMatch(result.status,/입력 분석 중단/);
});
test('existing model with blocked conversion is not mislabeled as failed input parsing',()=>{
 const result=messages({...job,slides:[{previewUrl:null}],slideCount:1});
 assert.match(result.status,/변환을 막는 항목/);assert.doesNotMatch(result.status,/입력 분석 중단/);
});
test('partial export remains explicit and actual preview uses normal success wording',()=>{
 const partial=messages({...job,slides:[{previewUrl:null}],partialExportAvailable:true});assert.match(partial.empty,/일부 장표는 다운로드/);
 const success=messages({...job,status:'analyzed',slides:[{previewUrl:'actual.png'}],slideCount:1,exportAvailable:true});assert.match(success.status,/1장을 확인/);assert.doesNotMatch(success.status,/실패|중단/);
});
test('quality note is hidden until a real preview exists and resets on each new input',()=>{
 assert.match(js,/const \$=id=>document.getElementById\(id\);\s*\$\('quality-note'\)\.hidden=true/);
 assert.match(js,/\$\('preview-area'\)\.hidden=!visible;\$\('quality-note'\)\.hidden=!visible/);
 const hiddenPreviews=[...js.matchAll(/\$\('preview-area'\)\.hidden=true;/g)];
 assert.ok(hiddenPreviews.length>=3);
 for(const match of hiddenPreviews)assert.ok(js.slice(match.index+match[0].length).startsWith("$('quality-note').hidden=true;"));
 assert.match(js,/if\(!slide\?\.previewUrl\)return;\s*\$\('quality-note'\)\.hidden=false/);
});
