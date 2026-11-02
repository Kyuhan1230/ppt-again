import test from 'node:test';
import assert from 'node:assert/strict';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
test('PPTX errors retain explicit slide and safe object context without a model',()=>{
 const d=publicDiagnostic({code:'UNSUPPORTED_OBJECT',severity:'error',slideId:'slide-3',slideNumber:3,objectName:'아이콘 7',objectType:'shape',reasonKey:'PPTX_SHAPE_EFFECT_UNSUPPORTED',message:'/private/source-secret',sourceRef:{path:'private'},nodeId:'secret'});
 assert.equal(d.slideNumber,3);assert.equal(d.objectName,'아이콘 7');assert.equal(d.objectType,'shape');assert.match(d.reason,/그라데이션/);assert.doesNotMatch(JSON.stringify(d),/private|secret/);assert.equal(d.sourceRef,undefined);assert.equal(d.nodeId,undefined);
});
test('selected font diagnostic states replacement without promising font installation',()=>{
 const d=publicDiagnostic({code:'PPTX_FONT_REPLACED',severity:'warning',reasonKey:'PPTX_FONT_SELECTED',fontFamily:'Malgun Gothic',slideNumber:1});
 assert.equal(d.code,'PPTX_FONT_REPLACED');assert.equal(d.fontFamily,'Malgun Gothic');assert.match(d.message,/파일은 포함하지 않/);assert.match(d.message,/달라질/);
});
test('public diagnostics redact paths, markup, raw exceptions and unknown reasons',()=>{
 for(const label of ['/private/font.ttf','C:\\secret\\file.ttf','https://example.test/private','<img src=x>','../private']){
  const d=publicDiagnostic({code:'PPTX_FONT_REPLACED',severity:'warning',reasonKey:'PPTX_FONT_SELECTED',objectName:label,fontFamily:label,objectType:'private-path',slideNumber:300});
  assert.equal(d.objectName,undefined);assert.equal(d.fontFamily,undefined);assert.equal(d.objectType,undefined);assert.equal(d.slideNumber,null);
 }
 const d=publicDiagnostic({code:'UNSUPPORTED_OBJECT',reasonKey:'SECRET_RAW',objectName:'secret',message:'raw stack secret'});
 assert.equal(d.objectName,undefined);assert.equal(d.reason,undefined);assert.doesNotMatch(JSON.stringify(d),/secret/i);
});
test('HTML-origin diagnostics do not newly disclose selectors or user text',()=>{
 const d=publicDiagnostic({code:'ASSET_MISSING',message:'sensitive-value.png',nodeId:'sensitive-value',sourceRef:{path:'sensitive-value'}});
 assert.doesNotMatch(JSON.stringify(d),/sensitive-value/);assert.equal(d.slideNumber,null);
});
test('placeholder and preview-unavailable warnings have accurate public wording',()=>{
 const placeholder=publicDiagnostic({code:'PPTX_OBJECT_PLACEHOLDER',severity:'warning',reasonKey:'PPTX_OBJECT_UNSUPPORTED',slideNumber:2});assert.match(placeholder.message,/표시 상자/);assert.equal(placeholder.slideNumber,2);
 const preview=publicDiagnostic({code:'PPTX_PREVIEW_UNAVAILABLE',severity:'warning',reasonKey:'PREVIEW_UNAVAILABLE'});assert.match(preview.message,/HTML/);assert.equal(preview.severity,'warning');
});
