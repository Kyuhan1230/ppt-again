import test from 'node:test';
import assert from 'node:assert/strict';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
import {summarizeJob} from '../src/html-to-pptx/server.mjs';

test('size diagnostics expose safe actionable reasons instead of generic messages',()=>{
  const mismatch=publicDiagnostic({code:'INPUT_INVALID',severity:'error',reasonKey:'SIZE_ORIGINAL_ASPECT_MISMATCH'});
  assert.match(mismatch.message,/출력 크기/);
  const override=publicDiagnostic({code:'UNSUPPORTED_OBJECT',severity:'warning',reasonKey:'SIZE_ORIGINAL_ASPECT_OVERRIDE'});
  assert.match(override.message,/지정한 출력 크기/);
  const target=publicDiagnostic({code:'INPUT_INVALID',severity:'error',reasonKey:'SIZE_TARGET_OUT_OF_RANGE'});
  assert.match(target.message,/1~56인치/);
});

test('job summary retains mismatch and physical scale while sanitizing slide IDs',()=>{
  const policy={mode:'contain',targetWidthEmu:1600*9525,targetHeightEmu:900*9525,marginColor:'#ffffff',source:'custom',originalAspectMismatch:true,slides:[{slideId:'private-id',scale:.75,physicalScale:1.2,offsetX:100,offsetY:0,offsetXEmu:1905000,offsetYEmu:0}]};
  const job={id:'job',fileName:'file.html',input:{profile:'static-slides-v1',source:{sha256:'a'.repeat(64)}},measured:{diagnostics:[],model:{slides:[]},measurements:[],environment:{},sizePolicy:policy}};
  const result=summarizeJob(job);
  assert.equal(result.sizePolicy.originalAspectMismatch,true);
  assert.deepEqual(result.sizePolicy.slides[0],{slideId:'slide-1',scale:.75,physicalScale:1.2,offsetX:100,offsetY:0,offsetXEmu:1905000,offsetYEmu:0});
});
