import test from 'node:test';
import assert from 'node:assert/strict';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {assetPath} from '../src/html-to-pptx/assets.mjs';

test('input scripts cannot run and external rendering requests are blocked',async()=>{
  const input=await readInput({html:'<section data-slide style="width:800px;height:600px"><p style="font:24px Arial">Local example</p></section><script>window.inputExecuted=true</script>',assets:[]});
  const runtime=await createBrowserRuntime();let view;
  try{
    view=await runtime.render(input);
    assert.equal(await view.page.evaluate(()=>window.inputExecuted),undefined);
    assert.match(runtime.environment.networkPolicy,/external requests blocked/);
    const blocked=await view.page.evaluate(async()=>{
      try{await fetch('https://network-check.invalid/document');return false;}catch{return true;}
    });
    assert.equal(blocked,true);
  }finally{await view?.context.close();await runtime.close();}
});

test('asset paths cannot escape supplied asset names',()=>{
  for(const path of ['../private.txt','/private.txt','C:/private.txt','https://outside.invalid/a.png'])assert.throws(()=>assetPath(path),undefined,path);
});
