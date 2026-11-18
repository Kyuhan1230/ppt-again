import test from 'node:test';
import assert from 'node:assert/strict';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {captureDocumentTiles} from '../src/html-to-pptx/html-pagination-tiles.mjs';

const viewport={width:800,height:600};
const html='<!doctype html><style>html,body{margin:0;width:1200px;height:9300px;background:white}#band{position:absolute;left:0;top:9000px;width:1200px;height:60px;background:linear-gradient(to right,red 0 50%,green 50% 100%)}@media(min-height:700px){#band{background:blue}}</style><div id="band"></div>';
const state=page=>page.evaluate(()=>({viewport:{width:innerWidth,height:innerHeight,dpr:devicePixelRatio,scrollX,scrollY},documentSize:{width:document.documentElement.scrollWidth,height:document.documentElement.scrollHeight},band:getComputedStyle(document.querySelector('#band')).backgroundImage}));

test('tiles preserve a fixed responsive viewport, overlap and fractional final document bounds beyond 9000px',async()=>{
  const runtime=await createBrowserRuntime({capture:false});
  try{
    const session=await runtime.render(await readInput({html}),{viewport});
    try{
      const before=await state(session.page),rect={x:0,y:0,width:1200,height:9200.25};
      const result=await captureDocumentTiles({page:session.page,rect,viewport,readState:()=>state(session.page)});
      assert.deepEqual(await state(session.page),before);
      assert.deepEqual(result.bounds,{x:0,y:0,width:1200,height:9201});assert.match(result.fingerprint,/^[a-f0-9]{64}$/);
      assert.ok(result.tiles.length>1);
      let ownedY=result.bounds.y;
      for(const tile of result.tiles){
        assert.equal(tile.owned.y,ownedY);assert.equal(tile.owned.x,0);assert.equal(tile.owned.width,1200);assert.ok(tile.owned.height>0&&tile.owned.height<=1024);
        assert.match(tile.sha256,/^[a-f0-9]{64}$/);ownedY+=tile.owned.height;
      }
      assert.equal(ownedY,9201);
      for(let i=1;i<result.tiles.length;i++)assert.equal(result.tiles[i-1].y+result.tiles[i-1].height-result.tiles[i].y,2);
      const samples=[];
      for(const clip of [{x:0,y:8990,width:800,height:100},{x:400,y:8990,width:800,height:100}]){
        const png=await session.page.screenshot({fullPage:true,clip,scale:'css'});
        samples.push(await session.page.evaluate(async data=>{const image=new Image();image.src='data:image/png;base64,'+data;await image.decode();const canvas=document.createElement('canvas');canvas.width=image.width;canvas.height=image.height;const context=canvas.getContext('2d');context.drawImage(image,0,0);return {width:image.width,height:image.height,pixels:[[100,20],[300,20],[500,20],[700,20],[100,90]].map(([x,y])=>[...context.getImageData(x,y,1,1).data])};},png.toString('base64')));
      }
      assert.deepEqual(samples.map(s=>[s.width,s.height]),[[800,100],[800,100]]);
      assert.deepEqual(samples[0].pixels[2],samples[1].pixels[0]);assert.deepEqual(samples[0].pixels[3],samples[1].pixels[1]);
      assert.deepEqual(samples[0].pixels[2],[255,0,0,255]);assert.deepEqual(samples[0].pixels[3],[0,128,0,255]);
      assert.deepEqual(samples[0].pixels[4],[255,255,255,255]);assert.deepEqual(await state(session.page),before);
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('missing or forged PNG tiles and changed observed state fail without a completed result',async()=>{
  const runtime=await createBrowserRuntime({capture:false});
  try{
    const session=await runtime.render(await readInput({html}),{viewport});
    try{
      const capture=session.page.screenshot.bind(session.page),rect={x:0,y:0,width:800,height:1500};
      let calls=0;session.page.screenshot=async options=>{if(++calls===2)return Buffer.alloc(0);return capture(options);};
      await assert.rejects(()=>captureDocumentTiles({page:session.page,rect,viewport,readState:()=>state(session.page)}),{code:'PAGINATION_UNSUPPORTED'});
      calls=0;session.page.screenshot=async options=>{
        const bytes=await capture(options);if(++calls!==2)return bytes;
        const encoded=await session.page.evaluate(async data=>{const image=new Image();image.src='data:image/png;base64,'+data;await image.decode();const canvas=document.createElement('canvas');canvas.width=image.width;canvas.height=image.height;const context=canvas.getContext('2d');context.drawImage(image,0,0);context.fillStyle='black';context.fillRect(0,0,2,2);return canvas.toDataURL('image/png').split(',')[1];},bytes.toString('base64'));
        return Buffer.from(encoded,'base64');
      };
      await assert.rejects(()=>captureDocumentTiles({page:session.page,rect,viewport,readState:()=>state(session.page)}),{code:'PAGINATION_UNSUPPORTED'});
      session.page.screenshot=capture;let reads=0;
      await assert.rejects(()=>captureDocumentTiles({page:session.page,rect,viewport,readState:async()=>({...await state(session.page),changed:++reads>2})}),{code:'PAGINATION_UNSUPPORTED'});
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('oversized tile requests fail before taking screenshots',async()=>{
  const runtime=await createBrowserRuntime({capture:false});
  try{
    const session=await runtime.render(await readInput({html}),{viewport});
    try{
      let captures=0;const original=session.page.screenshot.bind(session.page);session.page.screenshot=async options=>{captures++;return original(options);};
      for(const rect of [{x:0,y:0,width:8193,height:10},{x:0,y:0,width:1200,height:32769},{x:0,y:0,width:8192,height:8192}])
        await assert.rejects(()=>captureDocumentTiles({page:session.page,rect,viewport,readState:()=>state(session.page)}),{code:'PAGINATION_UNSUPPORTED'});
      assert.equal(captures,0);
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});
