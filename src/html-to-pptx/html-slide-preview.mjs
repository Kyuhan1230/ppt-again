// Capture an oversized deck through the original CSS viewport so that screen
// media queries, viewport units, and the measured PPT geometry still agree.
export function createDeckPreviewCapture({page,context}){
  const viewport=page.viewportSize();
  let imagePage;
  return async root=>{
    const bounds=await root.boundingBox();
    if(!bounds)return null;
    if(bounds.x>=0&&bounds.y>=0&&bounds.x+bounds.width<=viewport.width&&bounds.y+bounds.height<=viewport.height){
      return root.screenshot({animations:'disabled',scale:'css'});
    }
    const width=Math.round(bounds.width),height=Math.round(bounds.height);
    if(width<1||height<1||width>8192||height>8192||width*height>8_000_000||
       Math.ceil(width/viewport.width)*Math.ceil(height/viewport.height)>64)return null;

    const capture=await root.evaluateHandle(el=>{
      const canvas=el.parentElement;
      if(scrollX||scrollY||getComputedStyle(canvas).translate!=='none')return null;
      const box=el.getBoundingClientRect(),elements=[el,...el.querySelectorAll('*')];
      const layout=()=>{const origin=el.getBoundingClientRect();return elements.map(element=>{
        const r=element.getBoundingClientRect();
        return r.width||r.height?[r.x-origin.x,r.y-origin.y,r.width,r.height]:null;
      });};
      return {el,canvas,layout,before:layout(),x:Math.floor(box.x),y:Math.floor(box.y),pan:null};
    });
    try{
      const tiles=[];
      for(let y=0;y<height;y+=viewport.height)for(let x=0;x<width;x+=viewport.width){
        const tile={x,y,width:Math.min(viewport.width,width-x),height:Math.min(viewport.height,height-y)};
        const stable=await capture.evaluate((state,{x,y,viewport})=>{
          if(!state||innerWidth!==viewport.width||innerHeight!==viewport.height||scrollX||scrollY)return false;
          const translate=`${-state.x-x}px ${-state.y-y}px`;
          state.pan?.cancel();
          // Use a paused animation to move the canvas without changing its
          // author style attributes or CSS selector matches.
          state.pan=state.canvas.animate([{translate},{translate}],{duration:1,fill:'both'});
          state.pan.pause();
          state.pan.currentTime=0;
          const r=state.el.getBoundingClientRect(),after=state.layout();
          return r.x+x>=0&&r.x+x<1&&r.y+y>=0&&r.y+y<1&&state.before.every((box,i)=>{
            const other=after[i];
            return box&&other?box.every((value,j)=>Math.abs(value-other[j])<.05):box===other;
          });
        },{x,y,viewport});
        if(!stable)return null;
        const bytes=await page.screenshot({clip:{x:0,y:0,width:tile.width,height:tile.height},scale:'css'});
        tiles.push({...tile,png:bytes.toString('base64')});
      }
      // Only raster tiles enter this blank page. It has no input document or
      // input scripts. Its lifetime is bounded by the owning render context.
      imagePage??=await context.newPage();
      const png=await imagePage.evaluate(async({width,height,tiles})=>{
        const canvas=document.createElement('canvas');canvas.width=width;canvas.height=height;
        const ctx=canvas.getContext('2d');
        for(const tile of tiles){
          const image=new Image();image.src='data:image/png;base64,'+tile.png;await image.decode();
          if(image.width!==tile.width||image.height!==tile.height)return null;
          ctx.drawImage(image,tile.x,tile.y);
        }
        return canvas.toDataURL('image/png').slice('data:image/png;base64,'.length);
      },{width,height,tiles});
      return png?Buffer.from(png,'base64'):null;
    }catch{
      return null;
    }finally{
      try{await capture.evaluate(state=>state?.pan?.cancel());}
      finally{await capture.dispose();}
    }
  };
}
