const record=value=>value!==null&&typeof value==='object'&&!Array.isArray(value);
const label=value=>typeof value==='string'&&value.length<=100000&&value.isWellFormed()&&!/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/.test(value);
const format=value=>{const rounded=Number(value.toPrecision(10));return String(Number.isFinite(rounded)?rounded:value);};
const escape=value=>String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const number=value=>{if(!Number.isFinite(value))throw Object.assign(new Error('Invalid chart geometry'),{code:'INPUT_INVALID'});return String(Number(value.toFixed(6)));};

export function validateChart(content){
  const errors=[],check=(ok,path,message)=>{if(!ok)errors.push({path,message});};
  const keys=(value,allowed,path)=>check(Object.keys(value).every(key=>allowed.includes(key)),path,'Unsupported chart option');
  if(!record(content))return {valid:false,errors:[{path:'',message:'Expected chart content'}]};
  keys(content,['type','categories','series','title','legend','axes'],'');
  check(['bar','line','pie'].includes(content.type),'type','Expected bar, line or pie');
  check(label(content.title),'title','Expected chart title');check(typeof content.legend==='boolean','legend','Expected legend visibility');
  const categories=content.categories,series=content.series;
  check(Array.isArray(categories)&&categories.length>0&&categories.length<=1000&&categories.every(c=>label(c)&&c.trim()),'categories','Expected 1 to 1000 category labels');
  check(Array.isArray(series)&&series.length>0&&series.length<=10,'series','Expected 1 to 10 series');
  if(!Array.isArray(categories)||!Array.isArray(series)||categories.length>1000||series.length>10)return {valid:false,errors};
  check(categories.length*series.length<=1000,'series','Chart exceeds 1000 points');
  for(const [i,item] of series.entries()){
    const p=`series[${i}]`;if(!record(item)){check(false,p,'Expected series');continue;}
    keys(item,['name','values','color'],p);check(label(item.name)&&item.name.trim(),`${p}.name`,'Expected series name');
    check(typeof item.color==='string'&&/^#[a-f\d]{6}$/i.test(item.color),`${p}.color`,'Expected an RGB color');
    check(Array.isArray(item.values)&&item.values.length===categories.length&&item.values.every(Number.isFinite),`${p}.values`,'Expected one finite number per category');
  }
  const axes=content.axes;
  if(!record(axes)){check(false,'axes','Expected axes');return {valid:false,errors};}
  keys(axes,['xTitle','yTitle','min','max','majorUnit'],'axes');
  for(const key of ['xTitle','yTitle'])if(key in axes)check(label(axes[key]),`axes.${key}`,'Expected an axis title');
  for(const key of ['min','max','majorUnit'])if(key in axes)check(Number.isFinite(axes[key])&&(key!=='majorUnit'||axes[key]>0),`axes.${key}`,'Expected a finite axis value');
  if(errors.length)return {valid:false,errors};
  const values=series.flatMap(s=>s.values);
  if(content.type==='pie'){
    check(series.length===1&&values.every(v=>v>=0)&&values.some(v=>v>0),'series','Pie needs one nonnegative series with positive total');
    check(Object.keys(axes).length===0,'axes','Pie does not have value axes');
  }else{
    const axis=chartAxis(content);
    check(axis.min<axis.max&&axis.min<=Math.min(...values)&&axis.max>=Math.max(...values),'axes','Axis range must include every value');
    check(axis.ticks.length>0&&axis.ticks.length<=1000,'axes.majorUnit','Axis exceeds 1000 ticks');
  }
  check([content.title,...categories,...series.map(s=>s.name),axes.xTitle??'',axes.yTitle??''].reduce((sum,text)=>sum+text.length,0)<=2000000,'title','Chart exceeds the document text budget');
  return {valid:errors.length===0,errors};
}

// Simple linear, zero-inclusive defaults. The native writer uses these exact
// bounds and major unit too; no independent Office auto-axis inference.
export function chartAxis(content){
  if(content.type==='pie')return null;
  const values=content.series.flatMap(s=>s.values),axes=content.axes;
  let min=axes.min??Math.min(0,...values),max=axes.max??Math.max(0,...values);
  if(min===max&&axes.min===undefined&&axes.max===undefined){min=0;max=1;}
  const scale=Math.max(Math.abs(min),Math.abs(max),Number.MIN_VALUE),span=max/scale-min/scale;
  const majorUnit=axes.majorUnit??Math.max(Number.MIN_VALUE,span/5*scale),count=Math.floor(span/(majorUnit/scale)+1e-10);
  const ticks=Number.isFinite(count)&&count>=0&&count<1000?Array.from({length:count+1},(_,i)=>{
    const value=Math.min(max/scale,Math.max(min/scale,min/scale+i*(majorUnit/scale)))*scale;
    return Number(format(value));
  }):[];
  return {min,max,majorUnit,ticks};
}

// Font collection consumes this text with the chart node's resolved font.
// Data labels are included so enabling the native legend cannot lose glyphs.
export function chartText(content){
  return [content.title,...content.categories,...content.series.flatMap(s=>[s.name,...s.values.map(format)]),content.axes.xTitle??'',content.axes.yTitle??'',...(chartAxis(content)?.ticks.map(format)??[])].filter(Boolean);
}

export function renderChartSvg(content,{width,height,fontFamily='Arial',fontSize=12}={}){
  if(!validateChart(content).valid||![width,height,fontSize].every(v=>Number.isFinite(v)&&v>0)||!label(fontFamily))throw Object.assign(new Error('Invalid chart or display size'),{code:'INPUT_INVALID'});
  const unit=fontSize/12,left=Math.min(70*unit,width*.15),right=width-Math.min(20*unit,width*.05),top=Math.min((content.title?50:25)*unit,height*.2),bottom=height-Math.min((content.legend?75:50)*unit,height*.25),w=right-left,h=bottom-top;
  const parts=[],text=(x,y,value,extra='')=>`<text x="${number(x)}" y="${number(y)}"${extra}>${escape(value)}</text>`;
  if(content.title)parts.push(text(width/2,Math.min(22*unit,height*.09),content.title,' text-anchor="middle"'));
  if(content.type==='pie'){
    const values=content.series[0].values,largest=Math.max(...values),weights=values.map(v=>v/largest),total=weights.reduce((a,b)=>a+b,0),cx=(left+right)/2,cy=(top+bottom)/2,r=Math.min(w,h)*.46;
    let angle=-Math.PI/2;
    values.forEach((value,i)=>{
      const fraction=weights[i]/total,end=angle+fraction*2*Math.PI,mid=(angle+end)/2,color=content.series[0].color;
      if(fraction===1)parts.push(`<circle data-point-index="${i}" cx="${number(cx)}" cy="${number(cy)}" r="${number(r)}" fill="${color}"/>`);
      else parts.push(`<path data-point-index="${i}" d="M${number(cx)} ${number(cy)} L${number(cx+r*Math.cos(angle))} ${number(cy+r*Math.sin(angle))} A${number(r)} ${number(r)} 0 ${fraction>.5?1:0} 1 ${number(cx+r*Math.cos(end))} ${number(cy+r*Math.sin(end))} Z" fill="${color}" stroke="#ffffff" stroke-width="${number(unit)}"/>`);
      if(value>0)parts.push(text(cx+r*.6*Math.cos(mid),cy+r*.6*Math.sin(mid),content.categories[i],' text-anchor="middle" fill="#ffffff"'));
      angle=end;
    });
  }else{
    const axis=chartAxis(content),scale=Math.max(Math.abs(axis.min),Math.abs(axis.max),Number.MIN_VALUE),y=value=>bottom-(value/scale-axis.min/scale)/(axis.max/scale-axis.min/scale)*h;
    for(const tick of axis.ticks){const ty=y(tick);parts.push(`<line x1="${number(left)}" y1="${number(ty)}" x2="${number(right)}" y2="${number(ty)}" stroke="#cbd5e1" stroke-width="${number(unit)}"/>`,text(left-8*unit,ty+4*unit,format(tick),' text-anchor="end"'));}
    parts.push(`<path d="M${number(left)} ${number(top)} V${number(bottom)} H${number(right)}" fill="none" stroke="#475569" stroke-width="${number(unit)}"/>`);
    const step=w/content.categories.length,x=i=>left+step*(i+.5),baseline=y(Math.max(axis.min,Math.min(axis.max,0)));
    content.categories.forEach((category,i)=>parts.push(text(x(i),bottom+18*unit,category,' text-anchor="middle"')));
    content.series.forEach((series,s)=>{
      if(content.type==='bar'){
        const bar=step*.75/content.series.length;
        series.values.forEach((value,i)=>parts.push(`<rect data-series-index="${s}" data-point-index="${i}" x="${number(left+step*i+step*.125+s*bar)}" y="${number(Math.min(baseline,y(value)))}" width="${number(bar)}" height="${number(Math.abs(baseline-y(value)))}" fill="${series.color}"/>`));
      }else{
        parts.push(`<path d="${series.values.map((value,i)=>`${i?'L':'M'}${number(x(i))} ${number(y(value))}`).join(' ')}" fill="none" stroke="${series.color}" stroke-width="${number(2*unit)}"/>`);
        series.values.forEach((value,i)=>parts.push(`<circle data-series-index="${s}" data-point-index="${i}" cx="${number(x(i))}" cy="${number(y(value))}" r="${number(3*unit)}" fill="${series.color}"/>`));
      }
    });
    if(content.axes.xTitle)parts.push(text((left+right)/2,bottom+38*unit,content.axes.xTitle,' text-anchor="middle"'));
    if(content.axes.yTitle)parts.push(text(16*unit,(top+bottom)/2,content.axes.yTitle,` text-anchor="middle" transform="rotate(-90 ${number(16*unit)} ${number((top+bottom)/2)})"`));
  }
  if(content.legend){
    const items=content.type==='pie'?content.categories.map(name=>({name,color:content.series[0].color})):content.series;
    items.forEach((item,i)=>{const x=left+w*i/items.length,y=height-12*unit;parts.push(`<rect x="${number(x)}" y="${number(y-9*unit)}" width="${number(8*unit)}" height="${number(8*unit)}" fill="${item.color}"/>`,text(x+12*unit,y,item.name));});
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${number(width)}" height="${number(height)}" viewBox="0 0 ${number(width)} ${number(height)}" role="img" aria-label="${escape(content.title)}" font-family="${escape(fontFamily)}" font-size="${number(fontSize)}" font-weight="normal" font-style="normal" fill="#223845">${parts.join('')}</svg>`;
}
