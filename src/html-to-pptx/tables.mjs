const record=value=>value!==null&&typeof value==='object'&&!Array.isArray(value);
const positive=value=>Number.isFinite(value)&&value>0;

export function validParagraphBullet(bullet){
  if(!record(bullet)||!Number.isInteger(bullet.level)||bullet.level<0||bullet.level>8)return false;
  if(['followTextColor','followTextFont'].some(key=>bullet[key]!==undefined&&typeof bullet[key]!=='boolean'))return false;
  if(bullet.sourceSize!==undefined&&(!record(bullet.sourceSize)||!['buSzTx','buSzPct','buSzPts'].includes(bullet.sourceSize.kind)||bullet.sourceSize.kind!=='buSzTx'&&(!Number.isInteger(bullet.sourceSize.value)||bullet.sourceSize.value<=0||bullet.sourceSize.value>400000)))return false;
  if(bullet.type==='char')return typeof bullet.char==='string'&&[...bullet.char].length===1&&!/[\s\u0000-\u001f]/u.test(bullet.char);
  return bullet.type==='number'&&bullet.format==='arabicPeriod'&&Number.isInteger(bullet.start)&&bullet.start>=1&&bullet.start<=32767&&(bullet.restart===undefined||typeof bullet.restart==='boolean');
}

export function validSourceSpacing(spacing){
  return record(spacing)&&Object.entries(spacing).every(([name,value])=>['lnSpc','spcBef','spcAft'].includes(name)&&record(value)&&(name==='lnSpc'&&value.kind==='default'||['spcPct','spcPts'].includes(value.kind)&&Number.isInteger(value.value)&&value.value>=0&&value.value<=(value.kind==='spcPct'?13200000:158400)));
}

export function validSourceEnd(end){
  if(!record(end)||!['implicit','explicit'].includes(end.mode))return false;
  if(end.mode==='implicit')return end.attrs===undefined;
  if(!record(end.attrs))return false;
  return Object.entries(end.attrs).every(([key,value])=>{
    if(['b','i'].includes(key))return ['0','1','true','false'].includes(value);
    if(!['spc','sz','kern'].includes(key)||typeof value!=='string'||!/^[-+]?\d+$/.test(value))return false;
    const n=Number(value);return Number.isInteger(n)&&Math.abs(n)<=400000&&(key==='spc'||n>=0)&&(key!=='sz'||n>0);
  });
}

export function validateTable(content){
  const errors=[],check=(ok,path,message)=>{if(!ok)errors.push({path,message});};
  if(!record(content))return {valid:false,errors:[{path:'',message:'Expected table content'}]};
  const {rows,columns,cells,rowHeightsCssPx,columnWidthsCssPx}=content;
  check(Number.isInteger(rows)&&rows>0&&rows<=100,'rows','Expected 1 to 100 rows');
  check(Number.isInteger(columns)&&columns>0&&columns<=50,'columns','Expected 1 to 50 columns');
  check(rows*columns<=2000,'cells','Table exceeds 2000 cells');
  if(errors.length)return {valid:false,errors};
  for(const [key,values,count] of [['rowHeightsCssPx',rowHeightsCssPx,rows],['columnWidthsCssPx',columnWidthsCssPx,columns]])check(Array.isArray(values)&&values.length===count&&values.every(positive),key,'Expected positive sizes for every grid row or column');
  check(Array.isArray(cells)&&cells.length>0&&cells.length<=rows*columns,'cells','Expected start cells within the grid');
  if(!Array.isArray(cells)||cells.length>2000)return {valid:false,errors};
  const occupied=new Set();
  for(const [i,cell] of cells.entries()){
    const p=`cells[${i}]`;if(!record(cell)){check(false,p,'Expected cell');continue;}
    const rowValid=Number.isInteger(cell.row)&&cell.row>=0&&cell.row<rows,colValid=Number.isInteger(cell.col)&&cell.col>=0&&cell.col<columns;
    check(rowValid,`${p}.row`,'Cell row is outside the grid');check(colValid,`${p}.col`,'Cell column is outside the grid');
    const spanValid=Number.isInteger(cell.rowSpan)&&cell.rowSpan>0&&cell.row+cell.rowSpan<=rows&&Number.isInteger(cell.colSpan)&&cell.colSpan>0&&cell.col+cell.colSpan<=columns;
    check(spanValid,p,'Cell span is outside the grid');
    if(rowValid&&colValid&&spanValid)for(let row=cell.row;row<cell.row+cell.rowSpan;row++)for(let col=cell.col;col<cell.col+cell.colSpan;col++){
      const key=row*columns+col;check(!occupied.has(key),p,'Cell spans overlap');occupied.add(key);
    }
    check(typeof cell.text==='string',`${p}.text`,'Expected current cell text');
    const runs=cell.runs;check(Array.isArray(runs)&&runs.every(run=>record(run)&&typeof run.text==='string'&&record(run.style)),`${p}.runs`,'Expected text runs');
    if(Array.isArray(runs))check(runs.map(run=>run?.text??'').join('')===cell.text,`${p}.runs`,'Text runs do not match current cell text');
    if(cell.paragraphs!==undefined){
      const paragraphs=cell.paragraphs;
      check(Array.isArray(paragraphs)&&paragraphs.length>0&&paragraphs.length<=10000,`${p}.paragraphs`,'Expected text paragraphs');
      if(Array.isArray(paragraphs)){
        for(const [index,para] of paragraphs.entries()){
          const pp=`${p}.paragraphs[${index}]`;
          if(!record(para)){check(false,pp,'Expected paragraph');continue;}
          for(const key of ['beforePx','afterPx'])check(Number.isFinite(para[key])&&para[key]>=0&&para[key]<=20000,`${pp}.${key}`,'Invalid paragraph spacing');
          for(const key of ['leftPx','rightPx'])check(Number.isFinite(para[key])&&para[key]>=0&&para[key]<=5000,`${pp}.${key}`,'Invalid paragraph margin');
          check(Number.isFinite(para.indentPx)&&Math.abs(para.indentPx)<=5000,`${pp}.indentPx`,'Invalid paragraph indent');
          check(typeof para.lineHeight==='string'&&/^\d+(?:\.\d+)?(?:px)?$/.test(para.lineHeight)&&Number.parseFloat(para.lineHeight)>0&&Number.parseFloat(para.lineHeight)<=20000,`${pp}.lineHeight`,'Invalid paragraph line height');
          check(['left','center','right','justify'].includes(para.align),`${pp}.align`,'Invalid paragraph alignment');
          if(para.hangingPunct!==undefined)check(typeof para.hangingPunct==='boolean',`${pp}.hangingPunct`,'Invalid hanging punctuation');
          check(Array.isArray(para.runs)&&para.runs.length>0&&para.runs.every(r=>record(r)&&typeof r.text==='string'&&record(r.style)),`${pp}.runs`,'Invalid paragraph text');
          if(Array.isArray(para.runs))check(para.runs.every((r,i)=>r?.marker===undefined||r.marker===true&&i===0),`${pp}.runs`,'Invalid list marker');
          if(para.bullet!==undefined)check(validParagraphBullet(para.bullet),`${pp}.bullet`,'Invalid paragraph bullet');
          if(para.sourceSpacing!==undefined)check(validSourceSpacing(para.sourceSpacing),`${pp}.sourceSpacing`,'Invalid source spacing');
          if(para.sourceEnd!==undefined)check(validSourceEnd(para.sourceEnd),`${pp}.sourceEnd`,'Invalid paragraph end formatting');
        }
        if(paragraphs.every(para=>record(para)&&Array.isArray(para.runs)))check(paragraphs.map(para=>para.runs.map(r=>r?.text??'').join('')).join('\n')===cell.text,`${p}.paragraphs`,'Paragraphs do not match text');
      }
    }
    const style=cell.style;if(!record(style)){check(false,`${p}.style`,'Expected cell style');continue;}
    check(typeof style.fontFamily==='string'&&style.fontFamily.trim()&&positive(style.fontSize),`${p}.style`,'Expected cell font and size');
    check(['left','center','right','justify'].includes(style.align)&&['top','mid','bottom'].includes(style.valign),`${p}.style`,'Unsupported cell alignment');
    for(const side of ['top','right','bottom','left']){
      check(Number.isFinite(style.padding?.[side])&&style.padding[side]>=0,`${p}.style.padding.${side}`,'Expected nonnegative padding');
      const border=style.borders?.[side];check(record(border)&&Number.isFinite(border.width)&&border.width>=0&&['none','solid'].includes(border.style)&&typeof border.color==='string',`${p}.style.borders.${side}`,'Expected plain cell border');
    }
  }
  check(occupied.size===rows*columns,'cells','Every grid position must be covered exactly once');
  return {valid:errors.length===0,errors};
}
