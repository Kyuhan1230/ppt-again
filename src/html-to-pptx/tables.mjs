const record=value=>value!==null&&typeof value==='object'&&!Array.isArray(value);
const positive=value=>Number.isFinite(value)&&value>0;

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
