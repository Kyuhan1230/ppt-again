// FontFaceSet.check only reports load readiness; CDP identifies the actual
// platform fonts used to render each source paragraph, including fallbacks.
export async function readPaginationFontEvidence({context,page,selector,paragraphs,tables=[]}){
  const cdp=await context.newCDPSession(page);
  try{
    await cdp.send('DOM.enable');await cdp.send('CSS.enable');
    const document=(await cdp.send('DOM.getDocument')).root.nodeId;
    const root=(await cdp.send('DOM.querySelector',{nodeId:document,selector})).nodeId;
    if(!root)throw new Error('선택 문서를 확인할 수 없습니다.');
    const nodes=(await cdp.send('DOM.querySelectorAll',{nodeId:root,selector:':scope > :is(p,h1,h2,h3,h4,h5,h6)'})).nodeIds;
    if(nodes.length!==paragraphs.length)throw new Error('문단의 실제 글꼴 소유 관계가 다릅니다.');
    const verifyFont=async(nodeId,family)=>{
      const fonts=(await cdp.send('CSS.getPlatformFontsForNode',{nodeId})).fonts;
      if(!Array.isArray(fonts)||fonts.length!==1||fonts[0].isCustomFont||!Number.isSafeInteger(fonts[0].glyphCount)||fonts[0].glyphCount<1||fonts[0].familyName.toLowerCase()!==family.toLowerCase())throw new Error('실제 사용 글꼴과 지정 글꼴이 다르거나 글꼴이 혼합되어 있습니다.');
      const {familyName,postScriptName,isCustomFont,glyphCount}=fonts[0];
      return {familyName,postScriptName,isCustomFont,glyphCount};
    };
    const evidence=[];
    for(let i=0;i<nodes.length;i++){
      const font=await verifyFont(nodes[i],paragraphs[i].style.fontFamily);
      evidence.push({objectId:paragraphs[i].id,fonts:[font]});
    }
    for(const table of tables){
      for(let i=0;i<table.cells.length;i+=8){
        const cells=table.cells.slice(i,i+8).filter(cell=>cell.text);
        evidence.push(...await Promise.all(cells.map(async cell=>{
          if(!/^(?:td|th)@\d+$/.test(cell.cellPath))throw new Error('표 셀의 원본 위치를 확인할 수 없습니다.');
          const matches=(await cdp.send('DOM.querySelectorAll',{nodeId:root,selector:`[data-source-path="${cell.cellPath}"]`})).nodeIds;
          if(matches.length!==1)throw new Error('표 셀을 고유하게 확인할 수 없습니다.');
          return {objectId:table.id,sourceId:cell.sourceId,fonts:[await verifyFont(matches[0],cell.family)]};
        })));
      }
    }
    return evidence;
  }finally{await cdp.detach().catch(()=>{});}
}
