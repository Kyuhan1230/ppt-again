// PptxGenJS emits kern="0" when charSpacing is truthy. An epsilon rounds to
// zero in its XML, disabling kerning without changing character spacing.
export function pptCharacterSpacing(style,pointScale){
  const spacing=style.letterSpacing&&style.letterSpacing!=='normal'?parseFloat(style.letterSpacing)*pointScale:0;
  return spacing||(style.disableKerning===true?Number.EPSILON:0);
}
