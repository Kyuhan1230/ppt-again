// Presentation dimensions are in EMU. PowerPoint requires 1–56 inches.
export const MIN_SLIDE_EMU=914400;
export const MAX_SLIDE_EMU=56*914400;
export function validSizePolicy(value){
  if(!value||typeof value!=='object'||Array.isArray(value))return false;
  if(Object.keys(value).some(key=>!['mode','targetWidthEmu','targetHeightEmu','marginColor'].includes(key))||value.mode!=='contain')return false;
  const hasWidth=value.targetWidthEmu!==undefined,hasHeight=value.targetHeightEmu!==undefined;
  if(hasWidth!==hasHeight)return false;
  if(hasWidth&&(!Number.isSafeInteger(value.targetWidthEmu)||!Number.isSafeInteger(value.targetHeightEmu)||value.targetWidthEmu<MIN_SLIDE_EMU||value.targetHeightEmu<MIN_SLIDE_EMU||value.targetWidthEmu>MAX_SLIDE_EMU||value.targetHeightEmu>MAX_SLIDE_EMU))return false;
  return value.marginColor===undefined||typeof value.marginColor==='string'&&/^#[a-f0-9]{6}$/i.test(value.marginColor);
}
