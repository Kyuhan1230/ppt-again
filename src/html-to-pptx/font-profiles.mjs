import fs from 'node:fs';
import path from 'node:path';
import {sha256} from './assets.mjs';

// Font identity, not document/element identity. These bytes and native faces were
// independently exercised by experiment 003; any changed bytes disable mapping.
const profiles=[{webSha256:'9599f12fd42fc0bce1cd50b47a0c022e108d7aa64dd0d1bb0ed44f3282d900b4',family:'Pretendard Variable',version:'1.309',faces:{
  400:{family:'Pretendard',localName:'Pretendard Regular',file:'Pretendard-Regular.otf',sha256:'3ffbacde6ab8411f1d2db54bb9b1f0b3ee2a738932033722cf0388c06aed1c93'},
  500:{family:'Pretendard Medium',localName:'Pretendard Medium',file:'Pretendard-Medium.otf',sha256:'d39e50e4bb52b4993b6a4eeb821a171254745bd824446af01e1f616b89fface0'},
  600:{family:'Pretendard SemiBold',localName:'Pretendard SemiBold',file:'Pretendard-SemiBold.otf',sha256:'c89bc43027dc7cde5726e96223376f8eec09302b2fc1f8147fd5b57cfc376118'},
  700:{family:'Pretendard Bold',localName:'Pretendard Bold',file:'Pretendard-Bold.otf',sha256:'2e91915fab54df71cc9598ebf608b2bdb54c6fe3c066ac61dff0bc44fca71cc7'},
  750:{family:'Pretendard ExtraBold',localName:'Pretendard ExtraBold',file:'Pretendard-ExtraBold.otf',nativeWeight:800,approximate:true,sha256:'c35fe941b7568d52a96010561540e47f9d3948dfde66ba25bc1908233e0a40cd'},
  800:{family:'Pretendard ExtraBold',localName:'Pretendard ExtraBold',file:'Pretendard-ExtraBold.otf',sha256:'c35fe941b7568d52a96010561540e47f9d3948dfde66ba25bc1908233e0a40cd'}
}}];
export function identifyFontMappings(input){
  const directories=[process.env.SystemRoot&&path.join(process.env.SystemRoot,'Fonts'),process.env.LOCALAPPDATA&&path.join(process.env.LOCALAPPDATA,'Microsoft','Windows','Fonts')].filter(Boolean),result=[];
  for(const face of input.fontFaces??[]){
    const asset=input.assetRefs.filter(r=>r.usage==='font'&&face.css.includes(r.url)).map(r=>input.assets.find(a=>a.id===r.assetId)).find(Boolean);
    const profile=profiles.find(p=>p.webSha256===asset?.sha256);if(!profile)continue;
    // One alias must resolve to the same verified bytes throughout this document.
    // Other same-family declarations or local() alternatives can select a different face.
    const familyFaces=input.fontFaces.filter(f=>f.family.toLowerCase()===face.family.toLowerCase());
    if(familyFaces.some(f=>f.hasLocal||!f.assetIds?.length||f.assetIds.some(id=>input.assets.find(a=>a.id===id)?.sha256!==asset.sha256)))continue;
    for(const [weight,native] of Object.entries(profile.faces)){
      const file=directories.map(d=>path.join(d,native.file)).find(f=>fs.existsSync(f)&&sha256(fs.readFileSync(f))===native.sha256);if(!file)continue;
      if(result.some(m=>m.sourceFamily===face.family&&m.sourceWeight===Number(weight)))continue;
      result.push({sourceFamily:face.family,sourceWeight:Number(weight),sourceSha256:asset.sha256,nativeFamily:native.family,nativeLocalName:native.localName,nativeBold:false,nativeWeight:native.nativeWeight??Number(weight),approximate:native.approximate??false,path:file,sha256:native.sha256,version:profile.version,layoutCalibration:'unmeasured'});
    }
  }
  return result;
}

// A Windows family name need not be a valid CSS local() full face name. Only
// accept aliases backed by the same installed font bytes as our known profile.
export function installedNativeAliases(){
  const directories=[process.env.SystemRoot&&path.join(process.env.SystemRoot,'Fonts'),process.env.LOCALAPPDATA&&path.join(process.env.LOCALAPPDATA,'Microsoft','Windows','Fonts')].filter(Boolean),result=[];
  for(const profile of profiles)for(const native of Object.values(profile.faces)){
    if(native.family===native.localName||result.some(a=>a.family===native.family))continue;
    const file=directories.map(d=>path.join(d,native.file)).find(f=>fs.existsSync(f)&&sha256(fs.readFileSync(f))===native.sha256);
    if(file)result.push({family:native.family,localName:native.localName,sha256:native.sha256});
  }
  return result;
}
