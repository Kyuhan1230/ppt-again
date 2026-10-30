import decompress from 'wawoff2/decompress.js';
import {sha256} from './assets.mjs';
import {inspectFont,decideFontEmbedding} from './font-policy.mjs';

// Only decode explicitly supplied, bounded WOFF2 faces; never fetch fonts.
export async function normalizeWoff2Fonts(input){
  const replacements=new Map(),extra=[];let expanded=0;
  const referenced=new Set((input.fontFaces??[]).flatMap(face=>face.assetIds??[]));
  for(const asset of input.assets??[]){
    if(asset.mime!=='font/woff2'||!referenced.has(asset.id))continue;
    try{
      const bytes=Buffer.from(asset.bytes);
      if(bytes.length<48||bytes.toString('ascii',0,4)!=='wOF2'||bytes.readUInt32BE(8)!==bytes.length||sha256(bytes)!==asset.sha256)continue;
      const size=bytes.readUInt32BE(16);
      if(!size||size>16*1024*1024||expanded+size>32*1024*1024)continue;
      const decoded=Buffer.from(await decompress(bytes));
      if(decoded.length!==size)continue;
      const meta=inspectFont(decoded);
      if(!decideFontEmbedding(meta,{forEditing:true,subset:false}).allowed)continue;
      const hash=sha256(decoded),id=`asset-${hash}`;
      replacements.set(asset.id,id);expanded+=decoded.length;
      if(!extra.some(item=>item.id===id)&&!input.assets.some(item=>item.id===id))extra.push({id,sha256:hash,mime:'font/ttf',bytes:decoded});
    }catch{ /* Unsupported or invalid inputs keep the existing diagnostic path. */ }
  }
  return {...input,assets:[...input.assets,...extra],fontFaces:(input.fontFaces??[]).map(face=>({...face,assetIds:(face.assetIds??[]).map(id=>replacements.get(id)??id)}))};
}
