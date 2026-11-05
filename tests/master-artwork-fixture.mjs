import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
const image='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=';
const shape=(name,color,x)=>`<p:sp><p:nvSpPr><p:cNvPr id="99" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`;
const update=(ps,p,fn)=>{ps[p]=strToU8(fn(strFromU8(ps[p])));};
export async function masterFixture({slideHidden=false,layoutHidden=false,unsupported=false,placeholderPic=false,hiddenShape=false}={}){
 const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};s.addShape(p.ShapeType.rect,{x:3,y:0,w:1,h:1,fill:{color:'0000FF'},line:{transparency:100},objectName:'slide-shape'});s.addImage({data:image,x:5,y:1,w:1,h:1,objectName:'slide-picture'});
 const ps=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'})));
 const xml=strFromU8(ps['ppt/slides/slide1.xml']),pic=xml.match(/<p:pic>[\s\S]*?<\/p:pic>/)?.[0];if(!pic)throw Error('fixture missing picture');
 const rid=pic.match(/r:embed="([^"]+)"/)?.[1],rels=strFromU8(ps['ppt/slides/_rels/slide1.xml.rels']);
 const relation=rels.match(new RegExp(`<Relationship Id="${rid}"[^>]*Target="([^"]+)"`));if(!relation)throw Error('fixture missing image relation');
 let layoutPic=pic.replace('slide-picture','layout-picture').replace(`r:embed="${rid}"`,'r:embed="rIdLayoutImage"');
 if(placeholderPic)layoutPic=layoutPic.replace('<p:nvPr></p:nvPr>','<p:nvPr><p:ph type="pic"/></p:nvPr>');
 const layoutShape=unsupported?shape('layout-shape','00FF00',914400).replace('<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>','<a:gradFill/>'):shape('layout-shape','00FF00',914400);
 const hidden=hiddenShape?shape('hidden-artwork','888888',1828800).replace('name="hidden-artwork"','name="hidden-artwork" hidden="1"'):'';
 update(ps,'ppt/slideLayouts/slideLayout1.xml',x=>x.replace('</p:spTree>',layoutShape+layoutPic+hidden+'</p:spTree>').replace('<p:sldLayout ',`<p:sldLayout ${layoutHidden?'showMasterSp="0" ':''}`));
 update(ps,'ppt/slideLayouts/_rels/slideLayout1.xml.rels',x=>x.replace('</Relationships>',`<Relationship Id="rIdLayoutImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${relation[1]}"/></Relationships>`));
 update(ps,'ppt/slideMasters/slideMaster1.xml',x=>x.replace('</p:spTree>',shape('master-shape','FF0000',0)+'</p:spTree>'));
 if(slideHidden)update(ps,'ppt/slides/slide1.xml',x=>x.replace('<p:sld ','<p:sld showMasterSp="0" '));
 return zipSync(ps);
}
