import test from 'node:test';
import assert from 'node:assert/strict';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import PptxGenJS from 'pptxgenjs';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {strFromU8,strToU8,unzipSync,zipSync} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';

const child=(node,name)=>[...(node?.childNodes??[])].find(n=>n.nodeType===1&&n.localName===name);
const all=(node,name)=>[...node.getElementsByTagNameNS('*',name)];
const psLiteral=value=>`'${value.replaceAll("'","''")}'`;

test('installed PowerPoint preserves list markers, paragraph spacing and character formatting',
  {skip:process.platform!=='win32'||process.env.PPT_AGAIN_OFFICE_TEST!=='1',timeout:120000},async()=>{
    const dir=await fs.mkdtemp(path.join(os.tmpdir(),'ppt-again-text-office-'));
    const paths=Object.fromEntries(['zero','marker','table','numbering','roundtrip'].map(name=>[name,path.join(dir,`${name}.pptx`)]));
    let runtime;
    try{
      runtime=await createBrowserRuntime({capture:false});
      const convert=async html=>{
        const measured=await measureInput(await readInput({html}),runtime),result=await exportPptx(measured);
        assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));return result.pptxBytes;
      };
      const slide=(content,style='font:28px Arial;line-height:36px')=>`<section data-slide style="position:relative;width:1280px;height:720px;${style}">${content}</section>`;
      await fs.writeFile(paths.zero,await convert(slide('<ul id="zero" style="position:absolute;left:200px;top:120px;margin:0;padding:0;list-style:disc"><li>First item</li><li>Second item</li></ul>','font:32px Arial;line-height:42px')));
      await fs.writeFile(paths.marker,await convert(`<style>li::marker{content:"★ ";color:#c00000}</style>${slide('<ul id="marker"><li>First</li><li>Second</li></ul>')}`));
      await fs.writeFile(paths.table,await convert(slide('<table id="table" style="position:absolute;left:80px;top:80px;width:700px;border-collapse:collapse"><tr><td style="border:1px solid black"><ul style="margin:0;padding-left:24px;line-height:32px"><li style="margin-bottom:12px;letter-spacing:2px">First</li><li>Second</li></ul></td><td style="border:1px solid black"><ol start="3" style="margin:0"><li>Parent<ul style="margin:0"><li>Child</li></ul></li><li>Next</li></ol></td></tr></table>','font:24px Arial')));
      await fs.writeFile(paths.numbering,await convert(slide('<ol id="simple" start="3" style="position:absolute;left:80px;top:80px"><li>First</li><li>Next</li></ol><ol id="nested" start="3" style="position:absolute;left:80px;top:220px"><li>Parent<ul style="margin:0"><li>Child</li></ul></li><li>Next</li></ol>')));

      const ppt=new PptxGenJS(),s=ppt.addSlide();
      s.addText('Follow',{x:1,y:1,w:4,h:1,objectName:'follow',fontFace:'Arial',fontSize:20,bullet:{characterCode:'2022',indent:18}});
      s.addText('First\nSecond',{x:1,y:2,w:4,h:2,objectName:'spacing',fontFace:'Arial',fontSize:20,margin:0,paraSpaceAfter:6});
      s.addText('One\nTwo\nThree',{x:1,y:4,w:4,h:2,objectName:'numbered',fontFace:'Arial',fontSize:20,margin:0});
      s.addTable([['First\nSecond']],{x:6,y:1,w:4,h:2,objectName:'end-style',fontFace:'Arial',fontSize:20,bold:true});
      const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
      const doc=new DOMParser().parseFromString(strFromU8(parts[key]),'application/xml'),A='http://schemas.openxmlformats.org/drawingml/2006/main';
      const shape=all(doc,'sp').find(n=>all(n,'cNvPr')[0]?.getAttribute('name')==='follow');
      const follow=all(shape,'pPr')[0],size=child(follow,'buSzPct'),bullet=child(follow,'buChar');
      follow.insertBefore(doc.createElementNS(A,'a:buClrTx'),size);follow.insertBefore(doc.createElementNS(A,'a:buFontTx'),bullet);
      const numbered=all(doc,'sp').find(n=>all(n,'cNvPr')[0]?.getAttribute('name')==='numbered');
      for(const pPr of all(numbered,'pPr')){const none=child(pPr,'buNone');if(none)pPr.removeChild(none);const auto=doc.createElementNS(A,'a:buAutoNum');auto.setAttribute('type','arabicPeriod');auto.setAttribute('startAt','3');pPr.appendChild(auto);}
      const frame=all(doc,'graphicFrame').find(n=>all(n,'cNvPr')[0]?.getAttribute('name')==='end-style');
      const p=all(frame,'p')[0],props=child(p,'pPr'),none=child(props,'buNone');
      props.removeChild(none);props.setAttribute('marL','228600');props.setAttribute('indent','-228600');
      const marker=doc.createElementNS(A,'a:buChar');marker.setAttribute('char','•');props.appendChild(marker);
      const rPr=all(p,'rPr')[0];rPr.setAttribute('spc','40');rPr.setAttribute('b','1');
      const end=child(p,'endParaRPr');if(end)p.removeChild(end);
      parts[key]=strToU8(new XMLSerializer().serializeToString(doc));
      const parsed=await readPptx(zipSync(parts));assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
      const original=exportHtml(parsed.model).html;
      const edited=original.replace('>Second</span></span></span></div>','>Second</span>\n</span><span data-pptx-paragraph style="display:block;line-height:32px;text-align:left;padding:0;text-indent:0">Third</span></span></div>');
      assert.notEqual(edited,original);await fs.writeFile(paths.roundtrip,await convert(edited));
      await runtime.close();runtime=null;

      const command=`$ErrorActionPreference='Stop'
$mutex=New-Object System.Threading.Mutex($false,'Local\\ppt-again-chart-export')
$held=$false;$app=$null;$deck=$null;$existing=@(Get-Process POWERPNT -ErrorAction SilentlyContinue);$oldSecurity=$null
try{
 $held=$mutex.WaitOne(120000);if(!$held){throw 'Office is busy'}
 $app=New-Object -ComObject PowerPoint.Application;$oldSecurity=$app.AutomationSecurity;$app.AutomationSecurity=3
 $deck=$app.Presentations.Open(${psLiteral(paths.zero)},-1,0,0)
 $shape=$deck.Slides.Item(1).Shapes.Item('zero');$p=$shape.TextFrame2.TextRange.Paragraphs(2,1)
 if($p.ParagraphFormat.FirstLineIndent -ge 0 -or $p.BoundHeight -ge 40 -or [math]::Abs($p.BoundLeft-150) -gt 1){throw 'Outside list position or wrapping differs'}
 $deck.Close();$deck=$null
 $deck=$app.Presentations.Open(${psLiteral(paths.marker)},-1,0,0)
 $p=$deck.Slides.Item(1).Shapes.Item('marker').TextFrame2.TextRange.Paragraphs(1,1)
 if($p.ParagraphFormat.Bullet.Character -ne 9733){throw 'Custom marker differs'}
 $deck.Close();$deck=$null
 $deck=$app.Presentations.Open(${psLiteral(paths.table)},-1,0,0)
 $p=$deck.Slides.Item(1).Shapes.Item('table').Table.Cell(1,1).Shape.TextFrame2.TextRange.Paragraphs(1,1)
 $font=$p.Characters(1,$p.Text.TrimEnd([char]13,[char]10).Length).Font
 if($p.ParagraphFormat.Bullet.Character -ne 8226 -or $p.ParagraphFormat.Bullet.Visible -ne -1 -or [math]::Abs($p.ParagraphFormat.SpaceAfter-9) -gt .02 -or [math]::Abs($font.Spacing-1.5) -gt .02){throw 'Table list formatting differs'}
 $r=$deck.Slides.Item(1).Shapes.Item('table').Table.Cell(1,2).Shape.TextFrame2.TextRange
 if($r.Paragraphs(1,1).ParagraphFormat.Bullet.Number -ne 3 -or $r.Paragraphs(3,1).ParagraphFormat.Bullet.Number -ne 4){throw 'Nested table numbering differs'}
 $deck.Close();$deck=$null
 $deck=$app.Presentations.Open(${psLiteral(paths.numbering)},-1,0,0)
 $r=$deck.Slides.Item(1).Shapes.Item('simple').TextFrame2.TextRange
 if($r.Paragraphs(1,1).ParagraphFormat.Bullet.Number -ne 3 -or $r.Paragraphs(2,1).ParagraphFormat.Bullet.Number -ne 4){throw 'Simple list numbering differs'}
 $r=$deck.Slides.Item(1).Shapes.Item('nested').TextFrame2.TextRange
 if($r.Paragraphs(1,1).ParagraphFormat.Bullet.Number -ne 3 -or $r.Paragraphs(3,1).ParagraphFormat.Bullet.Number -ne 4){throw 'Nested list numbering differs'}
 $deck.Close();$deck=$null
 $deck=$app.Presentations.Open(${psLiteral(paths.roundtrip)},-1,0,0)
 $p=$deck.Slides.Item(1).Shapes.Item('follow').TextFrame2.TextRange.Paragraphs(1,1);$p.Font.Fill.ForeColor.RGB=255;$p.Font.Name='Arial'
 if($p.ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB -ne 255 -or $p.ParagraphFormat.Bullet.Font.Name -ne 'Arial'){throw 'Bullet does not follow text formatting'}
 $p=$deck.Slides.Item(1).Shapes.Item('spacing').TextFrame2.TextRange.Paragraphs(2,1)
 if([math]::Abs($p.ParagraphFormat.SpaceAfter) -gt .02){throw 'Previously hidden paragraph spacing returned'}
 $r=$deck.Slides.Item(1).Shapes.Item('numbered').TextFrame2.TextRange
 if($r.Paragraphs(1,1).ParagraphFormat.Bullet.Number -ne 3 -or $r.Paragraphs(2,1).ParagraphFormat.Bullet.Number -ne 4 -or $r.Paragraphs(3,1).ParagraphFormat.Bullet.Number -ne 5){throw 'PowerPoint numbering round trip differs'}
 $p=$deck.Slides.Item(1).Shapes.Item('end-style').Table.Cell(1,1).Shape.TextFrame2.TextRange.Paragraphs(1,1);$mark=$p.Characters($p.Text.Length,1).Font
 if([math]::Abs($mark.Spacing-.4) -gt .02 -or $mark.Bold -ne -1){throw 'Paragraph end character formatting differs'}
 Write-Output 'PowerPoint text-format checks passed'
}finally{
 if($deck){$deck.Close()}
 if($app){if($null -ne $oldSecurity -and $app.AutomationSecurity -eq 3){$app.AutomationSecurity=$oldSecurity};if($existing.Count -eq 0 -and $app.Presentations.Count -eq 0){$app.Quit()}}
 if($held){$mutex.ReleaseMutex()};$mutex.Dispose()
}`;
      const {stdout}=await promisify(execFile)('powershell.exe',['-NoProfile','-NonInteractive','-EncodedCommand',Buffer.from(command,'utf16le').toString('base64')],{windowsHide:true,timeout:120000,maxBuffer:65536});
      assert.match(stdout,/PowerPoint text-format checks passed/);
    }finally{
      await runtime?.close();
      for(const file of Object.values(paths))await fs.unlink(file).catch(()=>{});
      await fs.rmdir(dir).catch(()=>{});
    }
  });
