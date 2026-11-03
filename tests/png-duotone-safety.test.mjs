import test from 'node:test';
import assert from 'node:assert/strict';
import {deflateSync} from 'node:zlib';
import {DOMParser} from '@xmldom/xmldom';
import {applyPngDuotone,duotoneColors,validatePngStructure} from '../src/html-to-pptx/png-duotone.mjs';

const original=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAF0lEQVR4nGP4z8Dwn+E/QwMjmP7P8B8APeIHeyMDVGUAAAAASUVORK5CYII=','base64');
const endpoints=[[0,0,0],[255,255,255]];
function crc32(bytes){let c=0xffffffff;for(const byte of bytes){c^=byte;for(let bit=0;bit<8;bit++)c=c&1?0xedb88320^(c>>>1):c>>>1;}return (c^0xffffffff)>>>0;}
function chunk(type,data){const b=Buffer.alloc(12+data.length);b.writeUInt32BE(data.length,0);b.write(type,4,4,'ascii');data.copy(b,8);b.writeUInt32BE(crc32(b.subarray(4,8+data.length)),8+data.length);return b;}
function pngWith(header,compressed){return Buffer.concat([original.subarray(0,8),chunk('IHDR',header),chunk('IDAT',compressed),chunk('IEND',Buffer.alloc(0))]);}
test('DrawingML shade and saturation resolve in the correct order',()=>{
 const doc=new DOMParser().parseFromString('<a:duotone xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="E7E6E6"><a:shade val="45000"/><a:satMod val="135000"/></a:srgbClr><a:prstClr val="white"/></a:duotone>','application/xml');
 assert.deepEqual(duotoneColors(doc.documentElement,()=>null),[[162,161,161],[255,255,255]]);
});
test('PNG pixel and compressed expansion limits are enforced',()=>{
 const hdr=Buffer.from(original.subarray(16,29));hdr.writeUInt32BE(4_000_001,0);hdr.writeUInt32BE(1,4);
 assert.throws(()=>applyPngDuotone(pngWith(hdr,deflateSync(Buffer.from([0,0,0,0,0]))),endpoints),e=>e.code==='LIMIT_EXCEEDED');
 const small=Buffer.from(original.subarray(16,29)),tooMuch=Buffer.alloc(10000);
 assert.throws(()=>applyPngDuotone(pngWith(small,deflateSync(tooMuch)),endpoints));
});
test('duplicate PNG chunks, noncontiguous image data and invalid row filters are blocked',()=>{
 const header=Buffer.from(original.subarray(16,29));
 const duplicate=Buffer.concat([original.subarray(0,33),original.subarray(8,33),original.subarray(33)]);
 assert.throws(()=>applyPngDuotone(duplicate,endpoints));
 const compressed=deflateSync(Buffer.alloc(18));
 const split=Buffer.concat([original.subarray(0,8),chunk('IHDR',header),chunk('IDAT',compressed.subarray(0,2)),chunk('tEXt',Buffer.from('x\0y')),chunk('IDAT',compressed.subarray(2)),chunk('IEND',Buffer.alloc(0))]);
 assert.throws(()=>applyPngDuotone(split,endpoints));
 const invalidRows=Buffer.alloc(18);invalidRows[0]=5;
 assert.throws(()=>applyPngDuotone(pngWith(header,deflateSync(invalidRows)),endpoints));
});
test('CRC failures block decoding and valid unsupported interlacing is left untouched',()=>{
 const damaged=Buffer.from(original);damaged[damaged.length-1]^=1;assert.throws(()=>applyPngDuotone(damaged,endpoints));
 const hdr=Buffer.from(original.subarray(16,29));hdr[12]=1;
 const unsupported=pngWith(hdr,deflateSync(Buffer.alloc(18)));assert.equal(applyPngDuotone(unsupported,endpoints),null);
});
test('invalid methods, duplicate singleton metadata and trailing compressed bytes are blocked',()=>{
 const hdr=Buffer.from(original.subarray(16,29)),raw=Buffer.alloc(18),valid=deflateSync(raw);
 const badMethod=Buffer.from(hdr);badMethod[10]=1;assert.throws(()=>applyPngDuotone(pngWith(badMethod,valid),endpoints));
 const duplicate=Buffer.concat([original.subarray(0,8),chunk('IHDR',hdr),chunk('sRGB',Buffer.from([0])),chunk('sRGB',Buffer.from([0])),chunk('IDAT',valid),chunk('IEND',Buffer.alloc(0))]);
 assert.throws(()=>applyPngDuotone(duplicate,endpoints));
 assert.throws(()=>applyPngDuotone(pngWith(hdr,Buffer.concat([valid,Buffer.from([1,2,3])])),endpoints));
});
test('color-managed PNGs outside sRGB are not silently recolored',()=>{
 const hdr=Buffer.from(original.subarray(16,29)),valid=deflateSync(Buffer.alloc(18));
 const gamma=Buffer.alloc(4);gamma.writeUInt32BE(100000,0);
 const nonSrgb=Buffer.concat([original.subarray(0,8),chunk('IHDR',hdr),chunk('gAMA',gamma),chunk('IDAT',valid),chunk('IEND',Buffer.alloc(0))]);
 assert.equal(applyPngDuotone(nonSrgb,endpoints),null);
 const profile=Buffer.concat([original.subarray(0,8),chunk('IHDR',hdr),chunk('iCCP',Buffer.from('profile\0\0data')),chunk('IDAT',valid),chunk('IEND',Buffer.alloc(0))]);
 assert.equal(applyPngDuotone(profile,endpoints),null);
});

test('math fallback PNG validation accepts standard, palette and interlaced data',()=>{
 assert.doesNotThrow(()=>validatePngStructure(original));
 const header=Buffer.from(original.subarray(16,29)),interlaced=Buffer.from(header);interlaced[12]=1;
 // A 2x2 RGBA Adam7 image has three nonempty passes: 5, 5 and 9 bytes.
 assert.doesNotThrow(()=>validatePngStructure(pngWith(interlaced,deflateSync(Buffer.alloc(19)))));
 assert.throws(()=>validatePngStructure(pngWith(interlaced,deflateSync(Buffer.alloc(18)))));
 const palette=Buffer.from(header);palette[8]=1;palette[9]=3;
 const indexed=Buffer.concat([original.subarray(0,8),chunk('IHDR',palette),chunk('PLTE',Buffer.from([255,0,0,0,0,255])),chunk('IDAT',deflateSync(Buffer.from([0,0,0,128]))),chunk('IEND',Buffer.alloc(0))]);
 assert.doesNotThrow(()=>validatePngStructure(indexed));
 const extraPalette=Buffer.concat([original.subarray(0,8),chunk('IHDR',palette),chunk('PLTE',Buffer.alloc(9)),chunk('IDAT',deflateSync(Buffer.alloc(4))),chunk('IEND',Buffer.alloc(0))]);
 assert.throws(()=>validatePngStructure(extraPalette));
 const missingPaletteEntry=Buffer.concat([original.subarray(0,8),chunk('IHDR',palette),chunk('PLTE',Buffer.alloc(3)),chunk('IDAT',deflateSync(Buffer.from([0,128,0,0]))),chunk('IEND',Buffer.alloc(0))]);
 assert.throws(()=>validatePngStructure(missingPaletteEntry));
 const trailing=pngWith(header,Buffer.concat([deflateSync(Buffer.alloc(18)),Buffer.from([0,0,0,0])]));
 assert.doesNotThrow(()=>validatePngStructure(trailing));
 const tooLarge=Buffer.from(header);tooLarge.writeUInt32BE(100000,0);tooLarge.writeUInt32BE(100000,4);
 assert.throws(()=>validatePngStructure(pngWith(tooLarge,deflateSync(Buffer.alloc(18)))),e=>e.code==='LIMIT_EXCEEDED');
 const manyPixels=Buffer.from(header);manyPixels.writeUInt32BE(8000,0);manyPixels.writeUInt32BE(8000,4);manyPixels[8]=1;manyPixels[9]=0;
 assert.throws(()=>validatePngStructure(pngWith(manyPixels,deflateSync(Buffer.alloc(18)))),e=>e.code==='LIMIT_EXCEEDED');
});
