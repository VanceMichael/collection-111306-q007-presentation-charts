import {inflateRawSync} from 'node:zlib';
export function archiveEntries(bytes) {
 let end=-1;
 for(let offset=bytes.length-22;offset>=Math.max(0,bytes.length-65557);offset--){if(bytes.readUInt32LE(offset)===0x06054b50){end=offset;break;}}
 if(end<0)throw new Error('缺少资源目录');
 let cursor=bytes.readUInt32LE(end+16);const count=bytes.readUInt16LE(end+10);const entries=[];
 for(let index=0;index<count;index++){
  if(bytes.readUInt32LE(cursor)!==0x02014b50)throw new Error('资源目录损坏');
  const method=bytes.readUInt16LE(cursor+10);const size=bytes.readUInt32LE(cursor+20);const nameSize=bytes.readUInt16LE(cursor+28);const extraSize=bytes.readUInt16LE(cursor+30);const noteSize=bytes.readUInt16LE(cursor+32);const local=bytes.readUInt32LE(cursor+42);
  const name=bytes.subarray(cursor+46,cursor+46+nameSize).toString('utf8');
  const start=local+30+bytes.readUInt16LE(local+26)+bytes.readUInt16LE(local+28);
  const compressed=bytes.subarray(start,start+size);const content=method===0?compressed:method===8?inflateRawSync(compressed):null;
  if(!content)throw new Error('不支持的压缩类型');entries.push({name,content});cursor+=46+nameSize+extraSize+noteSize;
 }
 return entries;
}
export function preview(data) {return {package:data.package,selected:data.selected};}
