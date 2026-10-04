import {DetachError} from './formula.mjs';
const fail=message=>{throw new DetachError('XML',message)};
const validCode=c=>c===9||c===10||c===13||(c>=32&&c<=0xd7ff)||(c>=0xe000&&c<=0xfffd)||(c>=0x10000&&c<=0x10ffff);
export function xmlLexicalPreflight(text){
 for(const ch of text)if(!validCode(ch.codePointAt(0)))fail('Forbidden XML 1.0 character');
 for(const m of text.matchAll(/&#(?:x([0-9a-f]+)|([0-9]+));/gi))if(!validCode(parseInt(m[1]||m[2],m[1]?16:10)))fail('Forbidden XML character reference');
 let p=0,count=0,depth=0;const ws=c=>c===' '||c==='\t'||c==='\r'||c==='\n';
 while((p=text.indexOf('<',p))!==-1){if(text.startsWith('<!--',p)){const end=text.indexOf('-->',p+4);if(end<0)fail('Unclosed comment');p=end+3;continue}if(text.startsWith('<![CDATA[',p)){const end=text.indexOf(']]>',p+9);if(end<0)fail('Unclosed CDATA');p=end+3;continue}if(text.startsWith('<?',p)){const end=text.indexOf('?>',p+2);if(end<0)fail('Unclosed processing instruction');const decl=text.slice(p,end+2);if((p!==0&&!(p===1&&text.charCodeAt(0)===0xfeff))||!/^<\?xml\s+version=(['"])1\.0\1(?:\s+encoding=(['"])UTF-8\2)?(?:\s+standalone=(['"])(?:yes|no)\3)?\s*\?>$/i.test(decl))fail('Only the initial XML 1.0 UTF-8 declaration is supported');p=end+2;continue}if(text.startsWith('<!',p))fail('Unsupported XML declaration');if(text.startsWith('</',p)){const end=text.indexOf('>',p+2);if(end<0)fail('Unclosed end tag');depth--;p=end+1;continue}
  if(++count>100000||++depth>64)throw new DetachError('XML_LIMIT','XML depth/element limit exceeded');p++;const start=p;while(p<text.length&&!ws(text[p])&&text[p]!=='>'&&text[p]!=='/')p++;if(p===start)fail('Empty element name');const attrs=new Set();let ended=false;
  while(p<text.length){while(ws(text[p]))p++;if(text[p]==='>'){p++;ended=true;break}if(text[p]==='/'&&text[p+1]==='>'){p+=2;depth--;ended=true;break}const a=p;while(p<text.length&&!ws(text[p])&&!['=','>','/'].includes(text[p]))p++;const name=text.slice(a,p);if(!name||attrs.has(name))fail('Empty or duplicate XML attribute');attrs.add(name);if(attrs.size>1000||name.length>256)throw new DetachError('XML_LIMIT','XML attribute limit exceeded');while(ws(text[p]))p++;if(text[p++]!=='=')fail('Attribute requires equals');while(ws(text[p]))p++;const quote=text[p++];if(quote!=="'"&&quote!=='"')fail('Attribute must be quoted');const end=text.indexOf(quote,p);if(end<0)fail('Unclosed attribute');if(text.slice(p,end).includes('<'))fail('Less-than character in attribute');p=end+1;}
  if(!ended)fail('Unclosed start tag');
 }
}
