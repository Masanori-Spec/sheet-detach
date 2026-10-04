export class DetachError extends Error { constructor(code,message,path=[]){super(message);this.name='DetachError';this.code=code;this.path=path;} }
const fail=(code,message)=>{throw new DetachError(code,message)};
export function addressOf(raw){const m=/^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})$/.exec(raw);if(!m)fail('REFERENCE',`Unsupported A1 reference: ${raw}`);let col=0;for(const c of m[2].toUpperCase())col=col*26+c.charCodeAt(0)-64;const row=+m[4];if(col>16384||row>1048576)fail('REFERENCE',`Reference outside XLSX bounds: ${raw}`);return {address:m[2].toUpperCase()+row,row,col};}
export function columnName(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s;}
export function parseFormula(input){
 const source=input.startsWith('=')?input.slice(1):input;if(!source||source.length>8192)fail('FORMULA_SIZE','Formula must contain 1–8192 characters');
 let i=0;const tokens=[];
 while(i<source.length){if(/\s/.test(source[i])){i++;continue}const start=i,c=source[i];
  if('+-*/():,!'.includes(c)){tokens.push({type:c,raw:c,start,end:++i});continue}
  if(c==="'"){i++;let value='',closed=false;while(i<source.length){if(source[i]==="'"){if(source[i+1]==="'"){value+="'";i+=2}else{i++;closed=true;break}}else value+=source[i++];}if(!closed)fail('SYNTAX','Unclosed quoted sheet name');tokens.push({type:'quoted',value,raw:source.slice(start,i),start,end:i});continue}
  const num=/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(source.slice(i));
  if(num){i+=num[0].length;tokens.push({type:'number',value:Number(num[0]),raw:num[0],start,end:i});continue}
  const word=/^[\p{L}_$][\p{L}\p{N}_.$]*/u.exec(source.slice(i));
  if(word){i+=word[0].length;tokens.push({type:'word',raw:word[0],start,end:i});continue}
  fail('UNSUPPORTED_FORMULA',`Unsupported formula character at ${i+1}: ${c}`);
 }
 tokens.push({type:'eof',start:i,end:i});let p=0,depth=0;const peek=()=>tokens[p],take=()=>tokens[p++],expect=t=>{if(peek().type!==t)fail('SYNTAX',`Expected ${t} at ${peek().start+1}`);return take()};
 function primary(){if(++depth>128)fail('FORMULA_DEPTH','Formula nesting exceeds 128');let n,t=take();
  if(t.type==='number'){if(peek().type==='!'){take();const a=expect('word');n={kind:'ref',sheet:t.raw,...addressOf(a.raw),start:t.start,end:a.end};}else {if(!Number.isFinite(t.value))fail('NONFINITE','Non-finite numeric literal');n={kind:'number',value:t.value,start:t.start,end:t.end}}}
  else if(t.type==='+'||t.type==='-'){const value=primary();n={kind:'unary',op:t.type,value,start:t.start,end:value.end};}
  else if(t.type==='('){const value=expression();const close=expect(')');n={kind:'group',value,start:t.start,end:close.end};}
  else if(t.type==='quoted'||t.type==='word'){
   if(peek().type==='!'){take();const a=expect('word');n={kind:'ref',sheet:t.type==='quoted'?t.value:t.raw,...addressOf(a.raw),start:t.start,end:a.end};}
   else if(t.type==='word'&&peek().type==='('){if(t.raw.toUpperCase()!=='SUM')fail('FUNCTION',`Unsupported function: ${t.raw}`);take();const args=[];if(peek().type===')')fail('SUM_EMPTY','SUM requires numeric arguments');do{args.push(expression());if(peek().type!==',')break;take()}while(true);const close=expect(')');n={kind:'sum',args,start:t.start,end:close.end};}
   else if(t.type==='word'){n={kind:'ref',sheet:null,...addressOf(t.raw),start:t.start,end:t.end};}
   else fail('SYNTAX','Quoted names must qualify a cell reference');
  }else fail('SYNTAX',`Expected a numeric expression at ${t.start+1}`);
  if(peek().type===':'){if(n.kind!=='ref')fail('RANGE','Range endpoint must be an A1 cell');take();const other=primary();if(other.kind!=='ref')fail('RANGE','Range endpoint must be an A1 cell');n={kind:'range',from:n,to:other,start:n.start,end:other.end};}
  depth--;return n;
 }
 function product(){let n=primary();while(peek().type==='*'||peek().type==='/'){const op=take().type,right=primary();n={kind:'binary',op,left:n,right,start:n.start,end:right.end}}return n}
 function expression(){let n=product();while(peek().type==='+'||peek().type==='-'){const op=take().type,right=product();n={kind:'binary',op,left:n,right,start:n.start,end:right.end}}return n}
 const ast=expression();expect('eof');return {source,ast};
}
export function walkAst(node,fn,parent=null){const stack=[{n:node,parent}];while(stack.length){const item=stack.pop(),n=item.n;fn(n,item.parent);let nodes=[];if(n.kind==='binary')nodes=[n.left,n.right];else if(n.kind==='unary'||n.kind==='group')nodes=[n.value];else if(n.kind==='sum')nodes=n.args;else if(n.kind==='range')nodes=[n.from,n.to];for(let i=nodes.length-1;i>=0;i--)stack.push({n:nodes[i],parent:n});}}
export function rewriteReferences(parsed,replacements){const sorted=[...replacements].sort((a,b)=>a.start-b.start);let text='',end=0;for(const r of sorted){if(r.start<end)fail('INTERNAL','Overlapping AST references');text+=parsed.source.slice(end,r.start)+`(${Object.is(r.value,-0)?'0':String(r.value)})`;end=r.end;}return text+parsed.source.slice(end);}
