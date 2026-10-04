import {DetachError} from './formula.mjs';
import {children} from './xml.mjs';
const rule=(attrs='',kids='')=>({attrs:attrs.split(' ').filter(Boolean),kids:kids.split(' ').filter(Boolean)});
const colors='srgbClr sysClr schemeClr prstClr';
const transforms='alpha alphaMod alphaOff tint shade satMod satOff lumMod lumOff hue hueMod hueOff red redMod redOff green greenMod greenOff blue blueMod blueOff gamma invGamma comp inv gray';
const schema={
 theme:rule('name','themeElements objectDefaults extraClrSchemeLst'),themeElements:rule('','clrScheme fontScheme fmtScheme'),
 clrScheme:rule('name','dk1 lt1 dk2 lt2 accent1 accent2 accent3 accent4 accent5 accent6 hlink folHlink'),
 fontScheme:rule('name','majorFont minorFont'),majorFont:rule('','latin ea cs font'),minorFont:rule('','latin ea cs font'),latin:rule('typeface panose pitchFamily charset'),ea:rule('typeface panose pitchFamily charset'),cs:rule('typeface panose pitchFamily charset'),font:rule('script typeface'),
 fmtScheme:rule('name','fillStyleLst lnStyleLst effectStyleLst bgFillStyleLst'),fillStyleLst:rule('','solidFill gradFill noFill'),bgFillStyleLst:rule('','solidFill gradFill noFill'),lnStyleLst:rule('','ln'),effectStyleLst:rule('','effectStyle'),
 solidFill:rule('',colors),noFill:rule(),gradFill:rule('flip rotWithShape','gsLst lin path tileRect'),gsLst:rule('','gs'),gs:rule('pos',colors),lin:rule('ang scaled'),path:rule('path','fillToRect'),fillToRect:rule('l t r b'),tileRect:rule('l t r b'),
 ln:rule('w cap cmpd algn','solidFill gradFill noFill prstDash round bevel miter headEnd tailEnd'),prstDash:rule('val'),round:rule(),bevel:rule(),miter:rule('lim'),headEnd:rule('type w len'),tailEnd:rule('type w len'),
 effectStyle:rule('','effectLst scene3d sp3d'),effectLst:rule('','outerShdw innerShdw glow softEdge blur reflection'),outerShdw:rule('blurRad dist dir sx sy kx ky algn rotWithShape',colors),innerShdw:rule('blurRad dist dir',colors),glow:rule('rad',colors),softEdge:rule('rad'),blur:rule('rad grow'),reflection:rule('blurRad stA stPos endA endPos dist dir fadeDir sx sy kx ky algn rotWithShape'),
 scene3d:rule('','camera lightRig'),camera:rule('prst fov zoom','rot'),lightRig:rule('rig dir','rot'),rot:rule('lat lon rev'),sp3d:rule('z extrusionH contourW prstMaterial','bevelT bevelB extrusionClr contourClr'),bevelT:rule('w h prst'),bevelB:rule('w h prst'),extrusionClr:rule('',colors),contourClr:rule('',colors),
 srgbClr:rule('val',transforms),sysClr:rule('val lastClr',transforms),schemeClr:rule('val',transforms),prstClr:rule('val',transforms),
 objectDefaults:rule(),extraClrSchemeLst:rule()
};
for(const name of ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'])schema[name]=rule('',colors);
for(const name of transforms.split(' '))schema[name]=rule(['gamma','invGamma','comp','inv','gray'].includes(name)?'':'val');
export function validateDrawingTree(root){const stack=[root];while(stack.length){const node=stack.pop(),name=node.name.replace(/^a:/,''),s=schema[name];if(!s)throw new DetachError('THEME',`Unsupported DrawingML element: ${name}`);for(const attr of Object.keys(node.attributes||{}))if(!['xmlns','xmlns:a'].includes(attr)&&!s.attrs.includes(attr))throw new DetachError('THEME',`Unsupported ${name} attribute: ${attr}`);for(const c of children(node)){if(!s.kids.includes(c.name.replace(/^a:/,'')))throw new DetachError('THEME',`Unsupported ${name} child: ${c.name}`);stack.push(c);}if((node.elements||[]).some(e=>(e.type==='text'&&e.text.trim())||(e.type==='cdata'&&e.cdata.trim())))throw new DetachError('THEME',`Unexpected DrawingML text: ${name}`);}}
