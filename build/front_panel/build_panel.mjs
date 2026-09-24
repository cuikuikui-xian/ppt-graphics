import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const workspace='C:/Users/32933/Documents/ChatGPT/PPT制图';
const tmp=path.join(workspace,'build/front_panel');
const final=path.join(workspace,'output/前面板丝印规划_可编辑.pptx');
const logoPath=path.join(workspace,'build/xjtu_official_red.png');
const skill=process.env.SKILL_DIR;
await fs.mkdir(tmp,{recursive:true});
await fs.mkdir(path.dirname(final),{recursive:true});
const deck=Presentation.create({slideSize:{width:1200,height:675}});
const slide=deck.slides.add();
slide.background.fill='#EEF1F5';
const ink='#243446', navy='#173B62', muted='#52677A', red='#C72B31', metal='#FBFCFD', stroke='#697887';
const font='Microsoft YaHei';
function shape(name,geometry,x,y,w,h,fill='none',border=stroke,bw=1.6){
  return slide.shapes.add({geometry,name,position:{left:x,top:y,width:w,height:h},fill,line:{style:'solid',fill:border,width:bw}});
}
function label(name,x,y,w,h,value,fs=18,color=ink,bold=false,align='center'){
  const q=shape(name,'textbox',x,y,w,h,'none','none',0);
  q.text=value;
  q.text.style={typeface:font,fontSize:fs,color,bold,alignment:align,verticalAlignment:'middle',autoFit:'none'};
  return q;
}
function line(name,x,y,w,h,color=stroke,width=1.6){
  return slide.shapes.add({geometry:'line',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:'solid',fill:color,width}});
}
function connector(name,cx,cy,r){
  shape(name+'-outer','ellipse',cx-r,cy-r,2*r,2*r,metal,ink,2.2);
  shape(name+'-ring','ellipse',cx-r+9,cy-r+9,2*(r-9),2*(r-9),'none',ink,1.4);
  shape(name+'-socket','ellipse',cx-14,cy-14,28,28,'#F1F3F5',ink,1.4);
  for(let i=0;i<4;i++){
    const a=i*Math.PI/2;
    const px=cx+Math.cos(a)*(r-5),py=cy+Math.sin(a)*(r-5);
    shape(name+'-bolt-'+(i+1),'ellipse',px-3,py-3,6,6,metal,ink,1);
  }
}

// Panel geometry follows the first reference drawing, with no artwork in handle or connector clearances.
shape('panel','rect',80,45,1040,570,metal,ink,2.3);
for(const [x,y] of [[94,59],[1106,59],[94,601],[1106,601]]){
  shape('corner-screw-'+x+'-'+y,'ellipse',x-7,y-7,14,14,metal,ink,1.5);
  line('corner-slot-'+x+'-'+y,x-3,y,6,0,ink,1);
}
// Recessed side handles.
for(const [id,x] of [['left',180],['right',992]]){
  shape(id+'-handle-frame','roundRect',x,125,40,390,metal,ink,1.8);
  shape(id+'-handle-grip','roundRect',x+12,159,16,320,'#E2E7EB',ink,1.2);
  shape(id+'-handle-top-bolt','ellipse',x+15,136,10,10,metal,ink,1.2);
  shape(id+'-handle-bottom-bolt','ellipse',x+15,493,10,10,metal,ink,1.2);
}
// Original fixed openings / connectors (positions retained proportionally).
connector('inlet-quick-disconnect',480,188,54);
connector('outlet-quick-disconnect',865,333,54);
shape('L1L2-housing','rect',566,83,72,29,metal,ink,1.3);
for(const x of [588,616]) shape('L1L2-terminal-'+x,'ellipse',x-5,92,10,10,metal,ink,1.2);
shape('X83-flange','rect',842,501,48,48,metal,ink,1.3);
shape('X83-shell','ellipse',849,508,34,34,metal,ink,1.4);
shape('X83-core','ellipse',856,515,20,20,'#E9EEF2',ink,1.2);
for(const [dx,dy] of [[-4,-3],[3,-4],[-3,4],[4,3]]) shape('X83-pin-'+dx+'-'+dy,'ellipse',866+dx-1.8,525+dy-1.8,3.6,3.6,ink,ink,0.5);
for(const [x,y] of [[847,506],[885,506],[847,544],[885,544]]) shape('X83-bolt-'+x+'-'+y,'ellipse',x-2.3,y-2.3,4.6,4.6,metal,ink,0.8);

// Genuine Xi'an Jiaotong University mark, taken from its official VI system.
const logoBytes=await fs.readFile(logoPath);
slide.images.add({blob:logoBytes.buffer.slice(logoBytes.byteOffset,logoBytes.byteOffset+logoBytes.byteLength),contentType:'image/png',alt:'西安交通大学官方校徽',fit:'contain',position:{left:718,top:96,width:70,height:70}});
label('university-name',798,111,185,40,'西安交通大学',24,navy,true,'left');
line('university-underline',798,158,175,0,red,2.5);

// Device name and interface labels, with ample operational clearance.
label('module-title',300,307,500,54,'新型无泵式沸腾散热实验模块',30,navy,true,'center');
line('title-rule',332,367,436,0,red,2.3);
label('inlet-label',402,251,158,35,'工质入口  Y61',18,navy,true);
label('outlet-label',790,397,156,35,'工质出口  Y64',18,navy,true);
label('terminal-label',558,118,88,28,'L1 / L2',15,muted,false);
label('comm-label',791,554,155,34,'通讯接口  X83',17,navy,true);

// Project credits placed as one restrained, readable typographic block.
label('leaders-label',300,426,145,32,'项目负责人',17,muted,true,'left');
label('leaders-names',440,426,325,32,'杨小平    魏进家',19,ink,false,'left');
label('members-label',300,466,145,32,'项目成员',17,muted,true,'left');
label('members-names',440,466,325,32,'崔庆杰    王晋璞',19,ink,false,'left');
line('credits-rule',300,414,470,0,'#C5CDD3',1.2);

slide.speakerNotes.textFrame.setText('[Sources]\n- User-supplied front-panel engineering drawing: C:/Users/32933/Desktop/1.jpg.\n- User-supplied equipment photograph: codex-clipboard-1529e8ce-b47e-4bf4-8723-f4f996b9126b.jpg.\n- Official university emblem: https://vi.xjtu.edu.cn/images/a1-3jdxhred.png (accessed 2026-09-19).\nDesign concept only. Final screen-print positioning requires verification against physical panel, port clearances and manufacturer specifications.');

const preview=await deck.export({slide,format:'png',scale:1});
await fs.writeFile(path.join(tmp,'panel-preview.png'),new Uint8Array(await preview.arrayBuffer()));
const layout=await slide.export({format:'layout'});
await fs.writeFile(path.join(tmp,'panel-layout.json'),await layout.text());
const candidate=path.join(tmp,'candidate.pptx');
await (await PresentationFile.exportPptx(deck)).save(candidate);
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const result=await finalizePresentation({workspaceDir:workspace,candidatePath:candidate,finalPath:final,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu','11430000,6429375','--validate-bullet-geometry','--validate-heading-fit'],explicitTotalSlideCount:1,requiredNativeTableOwnerSlides:[],requiredNativeChartOwnerSlides:[],fontPolicy:{basis:'design',families:[font]},verifyArtifactToolImport:true,receiptPath:path.join(tmp,'validation.json')});
console.log(JSON.stringify({final,result},null,2));
