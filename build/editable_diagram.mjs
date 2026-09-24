import fs from 'node:fs/promises';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const out = 'C:/Users/32933/Documents/ChatGPT/PPT制图/可编辑_电场强化沸腾传热实验系统图.pptx';
const dir = 'C:/Users/32933/Documents/ChatGPT/PPT制图/build/rendered';
await fs.mkdir(dir, { recursive: true });

const deck = Presentation.create({ slideSize: { width: 1000, height: 720 } });
const slide = deck.slides.add();
slide.background.fill = '#FFFFFF';

const C={blue:'#3F69A8', orange:'#F5A128', red:'#D91F26', pale:'#D9D7F0', green:'#9BCB91', yellow:'#FFF51F', controller:'#DCE4F2', ink:'#111111'};
const font='Microsoft YaHei';
function box(name,x,y,w,h,fill,text,fs=24,bold=true,line='#333333'){
  const s=slide.shapes.add({geometry:'rect',name,position:{left:x,top:y,width:w,height:h},fill,line:{style:'solid',fill:line,width:2}});
  s.text=text; s.text.style={fontFamily:font,fontSize:fs,bold,color:C.ink,alignment:'center',verticalAlignment:'middle'}; return s;
}
function txt(name,x,y,w,h,text,fs=18,color=C.ink,bold=false,align='center'){
  const s=slide.shapes.add({geometry:'textbox',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:'solid',fill:'none',width:0}});
  s.text=text; s.text.style={fontFamily:font,fontSize:fs,bold,color,alignment:align,verticalAlignment:'middle'}; return s;
}
function line(name,x,y,w,h,color,width=3,dash='solid'){
  return slide.shapes.add({geometry:'line',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:dash,fill:color,width}});
}
function arrow(name,geom,x,y,w,h,color){ return slide.shapes.add({geometry:geom,name,position:{left:x,top:y,width:w,height:h},fill:color,line:{style:'solid',fill:color,width:1}}); }

// Connectors first so they remain behind all nodes.
// Left control loop.
line('left-orange-vertical',82,340,0,173,C.orange,9);
line('left-orange-horizontal',82,513,254,0,C.orange,9);
arrow('left-loop-arrow','rightArrow',315,500,35,27,C.orange);
arrow('left-up-arrow','upArrow',66,335,32,40,C.orange);
// Center RS422 link.
line('rs422-line',480,309,0,169,C.orange,9);
arrow('rs422-up','upArrow',464,303,32,38,C.orange);
arrow('rs422-down','downArrow',464,459,32,38,C.orange);
// Right control loop.
line('right-orange-horizontal',639,513,267,0,C.orange,9);
line('right-orange-vertical',906,258,0,255,C.orange,9);
arrow('right-loop-left','leftArrow',620,500,35,27,C.orange);
arrow('right-up-arrow','upArrow',890,246,32,40,C.orange);
// Controller data / power links.
arrow('data-up','upArrow',390,535,34,50,C.blue);
arrow('data-down','downArrow',390,575,34,50,C.blue);
arrow('power-up','upArrow',570,305,38,321,C.blue);
// Optical observation link.
line('ccd-dashed-link',620,132,204,0,C.blue,5,'dash');
arrow('ccd-left-arrow','leftArrow',607,120,31,25,C.blue);

// Left module group and dashed boundary.
const boundary=slide.shapes.add({geometry:'rect',name:'left-module-boundary',position:{left:7,top:2,width:190,height:316},fill:'none',line:{style:'dash',fill:C.red,width:3}});
box('small-flow-module',27,20,145,112,C.pale,'小流量供\n液模块',22,true);
box('large-flow-module',27,178,145,112,C.pale,'大流量供\n液模块',22,true);
line('y52-pipe',172,50,178,0,C.red,4); line('y51-pipe',172,93,178,0,C.red,4);
line('y54-pipe',172,210,178,0,C.red,4); line('y53-pipe',172,253,178,0,C.red,4);
txt('y52',177,16,55,36,'Y52',18,C.blue,true); txt('y61',286,16,58,36,'Y61',18,C.red,true);
txt('dn8a',233,51,70,30,'DN8',17,C.ink,true); txt('y51',177,72,55,36,'Y51',18,C.blue,true); txt('y64',286,72,58,36,'Y64',18,C.red,true); txt('dn8b',233,94,70,30,'DN8',17,C.ink,true);
txt('y54',177,176,55,36,'Y54',18,C.blue,true); txt('y62',286,176,58,36,'Y62',18,C.red,true); txt('dn12a',226,210,83,30,'DN12',17,C.ink,true);
txt('y53',177,219,55,36,'Y53',18,C.blue,true); txt('y63',286,219,58,36,'Y63',18,C.red,true); txt('dn12b',226,254,83,30,'DN12',17,C.ink,true);
box('left-precontrol',7,325,190,48,C.red,'前置控制单元',18,true,'#9A1820');

// Central experiment module.
box('experiment-module',350,2,257,305,C.green,'电场强化\n沸腾传热\n实验模块',31,true);
box('experiment-precontrol',382,254,192,48,C.red,'前置控制单元',18,true,'#9A1820');

// Optical platform.
box('optical-platform',826,73,139,111,C.pale,'光学观测\n平台',22,true);
box('optical-precontrol',798,184,194,48,C.red,'前置控制单元',18,true,'#9A1820');
txt('ccd-label',648,73,169,66,'左视细观CCD\n左视高速CCD',17,C.ink,true,'center');

// Main control and controller.
box('main-control',340,484,301,49,C.yellow,'两相柜实验主控单元',21,true,'#777000');
box('experiment-controller',338,626,280,88,C.controller,'实验柜控制器',22,true);
txt('rs422',397,376,70,30,'RS422',16,C.ink,true);
txt('x83',500,315,64,30,'X83',19,C.blue,true); txt('x70',500,453,64,30,'X70',19,'#C67E14',true);
txt('data-label',316,545,72,72,'数据\n通讯',18,C.red,true);
txt('power-label',616,543,118,70,'100V供电\n背插J95',18,C.red,true,'left');

slide.addNotes?.('[Sources]\n- User-provided reference image: codex-clipboard-c04cfa1c-a313-425e-bfee-3f65e55aba04.png (visual reconstruction only).');

// Slide 2: editable reconstruction of the thermal-management schematic.
const s2=deck.slides.add();
s2.background.fill='#FFFFFF';
function b2(name,x,y,w,h,text,fs=26){
  const s=s2.shapes.add({geometry:'rect',name,position:{left:x,top:y,width:w,height:h},fill:'#FFFFFF',line:{style:'solid',fill:'#111111',width:3}});
  s.text=text; s.text.style={fontFamily:font,fontSize:fs,bold:false,color:'#111111',alignment:'center',verticalAlignment:'middle'}; return s;
}
function t2(name,x,y,w,h,text,fs=20,bold=false,align='center'){
  const s=s2.shapes.add({geometry:'textbox',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:'solid',fill:'none',width:0}});
  s.text=text; s.text.style={fontFamily:font,fontSize:fs,bold,color:'#111111',alignment:align,verticalAlignment:'middle'}; return s;
}
function l2(name,x,y,w,h,dash='solid'){
  return s2.shapes.add({geometry:'line',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:dash,fill:'#111111',width:3}});
}
function a2(name,geom,x,y,w,h){
  return s2.shapes.add({geometry:geom,name,position:{left:x,top:y,width:w,height:h},fill:'#111111',line:{style:'solid',fill:'#111111',width:1}});
}
function fan(name,cx,cy){
  for(let i=0;i<5;i++){
    const angle=i*72;
    const rad=(angle-90)*Math.PI/180;
    const px=cx+Math.cos(rad)*25-10, py=cy+Math.sin(rad)*25-17;
    s2.shapes.add({geometry:'teardrop',name:`${name}-blade-${i+1}`,position:{left:px,top:py,width:20,height:34,rotation:angle+45},fill:'#FFFFFF',line:{style:'solid',fill:'#333333',width:2}});
  }
  s2.shapes.add({geometry:'ellipse',name:`${name}-hub`,position:{left:cx-9,top:cy-9,width:18,height:18},fill:'#FFFFFF',line:{style:'solid',fill:'#333333',width:3}});
}

// Pipes and thermal contact arrows are behind the equipment blocks.
l2('s2-small-supply',202,263,211,0); a2('s2-small-supply-head','rightArrow',394,252,28,22);
l2('s2-small-return',202,339,211,0); a2('s2-small-return-head','leftArrow',200,328,28,22);
l2('s2-large-supply',202,474,211,0); a2('s2-large-supply-head','rightArrow',394,463,28,22);
l2('s2-large-return',202,551,211,0); a2('s2-large-return-head','leftArrow',200,540,28,22);
l2('s2-contact-line',753,406,111,0); a2('s2-contact-left','leftArrow',744,395,28,22); a2('s2-contact-right','rightArrow',850,395,28,22);

// Liquid-cooled group.
s2.shapes.add({geometry:'rect',name:'s2-liquid-boundary',position:{left:28,top:181,width:198,height:444},fill:'none',line:{style:'dash',fill:'#111111',width:3}});
b2('s2-small-module',45,229,157,153,'小流量\n供液模块',28);
b2('s2-large-module',45,435,157,153,'大流量\n供液模块',28);
t2('s2-liquid-title',73,641,108,40,'液冷',27);

// Valve and pipe labels.
t2('s2-y52',222,218,55,36,'Y52',20); t2('s2-y61',332,218,55,36,'Y61',20);
t2('s2-dn8',268,284,70,34,'DN8',18,true);
t2('s2-y51',222,350,55,36,'Y51',20); t2('s2-y64',332,350,55,36,'Y64',20);
t2('s2-y54',222,424,55,36,'Y54',20); t2('s2-y62',332,424,55,36,'Y62',20);
t2('s2-dn12',266,496,76,34,'DN12',18,true);
t2('s2-y53',222,560,55,36,'Y53',20); t2('s2-y63',332,560,55,36,'Y63',20);

// Central experiment module and air cooling.
b2('s2-experiment-module',415,229,338,359,'电场强化\n沸腾传热\n实验模块',43);
t2('s2-air-title',517,70,136,45,'风冷',28);
fan('s2-fan-left',510,166); fan('s2-fan-right',656,166);

// Optical observation platform and contact cooling.
b2('s2-optical-platform',866,300,100,211,'光学\n观测\n平台',26);
t2('s2-contact-label',763,421,91,80,'接触面\n散热',21);
s2.addNotes?.('[Sources]\n- User-provided reference image: codex-clipboard-e75b9d22-ec5b-4d8e-a7a1-7bc0eae8b23c.png (visual reconstruction only).');

// Slide 3: editable experimental-project module architecture.
const s3=deck.slides.add();
s3.background.fill='#FFFFFF';
function b3(name,x,y,w,h,text,fs=16,bold=false){
  const s=s3.shapes.add({geometry:'rect',name,position:{left:x,top:y,width:w,height:h},fill:'#FFFFFF',line:{style:'solid',fill:'#111111',width:2}});
  s.text=text; s.text.style={fontFamily:font,fontSize:fs,bold,color:'#111111',alignment:'center',verticalAlignment:'middle'}; return s;
}
function l3(name,x,y,w,h,dash='solid',width=2){
  return s3.shapes.add({geometry:'line',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:dash,fill:'#111111',width}});
}
function ar3(name,geom,x,y,w,h){
  return s3.shapes.add({geometry:geom,name,position:{left:x,top:y,width:w,height:h},fill:'#111111',line:{style:'solid',fill:'#111111',width:1}});
}

// All connection lines are created first and remain behind the modules.
l3('s3-top-left-bus',155,78,0,55);
l3('s3-top-pool-bus',466,78,0,62);
l3('s3-top-measure-bus',839,78,0,60);
l3('s3-flow-group-down',155,240,0,35);
l3('s3-boiling-group-down',466,235,0,40);
l3('s3-measure-down',839,210,0,40);
l3('s3-flow-left-down',76,348,0,250);
l3('s3-flow-right-down',235,348,0,250);
l3('s3-small-flow-down',540,348,0,250);
l3('s3-ccd-left-down',145,445,0,153);
l3('s3-ccd-fast-down',394,445,0,153);
l3('s3-main-controller-down',748,565,0,33);
l3('s3-fan-down',923,348,0,250);

// Vertical supply arrows.
l3('s3-large-flow-up-line',76,397,0,-38); ar3('s3-large-flow-up','upArrow',65,350,22,28);
l3('s3-cooling-up-line',235,397,0,-38); ar3('s3-cooling-up','upArrow',224,350,22,28);
l3('s3-small-flow-up-line',540,397,0,-38); ar3('s3-small-flow-up','upArrow',529,350,22,28);
l3('s3-fan-up-line',923,397,0,-38); ar3('s3-fan-up','upArrow',912,350,22,28);

// Control hierarchy on the right.
l3('s3-hv-to-secondary',720,178,24,0); ar3('s3-hv-left','leftArrow',711,167,24,22);
l3('s3-secondary-to-hv',719,178,-14,0); ar3('s3-secondary-right','rightArrow',704,167,24,22);
l3('s3-secondary-sensor-down',748,322,0,26);
l3('s3-secondary-valve-down',748,322,55,26);
l3('s3-secondary-stem',748,322,0,43);
l3('s3-sensor-branch',693,348,110,0);
ar3('s3-sensor-down-head','downArrow',682,340,22,28); ar3('s3-valve-down-head','downArrow',792,340,22,28);
l3('s3-secondary-main-link',748,322,-115,0);
l3('s3-secondary-main-vertical',633,322,0,200);
l3('s3-main-horizontal',633,522,70,0);
ar3('s3-main-right-head','rightArrow',691,511,24,22);

// Top-level horizontal control arrows around high-voltage converter.
l3('s3-pool-to-hv',545,178,45,0); ar3('s3-pool-left-head','leftArrow',535,167,24,22);
l3('s3-hv-to-pool',590,178,-15,0); ar3('s3-pool-right-head','rightArrow',575,167,24,22);

// Bottom and top common bars.
b3('s3-project-header',12,18,976,60,'实验项目模块',25,false);
b3('s3-common-footer',12,650,976,58,'两相柜公共模块',24,false);

// Dashed experimental-unit group.
s3.shapes.add({geometry:'rect',name:'s3-experiment-group',position:{left:67,top:130,width:480,height:110},fill:'none',line:{style:'dash',fill:'#111111',width:2}});
b3('s3-flow-boiling-unit',88,148,135,75,'流动/射流复合\n沸腾实验单元',16,false);
b3('s3-pool-boiling-unit',403,145,132,78,'池沸腾\n实验单元',17,false);

// Experimental apparatus row.
b3('s3-flow-boiling-device',12,275,132,73,'流动/射流复合\n沸腾实验装置',16,false);
b3('s3-flow-heat-device',170,275,132,73,'流动/射流复合\n沸腾散热装置',16,false);
b3('s3-pool-device',329,275,132,73,'池沸腾\n实验装置',17,false);
b3('s3-pool-heat-device',473,275,132,73,'池沸腾\n散热装置',17,false);

// Supply/cooling and camera modules.
b3('s3-large-flow-module',12,397,132,73,'大流量\n供液模块',17,false);
b3('s3-cooling-device',170,397,132,73,'大流量供液\n冷却装置',16,false);
b3('s3-small-flow-module',473,397,132,73,'小流量\n供液模块',17,false);
b3('s3-micro-ccd',98,500,94,65,'细观CCD',17,false);
b3('s3-high-speed-ccd',347,500,94,65,'高速CCD',17,false);

// Measurement and control modules.
b3('s3-hv-converter',588,145,132,78,'高压电\n转换装置',17,false);
b3('s3-measure-control',778,145,122,78,'测量/控制模块',16,false);
b3('s3-secondary-controller',690,250,122,72,'次级控制器\n（电控盒）',16,false);
b3('s3-sensor-flowmeter',650,365,90,73,'传感器/\n流量计',16,false);
b3('s3-solenoid-valve',758,365,90,73,'电磁阀门',16,false);
b3('s3-main-controller',703,500,92,65,'主控制器',17,false);
b3('s3-controlled-heat',861,250,122,72,'电控盒\n散热结构',16,false);
b3('s3-fan-module',861,397,122,73,'风冷模块',17,false);

s3.addNotes?.('[Sources]\n- User-provided reference image: codex-clipboard-10ea398c-4897-4313-8f89-8ffff3cc53b1.png (visual reconstruction only).');

for(const [i,sl] of deck.slides.items.entries()){
  const stem='slide-'+String(i+1);
  const png=await deck.export({slide:sl,format:'png',scale:1});
  await fs.writeFile(dir+'/'+stem+'.png',new Uint8Array(await png.arrayBuffer()));
  const layout=await sl.export({format:'layout'});
  await fs.writeFile(dir+'/'+stem+'.layout.json',await layout.text());
}
const pptx=await PresentationFile.exportPptx(deck); await pptx.save(out);
console.log(out);
