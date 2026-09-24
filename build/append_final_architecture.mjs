import fs from 'node:fs/promises';
import { FileBlob, PresentationFile } from '@oai/artifact-tool';

const source='C:/Users/32933/Documents/ChatGPT/PPT制图/build/当前用户版本_备份.pptx';
const output='C:/Users/32933/Documents/ChatGPT/PPT制图/可编辑_电场强化沸腾传热实验系统图_追加框图.pptx';
const renderDir='C:/Users/32933/Documents/ChatGPT/PPT制图/build/final_append_render';
await fs.mkdir(renderDir,{recursive:true});
const deck=await PresentationFile.importPptx(await FileBlob.load(source));
const s=deck.slides.add();
s.background.fill='#FFFFFF';
const font='Microsoft YaHei';
function box(name,x,y,w,h,text,fs=16){
  const q=s.shapes.add({geometry:'rect',name,position:{left:x,top:y,width:w,height:h},fill:'#FFFFFF',line:{style:'solid',fill:'#111111',width:2}});
  q.text=text; q.text.style={fontFamily:font,fontSize:fs,color:'#111111',alignment:'center',verticalAlignment:'middle'}; return q;
}
function line(name,x,y,w,h,dash='solid',width=2){
  return s.shapes.add({geometry:'line',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:dash,fill:'#111111',width}});
}
function arrow(name,geometry,x,y,w,h){
  return s.shapes.add({geometry,name,position:{left:x,top:y,width:w,height:h},fill:'#111111',line:{style:'solid',fill:'#111111',width:1}});
}

// Connectors first.
line('top-left-bus',155,78,0,70); line('top-pool-bus',466,78,0,67); line('top-measure-bus',839,78,0,67);
line('group-left-down',155,223,0,52); line('group-pool-down',466,223,0,52); line('measure-down',839,223,0,27);
line('large-flow-bottom',76,470,0,180); line('cooling-bottom',235,470,0,180); line('small-flow-bottom',540,470,0,180);
line('micro-ccd-bottom',145,565,0,85); line('fast-ccd-bottom',394,565,0,85); line('main-bottom',748,565,0,85); line('fan-bottom',923,470,0,180);
line('large-flow-up',76,359,0,38); arrow('large-flow-arrow','upArrow',65,350,22,28);
line('cooling-up',235,359,0,38); arrow('cooling-arrow','upArrow',224,350,22,28);
line('small-flow-up',540,359,0,38); arrow('small-flow-arrow','upArrow',529,350,22,28);
line('fan-up',923,359,0,38); arrow('fan-arrow','upArrow',912,350,22,28);

line('pool-hv-link',535,184,53,0); arrow('pool-hv-left','leftArrow',526,173,24,22); arrow('pool-hv-right','rightArrow',575,173,24,22);
line('hv-secondary-link',720,184,15,0); arrow('hv-secondary-left','leftArrow',711,173,24,22); arrow('hv-secondary-right','rightArrow',722,173,24,22);
line('secondary-sensor-stem',751,322,0,26); line('sensor-valve-branch',696,348,110,0);
arrow('sensor-down','downArrow',685,340,22,28); arrow('valve-down','downArrow',795,340,22,28);
line('secondary-main-left',633,286,57,0); line('secondary-main-vertical',633,286,0,236); line('secondary-main-right',633,522,70,0); arrow('main-right','rightArrow',691,511,24,22);

box('project-header',12,18,976,60,'实验项目模块',25);
box('common-footer',12,650,976,58,'两相柜公共模块',24);
s.shapes.add({geometry:'rect',name:'experiment-unit-group',position:{left:67,top:130,width:480,height:110},fill:'none',line:{style:'dash',fill:'#111111',width:2}});
box('flow-boiling-unit',88,148,135,75,'流动/射流复合\n沸腾实验单元',16);
box('pool-boiling-unit',403,145,132,78,'池沸腾\n实验单元',17);
box('flow-boiling-device',12,275,132,73,'流动/射流复合\n沸腾实验装置',16);
box('flow-heat-device',170,275,132,73,'流动/射流复合\n沸腾散热装置',16);
box('pool-device',329,275,132,73,'池沸腾\n实验装置',17);
box('pool-heat-device',473,275,132,73,'池沸腾\n散热装置',17);
box('large-flow-module',12,397,132,73,'大流量\n供液模块',17);
box('cooling-device',170,397,132,73,'大流量供液\n冷却装置',16);
box('small-flow-module',473,397,132,73,'小流量\n供液模块',17);
box('micro-ccd',98,500,94,65,'细观CCD',17);
box('high-speed-ccd',347,500,94,65,'高速CCD',17);
box('hv-converter',588,145,132,78,'高压电\n转换装置',17);
box('measure-control',778,145,122,78,'测量/控制模块',16);
box('secondary-controller',690,250,122,72,'次级控制器\n（电控盒）',16);
box('sensor-flowmeter',650,365,90,73,'传感器/\n流量计',16);
box('solenoid-valve',758,365,90,73,'电磁阀门',16);
box('main-controller',703,500,92,65,'主控制器',17);
box('controlled-heat',861,250,122,72,'电控盒\n散热结构',16);
box('fan-module',861,397,122,73,'风冷模块',17);
s.addNotes?.('[Sources]\n- User-provided reference image: codex-clipboard-10ea398c-4897-4313-8f89-8ffff3cc53b1.png (visual reconstruction only).');

const png=await deck.export({slide:s,format:'png',scale:1});
await fs.writeFile(renderDir+'/last-slide.png',new Uint8Array(await png.arrayBuffer()));
const layout=await s.export({format:'layout'});
await fs.writeFile(renderDir+'/last-slide.layout.json',await layout.text());
const pptx=await PresentationFile.exportPptx(deck);
await pptx.save(output);
console.log('SOURCE_SLIDES='+String(deck.slides.items.length-1));
console.log('FINAL_SLIDES='+String(deck.slides.items.length));
console.log(output);
