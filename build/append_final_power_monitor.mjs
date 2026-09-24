import fs from 'node:fs/promises';
import { FileBlob, PresentationFile } from '@oai/artifact-tool';

const source='C:/Users/32933/Documents/ChatGPT/PPT制图/build/PPT绘图汇总_16页备份.pptx';
const output='C:/Users/32933/Documents/ChatGPT/PPT制图/PPT绘图汇总.pptx';
const renderDir='C:/Users/32933/Documents/ChatGPT/PPT制图/build/power_monitor_render';
await fs.mkdir(renderDir,{recursive:true});
const deck=await PresentationFile.importPptx(await FileBlob.load(source));
const s=deck.slides.add();
s.background.fill='#FFFFFF';
const font='Microsoft YaHei';
function box(name,x,y,w,h,value,fs=22){
  const q=s.shapes.add({geometry:'rect',name,position:{left:x,top:y,width:w,height:h},fill:'#FFFFFF',line:{style:'solid',fill:'#111111',width:3}});
  q.text=value; q.text.style={fontFamily:font,fontSize:fs,color:'#111111',alignment:'center',verticalAlignment:'middle'}; return q;
}
function text(name,x,y,w,h,value,fs=22){
  const q=s.shapes.add({geometry:'textbox',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:'solid',fill:'none',width:0}});
  q.text=value; q.text.style={fontFamily:font,fontSize:fs,color:'#111111',alignment:'center',verticalAlignment:'middle'}; return q;
}
function line(name,x,y,w,h,width=3){
  return s.shapes.add({geometry:'line',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:'solid',fill:'#111111',width}});
}
function arrow(name,geometry,x,y,w,h){
  return s.shapes.add({geometry,name,position:{left:x,top:y,width:w,height:h},fill:'#111111',line:{style:'solid',fill:'#111111',width:1}});
}

// Power circuit and measurement connections first.
line('supply-to-switch',280,385,100,0); arrow('supply-to-switch-head','rightArrow',363,373,28,24);
line('switch-to-clamp',520,385,85,0); arrow('switch-to-clamp-head','rightArrow',588,373,28,24);
line('clamp-to-module',635,385,125,0); arrow('clamp-to-module-head','rightArrow',743,373,28,24);
line('return-main',280,530,480,0); arrow('return-left-head','leftArrow',270,518,30,24);
line('module-return-down',760,385,0,145);

// Oscilloscope/current-clamp branch.
line('instrument-down',470,190,0,150);
line('instrument-horizontal',470,190,145,0);
line('clamp-branch',615,190,0,130);
arrow('instrument-left-head','leftArrow',456,178,30,24);

// Main editable equipment blocks.
box('power-supply',40,330,240,200,'供电电源',28);
box('air-switch',380,330,140,110,'空气开关',24);
box('boiling-module',760,330,200,200,'沸腾换热实验模块',24);
box('oscilloscope-clamp',335,85,190,105,'示波器＋\n电流钳',23);

// Current clamp symbol.
s.shapes.add({geometry:'ellipse',name:'current-clamp-ring',position:{left:600,top:330,width:30,height:105},fill:'#FFFFFF',line:{style:'solid',fill:'#111111',width:3}});

// Electrical labels.
text('positive-100v',385,270,130,45,'+100V',24);
text('return-label',345,535,190,45,'+100V回线',23);

s.addNotes?.('[Sources]\n- User-provided reference image: codex-clipboard-329c7cf6-5b2b-47df-9c6a-1358d7288492.png (transparent-background visual reconstruction only).');
const png=await deck.export({slide:s,format:'png',scale:1});
await fs.writeFile(renderDir+'/last-slide.png',new Uint8Array(await png.arrayBuffer()));
const layout=await s.export({format:'layout'});
await fs.writeFile(renderDir+'/last-slide.layout.json',await layout.text());
const pptx=await PresentationFile.exportPptx(deck);
await pptx.save(output);
console.log('SOURCE_SLIDES='+String(deck.slides.items.length-1));
console.log('FINAL_SLIDES='+String(deck.slides.items.length));
console.log(output);
