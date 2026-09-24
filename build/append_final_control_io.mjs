import fs from 'node:fs/promises';
import { FileBlob, PresentationFile } from '@oai/artifact-tool';

const source='C:/Users/32933/Documents/ChatGPT/PPT制图/build/当前15页版本_备份.pptx';
const output='C:/Users/32933/Documents/ChatGPT/PPT制图/可编辑_电场强化沸腾传热实验系统图_追加控制框图.pptx';
const renderDir='C:/Users/32933/Documents/ChatGPT/PPT制图/build/control_io_render';
await fs.mkdir(renderDir,{recursive:true});
const deck=await PresentationFile.importPptx(await FileBlob.load(source));
const s=deck.slides.add();
s.background.fill='#FFFFFF';
const font='Microsoft YaHei';
function box(name,x,y,w,h,text,fs=20){
  const q=s.shapes.add({geometry:'rect',name,position:{left:x,top:y,width:w,height:h},fill:'#FFFFFF',line:{style:'solid',fill:'#111111',width:3}});
  q.text=text; q.text.style={fontFamily:font,fontSize:fs,color:'#111111',alignment:'center',verticalAlignment:'middle'}; return q;
}
function text(name,x,y,w,h,value,fs=20){
  const q=s.shapes.add({geometry:'textbox',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:'solid',fill:'none',width:0}});
  q.text=value; q.text.style={fontFamily:font,fontSize:fs,color:'#111111',alignment:'center',verticalAlignment:'middle'}; return q;
}
function line(name,x,y,w,h,width=3){
  return s.shapes.add({geometry:'line',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:'solid',fill:'#111111',width}});
}
function arrow(name,geometry,x,y,w,h){
  return s.shapes.add({geometry,name,position:{left:x,top:y,width:w,height:h},fill:'#111111',line:{style:'solid',fill:'#111111',width:1}});
}

// Connectors first.
line('main-to-box-line',500,138,0,62); arrow('main-down-arrow','downArrow',488,177,24,28); arrow('main-up-arrow','upArrow',488,132,24,28);
line('control-box-stem',500,270,0,38);

// Left control bracket and four downward branches.
line('control-bus',235,308,265,0); line('control-left-drop',235,308,0,31); line('control-right-drop',500,308,0,31);
arrow('control-left-corner','downArrow',223,324,24,28);
line('control-group-top',82,338,320,0); line('control-group-left',82,338,0,20); line('control-group-right',402,338,0,20);
line('heater-command',82,420,0,36); arrow('heater-command-arrow','downArrow',70,443,24,28);
line('hv-command',192,420,0,36); arrow('hv-command-arrow','downArrow',180,443,24,28);
line('valve-command',302,420,0,36); arrow('valve-command-arrow','downArrow',290,443,24,28);
line('proportion-command',412,420,0,36); arrow('proportion-command-arrow','downArrow',400,443,24,28);

// Right read bracket and four upward measurement branches.
line('read-bus',500,308,265,0); line('read-right-up',765,308,0,31); arrow('read-up-arrow','upArrow',753,324,24,28);
line('read-group-top',618,338,320,0); line('read-group-left',618,338,0,20); line('read-group-right',938,338,0,20);
line('temperature-read',658,456,0,36); arrow('temperature-up','upArrow',646,420,24,28);
line('flow-read',768,456,0,36); arrow('flow-up','upArrow',756,420,24,28);
line('pressure-read',878,456,0,36); arrow('pressure-up','upArrow',866,420,24,28);
line('diff-pressure-read',958,456,0,36); arrow('diff-pressure-up','upArrow',946,420,24,28);

// Main hierarchy.
box('experiment-main-control',260,70,480,68,'实验主控单元',27);
box('experiment-electrical-box',260,200,480,70,'实验模块电控盒',25);
text('parameter-transfer-label',535,145,235,42,'传输指令及工参',23);

// Group labels.
text('control-label',145,298,90,40,'控制',23);
text('read-label',765,298,90,40,'读取',23);

// Control devices and commands.
box('heater',36,358,92,62,'加热片',18);
box('high-voltage-device',146,358,92,62,'高压装置',18);
box('normally-closed-valve',256,358,92,62,'常闭阀',18);
box('proportional-valve',366,358,92,62,'比例阀',18);
box('heater-code',36,492,92,82,'编号\n电压',20);
box('hv-code',146,492,92,82,'编号\n电压',20);
box('closed-valve-code',256,492,92,82,'编号\n开闭',20);
box('proportional-code',366,492,92,82,'编号\n开度',20);

// Measurement devices and values.
box('thermocouple',612,358,92,62,'热电偶',18);
box('flowmeter',722,358,92,62,'流量计',18);
box('pressure-gauge',832,358,92,62,'压力计',18);
box('differential-gauge',926,358,64,62,'差压计',14);
box('temperature-value',612,492,92,82,'温度',20);
box('flow-value',722,492,92,82,'流量',20);
box('pressure-value',832,492,92,82,'压力',20);
box('differential-value',926,492,64,82,'压差',20);

s.addNotes?.('[Sources]\n- User-provided reference image: codex-clipboard-de9e11d8-f8a9-4de9-9aa9-d8936f0a5604.png (visual reconstruction only).');
const png=await deck.export({slide:s,format:'png',scale:1});
await fs.writeFile(renderDir+'/last-slide.png',new Uint8Array(await png.arrayBuffer()));
const layout=await s.export({format:'layout'});
await fs.writeFile(renderDir+'/last-slide.layout.json',await layout.text());
const pptx=await PresentationFile.exportPptx(deck);
await pptx.save(output);
console.log('SOURCE_SLIDES='+String(deck.slides.items.length-1));
console.log('FINAL_SLIDES='+String(deck.slides.items.length));
console.log(output);
