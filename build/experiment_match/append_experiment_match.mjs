import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { FileBlob, PresentationFile } from '@oai/artifact-tool';

const workspaceDir='C:/Users/32933/Documents/ChatGPT/PPT制图';
const source=path.join(workspaceDir,'build/PPT绘图汇总_当前备份.pptx');
const tmpDir=path.join(workspaceDir,'build/experiment_match');
const finalPath=path.join(workspaceDir,'output/PPT绘图汇总_20页.pptx');
const skillDir=process.env.SKILL_DIR;
await fs.mkdir(tmpDir,{recursive:true});
await fs.mkdir(path.dirname(finalPath),{recursive:true});
const deck=await PresentationFile.importPptx(await FileBlob.load(source));
const s=deck.slides.add();
s.background.fill='#050505';
const blue='#4596D8', white='#F7F7F7', font='Microsoft YaHei';
function box(name,x,y,w,h,value,fs=17){
  const q=s.shapes.add({geometry:'rect',name,position:{left:x,top:y,width:w,height:h},fill:'#070707',line:{style:'solid',fill:blue,width:2}});
  q.text=value; q.text.style={fontFamily:font,fontSize:fs,color:blue,alignment:'center',verticalAlignment:'middle',autoFit:'shrinkText'}; return q;
}
function line(name,x,y,w,h,width=2){
  return s.shapes.add({geometry:'line',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:'solid',fill:blue,width}});
}
function arrow(name,geometry,x,y,w,h){
  return s.shapes.add({geometry,name,position:{left:x,top:y,width:w,height:h},fill:blue,line:{style:'solid',fill:blue,width:1}});
}

// Hierarchy connectors are behind all nodes.
line('root-to-branch',230,360,25,0);
line('left-vertical-bus',255,125,0,470);
line('pre-state-branch',255,125,45,0);
line('desktop-branch',255,360,45,0);
line('system-branch',255,595,45,0);

// Pre-match test sequence.
line('pre-to-test-bus',490,125,22,0);
line('pre-test-vertical',512,65,0,180);
line('pre-mech-branch',512,65,20,0);
line('pre-thermal-branch',512,135,20,0);
line('pre-electrical-branch',512,205,20,0);
line('mech-down',608,92,0,16); arrow('mech-down-head','downArrow',598,100,20,24);
line('thermal-down',608,162,0,16); arrow('thermal-down-head','downArrow',598,170,20,24);

// Desktop experiment sequence and branch.
line('desktop-to-sequence',490,360,22,0);
line('desktop-sequence-vertical',512,290,0,170);
line('desktop-power-branch',512,290,20,0);
line('desktop-match-branch',512,360,20,0);
line('desktop-condition-branch',512,430,20,0);
line('desktop-power-down',608,317,0,16); arrow('desktop-power-head','downArrow',598,325,20,24);
line('desktop-match-down',608,387,0,16); arrow('desktop-match-head','downArrow',598,395,20,24);

// System experiment sequence and branch.
line('system-to-sequence',490,595,22,0);
line('system-sequence-vertical',512,545,0,100);
line('system-power-branch',512,545,20,0);
line('system-condition-branch',512,615,20,0);
line('system-power-down',608,572,0,16); arrow('system-power-head','downArrow',598,580,20,24);

// Collection bus and arrows into summary.
line('collection-bus',710,65,0,580);
for(const [name,y] of [['mech',65],['thermal',135],['electrical',205],['desktop-power',290],['desktop-match',360],['desktop-condition',430],['system-power',545],['system-condition',615]]){
  line('collect-'+name,680,y,30,0); arrow('collect-head-'+name,'rightArrow',694,y-10,22,20);
}
line('bus-to-summary',710,275,45,0); arrow('bus-summary-head','rightArrow',742,265,24,20);

// Right-side handoff.
line('summary-to-report',850,325,0,40); arrow('summary-report-head','downArrow',840,350,20,24);
line('report-to-sample',850,455,0,40); arrow('report-sample-head','downArrow',840,480,20,24);

// Editable nodes.
box('engineering-match-root',40,325,190,70,'工程件匹配实验',23);
box('pre-match-state',300,85,190,80,'匹配前状态确认',23);
box('desktop-experiment',300,320,190,80,'桌面级实验',23);
box('system-experiment',300,555,190,80,'系统级实验',23);

box('mechanical-interface-test',532,38,148,54,'机械接口测试',18);
box('thermal-interface-test',532,108,148,54,'热接口测试',18);
box('electrical-interface-test',532,178,148,54,'电接口测试',18);
box('desktop-power-test',532,263,148,54,'桌面级系统上电\n与软件功能测试',16);
box('two-phase-cabinet-match',532,333,148,54,'与两相柜各模\n块匹配',16);
box('fourteen-conditions',532,403,148,54,'14个实验工况',18);
box('system-power-test',532,518,148,54,'系统级系统上电\n与软件功能测试',16);
box('six-conditions',532,588,148,54,'6个典型工况',18);

box('experiment-summary',755,235,190,90,'实验总结',24);
box('submit-report',755,365,190,90,'提交科学匹配实验\n报告',21);
box('prototype-build',755,495,190,90,'正样件研制',24);

s.speakerNotes.textFrame.setText('[Sources]\n- User-provided reference image: codex-clipboard-0ea210a6-7aa7-45d5-8baf-47462286ec59.png (visual reconstruction only).');
const preview=await deck.export({slide:s,format:'png',scale:1});
await fs.writeFile(path.join(tmpDir,'last-slide.png'),new Uint8Array(await preview.arrayBuffer()));
const layout=await s.export({format:'layout'});
await fs.writeFile(path.join(tmpDir,'last-slide.layout.json'),await layout.text());

const candidatePath=path.join(tmpDir,'candidate.pptx');
await (await PresentationFile.exportPptx(deck)).save(candidatePath);
const { finalizePresentation } = await import(pathToFileURL(path.join(skillDir,'container_tools/artifact_tool_utils.mjs')).href);
const result=await finalizePresentation({
  workspaceDir,
  candidatePath,
  finalPath,
  pythonExecutable:process.env.RUNTIME_PYTHON,
  integrityValidatorPath:path.join(skillDir,'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath:path.join(skillDir,'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs:['--expected-slide-size-emu','9525000,6858000','--validate-bullet-geometry','--validate-heading-fit'],
  explicitTotalSlideCount:20,
  requiredNativeTableOwnerSlides:[],
  requiredNativeChartOwnerSlides:[],
  verifyArtifactToolImport:true,
  receiptPath:path.join(tmpDir,'validation.json'),
});
console.log(JSON.stringify({sourceSlides:19,finalSlides:deck.slides.items.length,finalPath,result},null,2));
