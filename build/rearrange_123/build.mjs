import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const root='C:/Users/32933/Documents/ChatGPT/PPT制图';
const dir=path.join(root,'build/rearrange_123');
const final=path.join(root,'output/123前面板_重新排布_可编辑.pptx');
const skill=process.env.SKILL_DIR;
const bg=await fs.readFile(path.join(dir,'fixed_panel.png'));
const seal=await fs.readFile(path.join(dir,'school_seal.png'));
const deck=Presentation.create({slideSize:{width:1800,height:1273}});
const slide=deck.slides.add();
slide.background.fill='#FFFFFF';
function addImage(name,bytes,x,y,w,h){
  slide.images.add({blob:bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength),contentType:'image/png',alt:name,fit:'contain',position:{left:x,top:y,width:w,height:h}});
}
function shape(name,geometry,x,y,w,h,fill='none',border='none',bw=0){
  return slide.shapes.add({geometry,name,position:{left:x,top:y,width:w,height:h},fill,line:{style:'solid',fill:border,width:bw}});
}
function label(name,x,y,w,h,value,size=35,bold=false,align='left'){
  const s=shape(name,'textbox',x,y,w,h);
  s.text=value;
  s.text.style={typeface:'Microsoft YaHei',fontSize:size,color:'#111111',bold,alignment:align,verticalAlignment:'middle',autoFit:'none'};
  return s;
}
function rule(name,x,y,w,color='#111111',thick=1.5){
  shape(name,'line',x,y,w,0,'none',color,thick);
}
addImage('固定面板底图（接口与加工轮廓不可移动）',bg,0,0,1800,1273);
// The seal and every text/design component below is a distinct selectable object.
addImage('西安交通大学校徽—可移动可缩放',seal,343,655,108,108);
label('设备标题—可移动',482,663,720,87,'新型无泵式沸腾散热实验模块',52,true);
rule('标题下划线',483,758,707,'#1A1A1A',1.8);

// Certificate diamond: editable border and editable text, not a flattened image.
const diamond=shape('TGMTYY型号菱形边框—可移动','diamond',605,823,310,150,'none','#111111',1.3);
label('菱形上行型号—可移动',667,837,186,43,'Z01-1',31,false,'center');
label('菱形中行编号—可移动',618,872,284,45,'TGMTYY14027',32,false,'center');
label('菱形下行年份—可移动',690,914,139,42,'2026',31,false,'center');

rule('署名上方分隔线',350,1010,705,'#888888',1);
label('负责人标签—可移动可缩放',375,1027,260,50,'项目负责人：',32,true);
label('负责人姓名—可移动可缩放',637,1027,405,50,'杨小平    魏进家',35,false);
label('成员标签—可移动可缩放',375,1083,260,50,'项目成员：',32,true);
label('成员姓名—可移动可缩放',637,1083,405,50,'崔庆杰    王晋璞',35,false);

slide.speakerNotes.textFrame.setText('[Sources]\n- User-provided C:/Users/32933/Desktop/123.pdf.\nThe fixed engineering geometry, connector markings, and Innovation Harbour graphic are retained from the source PDF as a non-editable background. Only the explicitly permitted elements are reconstructed as movable slide objects. Confirm final silk-screen clearance against the engineering source before manufacture.');
const preview=await deck.export({slide,format:'png',scale:1});
await fs.writeFile(path.join(dir,'preview.png'),new Uint8Array(await preview.arrayBuffer()));
const candidate=path.join(dir,'candidate.pptx');
await (await PresentationFile.exportPptx(deck)).save(candidate);
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const result=await finalizePresentation({workspaceDir:root,candidatePath:candidate,finalPath:final,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu','17145000,12125325'],explicitTotalSlideCount:1,requiredNativeTableOwnerSlides:[],requiredNativeChartOwnerSlides:[],fontPolicy:{basis:'design',families:['Microsoft YaHei']},verifyArtifactToolImport:true,receiptPath:path.join(dir,'validation.json')});
console.log(JSON.stringify({final,result},null,2));
