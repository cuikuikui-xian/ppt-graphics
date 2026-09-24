import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const { SKILL_DIR, TMP_DIR, FINAL_PPTX, RUNTIME_PYTHON } = process.env;
if (![SKILL_DIR, TMP_DIR, FINAL_PPTX, RUNTIME_PYTHON].every(v => path.isAbsolute(v ?? ""))) {
  throw new Error("Required absolute runtime paths are missing");
}
const workspaceDir = "C:\\Users\\32933\\Documents\\ChatGPT\\PPT制图";
const buildDir = "C:\\Users\\32933\\Documents\\ChatGPT\\PPT制图\\mechanical_rebuild";
const assetDir = path.join(buildDir, "assets");
const previewDir = path.join(TMP_DIR, "preview");
const templatePath = path.join(buildDir, "template.pptx");
const templateSha = "f0ab9d2e02b02b480d021007c6fe8e3da53bfe9e93290b89cf04ba07fbac5f06";
const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href);

const W = 1280, H = 720;
const F = "微软雅黑";
const E = F;
const C = {
  ink: "#303235", dark: "#4E5358", body: "#60676E", blue: "#637A8A",
  blue2: "#8195A2", pale: "#E7EAEC", panel: "#F5F6F6", white: "#FFFFFF",
  line: "#B7BDC1", red: "#A94A43"
};

async function bytes(name) {
  const b = await fs.readFile(path.join(assetDir, name));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
}
const bgBytes = await bytes("metal-bg.jpg");
const buttonBytes = await bytes("hero-button.png");
const caseBytes = {
  door: await bytes("door.png"), kettle: await bytes("kettle.png"),
  knob: await bytes("knob.png"), emergency: await bytes("emergency.png")
};

function p(left, top, width, height) { return { left, top, width, height }; }
function addText(slide, text, position, style = {}, name = "text") {
  const sh = slide.shapes.add({
    geometry: "textbox", name, position,
    fill: style.fill ?? "none",
    line: style.line ? { style: "solid", fill: style.line, width: style.lineWidth ?? 1 } : { fill: "none", width: 0 }
  });
  sh.text = text;
  sh.text.style = {
    typeface: style.typeface ?? F,
    fontSize: style.fontSize ?? 22,
    bold: style.bold ?? false,
    color: style.color ?? C.ink,
    alignment: style.alignment ?? "left",
    verticalAlignment: style.verticalAlignment ?? "top",
    autoFit: style.autoFit ?? "shrinkText",
    wrap: "square",
    insets: style.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return sh;
}
function addLine(slide, left, top, width, color = C.line, weight = 1) {
  return slide.shapes.add({
    geometry: "straightConnector1", name: "line", position: p(left, top, width, 0), fill: "none",
    line: { style: "solid", fill: color, width: weight }
  });
}
function addBackground(slide) {
  slide.images.add({ blob: bgBytes, contentType: "image/jpeg", alt: "银色拉丝金属背景", fit: "cover", position: p(0,0,W,H) });
}
function addImage(slide, blob, position, alt, fit = "cover", crop) {
  return slide.images.add({ blob, contentType: "image/png", alt, fit, position, ...(crop ? {crop} : {}) });
}
function sectionHeader(slide, n, title, subtitle = "") {
  addText(slide, `◆  ${String(n).padStart(2,"0")}`, p(42,30,110,34), { fontSize: 18, bold: true, color: C.dark, typeface: E }, `section-number-${n}`);
  addText(slide, title, p(150,30,780,46), { fontSize: 30, bold: true, color: C.ink, autoFit: "none" }, `slide-title-${n}`);
  if (subtitle) addText(slide, subtitle, p(946,36,290,26), { fontSize: 15, color: C.body, alignment: "right" }, `subtitle-${n}`);
  addLine(slide, 42, 88, 1196, C.line, 1);
}
function panelText(slide, text, position, options = {}, name = "panel-text") {
  return addText(slide, text, position, {
    fontSize: options.fontSize ?? 21,
    bold: options.bold ?? false,
    color: options.color ?? C.ink,
    fill: options.fill ?? C.panel,
    line: options.line ?? "#D3D7D9",
    lineWidth: options.lineWidth ?? 1,
    verticalAlignment: options.verticalAlignment ?? "middle",
    alignment: options.alignment ?? "left",
    insets: options.insets ?? { top: 16, right: 18, bottom: 14, left: 18 },
    autoFit: options.autoFit ?? "shrinkText",
  }, name);
}
function notes(slide, text, sources = []) {
  slide.speakerNotes.textFrame.setText(`${text}\n\n[Sources]\n${sources.length ? sources.map(s => `- ${s}`).join("\n") : "- Original course-report synthesis."}`);
}
function page(slide, n) {
  addText(slide, String(n).padStart(2,"0"), p(1182,670,54,20), { fontSize: 12, color: C.body, alignment: "right", typeface: E }, `page-${n}`);
}
function photoStrip(slide, top, items, options = {}) {
  const left = options.left ?? 62, gap = options.gap ?? 14;
  const totalWidth = options.width ?? 1156;
  const imageHeight = options.imageHeight ?? 92;
  const itemWidth = (totalWidth - gap * (items.length - 1)) / items.length;
  items.forEach((it, i) => {
    const x = left + i * (itemWidth + gap);
    addImage(slide, caseBytes[it.key], p(x, top, itemWidth, imageHeight), it.alt, "cover");
    if (it.label) addText(slide, it.label, p(x, top + imageHeight + 5, itemWidth, 24), { fontSize: 14, bold: true, color: C.dark, alignment: "center" }, `photo-label-${top}-${i}`);
  });
}

const deck = Presentation.create({ slideSize: { width: W, height: H } });

// 1. Cover
{
  const s = deck.slides.add(); addBackground(s);
  addText(s, "PRODUCT FORM SEMANTICS", p(70,148,650,34), { fontSize: 22, bold: true, color: C.blue, typeface: E }, "cover-en");
  addText(s, "产品形态语义设计", p(68,212,690,80), { fontSize: 54, bold: true, color: C.ink, autoFit: "none" }, "cover-title");
  addText(s, "四类日常产品案例分析", p(72,312,570,42), { fontSize: 29, color: C.dark, autoFit: "none" }, "cover-subtitle");
  addLine(s, 72, 382, 520, C.blue2, 2);
  addText(s, "课程汇报    汇报人：________", p(72,500,500,32), { fontSize: 18, color: C.body }, "cover-info");
  addImage(s, buttonBytes, p(760,55,510,640), "替代模板大轴承的工业大按钮", "contain");
  notes(s, "封面沿用机械行业模板的金属背景和大型机械部件构图。", ["Background extracted from user-provided 机械行业.pptx.", "OpenAI-generated hero-button.png, 2026-09-06."]);
}

// 2. Definition
{
  const s = deck.slides.add(); addBackground(s); sectionHeader(s,1,"形态语义是什么","基本概念");
  addText(s,"FORM\nSEMANTICS",p(62,145,400,138),{fontSize:52,bold:true,color:"#C4C9CC",typeface:E,autoFit:"none"},"watermark");
  addText(s,"产品的外观和结构会传出使用信息。用户看到轮廓、接触位置和材质后，会判断哪里能操作、应该怎样操作，以及是否存在风险。",p(470,136,690,124),{fontSize:25,bold:true,color:C.ink},"definition");
  const items=[
    ["轮廓与尺度","决定部件是否显得可按、可握或可转"],
    ["接触方式","平面、凹面、纹理对应不同手部动作"],
    ["位置与方向","建立控制动作与产品结果之间的关系"],
    ["色彩与状态","提示工作状态，并突出安全信息"]
  ];
  items.forEach((it,i)=>{
    const x=70+(i%2)*585, y=330+Math.floor(i/2)*138;
    addText(s,`0${i+1}`,p(x,y,52,30),{fontSize:18,bold:true,color:C.blue,typeface:E},`def-num-${i}`);
    addText(s,it[0],p(x+70,y-2,460,32),{fontSize:24,bold:true},`def-head-${i}`);
    addText(s,it[1],p(x+70,y+42,470,54),{fontSize:18,color:C.body},`def-body-${i}`);
  });
  photoStrip(s,535,[
    {key:"door",alt:"门推板与拉手",label:"门五金"},
    {key:"kettle",alt:"电热水壶把手与壶嘴",label:"电热水壶"},
    {key:"knob",alt:"带刻度的旋钮",label:"控制旋钮"},
    {key:"emergency",alt:"红色急停按钮",label:"急停按钮"}
  ],{imageHeight:82});
  page(s,2); notes(s,"说明形态语义的范围，后面四个案例都用这一概念分析。", ["Nielsen Norman Group: https://www.nngroup.com/articles/clickable-elements/"]);
}

// 3. Framework
{
  const s=deck.slides.add(); addBackground(s); sectionHeader(s,2,"分析框架","四个问题");
  const qs=[
    ["01","能否操作","部件是否容易被发现和触达"],
    ["02","动作是什么","推、拉、握、转或按压"],
    ["03","结果怎样变化","动作方向和状态变化是否对应"],
    ["04","风险是否清楚","高风险操作是否有醒目的提示"]
  ];
  qs.forEach((q,i)=>{
    const y=145+i*126;
    addText(s,q[0],p(82,y,74,56),{fontSize:31,bold:true,color:i===3?C.red:C.blue,typeface:E,alignment:"center",verticalAlignment:"middle",fill:"#E1E5E7",line:"#C7CCCF",insets:{top:0,right:0,bottom:0,left:0}},`q-num-${i}`);
    addText(s,q[1],p(190,y+4,290,40),{fontSize:25,bold:true},`q-head-${i}`);
    addText(s,q[2],p(480,y+6,500,50),{fontSize:19,color:C.body},`q-body-${i}`);
    addImage(s,caseBytes[["door","kettle","knob","emergency"][i]],p(1015,y-10,160,84),q[1],"cover");
    addLine(s,190,y+76,990,"#CFD3D5",1);
  });
  addText(s,"汇报时可直接对着照片回答四个问题：先说看到了什么，再说明它为什么会让人产生某种操作判断。",p(190,640,900,28),{fontSize:16,color:C.body},"framework-tip");
  page(s,3); notes(s,"四个问题用于统一比较不同产品，不依赖具体品类。", ["Nielsen Norman Group: https://www.nngroup.com/articles/clickable-elements/"]);
}

// 4. Categories
{
  const s=deck.slides.add(); addBackground(s); sectionHeader(s,3,"四类产品","案例范围");
  const rows=[
    ["01","门五金","推与拉","接触面积和抓握空间提示开门动作"],
    ["02","电热水壶","握持与倾倒","把手、壶嘴和重心共同限定动作顺序"],
    ["03","控制旋钮","旋转调节","滚花、指针和刻度表示连续变化"],
    ["04","急停按钮","紧急停止","尺寸、颜色和防护结构提示风险等级"]
  ];
  rows.forEach((r,i)=>{
    const y=125+i*132;
    const key=["door","kettle","knob","emergency"][i];
    addImage(s,caseBytes[key],p(62,y,210,104),`${r[1]}案例照片`,"cover");
    addText(s,r[0],p(300,y,58,28),{fontSize:18,bold:true,color:C.blue,typeface:E},`cat-num-${i}`);
    addText(s,r[1],p(365,y-2,210,36),{fontSize:24,bold:true},`cat-product-${i}`);
    addText(s,r[2],p(610,y,185,32),{fontSize:20,bold:true,color:C.dark},`cat-action-${i}`);
    addText(s,r[3],p(365,y+48,790,48),{fontSize:18,color:C.body},`cat-desc-${i}`);
    addLine(s,300,y+112,870,"#C9CED0",1);
  });
  page(s,4); notes(s,"四类按操作信息分类，保留原汇报的四个产品案例。", []);
}

function caseSlide({slideNo, sectionNo, title, key, alt, imageLeft, points, source}) {
  const s=deck.slides.add(); addBackground(s); sectionHeader(s,sectionNo,title,"案例分析");
  const imagePos=imageLeft?p(58,126,560,322):p(662,126,560,322);
  const textX=imageLeft?660:58;
  addImage(s,caseBytes[key],imagePos,alt,"cover");
  const photoX=imageLeft?58:662;
  addImage(s,caseBytes[key],p(photoX,466,270,154),`${alt}的接触区域特写`,"cover");
  addImage(s,caseBytes[key],p(photoX+290,466,270,154),`${alt}的方向和状态特写`,"cover");
  addText(s,"接触区域",p(photoX+10,586,110,24),{fontSize:14,bold:true,color:C.white,fill:"#51575B",insets:{top:3,right:6,bottom:2,left:6}},`detail-a-${slideNo}`);
  addText(s,"方向与状态",p(photoX+300,586,120,24),{fontSize:14,bold:true,color:C.white,fill:"#51575B",insets:{top:3,right:6,bottom:2,left:6}},`detail-b-${slideNo}`);
  addText(s,`0${sectionNo-3}  观察重点`,p(textX,136,180,32),{fontSize:19,bold:true,color:key==="emergency"?C.red:C.blue,typeface:E},`case-index-${slideNo}`);
  points.forEach((pt,i)=>{
    const y=190+i*142;
    addText(s,pt[0],p(textX,y,510,34),{fontSize:22,bold:true,color:C.ink},`case-head-${slideNo}-${i}`);
    addText(s,pt[1],p(textX,y+42,520,74),{fontSize:18,color:C.body},`case-body-${slideNo}-${i}`);
  });
  page(s,slideNo); notes(s,"按形态细节、传递信息和用户动作进行分析。", [`Retained case image from prior deck: ${key}.png.`, source]);
}

caseSlide({slideNo:5,sectionNo:4,title:"推与拉：门五金",key:"door",alt:"玻璃门上的推板和拉手",imageLeft:false,points:[
  ["平板提示向前推","表面宽而平，没有抓握空间，适合用手掌施力。"],
  ["拉手提示向后拉","立体拉手为手指留出空间，抓握后自然向身体方向用力。"],
  ["容易出现的误用","如果两侧都安装相同拉手，用户会先拉门，动作与真实开启方向可能冲突。"]
],source:"Nielsen Norman Group: https://www.nngroup.com/articles/clickable-elements/"});

caseSlide({slideNo:6,sectionNo:5,title:"握持与倾倒：电热水壶",key:"kettle",alt:"突出把手和壶嘴的电热水壶",imageLeft:true,points:[
  ["把手规定握持位置","弧形空间容纳手指，隔热材料说明这里可以安全接触。"],
  ["壶嘴说明出水方向","壶嘴朝向与壶身重心一起提示倾倒方向。"],
  ["细节关系","把手与壶嘴保持距离，手腕转动时更容易控制水流，也能减少接近蒸汽的位置。"]
],source:"Original visual analysis of a generic everyday product."});

caseSlide({slideNo:7,sectionNo:6,title:"旋转调节：控制旋钮",key:"knob",alt:"带滚花、指针和刻度的控制旋钮",imageLeft:false,points:[
  ["圆形和滚花提示旋转","圆周轮廓允许连续转动，边缘纹理增加抓握摩擦。"],
  ["指针和刻度建立映射","指针位置对应刻度变化，用户能够判断增减方向。"],
  ["反馈帮助确认结果","刻度提供视觉反馈，转动阻尼或卡点提供触觉反馈，两者共同减少调错。"]
],source:"Original visual analysis of a generic rotary control."});

caseSlide({slideNo:8,sectionNo:7,title:"紧急停止：急停按钮",key:"emergency",alt:"红色蘑菇头急停按钮",imageLeft:true,points:[
  ["大尺寸便于快速按下","蘑菇头突出且接触面积大，紧急情况下更容易命中。"],
  ["红黄配色突出风险","高对比颜色显示重要性，防护结构同时减少误触。"],
  ["复位方式需要区别","旋转复位或拉起复位应与按下动作明显不同，避免停机后被无意恢复。"]
],source:"ISO 3864-1:2011: https://www.iso.org/standard/51021.html"});

// 9. Native comparison table
{
  const s=deck.slides.add(); addBackground(s); sectionHeader(s,8,"四个案例对比","形态线索与信息");
  const values=[
    ["产品","主要信息","形态线索","对应关系","反馈或约束"],
    ["门五金","推、拉","平板、拉手","施力方式与开门方向","抓握空间"],
    ["电热水壶","握持、倾倒","把手、壶嘴、重心","握姿与出水方向","隔热和防烫"],
    ["控制旋钮","连续调节","圆形、滚花、指针","旋转方向与增减","刻度和阻尼"],
    ["急停按钮","紧急停止","大体量、红黄对比","拍击与停机","防护圈和复位"]
  ];
  photoStrip(s,116,[
    {key:"door",alt:"门五金照片",label:"推 / 拉"},{key:"kettle",alt:"水壶照片",label:"握 / 倾倒"},
    {key:"knob",alt:"旋钮照片",label:"旋转"},{key:"emergency",alt:"急停按钮照片",label:"按压"}
  ],{imageHeight:72});
  const table=s.tables.add({rows:5,columns:5,left:58,top:225,width:1164,height:350,columnWidths:[180,190,285,285,224],values});
  table.borders.assign({style:"solid",fill:"#AEB5B9",width:1});
  table.cells.block({row:0,column:0,rowCount:1,columnCount:5}).assign({fill:C.dark,textStyle:{fontSize:18,bold:true,color:C.white,typeface:F},margins:{top:10,right:10,bottom:8,left:10},anchor:"middle"});
  table.cells.block({row:1,column:0,rowCount:4,columnCount:5}).assign({fill:"#F4F5F5",textStyle:{fontSize:17,color:C.ink,typeface:F},margins:{top:10,right:10,bottom:8,left:10},anchor:"middle"});
  table.cells.block({row:1,column:0,rowCount:4,columnCount:1}).textStyle.bold=true;
  page(s,9); notes(s,"表格对比四类产品的形态线索、映射和约束。", ["IEC/ISO symbols general description: https://webstore.iec.ch/en/iec_catalog/product/preview/?id=L3B1Yi9wZGYvcHJldmlldy9pbmZvX2llYzYwNDE3e2VkMS4wfWIucGRm"]);
}

// 10. Common issues
{
  const s=deck.slides.add(); addBackground(s); sectionHeader(s,9,"常见问题","形态语义失效的原因");
  photoStrip(s,125,[
    {key:"door",alt:"门五金",label:"动作冲突"},{key:"knob",alt:"旋钮",label:"方向相反"},
    {key:"emergency",alt:"急停按钮",label:"风险不明显"},{key:"kettle",alt:"水壶",label:"只靠文字说明"}
  ],{imageHeight:105});
  const issues=[
    ["动作冲突","外形让人想拉，实际结构只能推。门五金最常见。"],
    ["方向相反","旋钮顺时针旋转后数值反而减小，容易造成误判。"],
    ["风险不明显","急停与普通按钮的大小、位置和颜色接近，紧急时难以识别。"],
    ["只靠文字说明","标签被遮挡或光线不足时，用户无法仅凭外形判断功能。"]
  ];
  issues.forEach((it,i)=>{
    const y=292+i*76;
    addText(s,`0${i+1}`,p(72,y,52,28),{fontSize:17,bold:true,color:i===2?C.red:C.blue,typeface:E},`issue-num-${i}`);
    addText(s,it[0],p(140,y-2,210,32),{fontSize:21,bold:true,color:i===2?C.red:C.ink},`issue-head-${i}`);
    addText(s,it[1],p(370,y,780,42),{fontSize:17,color:C.body},`issue-body-${i}`);
  });
  addText(s,"判断标准：去掉文字标签后，外形仍然能够提示真实动作，并能区分普通操作和高风险操作。",p(72,615,1030,34),{fontSize:19,bold:true,color:C.blue},"issue-standard");
  page(s,10); notes(s,"通过常见错误说明形态和操作不一致时会产生误用。", ["Nielsen Norman Group: https://www.nngroup.com/articles/clickable-elements/"]);
}

// 11. Design steps
{
  const s=deck.slides.add(); addBackground(s); sectionHeader(s,10,"设计步骤","从信息到形态");
  const steps=[
    ["01","定义信息","先确定产品需要说明的动作、状态或风险。"],
    ["02","选择线索","用轮廓、位置、材质和颜色表达这些信息。"],
    ["03","制作样机","观察用户第一眼如何理解和操作。"],
    ["04","去掉文字验证","隐藏标签，再检查功能是否仍然清楚。"]
  ];
  steps.forEach((st,i)=>{
    const y=126+i*132;
    addImage(s,caseBytes[["door","kettle","knob","emergency"][i]],p(72,y,174,98),`${st[1]}示例照片`,"cover");
    addText(s,st[0],p(275,y,58,34),{fontSize:21,bold:true,color:C.blue,typeface:E,alignment:"center",verticalAlignment:"middle",fill:"#E1E5E7",line:"#C7CCCF",insets:{top:0,right:0,bottom:0,left:0}},`step-num-${i}`);
    addText(s,st[1],p(358,y-2,220,36),{fontSize:23,bold:true},`step-head-${i}`);
    addText(s,st[2],p(600,y,570,48),{fontSize:18,color:C.body},`step-body-${i}`);
    addText(s,["确定需要用户理解什么","把信息落实到可见、可触的细节","记录第一次接触时的真实反应","确认产品脱离说明书仍可理解"][i],p(358,y+52,800,35),{fontSize:16,color:C.blue},`step-note-${i}`);
    if(i<3) addLine(s,275,y+112,900,"#C9CED0",1);
  });
  page(s,11); notes(s,"四个步骤可用于后续产品方案分析和形态推敲。", ["Nielsen Norman Group: https://www.nngroup.com/articles/clickable-elements/"]);
}

// 12. Conclusion
{
  const s=deck.slides.add(); addBackground(s);
  addText(s,"CONCLUSION",p(72,76,520,34),{fontSize:22,bold:true,color:C.blue,typeface:E},"end-en");
  addText(s,"结论",p(70,132,400,64),{fontSize:48,bold:true,color:C.ink,autoFit:"none"},"end-title");
  const conclusions=[
    "形态首先说明哪里可以操作。",
    "接触方式提示推、拉、握、转等动作。",
    "控制方向应与结果变化保持一致。",
    "高风险信息需要更强的视觉和结构提示。"
  ];
  conclusions.forEach((t,i)=>{
    addText(s,`0${i+1}`,p(78,240+i*78,52,30),{fontSize:18,bold:true,color:C.blue,typeface:E},`end-num-${i}`);
    addText(s,t,p(148,236+i*78,510,40),{fontSize:22,color:C.ink,bold:i===3},`end-text-${i}`);
  });
  addText(s,"产品形态语义设计",p(78,600,500,28),{fontSize:18,color:C.body},"end-course");
  addImage(s,buttonBytes,p(760,72,500,590),"工业大按钮，替代模板大轴承", "contain");
  page(s,12); notes(s,"结尾概括四个案例的共同设计原则。", ["OpenAI-generated hero-button.png, 2026-09-06.", "Nielsen Norman Group: https://www.nngroup.com/articles/clickable-elements/", "ISO 3864-1:2011: https://www.iso.org/standard/51021.html"]);
}

await fs.mkdir(TMP_DIR,{recursive:true});
await fs.mkdir(previewDir,{recursive:true});
for (let i=0;i<deck.slides.items.length;i++) {
  const slide=deck.slides.items[i];
  const png=await deck.export({slide,format:"png",scale:1});
  await fs.writeFile(path.join(previewDir,`slide-${i+1}.png`),new Uint8Array(await png.arrayBuffer()));
  const layout=await slide.export({format:"layout"});
  await fs.writeFile(path.join(previewDir,`slide-${i+1}.layout.json`),await layout.text());
}
const requirements={
  explicitTotalSlideCount:12,
  requiredNativeTableOwnerSlides:[9],
  requiredNativeChartOwnerSlides:[],
};
const fontPolicy={basis:"reference",families:[F],referencePath:templatePath,referenceSha256:templateSha};
const stagingDir=path.join(workspaceDir,".codex-finalizer-mechanical");
await fs.mkdir(stagingDir,{recursive:true});
await fs.mkdir(path.dirname(FINAL_PPTX),{recursive:true});
const candidatePath=path.join(stagingDir,"candidate-mechanical.pptx");
await (await PresentationFile.exportPptx(deck)).save(candidatePath);
const result=await finalizePresentation({
  ...requirements,
  workspaceDir,
  candidatePath,
  finalPath:FINAL_PPTX,
  pythonExecutable:RUNTIME_PYTHON,
  integrityValidatorPath:path.join(SKILL_DIR,"container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath:path.join(SKILL_DIR,"container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs:["--expected-slide-size-emu","12192000,6858000","--validate-heading-fit","--require-native-table-slide","9"],
  requiredNativeTableOwnerSlides:[9],
  fontPolicy,
  verifyArtifactToolImport:true,
  receiptPath:path.join(stagingDir,`${path.basename(FINAL_PPTX)}.validation.json`),
});
console.log(JSON.stringify(result));
