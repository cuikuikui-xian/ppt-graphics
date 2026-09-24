import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const W = 1280, H = 720;
const C = {
  white: "#FFFFFF", ink: "#0A0D12", muted: "#596273", rule: "#B8BCC4",
  panel: "#EDEDED", blue: "#3D8DFF", pale: "#EAF5FB", red: "#E53935",
  yellow: "#FFC928", green: "#3AA981"
};
const FONT = "Microsoft YaHei";
const buildDir = path.dirname(fileURLToPath(import.meta.url));
const assets = path.join(buildDir, "assets");
const outDir = path.join(buildDir, "rendered");
const finalPath = path.resolve(buildDir, "..", "产品形态语义设计_课程汇报.pptx");

function pos(left, top, width, height) { return { left, top, width, height }; }

function box(slide, p, fill = C.white, line = "none", radius = 0, name = "box") {
  return slide.shapes.add({
    geometry: radius ? "roundRect" : "rect", name, position: p, fill,
    line: line === "none" ? { style: "solid", fill: "none", width: 0 } : { style: "solid", fill: line, width: 1 },
    ...(radius ? { borderRadius: radius } : {})
  });
}

function textBox(slide, text, p, style = {}, name = "text") {
  const s = slide.shapes.add({
    geometry: "textbox", name, position: p, fill: "none",
    line: { style: "solid", fill: "none", width: 0 }
  });
  s.text = text;
  s.text.style = {
    fontSize: style.fontSize ?? 24,
    typeface: FONT,
    color: style.color ?? C.ink,
    bold: style.bold ?? false,
    alignment: style.alignment ?? "left",
    verticalAlignment: style.verticalAlignment ?? "top",
    autoFit: style.autoFit ?? "shrinkText",
    wrap: "square",
    insets: style.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
    ...(style.italic ? { italic: true } : {})
  };
  return s;
}

function rule(slide, left, top, width, color = C.rule, thickness = 1) {
  return slide.shapes.add({
    geometry: "straightConnector1", name: "rule", position: pos(left, top, width, 0), fill: "none",
    line: { style: "solid", fill: color, width: thickness }
  });
}

function footer(slide, n, section = "产品形态语义设计") {
  textBox(slide, section, pos(42, 664, 420, 20), { fontSize: 14, color: C.muted }, `footer-${n}`);
  textBox(slide, String(n).padStart(2, "0"), pos(1180, 660, 58, 24), { fontSize: 14, alignment: "right", color: C.muted }, `page-${n}`);
}

function title(slide, t, n, kicker = "") {
  if (kicker) textBox(slide, kicker, pos(42, 34, 420, 24), { fontSize: 16, bold: true, color: C.blue }, `kicker-${n}`);
  textBox(slide, t, pos(42, 68, 1196, 70), { fontSize: 48, bold: true, autoFit: "none" }, `title-${n}`);
  rule(slide, 42, 154, 1196, C.rule, 1);
}

function notes(slide, body, sources) {
  slide.speakerNotes.textFrame.setText(`${body}\n\n[Sources]\n${sources.map(s => `- ${s}`).join("\n")}`);
}

async function imageBytes(name) {
  const b = await fs.readFile(path.join(assets, name));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
}

function addImage(slide, bytes, p, alt, crop) {
  return slide.images.add({
    blob: bytes, contentType: "image/png", alt, fit: "cover", position: p,
    geometry: "roundRect", borderRadius: 18, ...(crop ? { crop } : {})
  });
}

function addCaseNumber(slide, num, x, y, color = C.blue) {
  textBox(slide, String(num).padStart(2, "0"), pos(x, y, 90, 44), { fontSize: 22, bold: true, color }, `case-num-${num}`);
}

function addBullet(slide, y, heading, body, x = 42, width = 520, color = C.blue) {
  box(slide, pos(x, y + 5, 10, 10), color, "none", 5, `dot-${heading}`);
  textBox(slide, heading, pos(x + 26, y, width - 26, 30), { fontSize: 24, bold: true }, `head-${heading}`);
  textBox(slide, body, pos(x + 26, y + 38, width - 26, 58), { fontSize: 19, color: C.muted }, `body-${heading}`);
}

const deck = Presentation.create({ slideSize: { width: W, height: H } });

// 1 — cover: sparse stacked-text-flow silhouette
{
  const s = deck.slides.add(); s.background.fill = C.white;
  box(s, pos(42, 42, 8, 596), C.blue, "none", 4, "accent-bar");
  textBox(s, "产品形态语义设计", pos(76, 46, 520, 34), { fontSize: 20, bold: true, color: C.blue }, "course");
  textBox(s, "形态会说话", pos(76, 184, 900, 98), { fontSize: 80, bold: true, autoFit: "none" }, "cover-title-1");
  textBox(s, "4类日常产品中的信息传递", pos(76, 290, 860, 70), { fontSize: 42, color: C.ink, autoFit: "none" }, "cover-title-2");
  textBox(s, "从动作暗示、方向映射到安全警示", pos(76, 472, 650, 38), { fontSize: 24, color: C.muted }, "cover-subtitle");
  textBox(s, "课程汇报  ·  汇报人：________", pos(76, 580, 600, 28), { fontSize: 18, color: C.muted }, "presenter");
  notes(s, "开场：我们经常在读到文字之前，就已经根据产品形态做出了动作判断。", ["Original synthesis; no external visual asset."]);
}

// 2 — thesis
{
  const s = deck.slides.add(); s.background.fill = C.white; title(s, "形态是界面的第一层语言", 2, "核心观点");
  textBox(s, "“看起来怎么用”\n往往先于“说明它怎么用”", pos(80, 214, 720, 190), { fontSize: 56, bold: true, autoFit: "none" }, "thesis");
  box(s, pos(868, 206, 306, 306), C.pale, "none", 153, "semantic-circle");
  textBox(s, "形", pos(868, 264, 306, 86), { fontSize: 70, bold: true, color: C.blue, alignment: "center" }, "form-char");
  textBox(s, "动作暗示\n方向映射\n状态反馈\n风险等级", pos(868, 364, 306, 116), { fontSize: 23, alignment: "center", color: C.ink }, "semantic-list");
  textBox(s, "产品的轮廓、尺寸、位置、材质和颜色，共同构成可感知的“符号”。", pos(80, 520, 760, 60), { fontSize: 22, color: C.muted }, "thesis-caption");
  footer(s, 2);
  notes(s, "强调形态语义不是装饰风格，而是用户在接触产品时读到的第一层信息。", ["Nielsen Norman Group: https://www.nngroup.com/articles/clickable-elements/"]);
}

// 3 — three judgments, timeline silhouette
{
  const s = deck.slides.add(); s.background.fill = C.white; title(s, "用户会先做三个判断", 3, "认知路径");
  rule(s, 104, 340, 1030, C.ink, 2);
  const xs = [112, 500, 888];
  const labels = ["01  能不能操作", "02  应该怎么操作", "03  结果与风险"];
  const bodies = ["可触达、可按压、可握持", "推 / 拉 / 转 / 倾倒", "状态变化、力度边界、安全后果"];
  xs.forEach((x, i) => {
    box(s, pos(x, 327, 26, 26), i === 2 ? C.blue : C.ink, "none", 13, `node-${i}`);
    textBox(s, labels[i], pos(x, 236, 306, 42), { fontSize: 25, bold: true }, `label-${i}`);
    textBox(s, bodies[i], pos(x, 394, 306, 76), { fontSize: 21, color: C.muted }, `body-${i}`);
  });
  textBox(s, "设计目标：让判断连续发生，而不是让用户停下来猜。", pos(104, 540, 940, 46), { fontSize: 30, bold: true, color: C.blue }, "goal");
  footer(s, 3);
  notes(s, "用三个问题建立后续四个案例的统一分析框架。", ["Nielsen Norman Group: https://www.nngroup.com/articles/clickable-elements/"]);
}

// 4 — four-category grid silhouette
{
  const s = deck.slides.add(); s.background.fill = C.white; title(s, "四类产品，传递四种核心信息", 4, "案例地图");
  const items = [
    ["01", "推 / 拉", "门五金", "用接触面积与抓握空间暗示动作"],
    ["02", "握持 / 倾倒", "电热水壶", "用重心、隔热与方向组织动作链"],
    ["03", "旋转 / 调节", "控制旋钮", "用纹理、指针与刻度表达连续变化"],
    ["04", "紧急 / 停止", "急停按钮", "用体量、颜色与防护结构编码风险"]
  ];
  const coords = [[42,214],[656,214],[42,424],[656,424]];
  items.forEach((it, i) => {
    const [x,y]=coords[i];
    textBox(s, it[0], pos(x,y,60,34), {fontSize:18,bold:true,color:C.blue}, `map-num-${i}`);
    textBox(s, it[1], pos(x+76,y-2,250,38), {fontSize:29,bold:true}, `map-title-${i}`);
    textBox(s, it[2], pos(x+76,y+50,180,28), {fontSize:19,bold:true,color:C.muted}, `map-product-${i}`);
    textBox(s, it[3], pos(x+76,y+86,460,52), {fontSize:19,color:C.muted}, `map-body-${i}`);
  });
  footer(s, 4);
  notes(s, "四类不是按行业分，而是按形态需要传递的核心信息来分。", ["Original classification for this course report."]);
}

async function caseSlide({n, kicker, headline, imageName, alt, reverse=false, bullets, takeaway, imageCrop, source}) {
  const s = deck.slides.add(); s.background.fill = C.white;
  title(s, headline, n, kicker);
  const imgP = reverse ? pos(42,190,560,424) : pos(678,190,560,424);
  const tx = reverse ? 656 : 42;
  const bytes = await imageBytes(imageName);
  addImage(s, bytes, imgP, alt, imageCrop);
  addCaseNumber(s, n-4, tx, 198, n===8 ? C.red : C.blue);
  bullets.forEach((b,i)=>addBullet(s, 264+i*112, b[0], b[1], tx, 540, n===8 ? C.red : C.blue));
  box(s, pos(tx, 566, 540, 50), n===8 ? "#FFF0ED" : C.pale, "none", 12, `takeaway-box-${n}`);
  textBox(s, takeaway, pos(tx+18, 578, 504, 26), { fontSize: 18, bold: true, color: n===8 ? C.red : C.blue, alignment: "center" }, `takeaway-${n}`);
  footer(s, n);
  notes(s, "按‘可见细节—传递信息—用户动作’顺序讲解。", [`OpenAI-generated visual asset: ${imageName}.`, source]);
}

await caseSlide({
  n:5, kicker:"案例 01 · 推 / 拉", headline:"门五金：接触方式本身就在给指令",
  imageName:"door.png", alt:"玻璃门上的推板与拉手对比", bullets:[
    ["平板＝施力面", "宽、平、无抓握空间，邀请手掌向前推。"],
    ["拉手＝抓握通道", "立体、可环握、与门面保持距离，邀请向身体方向拉。"]
  ], takeaway:"语义关键：动作机会应与真实开启方式一致", source:"Nielsen Norman Group: https://www.nngroup.com/articles/clickable-elements/"
});

await caseSlide({
  n:6, kicker:"案例 02 · 握持 / 倾倒", headline:"电热水壶：形态把连续动作排成顺序",
  imageName:"kettle.png", alt:"突出把手、壶嘴和顶盖按键的电热水壶", reverse:true, bullets:[
    ["把手决定握姿", "弧形空间容纳手指，隔热材质区分安全接触区。"],
    ["壶嘴决定方向", "尖端朝向给出液体出口，壶身重心支持倾倒轨迹。"]
  ], takeaway:"语义关键：多个细节要共同指向同一动作链", source:"Original visual analysis of a generic everyday product."
});

await caseSlide({
  n:7, kicker:"案例 03 · 旋转 / 调节", headline:"控制旋钮：圆形暗示旋转，刻度表达连续性",
  imageName:"knob.png", alt:"带滚花、指针和渐进刻度的控制旋钮", bullets:[
    ["滚花＝需要扭矩", "边缘纹理增加摩擦，也在视觉上强化可抓握感。"],
    ["指针＋弧线＝映射", "旋转方向与刻度空间对应，用户能预判增减关系。"]
  ], takeaway:"语义关键：控制方向必须与结果变化自然对应", source:"Original visual analysis of a generic rotary control."
});

await caseSlide({
  n:8, kicker:"案例 04 · 紧急 / 停止", headline:"急停按钮：高风险信息必须先于文字被看见",
  imageName:"emergency.png", alt:"红色蘑菇头工业急停按钮与黄色防护结构", reverse:true, bullets:[
    ["蘑菇头＝快速拍击", "大直径、突出表面，让紧急状态下也能迅速命中。"],
    ["红＋黄＝风险编码", "高对比色与防护圈同时强调重要性并减少误触。"]
  ], takeaway:"语义关键：风险越高，信号越要冗余、醒目、可确认", source:"ISO 3864-1:2011: https://www.iso.org/standard/51021.html"
});

// 9 — comparison matrix, flat not card-heavy
{
  const s = deck.slides.add(); s.background.fill = C.white; title(s, "四个案例的共性：细节必须形成合力", 9, "横向比较");
  const x0=42, cols=[0,290,520,770,1000], widths=[290,230,250,230,196];
  const headers=["案例","主要信息","形态线索","映射关系","反馈 / 约束"];
  box(s,pos(x0,202,1196,54),C.ink,"none",0,"table-head");
  headers.forEach((h,i)=>textBox(s,h,pos(x0+cols[i]+12,216,widths[i]-20,26),{fontSize:17,bold:true,color:C.white},`th-${i}`));
  const rows=[
    ["门五金","推 / 拉","平板 vs. 拉手","施力方式↔开门方向","物理抓握约束"],
    ["电热水壶","握持 / 倾倒","把手、壶嘴、重心","握姿↔出水方向","隔热与防烫"],
    ["控制旋钮","连续调节","圆形、滚花、指针","旋转↔增减","刻度与阻尼"],
    ["急停按钮","紧急停止","大体量、红黄对比","拍击↔停机","防护圈与复位"]
  ];
  rows.forEach((r,ri)=>{
    const y=256+ri*82;
    if(ri%2===1) box(s,pos(x0,y,1196,82),"#F5F6F8","none",0,`row-bg-${ri}`);
    rule(s,x0,y+81,1196,C.rule,1);
    r.forEach((v,i)=>textBox(s,v,pos(x0+cols[i]+12,y+18,widths[i]-20,48),{fontSize:i===0?20:17,bold:i===0,color:i===0?C.blue:C.ink},`cell-${ri}-${i}`));
  });
  textBox(s,"结论：一个清晰语义通常不是靠单一符号，而是靠轮廓、位置、材质和反馈相互验证。",pos(42,610,1130,36),{fontSize:21,bold:true,color:C.blue},"matrix-conclusion");
  footer(s,9);
  notes(s,"横向比较四个案例，指出有效语义的共同机制是多线索一致。", ["IEC/ISO symbols general description: https://webstore.iec.ch/en/iec_catalog/product/preview/?id=L3B1Yi9wZGYvcHJldmlldy9pbmZvX2llYzYwNDE3e2VkMS4wfWIucGRm", "ISO 3864-1:2011: https://www.iso.org/standard/51021.html"]);
}

// 10 — anti-pattern comparison
{
  const s = deck.slides.add(); s.background.fill = C.white; title(s, "当造型和动作冲突，用户只能靠猜", 10, "反例与修正");
  textBox(s,"冲突的形态",pos(42,206,520,38),{fontSize:29,bold:true,color:C.red},"bad-title");
  textBox(s,"门上安装可抓握的拉手，实际却只能向前推；\n旋钮顺时针旋转，结果却对应数值降低；\n危险控制与普通按键尺寸、颜色完全相同。",pos(42,270,520,190),{fontSize:23,color:C.ink},"bad-copy");
  box(s,pos(616,202,2,360),C.rule,"none",0,"divider");
  textBox(s,"一致的形态",pos(674,206,520,38),{fontSize:29,bold:true,color:C.green},"good-title");
  textBox(s,"让可见形态与真实动作一致；\n让控制方向与结果变化一致；\n让视觉显著性与风险等级一致。",pos(674,270,520,190),{fontSize:23,color:C.ink},"good-copy");
  box(s,pos(42,520,1152,76),C.pale,"none",10,"anti-takeaway");
  textBox(s,"问题不在于用户“不会用”，而在于产品给出了相互矛盾的信息。",pos(70,542,1096,34),{fontSize:25,bold:true,alignment:"center",color:C.blue},"anti-summary");
  footer(s,10);
  notes(s,"反例页用于把‘形态语义’转化为可评价的设计准则。", ["Nielsen Norman Group: https://www.nngroup.com/articles/clickable-elements/"]);
}

// 11 — process timeline
{
  const s = deck.slides.add(); s.background.fill = C.white; title(s, "先定义信息，再设计形态细节", 11, "设计方法");
  rule(s,88,346,1080,C.ink,2);
  const xs=[96,474,852];
  const steps=[
    ["01", "定义意图", "用户需要知道什么？\n动作、方向、状态还是风险？"],
    ["02", "选择线索", "轮廓、尺寸、位置、材质、\n颜色分别承担什么信息？"],
    ["03", "验证一致性", "不看说明能否预判？\n操作结果是否符合预期？"]
  ];
  steps.forEach((st,i)=>{
    box(s,pos(xs[i],333,26,26),i===1?C.blue:C.ink,"none",13,`process-node-${i}`);
    textBox(s,st[0],pos(xs[i],220,80,28),{fontSize:18,bold:true,color:C.blue},`process-num-${i}`);
    textBox(s,st[1],pos(xs[i],258,300,42),{fontSize:30,bold:true},`process-title-${i}`);
    textBox(s,st[2],pos(xs[i],396,306,88),{fontSize:21,color:C.muted},`process-body-${i}`);
  });
  textBox(s,"检验问题：去掉文字标签，产品还会不会“说话”？",pos(96,548,980,42),{fontSize:29,bold:true,color:C.blue},"test-question");
  footer(s,11);
  notes(s,"把案例分析收束为一个可复用的三步设计方法。", ["Nielsen Norman Group: https://www.nngroup.com/articles/clickable-elements/"]);
}

// 12 — close
{
  const s = deck.slides.add(); s.background.fill = C.white;
  textBox(s,"结论",pos(42,42,180,28),{fontSize:18,bold:true,color:C.blue},"close-kicker");
  textBox(s,"好形态不是“更像”，\n而是让人“更少猜”。",pos(42,146,1040,178),{fontSize:66,bold:true,autoFit:"none"},"close-title");
  rule(s,42,374,1196,C.ink,2);
  const points=[
    ["动作可读", "形态先给出可操作性"],
    ["映射自然", "方向与结果彼此对应"],
    ["风险醒目", "重要信息形成多重验证"]
  ];
  points.forEach((p,i)=>{
    const x=42+i*398;
    textBox(s,`0${i+1}`,pos(x,420,50,28),{fontSize:18,bold:true,color:C.blue},`close-num-${i}`);
    textBox(s,p[0],pos(x,462,310,38),{fontSize:28,bold:true},`close-head-${i}`);
    textBox(s,p[1],pos(x,512,330,54),{fontSize:20,color:C.muted},`close-body-${i}`);
  });
  textBox(s,"讨论：你身边还有哪些产品，去掉文字后仍然能让人一眼看懂？",pos(42,624,1120,32),{fontSize:22,bold:true,color:C.blue},"discussion");
  textBox(s,"12",pos(1180,660,58,24),{fontSize:14,alignment:"right",color:C.muted},"page-12");
  notes(s, "结束时回到开场问题，并邀请同学用同一框架寻找新的案例。", ["Original synthesis based on the preceding analysis."]);
}

await fs.mkdir(outDir, { recursive: true });
for (const [i, slide] of deck.slides.items.entries()) {
  const stem = `slide-${String(i+1).padStart(2,"0")}`;
  const png = await deck.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(path.join(outDir, `${stem}.png`), new Uint8Array(await png.arrayBuffer()));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(path.join(outDir, `${stem}.layout.json`), await layout.text());
}
const montage = await deck.export({ format: "webp", montage: true, scale: 1 });
await fs.writeFile(path.join(outDir, "montage.webp"), new Uint8Array(await montage.arrayBuffer()));
const pptx = await PresentationFile.exportPptx(deck);
await pptx.save(finalPath);
console.log(finalPath);
