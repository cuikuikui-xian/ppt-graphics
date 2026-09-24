#target photoshop
app.displayDialogs = DialogModes.NO;

var scriptFile = new File($.fileName);
var folder = scriptFile.parent;
var names = [
  "01_环境背景.png",
  "02_主板与电子元件.png",
  "03_黑色显示承载区.png",
  "04_橙铜超薄环路热管.png",
  "05_蓝色工质区域.png",
  "06_芯片红色热区.png"
];
var layerNames = [
  "环境背景",
  "主板与电子元件",
  "黑色显示/承载区域",
  "橙铜超薄环路热管",
  "蓝色工质区域",
  "芯片红色热区"
];

var target = app.documents.add(1672, 941, 300, "超薄环路热管_语义分层", NewDocumentMode.RGB, DocumentFill.TRANSPARENT);
for (var i = 0; i < names.length; i++) {
  var f = new File(folder.fsName + "/" + names[i]);
  var d = app.open(f);
  d.activeLayer.name = layerNames[i];
  d.activeLayer.duplicate(target, ElementPlacement.PLACEATBEGINNING);
  d.close(SaveOptions.DONOTSAVECHANGES);
}
app.activeDocument = target;
target.activeLayer = target.layers[0];
alert("分层文档已建立。请使用“文件 > 另存为”保存为 PSD。");
