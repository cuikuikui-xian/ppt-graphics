import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const base = "C:\\Users\\32933\\Documents\\ChatGPT\\PPT制图\\mechanical_rebuild";
const sources = [
  ["template", path.join(base, "template.pptx")],
  ["original", "C:\\Users\\32933\\Documents\\ChatGPT\\PPT制图\\产品形态语义设计_课程汇报.pptx"],
];

for (const [name, sourcePath] of sources) {
  const deck = await PresentationFile.importPptx(await FileBlob.load(sourcePath));
  const snapshot = await deck.inspect({
    kind: "deck,slide,textbox,shape,image,table,chart,notes,layout",
    include: "id,slide,name,title,text,textPreview,textChars,textLines,bbox,bboxUnit,alt,prompt,isPlaceholder,placeholders",
    maxChars: 200000,
  });
  await fs.writeFile(path.join(base, `${name}-inspect.ndjson`), snapshot.ndjson);
  const summary = {
    slideCount: deck.slides.items.length,
    masters: deck.masters.items.map((m, i) => ({
      index: i,
      id: m.id,
      name: m.name,
      placeholderSummary: m.placeholders?.summary?.() ?? null,
      elementCount: m.elements?.items?.length ?? m.shapes?.items?.length ?? null,
    })),
    layouts: deck.layouts.items.map((l, i) => ({
      index: i,
      id: l.id,
      name: l.name,
      parentLayoutId: l.parentLayoutId,
      placeholderSummary: l.placeholders?.summary?.() ?? null,
      elementCount: l.elements?.items?.length ?? l.shapes?.items?.length ?? null,
    })),
  };
  await fs.writeFile(path.join(base, `${name}-structure.json`), JSON.stringify(summary, null, 2));
  for (let i = 0; i < deck.slides.items.length; i++) {
    const slide = deck.slides.items[i];
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(base, `${name}-slide-${i+1}.layout.json`), await layout.text());
  }
}
