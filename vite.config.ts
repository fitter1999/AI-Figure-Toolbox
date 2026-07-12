import { randomUUID } from "node:crypto";
import { execFile } from "node:child_process";
import { access, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";
import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import PptxGenJS from "pptxgenjs";

export default defineConfig({
  plugins: [react(), illustratorBridge()],
});

const execFileAsync = promisify(execFile);

type PptxTextLayer = {
  text: string;
  left: number;
  top: number;
  width: number;
  height: number;
  fontSize: number;
  color: string;
  maskColor?: string;
  visible: boolean;
};

type PptxAssetLayer = {
  dataUrl: string;
  vectorSvg?: string;
  left: number;
  top: number;
  width: number;
  height: number;
  visible: boolean;
};

type SvgPathItem = {
  markup: string;
  box: Box;
  fill: string;
  complexity: number;
  index?: number;
};

type SvgChunk = {
  svg: string;
  box: Box;
};

type CroppedSvg = {
  svg: string;
  box: Box;
};

type Box = {
  minX: number;
  minY: number;
  maxX: number;
  maxY: number;
};

function illustratorBridge(): Plugin {
  return {
    name: "illustrator-bridge",
    configureServer(server) {
      server.middlewares.use("/api/illustrator/status", async (_request, response) => {
        const illustrator = await findIllustrator();
        sendJson(response, { available: Boolean(illustrator), illustrator });
      });

      server.middlewares.use("/api/illustrator/trace", async (request, response) => {
        if (request.method !== "POST") {
          response.statusCode = 405;
          sendJson(response, { error: "仅支持 POST" });
          return;
        }

        try {
          const payload = JSON.parse(await readBody(request)) as {
            dataUrl?: string;
            name?: string;
            width?: number;
            height?: number;
          };
          if (!payload.dataUrl?.startsWith("data:image/")) throw new Error("缺少图片数据");

          const svg = await traceWithIllustrator(payload.dataUrl, Number(payload.width) || 0, Number(payload.height) || 0);
          sendJson(response, { svg, name: payload.name ?? "illustrator-trace.svg" });
        } catch (error) {
          response.statusCode = 500;
          sendJson(response, { error: error instanceof Error ? error.message : "Illustrator 转换失败" });
        }
      });

      server.middlewares.use("/api/illustrator/layered-pptx", async (request, response) => {
        if (request.method !== "POST") {
          response.statusCode = 405;
          sendJson(response, { error: "仅支持 POST" });
          return;
        }

        try {
          const payload = JSON.parse(await readBody(request)) as {
            dataUrl?: string;
            name?: string;
            width?: number;
            height?: number;
            traceDataUrl?: string;
            fallbackSvg?: string;
            texts?: PptxTextLayer[];
            assets?: PptxAssetLayer[];
            textBackgroundDataUrl?: string;
          };
          if (!payload.dataUrl?.startsWith("data:image/")) throw new Error("缺少图片数据");
          const width = Math.max(1, Number(payload.width) || 16);
          const height = Math.max(1, Number(payload.height) || 9);
          const traceDataUrl = payload.traceDataUrl?.startsWith("data:image/")
            ? payload.traceDataUrl
            : payload.dataUrl;
          const tracedSvg = await traceWithIllustratorOrFallback(traceDataUrl, width, height, payload.fallbackSvg);
          const pptx = await buildLayeredPptx(
            tracedSvg,
            width,
            height,
            payload.texts ?? [],
            payload.dataUrl,
            payload.assets ?? [],
            payload.textBackgroundDataUrl,
          );
          sendBinary(
            response,
            pptx,
            `${stripExtension(payload.name || "illustrator-layered")}-illustrator-layered.pptx`,
          );
        } catch (error) {
          response.statusCode = 500;
          sendJson(response, { error: error instanceof Error ? error.message : "Illustrator 分层 PPTX 生成失败" });
        }
      });
    },
  };
}

async function traceWithIllustratorOrFallback(
  dataUrl: string,
  imageWidth: number,
  imageHeight: number,
  fallbackSvg?: string,
) {
  try {
    return await traceWithIllustrator(dataUrl, imageWidth, imageHeight);
  } catch (error) {
    console.warn("Illustrator trace failed, using fallback SVG.", error);
    if (fallbackSvg?.includes("<svg") && hasVectorContent(fallbackSvg)) {
      return normalizeSvgSize(fallbackSvg, imageWidth, imageHeight);
    }
    throw new Error("Illustrator failed and no vector fallback SVG was available.");
  }
}

async function traceWithIllustrator(dataUrl: string, imageWidth = 0, imageHeight = 0) {
  const illustrator = await findIllustrator();
  if (!illustrator) throw new Error("未找到 Adobe Illustrator 2024");

  const workDir = join(tmpdir(), `image-editor-ai-${randomUUID()}`);
  const inputPath = join(workDir, "source.png");
  const outputPath = join(workDir, "trace.svg");
  const scriptPath = join(workDir, "trace.jsx");

  await mkdir(workDir, { recursive: true });
  try {
    await writeFile(inputPath, Buffer.from(dataUrl.split(",")[1] ?? "", "base64"));
    await writeFile(scriptPath, buildTraceScript(inputPath, outputPath, imageWidth, imageHeight), "ascii");

    await runIllustratorScript(scriptPath);
    const svg = await readFile(outputPath, "utf8");
    if (!svg.includes("<svg")) throw new Error("Illustrator 没有导出有效 SVG");
    return svg;
  } finally {
    rm(workDir, { recursive: true, force: true }).catch(() => {});
  }
}

function buildTraceScript(inputPath: string, outputPath: string, imageWidth: number, imageHeight: number) {
  const input = toIllustratorPath(inputPath);
  const output = toIllustratorPath(outputPath);
  const width = Math.max(1, Math.round(imageWidth));
  const height = Math.max(1, Math.round(imageHeight));
  return `
app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
var doc = app.documents.add(DocumentColorSpace.RGB, ${width}, ${height});
try {
  var placed = doc.placedItems.add();
  placed.file = new File("${input}");
  placed.position = [0, ${height}];
  placed.width = ${width};
  placed.height = ${height};
  var tracing = placed.trace();
  var presets = app.tracingPresetsList;
  if (presets && presets.length > 0) {
    var presetName = "";
    var preferred = ["High Fidelity Photo", "高保真度照片", "High Fidelity", "高保真", "Photo"];
    for (var p = 0; p < preferred.length && presetName === ""; p++) {
      for (var i = 0; i < presets.length; i++) {
        if (String(presets[i]).indexOf(preferred[p]) >= 0) {
          presetName = presets[i];
          break;
        }
      }
    }
    tracing.tracing.tracingOptions.loadFromPreset(presetName || presets[0]);
  }
  try { tracing.tracing.tracingOptions.ignoreWhite = true; } catch (e) {}
  try { tracing.tracing.tracingOptions.tracingColors = 96; } catch (e) {}
  try { tracing.tracing.tracingOptions.pathFidelity = 70; } catch (e) {}
  try { tracing.tracing.tracingOptions.cornerFidelity = 80; } catch (e) {}
  try { tracing.tracing.tracingOptions.noiseFidelity = 2; } catch (e) {}
  app.redraw();
  tracing.tracing.expandTracing();
  app.redraw();
  var options = new ExportOptionsSVG();
  options.embedRasterImages = false;
  options.fontType = SVGFontType.OUTLINEFONT;
  options.coordinatePrecision = 3;
  doc.exportFile(new File("${output}"), ExportType.SVG, options);
} finally {
  doc.close(SaveOptions.DONOTSAVECHANGES);
}
`;
}

async function runIllustratorScript(scriptPath: string) {
  const escaped = scriptPath.replace(/'/g, "''");
  const command = `$app = New-Object -ComObject Illustrator.Application; $app.DoJavaScriptFile('${escaped}')`;
  await execFileAsync("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command], {
    timeout: 120000,
    windowsHide: true,
  });
}

async function buildLayeredPptx(
  textlessSvg: string,
  imageWidth: number,
  imageHeight: number,
  texts: PptxTextLayer[],
  backgroundDataUrl: string,
  assets: PptxAssetLayer[],
  textBackgroundDataUrl?: string,
) {
  const pptx = new PptxGenJS();
  const slideWidth = 13.333;
  const slideHeight = slideWidth * (imageHeight / imageWidth);
  const layoutName = "ILLUSTRATOR_LAYERED";
  pptx.defineLayout({ name: layoutName, width: slideWidth, height: slideHeight });
  pptx.layout = layoutName;
  pptx.author = "钳工的美术箱";

  const slide = pptx.addSlide();
  const editableTexts = texts
    .filter((text) => text.visible && text.text.trim().length > 0)
    .slice(0, 900);
  slide.addImage({ data: backgroundDataUrl, x: 0, y: 0, w: slideWidth, h: slideHeight });

  const scaleX = slideWidth / imageWidth;
  const scaleY = slideHeight / imageHeight;

  const previewSlide = pptx.addSlide();
  previewSlide.addImage({ data: backgroundDataUrl, x: 0, y: 0, w: slideWidth, h: slideHeight });

  const visibleAssets = assets.filter((asset) => asset.visible && asset.dataUrl.startsWith("data:image/"));
  const svgSlide = pptx.addSlide();
  const assetSvgCount = addAssetSvgs(svgSlide, visibleAssets, scaleX, scaleY);
  if (assetSvgCount === 0) {
    addAssetImages(svgSlide, visibleAssets, scaleX, scaleY);
  }
  editableTexts.forEach((text) => addEditableText(svgSlide, text, scaleX, scaleY));
  if (visibleAssets.length > 0) {
    const pngSlide = pptx.addSlide();
    addAssetImages(pngSlide, visibleAssets, scaleX, scaleY);
    editableTexts.forEach((text) => addEditableText(pngSlide, text, scaleX, scaleY));
  }

  return Buffer.from((await pptx.write({ outputType: "nodebuffer" })) as Buffer);
}

function addSvgChunk(
  slide: any,
  chunk: SvgChunk,
  slideWidth: number,
  slideHeight: number,
  imageWidth: number,
  imageHeight: number,
) {
  slide.addImage({
    data: svgToDataUri(chunk.svg),
    x: chunk.box.minX * (slideWidth / imageWidth),
    y: chunk.box.minY * (slideHeight / imageHeight),
    w: Math.max(0.01, boxWidth(chunk.box) * (slideWidth / imageWidth)),
    h: Math.max(0.01, boxHeight(chunk.box) * (slideHeight / imageHeight)),
  });
}

function addCroppedSvgImage(slide: any, cropped: CroppedSvg, scaleX: number, scaleY: number) {
  slide.addImage({
    data: svgToDataUri(cropped.svg),
    x: cropped.box.minX * scaleX,
    y: cropped.box.minY * scaleY,
    w: Math.max(0.01, boxWidth(cropped.box) * scaleX),
    h: Math.max(0.01, boxHeight(cropped.box) * scaleY),
  });
}

function cropSvgToContent(svg: string, imageWidth: number, imageHeight: number): CroppedSvg {
  const bounds = svgContentBox(svg);
  if (!bounds) {
    return {
      svg: normalizeSvgSize(svg, imageWidth, imageHeight),
      box: { minX: 0, minY: 0, maxX: imageWidth, maxY: imageHeight },
    };
  }

  const fullBox = { minX: 0, minY: 0, maxX: imageWidth, maxY: imageHeight };
  const croppedBox = padBox(bounds, fullBox, Math.max(1, Math.min(imageWidth, imageHeight) * 0.003));
  const croppedSvg = normalizeSvgSize(svg, imageWidth, imageHeight)
    .replace(/<svg\b([^>]*)>/i, (match) => {
      let next = match
        .replace(/\swidth="[^"]*"/i, "")
        .replace(/\sheight="[^"]*"/i, "")
        .replace(/\sviewBox="[^"]*"/i, "");
      return next.replace(
        /<svg\b/i,
        `<svg width="${boxWidth(croppedBox)}" height="${boxHeight(croppedBox)}" viewBox="${croppedBox.minX} ${croppedBox.minY} ${boxWidth(croppedBox)} ${boxHeight(croppedBox)}"`,
      );
    });

  return { svg: croppedSvg, box: croppedBox };
}

function svgContentBox(svg: string) {
  const boxes: Box[] = [];
  Array.from(svg.matchAll(/<path\b[^>]*>/gi)).forEach((match) => {
    const box = pathBox(match[0]);
    if (box && countsForSvgCrop(match[0])) boxes.push(box);
  });
  Array.from(svg.matchAll(/<(rect|image)\b[^>]*>/gi)).forEach((match) => {
    const box = boxFromSvgRectLike(match[0]);
    if (box && countsForSvgCrop(match[0])) boxes.push(box);
  });
  Array.from(svg.matchAll(/<(circle|ellipse)\b[^>]*>/gi)).forEach((match) => {
    const box = boxFromSvgEllipse(match[0]);
    if (box && countsForSvgCrop(match[0])) boxes.push(box);
  });
  Array.from(svg.matchAll(/<line\b[^>]*>/gi)).forEach((match) => {
    const box = boxFromSvgLine(match[0]);
    if (box && countsForSvgCrop(match[0])) boxes.push(box);
  });
  Array.from(svg.matchAll(/<(polygon|polyline)\b[^>]*>/gi)).forEach((match) => {
    const box = boxFromSvgPoints(match[0]);
    if (box && countsForSvgCrop(match[0])) boxes.push(box);
  });
  Array.from(svg.matchAll(/<text\b[^>]*>([\s\S]*?)<\/text>/gi)).forEach((match) => {
    const box = boxFromSvgText(match[0], match[1] ?? "");
    if (box && countsForSvgCrop(match[0])) boxes.push(box);
  });

  return boxes.length > 0 ? boxes.reduce(mergeBoxes) : null;
}

function isVisibleSvgElement(markup: string) {
  if (/display\s*:\s*none/i.test(markup) || /visibility\s*:\s*hidden/i.test(markup)) return false;
  if (/\sopacity="0(?:\.0*)?"/i.test(markup) || /opacity\s*:\s*0(?:\.0*)?/i.test(markup)) return false;
  const fill = pathFill(markup);
  const stroke = attrValue(markup, "stroke") ?? styleValue(markup, "stroke");
  return fill !== "NONE" || (Boolean(stroke) && stroke.toLowerCase() !== "none");
}

function countsForSvgCrop(markup: string) {
  if (!isVisibleSvgElement(markup)) return false;
  if (isNearWhiteSvgElement(markup)) return false;
  const strokeWidth = svgStrokeWidth(markup);
  return strokeWidth < 200;
}

function isNearWhiteSvgElement(markup: string) {
  const fill = pathFill(markup);
  const stroke = attrValue(markup, "stroke") ?? styleValue(markup, "stroke") ?? "";
  const fillRgb = svgColorToRgb(fill);
  const strokeRgb = svgColorToRgb(stroke);
  const colors = [fillRgb, strokeRgb].filter((color): color is { r: number; g: number; b: number } => Boolean(color));
  if (colors.length === 0) return false;
  return colors.every((color) => {
    const max = Math.max(color.r, color.g, color.b);
    const min = Math.min(color.r, color.g, color.b);
    const luminance = color.r * 0.299 + color.g * 0.587 + color.b * 0.114;
    return luminance > 232 && max - min < 28;
  });
}

function boxFromSvgRectLike(markup: string) {
  const x = numberAttr(markup, "x") ?? 0;
  const y = numberAttr(markup, "y") ?? 0;
  const width = numberAttr(markup, "width");
  const height = numberAttr(markup, "height");
  if (!Number.isFinite(width) || !Number.isFinite(height) || width === null || height === null) return null;
  return padForStroke({ minX: x, minY: y, maxX: x + width, maxY: y + height }, markup);
}

function boxFromSvgEllipse(markup: string) {
  const tag = markup.match(/^<\s*(circle|ellipse)\b/i)?.[1]?.toLowerCase();
  const cx = numberAttr(markup, "cx") ?? 0;
  const cy = numberAttr(markup, "cy") ?? 0;
  const rx = tag === "circle" ? numberAttr(markup, "r") : numberAttr(markup, "rx");
  const ry = tag === "circle" ? numberAttr(markup, "r") : numberAttr(markup, "ry");
  if (rx === null || ry === null || !Number.isFinite(rx) || !Number.isFinite(ry)) return null;
  return padForStroke({ minX: cx - rx, minY: cy - ry, maxX: cx + rx, maxY: cy + ry }, markup);
}

function boxFromSvgLine(markup: string) {
  const x1 = numberAttr(markup, "x1") ?? 0;
  const y1 = numberAttr(markup, "y1") ?? 0;
  const x2 = numberAttr(markup, "x2") ?? 0;
  const y2 = numberAttr(markup, "y2") ?? 0;
  return padForStroke({ minX: Math.min(x1, x2), minY: Math.min(y1, y2), maxX: Math.max(x1, x2), maxY: Math.max(y1, y2) }, markup);
}

function boxFromSvgPoints(markup: string) {
  const points = attrValue(markup, "points");
  if (!points) return null;
  const values = Array.from(points.matchAll(/[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?/g), (match) => Number(match[0]));
  if (values.length < 2) return null;
  let box: Box | null = null;
  for (let index = 0; index < values.length - 1; index += 2) {
    const point = { minX: values[index], minY: values[index + 1], maxX: values[index], maxY: values[index + 1] };
    box = box ? mergeBoxes(box, point) : point;
  }
  return box ? padForStroke(box, markup) : null;
}

function boxFromSvgText(markup: string, text: string) {
  const x = numberAttr(markup, "x") ?? 0;
  const y = numberAttr(markup, "y") ?? 0;
  const styleFontSize = Number(styleValue(markup, "font-size")?.replace(/[^\d.]/g, ""));
  const fontSize = numberAttr(markup, "font-size") ?? (Number.isFinite(styleFontSize) && styleFontSize > 0 ? styleFontSize : 12);
  const width = numberAttr(markup, "data-width") ?? Math.max(1, text.replace(/<[^>]+>/g, "").length * fontSize * 0.6);
  return { minX: x, minY: y - fontSize, maxX: x + width, maxY: y + fontSize * 0.25 };
}

function padForStroke(box: Box, markup: string) {
  const styleStrokeWidth = Number(styleValue(markup, "stroke-width")?.replace(/[^\d.]/g, ""));
  const strokeWidth = numberAttr(markup, "stroke-width") ?? (Number.isFinite(styleStrokeWidth) ? styleStrokeWidth : 0);
  return padFreeBox(box, Math.max(0, strokeWidth / 2));
}

function attrValue(markup: string, name: string) {
  const pattern = new RegExp(`\\s${name}="([^"]*)"`, "i");
  return markup.match(pattern)?.[1] ?? null;
}

function styleValue(markup: string, name: string) {
  const style = attrValue(markup, "style");
  if (!style) return null;
  const pattern = new RegExp(`${name}\\s*:\\s*([^;]+)`, "i");
  return style.match(pattern)?.[1]?.trim() ?? null;
}

function numberAttr(markup: string, name: string) {
  const value = attrValue(markup, name);
  if (value === null) return null;
  const number = Number(value.replace(/[a-z%]+$/i, ""));
  return Number.isFinite(number) ? number : null;
}

function addAssetImages(slide: any, assets: PptxAssetLayer[], scaleX: number, scaleY: number) {
  assets
    .filter((asset) => asset.visible && asset.dataUrl.startsWith("data:image/"))
    .slice(0, 240)
    .forEach((asset) => {
      slide.addImage({
        data: asset.dataUrl,
        x: asset.left * scaleX,
        y: asset.top * scaleY,
        w: Math.max(0.01, asset.width * scaleX),
        h: Math.max(0.01, asset.height * scaleY),
      });
    });
}

function addAssetSvgs(slide: any, assets: PptxAssetLayer[], scaleX: number, scaleY: number) {
  let count = 0;
  assets
    .filter((asset) => asset.visible && asset.vectorSvg?.includes("<svg"))
    .slice(0, 240)
    .forEach((asset) => {
      slide.addImage({
        data: svgToDataUri(normalizeSvgSize(asset.vectorSvg || "", asset.width, asset.height)),
        x: asset.left * scaleX,
        y: asset.top * scaleY,
        w: Math.max(0.01, asset.width * scaleX),
        h: Math.max(0.01, asset.height * scaleY),
      });
      count += 1;
    });
  return count;
}

function addEditableText(slide: any, text: PptxTextLayer, scaleX: number, scaleY: number) {
  slide.addText(text.text, {
    x: text.left * scaleX,
    y: text.top * scaleY,
    w: Math.max(text.width, text.text.length * text.fontSize * 0.5) * scaleX,
    h: Math.max(text.height, text.fontSize * 1.2) * scaleY,
    margin: 0,
    fontFace: "Times New Roman",
    fontSize: Math.max(5, text.fontSize * scaleY * 72),
    color: normalizePptColor(text.color),
    fill: { color: "FFFFFF", transparency: 100 },
    line: { color: "FFFFFF", transparency: 100 },
    breakLine: false,
    fit: "shrink",
  });
}

function assetBoxKey(box: Box) {
  return [box.minX, box.minY, box.maxX, box.maxY].map((value) => Math.round(value * 100) / 100).join(",");
}

function addTextMask(slide: any, text: PptxTextLayer, scaleX: number, scaleY: number) {
  const padding = Math.max(1, text.fontSize * 0.12);
  slide.addShape("rect", {
    x: Math.max(0, (text.left - padding) * scaleX),
    y: Math.max(0, (text.top - padding) * scaleY),
    w: Math.max(0.01, (text.width + padding * 2) * scaleX),
    h: Math.max(0.01, (text.height + padding * 2) * scaleY),
    fill: { color: normalizePptColor(text.maskColor || "FFFFFF"), transparency: 0 },
    line: { color: normalizePptColor(text.maskColor || "FFFFFF"), transparency: 100 },
  });
}

function assetClusterSvgChunks(
  svg: string,
  imageWidth: number,
  imageHeight: number,
  assets: PptxAssetLayer[],
  excludedBoxes: Box[] = [],
) {
  const fullBox = { minX: 0, minY: 0, maxX: imageWidth, maxY: imageHeight };
  const assetBoxes = assets
    .filter((asset) => asset.visible && asset.width > 8 && asset.height > 8)
    .slice(0, 80)
    .map((asset) =>
      padBox(
        {
          minX: asset.left,
          minY: asset.top,
          maxX: asset.left + asset.width,
          maxY: asset.top + asset.height,
        },
        fullBox,
        Math.max(2, Math.min(asset.width, asset.height) * 0.035),
      ),
    );
  const parsed = parseVisibleSvgPaths(svg, excludedBoxes, assetBoxes);
  if (parsed.paths.length === 0) return [];

  const extras = svgExtras(svg);

  const groups =
    assetBoxes.length > 0
      ? groupsFromAssetBoxes(parsed.paths, assetBoxes)
      : mergeVideoPathGroups(parsed.paths, parsed.fullBox).flatMap((group) => splitOversizedVideoGroup(group, parsed.fullBox, 0));

  return groupsToLimitedChunks(groups, parsed.fullBox, extras, 80, 1_500_000);
}

function parseVisibleSvgPaths(svg: string, excludedBoxes: Box[] = [], targetBoxes: Box[] = []) {
  const fullBox = svgViewBox(svg) ?? { minX: 0, minY: 0, maxX: 1, maxY: 1 };
  const fullArea = Math.max(1, boxArea(fullBox));
  const paths: SvgPathItem[] = [];
  let index = 0;
  const maxPaths = targetBoxes.length > 0 ? 24000 : 18000;

  for (const match of svg.matchAll(/<path\b[^>]*>/gi)) {
    const box = pathBox(match[0]);
    if (!box) continue;
    const path = { markup: match[0], box, fill: pathFill(match[0]), complexity: pathComplexity(match[0]), index };
    index += 1;
    const area = boxArea(path.box);
    const inTarget = targetBoxes.length === 0 || targetBoxes.some((target) => overlapArea(path.box, target) > 0);
    if (
      inTarget &&
      area > 0.05 &&
      area < fullArea * 0.92 &&
      !isExcludedTextPath(path.box, excludedBoxes) &&
      !isVideoWhitePath(path) &&
      !isLargeLightPath(path, fullArea) &&
      (targetBoxes.length > 0 || !isLargeFilledBlock(path, fullBox))
    ) {
      paths.push(path);
      if (paths.length >= maxPaths) break;
    }
  }

  return { paths, fullBox };
}

function svgViewBox(svg: string) {
  const match = svg.match(/viewBox="([^"]+)"/i);
  if (!match) return null;
  const values = match[1].split(/[\s,]+/).map(Number).filter(Number.isFinite);
  if (values.length < 4 || values[2] <= 0 || values[3] <= 0) return null;
  return { minX: values[0], minY: values[1], maxX: values[0] + values[2], maxY: values[1] + values[3] };
}

function svgExtras(svg: string) {
  const defs = Array.from(svg.matchAll(/<defs\b[\s\S]*?<\/defs>/gi), (match) => match[0]).join("");
  const styles = Array.from(svg.matchAll(/<style\b[\s\S]*?<\/style>/gi), (match) => match[0]).join("");
  return [defs, styles];
}

function groupsFromAssetBoxes(paths: SvgPathItem[], boxes: Box[]) {
  const used = new Set<number>();
  const groups: SvgPathItem[][] = [];

  boxes.forEach((box) => {
    const group = paths.filter((path) => {
      if (used.has(pathOrderIndex(path))) return false;
      const area = Math.max(1, boxArea(path.box));
      const overlap = overlapArea(path.box, box);
      return centerInside(path.box, box) || overlap / area > 0.45;
    });
    if (group.length === 0) return;
    group.forEach((path) => used.add(pathOrderIndex(path)));
    groups.push(group);
  });

  const rest = paths.filter((path) => !used.has(pathOrderIndex(path)));
  if (rest.length > 0) {
    groups.push(...makePathGroups(rest, groupBounds(rest)).flatMap((group) => splitOversizedVideoGroup(group, groupBounds(rest), 0)));
  }

  return groups;
}

function groupsToLimitedChunks(
  groups: SvgPathItem[][],
  fullBox: Box,
  extras: string[],
  maxChunks: number,
  maxChunkChars: number,
) {
  const chunks: SvgChunk[] = [];
  const ordered = groups
    .filter((group) => group.length > 0)
    .sort((a, b) => boxArea(groupBounds(b)) - boxArea(groupBounds(a)));

  for (const group of ordered) {
    if (chunks.length >= maxChunks) break;
    for (const part of splitGroupBySize(group, maxChunkChars)) {
      if (chunks.length >= maxChunks) break;
      const box = padBox(groupBounds(part), fullBox, 1);
      chunks.push({
        svg: wrapSvg(part.sort((a, b) => pathOrderIndex(a) - pathOrderIndex(b)).map((path) => path.markup).join("\n"), box, extras),
        box,
      });
    }
  }

  return chunks.sort((a, b) => a.box.minY - b.box.minY || a.box.minX - b.box.minX);
}

function splitGroupBySize(group: SvgPathItem[], maxChars: number) {
  const result: SvgPathItem[][] = [];
  let current: SvgPathItem[] = [];
  let currentSize = 0;

  group
    .sort((a, b) => pathOrderIndex(a) - pathOrderIndex(b))
    .forEach((path) => {
      if (current.length > 0 && currentSize + path.markup.length > maxChars) {
        result.push(current);
        current = [];
        currentSize = 0;
      }
      current.push(path);
      currentSize += path.markup.length;
    });

  if (current.length > 0) result.push(current);
  return result;
}

function clusterSvgPaths(svg: string, excludedBoxes: Box[] = []) {
  const parsed = Array.from(svg.matchAll(/<path\b[^>]*>/gi), (match, index) => {
    const box = pathBox(match[0]);
    return box ? { markup: match[0], box, fill: pathFill(match[0]), complexity: pathComplexity(match[0]), index } : null;
  }).filter((item): item is SvgPathItem => Boolean(item));
  if (parsed.length === 0) return [];

  const fullBox = parsed.reduce((box, path) => mergeBoxes(box, path.box), parsed[0].box);
  const fullArea = boxArea(fullBox);
  const paths = parsed.filter((path) => {
    const area = boxArea(path.box);
    return (
      area > 1 &&
      area < fullArea * 0.92 &&
      !isExcludedTextPath(path.box, excludedBoxes) &&
      !isLargeLightPath(path, fullArea) &&
      !isLargeFilledBlock(path, fullBox)
    );
  });
  if (paths.length === 0) return [{ svg: wrapSvg(svg, fullBox, []), box: fullBox }];

  const defs = Array.from(svg.matchAll(/<defs\b[\s\S]*?<\/defs>/gi), (match) => match[0]).join("");
  const styles = Array.from(svg.matchAll(/<style\b[\s\S]*?<\/style>/gi), (match) => match[0]).join("");

  const groups = makePathGroups(paths, fullBox).flatMap((group) => splitLargeGroup(group, fullBox, 0));

  const usefulGroups = groups
    .filter((group) => isUsefulGroup(group, fullBox))
    .sort((a, b) => boxArea(groupBounds(b)) - boxArea(groupBounds(a)));
  const finalGroups = [
    ...usefulGroups.slice(0, 260),
    ...packDetailGroups(usefulGroups.slice(260), fullBox),
  ];

  return finalGroups
    .sort(readingOrder)
    .map((group) => {
      const box = padBox(groupBounds(group), fullBox, 2);
      return {
        svg: wrapSvg(group.map((path) => path.markup).join("\n"), box, [defs, styles]),
        box,
      };
    });
}

function videoMethodSvgChunks(svg: string, excludedBoxes: Box[] = []) {
  const parsed = Array.from(svg.matchAll(/<path\b[^>]*>/gi), (match, index) => {
    const box = pathBox(match[0]);
    return box ? { markup: match[0], box, fill: pathFill(match[0]), complexity: pathComplexity(match[0]), index } : null;
  }).filter((item): item is SvgPathItem => Boolean(item));
  if (parsed.length === 0) return [];

  const fullBox = parsed.reduce((box, path) => mergeBoxes(box, path.box), parsed[0].box);
  const fullArea = Math.max(1, boxArea(fullBox));
  const defs = Array.from(svg.matchAll(/<defs\b[\s\S]*?<\/defs>/gi), (match) => match[0]).join("");
  const styles = Array.from(svg.matchAll(/<style\b[\s\S]*?<\/style>/gi), (match) => match[0]).join("");

  const paths = parsed
    .filter((path) => {
      const area = boxArea(path.box);
      return (
        area > 0.05 &&
        area < fullArea * 0.92 &&
        !isExcludedTextPath(path.box, excludedBoxes) &&
        !isVideoWhitePath(path)
      );
    })
    .slice(0, 9000);

  const groups = mergeVideoPathGroups(paths, fullBox).flatMap((group) => splitOversizedVideoGroup(group, fullBox, 0));

  return groups
    .filter((group) => group.length > 0)
    .sort((a, b) => groupOrderIndex(a) - groupOrderIndex(b))
    .map((group) => {
      const box = padBox(groupBounds(group), fullBox, 1);
      return {
        svg: wrapSvg(pathsForBox(parsed, box, fullArea).map((path) => path.markup).join("\n"), box, [defs, styles]),
        box,
      };
    });
}

function pathsForBox(paths: SvgPathItem[], box: Box, fullArea: number) {
  const padded = padFreeBox(box, Math.max(2, Math.sqrt(fullArea) * 0.002));
  const chunkArea = Math.max(1, boxArea(padded));
  return paths
    .filter((path) => {
      const area = boxArea(path.box);
      if (area <= 0.05 || area >= fullArea * 0.92 || isChunkBackgroundPath(path, padded)) return false;
      const overlap = overlapArea(path.box, padded);
      if (overlap <= 0) return false;
      return centerInside(path.box, padded) || overlap / Math.min(area, chunkArea) > 0.22;
    })
    .sort((a, b) => pathOrderIndex(a) - pathOrderIndex(b));
}

function isChunkBackgroundPath(path: SvgPathItem, box: Box) {
  const rgb = hexToRgb(path.fill);
  if (!rgb) return false;
  const max = Math.max(rgb.r, rgb.g, rgb.b);
  const min = Math.min(rgb.r, rgb.g, rgb.b);
  const luminance = rgb.r * 0.299 + rgb.g * 0.587 + rgb.b * 0.114;
  return luminance > 224 && max - min < 38 && boxArea(path.box) > boxArea(box) * 0.45 && path.complexity < 80;
}

function centerInside(inner: Box, outer: Box) {
  const x = (inner.minX + inner.maxX) / 2;
  const y = (inner.minY + inner.maxY) / 2;
  return x >= outer.minX && x <= outer.maxX && y >= outer.minY && y <= outer.maxY;
}

function padFreeBox(box: Box, padding: number) {
  return {
    minX: box.minX - padding,
    minY: box.minY - padding,
    maxX: box.maxX + padding,
    maxY: box.maxY + padding,
  };
}

function groupOrderIndex(group: SvgPathItem[]) {
  return Math.min(...group.map(pathOrderIndex));
}

function pathOrderIndex(path: SvgPathItem) {
  return path.index ?? 0;
}

function isVideoWhitePath(path: SvgPathItem) {
  const rgb = hexToRgb(path.fill);
  if (!rgb) return false;
  const max = Math.max(rgb.r, rgb.g, rgb.b);
  const min = Math.min(rgb.r, rgb.g, rgb.b);
  const luminance = rgb.r * 0.299 + rgb.g * 0.587 + rgb.b * 0.114;
  return luminance > 232 && max - min < 28;
}

function mergeVideoPathGroups(paths: SvgPathItem[], fullBox: Box) {
  const groups: SvgPathItem[][] = [];
  const maxGap = Math.max(3, Math.min(boxWidth(fullBox), boxHeight(fullBox)) * 0.005);

  for (const path of paths.sort((a, b) => a.box.minY - b.box.minY || a.box.minX - b.box.minX)) {
    let target = -1;
    let targetGap = Number.POSITIVE_INFINITY;

    for (let index = 0; index < groups.length; index += 1) {
      const groupBox = groupBounds(groups[index]);
      const gap = boxGap(path.box, groupBox);
      if (gap <= maxGap && gap < targetGap && canMergeVideoGroup(path.box, groupBox, fullBox)) {
        target = index;
        targetGap = gap;
      }
    }

    if (target >= 0) {
      groups[target].push(path);
    } else {
      groups.push([path]);
    }
  }

  return groups;
}

function canMergeVideoGroup(pathBox: Box, groupBox: Box, fullBox: Box) {
  const merged = mergeBoxes(pathBox, groupBox);
  const fullArea = Math.max(1, boxArea(fullBox));
  const areaRatio = boxArea(merged) / fullArea;
  const widthRatio = boxWidth(merged) / Math.max(1, boxWidth(fullBox));
  const heightRatio = boxHeight(merged) / Math.max(1, boxHeight(fullBox));
  const longLine = (widthRatio > 0.18 && heightRatio < 0.045) || (heightRatio > 0.18 && widthRatio < 0.045);

  return (areaRatio < 0.025 && widthRatio < 0.28 && heightRatio < 0.3) || (longLine && areaRatio < 0.012);
}

function splitOversizedVideoGroup(group: SvgPathItem[], fullBox: Box, depth: number): SvgPathItem[][] {
  const box = groupBounds(group);
  const tooLarge =
    boxArea(box) > boxArea(fullBox) * 0.045 ||
    boxWidth(box) > boxWidth(fullBox) * 0.38 ||
    boxHeight(box) > boxHeight(fullBox) * 0.42 ||
    group.length > 700;

  if (!tooLarge || depth >= 5 || group.length <= 8) return [group];

  const split = findBestSplit(group);
  if (!split) return [group];
  return [
    ...splitOversizedVideoGroup(split[0], fullBox, depth + 1),
    ...splitOversizedVideoGroup(split[1], fullBox, depth + 1),
  ];
}

function textLayerBox(text: PptxTextLayer): Box {
  const width = Math.max(text.width, text.text.length * text.fontSize * 0.5);
  const height = Math.max(text.height, text.fontSize * 1.2);
  const padding = Math.max(3, text.fontSize * 0.35);
  return {
    minX: text.left - padding,
    minY: text.top - padding,
    maxX: text.left + width + padding,
    maxY: text.top + height + padding,
  };
}

function isExcludedTextPath(box: Box, textBoxes: Box[]) {
  if (textBoxes.length === 0) return false;
  const centerX = (box.minX + box.maxX) / 2;
  const centerY = (box.minY + box.maxY) / 2;
  const area = Math.max(1, boxArea(box));
  const width = boxWidth(box);
  const height = boxHeight(box);
  return textBoxes.some((textBox) => {
    const textHeight = boxHeight(textBox);
    const textWidth = boxWidth(textBox);
    const looksLikeGlyph =
      height <= textHeight * 1.35 &&
      width <= textWidth * 0.8 &&
      area <= boxArea(textBox) * 0.35;
    if (!looksLikeGlyph) return false;
    const centerInside =
      centerX >= textBox.minX &&
      centerX <= textBox.maxX &&
      centerY >= textBox.minY &&
      centerY <= textBox.maxY;
    return centerInside || overlapArea(box, textBox) / area > 0.65;
  });
}

function overlapArea(a: Box, b: Box) {
  const width = Math.max(0, Math.min(a.maxX, b.maxX) - Math.max(a.minX, b.minX));
  const height = Math.max(0, Math.min(a.maxY, b.maxY) - Math.max(a.minY, b.minY));
  return width * height;
}

function packDetailGroups(groups: SvgPathItem[][], fullBox: Box) {
  const paths = groups.flat();
  if (paths.length === 0) return [];

  const columns = 6;
  const rows = 4;
  const tiles = new Map<string, SvgPathItem[]>();
  paths.forEach((path) => {
    const centerX = (path.box.minX + path.box.maxX) / 2;
    const centerY = (path.box.minY + path.box.maxY) / 2;
    const column = Math.max(0, Math.min(columns - 1, Math.floor(((centerX - fullBox.minX) / boxWidth(fullBox)) * columns)));
    const row = Math.max(0, Math.min(rows - 1, Math.floor(((centerY - fullBox.minY) / boxHeight(fullBox)) * rows)));
    const key = `${row}:${column}`;
    tiles.set(key, [...(tiles.get(key) ?? []), path]);
  });

  const detailGroups: SvgPathItem[][] = [];
  tiles.forEach((tilePaths) => {
    const sorted = tilePaths.sort((a, b) => a.box.minY - b.box.minY || a.box.minX - b.box.minX);
    for (let index = 0; index < sorted.length; index += 300) {
      detailGroups.push(sorted.slice(index, index + 300));
    }
  });
  return detailGroups;
}

function makePathGroups(paths: SvgPathItem[], fullBox: Box) {
  const groups: SvgPathItem[][] = [];
  const fullArea = boxArea(fullBox);
  const fullWidth = boxWidth(fullBox);
  const fullHeight = boxHeight(fullBox);
  const maxGap = Math.max(10, Math.min(fullWidth, fullHeight) * 0.014);

  for (const path of paths.sort((a, b) => a.box.minY - b.box.minY || a.box.minX - b.box.minX)) {
    let target = -1;
    let targetGap = Number.POSITIVE_INFINITY;
    for (let index = 0; index < groups.length; index += 1) {
      const groupBox = groupBounds(groups[index]);
      const gap = boxGap(path.box, groupBox);
      const merged = mergeBoxes(path.box, groupBox);
      if (
        gap < targetGap &&
        gap <= maxGap &&
        boxArea(merged) < fullArea * 0.09 &&
        boxWidth(merged) < fullWidth * 0.45 &&
        boxHeight(merged) < fullHeight * 0.45
      ) {
        target = index;
        targetGap = gap;
      }
    }

    if (target >= 0) {
      groups[target].push(path);
    } else {
      groups.push([path]);
    }
  }

  return groups;
}

function splitLargeGroup(group: SvgPathItem[], fullBox: Box, depth: number): SvgPathItem[][] {
  const box = groupBounds(group);
  const fullArea = boxArea(fullBox);
  const tooLarge =
    boxArea(box) > fullArea * 0.07 ||
    boxWidth(box) > boxWidth(fullBox) * 0.4 ||
    boxHeight(box) > boxHeight(fullBox) * 0.45 ||
    group.length > 420;

  if (!tooLarge || depth >= 8 || group.length <= 5) return [group];

  const split = findBestSplit(group);
  if (!split) return [group];

  return [
    ...splitLargeGroup(split[0], fullBox, depth + 1),
    ...splitLargeGroup(split[1], fullBox, depth + 1),
  ];
}

function findBestSplit(group: SvgPathItem[]) {
  const byX = findAxisSplit(group, "x");
  const byY = findAxisSplit(group, "y");
  if (byX && byY) return byX.score > byY.score ? byX.parts : byY.parts;
  return byX?.parts ?? byY?.parts ?? null;
}

function findAxisSplit(group: SvgPathItem[], axis: "x" | "y") {
  const sorted = [...group].sort((a, b) =>
    axis === "x" ? a.box.minX - b.box.minX : a.box.minY - b.box.minY,
  );
  const box = groupBounds(group);
  const axisLength = axis === "x" ? boxWidth(box) : boxHeight(box);
  const minGap = Math.max(4, axisLength * 0.01);
  let best: { gap: number; index: number } | null = null;
  let runningMax = axis === "x" ? sorted[0].box.maxX : sorted[0].box.maxY;

  for (let index = 0; index < sorted.length - 1; index += 1) {
    runningMax = Math.max(runningMax, axis === "x" ? sorted[index].box.maxX : sorted[index].box.maxY);
    const nextMin = axis === "x" ? sorted[index + 1].box.minX : sorted[index + 1].box.minY;
    const gap = nextMin - runningMax;
    if (gap > minGap && (!best || gap > best.gap)) {
      best = { gap, index };
    }
  }

  if (!best && sorted.length >= 12) {
    const middle = Math.floor(sorted.length / 2);
    return {
      score: 0.001,
      parts: [sorted.slice(0, middle), sorted.slice(middle)] as [SvgPathItem[], SvgPathItem[]],
    };
  }
  if (!best) return null;
  const first = sorted.slice(0, best.index + 1);
  const second = sorted.slice(best.index + 1);
  if (first.length < 2 || second.length < 2) return null;
  return { score: best.gap / Math.max(1, axisLength), parts: [first, second] as [SvgPathItem[], SvgPathItem[]] };
}

function isUsefulGroup(group: SvgPathItem[], fullBox: Box) {
  const box = groupBounds(group);
  return group.length > 1 || boxArea(box) > boxArea(fullBox) * 0.00004;
}

function readingOrder(a: SvgPathItem[], b: SvgPathItem[]) {
  const boxA = groupBounds(a);
  const boxB = groupBounds(b);
  return boxA.minY - boxB.minY || boxA.minX - boxB.minX;
}

function isLargeLightPath(path: SvgPathItem, fullArea: number) {
  const area = boxArea(path.box);
  if (area < fullArea * 0.00018) return false;
  const rgb = hexToRgb(path.fill);
  if (!rgb) return false;
  const max = Math.max(rgb.r, rgb.g, rgb.b);
  const min = Math.min(rgb.r, rgb.g, rgb.b);
  const luminance = rgb.r * 0.299 + rgb.g * 0.587 + rgb.b * 0.114;
  return luminance > 185 && max - min < 44;
}

function isLargeFilledBlock(path: SvgPathItem, fullBox: Box) {
  const area = boxArea(path.box);
  const fullArea = boxArea(fullBox);
  if (area < fullArea * 0.008) return false;

  const widthRatio = boxWidth(path.box) / Math.max(1, boxWidth(fullBox));
  const heightRatio = boxHeight(path.box) / Math.max(1, boxHeight(fullBox));
  const longThin = (widthRatio > 0.18 && heightRatio < 0.025) || (heightRatio > 0.18 && widthRatio < 0.025);
  if (longThin) return false;

  const rgb = hexToRgb(path.fill);
  const simplePath = path.complexity < 18 || path.markup.length < 900;
  const largeShape = area > fullArea * 0.009 || widthRatio > 0.16 || heightRatio > 0.2;
  if (!largeShape) return false;
  if (!rgb) return simplePath && area > fullArea * 0.03;

  const max = Math.max(rgb.r, rgb.g, rgb.b);
  const min = Math.min(rgb.r, rgb.g, rgb.b);
  const luminance = rgb.r * 0.299 + rgb.g * 0.587 + rgb.b * 0.114;
  const dark = luminance < 95;
  const highSaturation = max - min > 80;

  return simplePath && (!dark || area > fullArea * 0.028) && (!highSaturation || area > fullArea * 0.024);
}

function pathFill(path: string) {
  const hex =
    path.match(/fill:\s*(#[0-9a-fA-F]{6})/i)?.[1] ??
    path.match(/\sfill="(#[0-9a-fA-F]{6})"/i)?.[1];
  if (hex) return hex.toUpperCase();

  if (/\sfill="none"/i.test(path) || /fill:\s*none/i.test(path)) return "NONE";

  const rgb =
    path.match(/fill:\s*rgb\((\d+),\s*(\d+),\s*(\d+)\)/i) ??
    path.match(/\sfill="rgb\((\d+),\s*(\d+),\s*(\d+)\)"/i);
  if (!rgb) return "";
  return `#${[rgb[1], rgb[2], rgb[3]]
    .map((value) => Number(value).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()}`;
}

function svgColorToRgb(color: string) {
  const clean = color.trim();
  if (!clean || clean.toLowerCase() === "none") return null;
  const hex = hexToRgb(clean);
  if (hex) return hex;
  const rgb = clean.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/i);
  if (!rgb) return null;
  return { r: Number(rgb[1]), g: Number(rgb[2]), b: Number(rgb[3]) };
}

function svgStrokeWidth(markup: string) {
  const styleStrokeWidth = Number(styleValue(markup, "stroke-width")?.replace(/[^\d.]/g, ""));
  const strokeWidth = numberAttr(markup, "stroke-width") ?? (Number.isFinite(styleStrokeWidth) ? styleStrokeWidth : 0);
  return Number.isFinite(strokeWidth) ? strokeWidth : 0;
}

function pathComplexity(path: string) {
  return (path.match(/[a-zA-Z]/g) ?? []).length;
}

function hexToRgb(hex: string) {
  const clean = hex.replace("#", "");
  if (!/^[0-9a-fA-F]{6}$/.test(clean)) return null;
  return {
    r: parseInt(clean.slice(0, 2), 16),
    g: parseInt(clean.slice(2, 4), 16),
    b: parseInt(clean.slice(4, 6), 16),
  };
}

function pathBox(path: string) {
  const d = path.match(/\sd="([\s\S]*?)"/i)?.[1];
  if (!d) return null;
  const tokens = Array.from(d.matchAll(/[a-zA-Z]|[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?/g), (match) => match[0]);
  let index = 0;
  let command = "";
  let currentX = 0;
  let currentY = 0;
  let startX = 0;
  let startY = 0;
  let box: Box | null = null;

  const isCommand = (value: string) => /^[a-zA-Z]$/.test(value);
  const read = () => Number(tokens[index++]);
  const hasNumber = () => index < tokens.length && !isCommand(tokens[index]);
  const addPoint = (x: number, y: number) => {
    if (!Number.isFinite(x) || !Number.isFinite(y)) return;
    box = box ? mergeBoxes(box, { minX: x, minY: y, maxX: x, maxY: y }) : { minX: x, minY: y, maxX: x, maxY: y };
  };

  while (index < tokens.length) {
    if (isCommand(tokens[index])) command = tokens[index++];
    const relative = command === command.toLowerCase();
    switch (command.toLowerCase()) {
      case "m":
        while (hasNumber()) {
          const x = read();
          const y = read();
          currentX = relative ? currentX + x : x;
          currentY = relative ? currentY + y : y;
          startX = currentX;
          startY = currentY;
          addPoint(currentX, currentY);
          command = relative ? "l" : "L";
        }
        break;
      case "l":
      case "t":
        while (hasNumber()) {
          const x = read();
          const y = read();
          currentX = relative ? currentX + x : x;
          currentY = relative ? currentY + y : y;
          addPoint(currentX, currentY);
        }
        break;
      case "h":
        while (hasNumber()) {
          const x = read();
          currentX = relative ? currentX + x : x;
          addPoint(currentX, currentY);
        }
        break;
      case "v":
        while (hasNumber()) {
          const y = read();
          currentY = relative ? currentY + y : y;
          addPoint(currentX, currentY);
        }
        break;
      case "c":
        while (hasNumber()) {
          for (let count = 0; count < 3; count += 1) {
            const x = read();
            const y = read();
            const pointX = relative ? currentX + x : x;
            const pointY = relative ? currentY + y : y;
            addPoint(pointX, pointY);
            if (count === 2) {
              currentX = pointX;
              currentY = pointY;
            }
          }
        }
        break;
      case "s":
      case "q":
        while (hasNumber()) {
          for (let count = 0; count < 2; count += 1) {
            const x = read();
            const y = read();
            const pointX = relative ? currentX + x : x;
            const pointY = relative ? currentY + y : y;
            addPoint(pointX, pointY);
            if (count === 1) {
              currentX = pointX;
              currentY = pointY;
            }
          }
        }
        break;
      case "a":
        while (hasNumber()) {
          read();
          read();
          read();
          read();
          read();
          const x = read();
          const y = read();
          currentX = relative ? currentX + x : x;
          currentY = relative ? currentY + y : y;
          addPoint(currentX, currentY);
        }
        break;
      case "z":
        currentX = startX;
        currentY = startY;
        addPoint(currentX, currentY);
        break;
      default:
        index += 1;
        break;
    }
  }

  return box;
}

function groupBounds(group: SvgPathItem[]) {
  return group.reduce((box, path) => mergeBoxes(box, path.box), group[0].box);
}

function mergeBoxes(a: Box, b: Box) {
  return {
    minX: Math.min(a.minX, b.minX),
    minY: Math.min(a.minY, b.minY),
    maxX: Math.max(a.maxX, b.maxX),
    maxY: Math.max(a.maxY, b.maxY),
  };
}

function padBox(box: Box, outer: Box, padding: number) {
  return {
    minX: Math.max(outer.minX, box.minX - padding),
    minY: Math.max(outer.minY, box.minY - padding),
    maxX: Math.min(outer.maxX, box.maxX + padding),
    maxY: Math.min(outer.maxY, box.maxY + padding),
  };
}

function boxWidth(box: Box) {
  return Math.max(0, box.maxX - box.minX);
}

function boxHeight(box: Box) {
  return Math.max(0, box.maxY - box.minY);
}

function boxArea(box: Box) {
  return boxWidth(box) * boxHeight(box);
}

function boxGap(a: Box, b: Box) {
  const dx = Math.max(0, Math.max(a.minX, b.minX) - Math.min(a.maxX, b.maxX));
  const dy = Math.max(0, Math.max(a.minY, b.minY) - Math.min(a.maxY, b.maxY));
  return Math.hypot(dx, dy);
}

function wrapSvg(inner: string, box: Box, extras: string[]) {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${box.minX} ${box.minY} ${boxWidth(box)} ${boxHeight(box)}">
${extras.join("\n")}
${inner}
</svg>`;
}

function svgToDataUri(svg: string) {
  return `data:image/svg+xml;base64,${Buffer.from(svg, "utf8").toString("base64")}`;
}

function normalizeSvgSize(svg: string, width: number, height: number) {
  const hasViewBox = /viewBox="[^"]+"/i.test(svg);
  return svg.replace(/<svg\b([^>]*)>/i, (match) => {
    let next = match.replace(/\swidth="[^"]*"/i, "").replace(/\sheight="[^"]*"/i, "");
    if (!hasViewBox) next = next.replace(/<svg\b/i, `<svg viewBox="0 0 ${width} ${height}"`);
    return next.replace(/<svg\b/i, `<svg width="${width}" height="${height}"`);
  });
}

function hasVectorContent(svg: string) {
  return /<(path|rect|circle|ellipse|polygon|polyline|line|text)\b/i.test(svg);
}

function normalizePptColor(color: string) {
  const clean = color.replace("#", "").trim();
  return /^[0-9a-fA-F]{6}$/.test(clean) ? clean : "111111";
}

async function findIllustrator() {
  const paths = [
    "D:\\Program Files\\Adobe Illustrator 2024\\Support Files\\Contents\\Windows\\Illustrator.exe",
    "C:\\Program Files\\Adobe\\Adobe Illustrator 2024\\Support Files\\Contents\\Windows\\Illustrator.exe",
  ];

  for (const path of paths) {
    try {
      await access(path);
      return path;
    } catch {
      // Try next common install path.
    }
  }
  return "";
}

function readBody(request: NodeJS.ReadableStream) {
  return new Promise<string>((resolve, reject) => {
    let body = "";
    request.setEncoding("utf8");
    request.on("data", (chunk) => {
      body += chunk;
    });
    request.on("end", () => resolve(body));
    request.on("error", reject);
  });
}

function sendJson(response: { setHeader(name: string, value: string): void; end(value: string): void }, value: unknown) {
  response.setHeader("Content-Type", "application/json; charset=utf-8");
  response.end(JSON.stringify(value));
}

function sendBinary(
  response: { setHeader(name: string, value: string): void; end(value: Buffer): void },
  value: Buffer,
  fileName: string,
) {
  response.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  response.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(fileName)}"`);
  response.end(value);
}

function toIllustratorPath(path: string) {
  return path.replace(/\\/g, "/");
}

function stripExtension(fileName: string) {
  return fileName.replace(/\.[^.]+$/, "");
}
