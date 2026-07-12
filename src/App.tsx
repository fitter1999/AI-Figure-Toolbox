import { InputHTMLAttributes, KeyboardEvent, PointerEvent, useMemo, useRef, useState } from "react";
import {
  CheckSquare,
  Crop,
  Download,
  FileImage,
  FolderOpen,
  ImagePlus,
  MousePointer2,
  Play,
  Square,
  Trash2,
} from "lucide-react";
import PptxGenJS from "pptxgenjs";
import { createWorker } from "tesseract.js";

type Mode = "single" | "multi";
type ToolMode = "select" | "asset" | "text";
type JobStatus = "idle" | "reading" | "ocr" | "ready" | "error";
type ResizeHandle = "nw" | "ne" | "sw" | "se";

type TextLayer = {
  id: string;
  sourceRegionId?: string;
  text: string;
  confidence?: number;
  sourceLeft: number;
  sourceTop: number;
  sourceWidth: number;
  sourceHeight: number;
  left: number;
  top: number;
  width: number;
  height: number;
  fontSize: number;
  color: string;
  maskColor: string;
  visible: boolean;
  dirty: boolean;
  manual?: boolean;
  pending?: boolean;
  maskOnly?: boolean;
};

type AssetLayer = {
  id: string;
  name: string;
  dataUrl: string;
  vectorSvg: string;
  maskColor: string;
  sourceLeft: number;
  sourceTop: number;
  sourceWidth: number;
  sourceHeight: number;
  left: number;
  top: number;
  width: number;
  height: number;
  visible: boolean;
};

type ImageJob = {
  id: string;
  file: File;
  name: string;
  url: string;
  dataUrl: string;
  width: number;
  height: number;
  selected: boolean;
  status: JobStatus;
  progress: number;
  message: string;
  texts: TextLayer[];
  assets: AssetLayer[];
  vectorSvg: string;
};

type DragState =
  | {
      kind: "text";
      textId: string;
      startX: number;
      startY: number;
      originLeft: number;
      originTop: number;
    }
  | {
      kind: "asset";
      assetId: string;
      startX: number;
      startY: number;
      originLeft: number;
      originTop: number;
    }
  | {
      kind: "asset-resize";
      assetId: string;
      handle: ResizeHandle;
      startX: number;
      startY: number;
      originLeft: number;
      originTop: number;
      originWidth: number;
      originHeight: number;
    }
  | {
      kind: "text-resize";
      textId: string;
      handle: ResizeHandle;
      startX: number;
      startY: number;
      originLeft: number;
      originTop: number;
      originWidth: number;
      originHeight: number;
    };

type CropState = {
  startX: number;
  startY: number;
  currentX: number;
  currentY: number;
};

type FolderInputProps = InputHTMLAttributes<HTMLInputElement> & {
  webkitdirectory?: string;
  directory?: string;
};

type IllustratorTraceResponse = {
  svg?: string;
  error?: string;
};

const SAMPLE_DIR = "选择 PNG/JPG 图片，或拖入包含图片的文件夹";
let openCvPromise: Promise<any> | null = null;
let imageTracerPromise: Promise<any> | null = null;

function App() {
  const [mode, setMode] = useState<Mode>("single");
  const [toolMode, setToolMode] = useState<ToolMode>("select");
  const [jobs, setJobs] = useState<ImageJob[]>([]);
  const [activeJobId, setActiveJobId] = useState<string>("");
  const [selectedTextId, setSelectedTextId] = useState<string>("");
  const [selectedAssetId, setSelectedAssetId] = useState<string>("");
  const [dragState, setDragState] = useState<DragState | null>(null);
  const [cropState, setCropState] = useState<CropState | null>(null);
  const [traceColorCount, setTraceColorCount] = useState(20);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);
  const svgRef = useRef<SVGSVGElement>(null);
  const folderInputProps: FolderInputProps = {
    webkitdirectory: "",
    directory: "",
  };

  const activeJob = jobs.find((job) => job.id === activeJobId) ?? jobs[0];
  const selectedText = activeJob?.texts.find((item) => item.id === selectedTextId);
  const selectedAsset = activeJob?.assets.find((item) => item.id === selectedAssetId);
  const selectedJobs = jobs.filter((job) => job.selected);
  const canConvertSelected = selectedJobs.some((job) => job.status === "idle" || job.status === "error");

  const statusText = useMemo(() => {
    if (!activeJob) return "请选择图片";
    if (activeJob.status === "ready") {
      return `已生成全图矢量 SVG，识别 ${activeJob.texts.length} 个文字对象，辅助素材 ${activeJob.assets.length} 个`;
    }
    if (activeJob.status === "ocr") return `${activeJob.message} ${Math.round(activeJob.progress * 100)}%`;
    if (activeJob.status === "error") return activeJob.message;
    return activeJob.message || "等待转换";
  }, [activeJob]);

  async function handleFiles(inputFiles: FileList | null, selected = true) {
    if (!inputFiles) return;

    const imageFiles = Array.from(inputFiles).filter((file) => file.type.startsWith("image/"));
    const nextJobs = await Promise.all(imageFiles.map((file) => createImageJob(file, selected)));

    setJobs((current) => {
      const merged = mode === "single" ? nextJobs.slice(0, 1) : [...current, ...nextJobs];
      if (!activeJobId && merged[0]) setActiveJobId(merged[0].id);
      if (mode === "single" && merged[0]) setActiveJobId(merged[0].id);
      return merged;
    });
    clearSelection();
  }

  async function createImageJob(file: File, selected: boolean): Promise<ImageJob> {
    const dataUrl = await readAsDataUrl(file);
    const size = await readImageSize(dataUrl);
    return {
      id: crypto.randomUUID(),
      file,
      name: file.name,
      url: URL.createObjectURL(file),
      dataUrl,
      width: size.width,
      height: size.height,
      selected,
      status: "idle",
      progress: 0,
      message: "等待转换",
      texts: [],
      assets: [],
      vectorSvg: "",
    };
  }

  async function convertJob(jobId: string): Promise<ImageJob | null> {
    const job = jobs.find((item) => item.id === jobId);
    if (!job) return null;

    updateJob(jobId, { status: "ocr", progress: 0, message: "正在加载 OCR" });

    try {
      const worker = await createWorker("eng", 1, {
        logger: (message) => {
          if (message.status === "recognizing text") {
            updateJob(jobId, {
              status: "ocr",
              progress: message.progress,
              message: "正在识别文字",
            });
          }
        },
      });

      const ocrDataUrl = await buildOcrDataUrl(job.dataUrl);
      const result = await worker.recognize(ocrDataUrl.dataUrl);
      await worker.terminate();

      const textLayers = await buildTextLayers(job.dataUrl, result.data, ocrDataUrl.scale);
      updateJob(jobId, {
        status: "ocr",
        progress: 0.92,
        message: "正在生成全图矢量 SVG",
      });
      const vectorSvg = await vectorizeImage(job);
      updateJob(jobId, {
        status: "ocr",
        progress: 0.96,
        message: "正在用 OpenCV 抠出辅助素材",
      });
      const assets = await autoExtractAssets(job, textLayers.filter(shouldBlockForAssetExtraction));
      const readyJob: ImageJob = {
        ...job,
        status: "ready",
        progress: 1,
        message: "转换完成",
        texts: textLayers,
        assets,
        vectorSvg,
      };
      updateJob(jobId, readyJob);

      if (!activeJobId) setActiveJobId(jobId);
      return readyJob;
    } catch (error) {
      updateJob(jobId, {
        status: "error",
        progress: 0,
        message: error instanceof Error ? error.message : "转换失败",
      });
      return null;
    }
  }

  async function convertSelectedJobs() {
    const targets = selectedJobs.filter((job) => job.status === "idle" || job.status === "error");
    for (const job of targets) {
      await convertJob(job.id);
    }
  }

  async function traceActiveWithIllustrator() {
    if (!activeJob) return;
    const job = activeJob;
    updateJob(job.id, {
      status: "ocr",
      progress: 0.05,
      message: "正在调用 Illustrator 图像描摹",
    });

    try {
      const response = await fetch("/api/illustrator/trace", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ dataUrl: job.dataUrl, name: job.name, width: job.width, height: job.height }),
      });
      const result = (await response.json()) as IllustratorTraceResponse;
      if (!response.ok || !result.svg) throw new Error(result.error || "Illustrator 没有返回 SVG");

      updateJob(job.id, {
        status: "ready",
        progress: 1,
        message: "Illustrator 描摹完成",
        vectorSvg: normalizeSvgSize(result.svg, job.width, job.height),
      });
    } catch (error) {
      updateJob(job.id, {
        status: "error",
        progress: 0,
        message: error instanceof Error ? error.message : "Illustrator 描摹失败",
      });
    }
  }

  async function downloadIllustratorLayeredPptx() {
    if (!activeJob) return;
    let job = activeJob;
    if (job.status !== "ready" || job.texts.length === 0 || job.assets.length === 0) {
      const preparedJob = await convertJob(job.id);
      if (!preparedJob) return;
      job = preparedJob;
    }
    updateJob(job.id, {
      status: "ocr",
      progress: 0.05,
      message: "正在用 Illustrator 生成分层 PPTX",
    });

    try {
      job = await prepareTextRegionsForExport(job);
      updateJob(job.id, { texts: job.texts });
      const exportTexts = job.texts.filter(shouldExportLayeredText);
      const textMasks = job.texts.filter(shouldRemoveFromImage);
      const textBackgroundDataUrl =
        textMasks.length > 0 ? await buildTextReplacementBackgroundDataUrl(job, textMasks) : undefined;
      const traceDataUrl = textBackgroundDataUrl ?? job.dataUrl;
      const fallbackSvg = textBackgroundDataUrl
        ? await vectorizeDataUrl(textBackgroundDataUrl, job.width, job.height, traceColorCount)
        : job.vectorSvg;
      updateJob(job.id, {
        progress: 0.28,
        message: "正在分别矢量化素材",
      });
      const layeredAssets = await prepareLayeredAssets(job, traceColorCount, textBackgroundDataUrl);
      const response = await fetch("/api/illustrator/layered-pptx", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          dataUrl: job.dataUrl,
          name: job.name,
          width: job.width,
          height: job.height,
          traceDataUrl,
          fallbackSvg,
          textBackgroundDataUrl,
          texts: exportTexts
            .map(({ text, left, top, width, height, fontSize, color, maskColor, visible }) => ({
              text,
              left,
              top,
              width,
              height,
              fontSize,
              color,
              maskColor,
              visible,
            })),
          assets: layeredAssets
            .filter((asset) => asset.visible)
            .map(({ dataUrl, vectorSvg, left, top, width, height, visible }) => ({
              dataUrl,
              vectorSvg,
              left,
              top,
              width,
              height,
              visible,
            })),
        }),
      });
      if (!response.ok) {
        const result = (await response.json()) as IllustratorTraceResponse;
        throw new Error(result.error || "Illustrator 分层 PPTX 生成失败");
      }

      const blob = await response.blob();
      downloadBlob(blob, `${stripExtension(job.name)}-illustrator-layered.pptx`);
      updateJob(job.id, {
        status: "ready",
        progress: 1,
        message: "Illustrator 分层 PPTX 已生成",
      });
    } catch (error) {
      updateJob(job.id, {
        status: "error",
        progress: 0,
        message: error instanceof Error ? error.message : "Illustrator 分层 PPTX 生成失败",
      });
    }
  }

  function updateJob(jobId: string, patch: Partial<ImageJob>) {
    setJobs((current) => current.map((job) => (job.id === jobId ? { ...job, ...patch } : job)));
  }

  function updateText(textId: string, patch: Partial<TextLayer>, dirty?: boolean) {
    if (!activeJob) return;
    setJobs((current) =>
      current.map((job) =>
        job.id === activeJob.id
          ? {
              ...job,
              texts: job.texts.map((item) =>
                item.id === textId ? { ...item, ...patch, dirty: dirty ?? item.dirty } : item,
              ),
            }
          : job,
      ),
    );
  }

  function updateAsset(assetId: string, patch: Partial<AssetLayer>) {
    if (!activeJob) return;
    const invalidatedPatch =
      "left" in patch || "top" in patch || "width" in patch || "height" in patch || patch.visible === false
        ? { ...patch, dataUrl: "", vectorSvg: "" }
        : patch;
    setJobs((current) =>
      current.map((job) =>
        job.id === activeJob.id
          ? {
              ...job,
              assets: job.assets.map((item) => (item.id === assetId ? { ...item, ...invalidatedPatch } : item)),
            }
          : job,
      ),
    );
  }

  function removeText() {
    if (!selectedText) return;
    const textId = selectedText.id;
    const sourceRegionId = selectedText.sourceRegionId;
    setJobs((current) =>
      current.map((job) =>
        job.id === activeJob?.id
          ? {
              ...job,
              texts: removeTextAndOrphanMask(job.texts, textId, sourceRegionId),
            }
          : job,
      ),
    );
    clearSelection();
  }

  function ignoreText() {
    if (!selectedText) return;
    updateText(selectedText.id, { visible: false, dirty: false }, false);
    clearSelection();
  }

  function showAllTextMarks() {
    if (!activeJob) return;
    setJobs((current) =>
      current.map((job) =>
        job.id === activeJob.id
          ? { ...job, texts: job.texts.map((text) => (text.maskOnly ? text : { ...text, visible: true })) }
          : job,
      ),
    );
  }

  async function prepareTextRegionsForExport(job: ImageJob): Promise<ImageJob> {
    const regions = job.texts.filter((text) => text.visible && text.pending);
    if (regions.length === 0) return job;

    updateJob(job.id, {
      status: "ocr",
      progress: 0.16,
      message: "正在识别文字区域",
    });

    const image = await loadImage(job.dataUrl);
    const sampler = createImageSampler(image);
    const recognizedTexts: TextLayer[] = [];

    for (const region of regions) {
      const recognized = await recognizeManualText(job, region.left, region.top, region.width, region.height);
      const items =
        recognized.items.length > 0
          ? recognized.items
          : recognized.text.trim().length > 0
            ? [
                {
                  text: recognized.text,
                  confidence: recognized.confidence,
                  bbox: { x0: 0, y0: 0, x1: region.width, y1: region.height },
                },
              ]
            : [];
      recognizedTexts.push(...items.map((item) => textItemToLayer(region, item, sampler)));
    }

    return {
      ...job,
      texts: [
        ...job.texts.map((text) =>
          text.pending ? { ...text, text: "", visible: false, dirty: false, pending: false, maskOnly: true } : text,
        ),
        ...recognizedTexts,
      ],
    };
  }

  function removeTextAndOrphanMask(texts: TextLayer[], textId: string, sourceRegionId?: string) {
    const withoutText = texts.filter((text) => text.id !== textId);
    if (!sourceRegionId) return withoutText;
    const stillUsesMask = withoutText.some((text) => text.sourceRegionId === sourceRegionId && text.visible);
    return stillUsesMask ? withoutText : withoutText.filter((text) => text.id !== sourceRegionId);
  }

  function handleKeyDown(event: KeyboardEvent<HTMLDivElement>) {
    const target = event.target as HTMLElement;
    if (target.tagName === "INPUT" || target.tagName === "TEXTAREA") return;
    if ((event.key === "Delete" || event.key === "Backspace") && selectedText) {
      event.preventDefault();
      removeText();
      return;
    }
    if ((event.key === "Delete" || event.key === "Backspace") && selectedAsset) {
      event.preventDefault();
      removeAsset();
    }
  }

  function removeAsset() {
    if (!selectedAsset) return;
    updateAsset(selectedAsset.id, { visible: false });
    clearSelection();
  }

  function toggleJob(jobId: string) {
    setJobs((current) =>
      current.map((job) => (job.id === jobId ? { ...job, selected: !job.selected } : job)),
    );
  }

  function clearJobs() {
    jobs.forEach((job) => URL.revokeObjectURL(job.url));
    setJobs([]);
    setActiveJobId("");
    clearSelection();
  }

  function clearSelection() {
    setSelectedTextId("");
    setSelectedAssetId("");
  }

  function svgPoint(event: PointerEvent<SVGSVGElement>) {
    if (!activeJob || !svgRef.current) return { x: 0, y: 0 };
    const rect = svgRef.current.getBoundingClientRect();
    return {
      x: ((event.clientX - rect.left) / rect.width) * activeJob.width,
      y: ((event.clientY - rect.top) / rect.height) * activeJob.height,
    };
  }

  function startTextDrag(event: PointerEvent<SVGGElement>, text: TextLayer) {
    if (!activeJob || (toolMode !== "select" && toolMode !== "text")) return;
    event.stopPropagation();
    const point = svgPoint(event as unknown as PointerEvent<SVGSVGElement>);
    setSelectedTextId(text.id);
    setSelectedAssetId("");
    setDragState({
      kind: "text",
      textId: text.id,
      startX: point.x,
      startY: point.y,
      originLeft: text.left,
      originTop: text.top,
    });
  }

  function startTextResize(event: PointerEvent<SVGRectElement>, text: TextLayer, handle: ResizeHandle) {
    if (!activeJob || (toolMode !== "select" && toolMode !== "text")) return;
    event.stopPropagation();
    const point = svgPoint(event as unknown as PointerEvent<SVGSVGElement>);
    setSelectedTextId(text.id);
    setSelectedAssetId("");
    setDragState({
      kind: "text-resize",
      textId: text.id,
      handle,
      startX: point.x,
      startY: point.y,
      originLeft: text.left,
      originTop: text.top,
      originWidth: text.width,
      originHeight: text.height,
    });
  }

  function startAssetDrag(event: PointerEvent<SVGGElement>, asset: AssetLayer) {
    if (!activeJob || (toolMode !== "select" && toolMode !== "asset")) return;
    event.stopPropagation();
    const point = svgPoint(event as unknown as PointerEvent<SVGSVGElement>);
    setSelectedAssetId(asset.id);
    setSelectedTextId("");
    setDragState({
      kind: "asset",
      assetId: asset.id,
      startX: point.x,
      startY: point.y,
      originLeft: asset.left,
      originTop: asset.top,
    });
  }

  function startAssetResize(event: PointerEvent<SVGRectElement>, asset: AssetLayer, handle: ResizeHandle) {
    if (!activeJob || (toolMode !== "select" && toolMode !== "asset")) return;
    event.stopPropagation();
    const point = svgPoint(event as unknown as PointerEvent<SVGSVGElement>);
    setSelectedAssetId(asset.id);
    setSelectedTextId("");
    setDragState({
      kind: "asset-resize",
      assetId: asset.id,
      handle,
      startX: point.x,
      startY: point.y,
      originLeft: asset.left,
      originTop: asset.top,
      originWidth: asset.width,
      originHeight: asset.height,
    });
  }

  function handlePointerDown(event: PointerEvent<SVGSVGElement>) {
    if ((toolMode !== "asset" && toolMode !== "text") || !activeJob) {
      clearSelection();
      return;
    }

    const point = svgPoint(event);
    setCropState({ startX: point.x, startY: point.y, currentX: point.x, currentY: point.y });
    clearSelection();
  }

  function handlePointerMove(event: PointerEvent<SVGSVGElement>) {
    const point = svgPoint(event);

    if (cropState) {
      setCropState({ ...cropState, currentX: point.x, currentY: point.y });
      return;
    }

    if (!dragState) return;
    if (dragState.kind === "text") {
      const left = dragState.originLeft + point.x - dragState.startX;
      const top = dragState.originTop + point.y - dragState.startY;
      updateText(dragState.textId, {
        left,
        top,
        sourceLeft: left,
        sourceTop: top,
      });
    } else if (dragState.kind === "text-resize") {
      updateText(dragState.textId, resizedTextBox(dragState, point.x, point.y));
    } else if (dragState.kind === "asset-resize") {
      updateAsset(dragState.assetId, resizedAssetBox(dragState, point.x, point.y));
    } else {
      const left = dragState.originLeft + point.x - dragState.startX;
      const top = dragState.originTop + point.y - dragState.startY;
      updateAsset(dragState.assetId, {
        left,
        top,
        sourceLeft: left,
        sourceTop: top,
      });
    }
  }

  function resizedTextBox(state: Extract<DragState, { kind: "text-resize" }>, x: number, y: number) {
    const minSize = 6;
    const dx = x - state.startX;
    const dy = y - state.startY;
    let left = state.originLeft;
    let top = state.originTop;
    let width = state.originWidth;
    let height = state.originHeight;

    if (state.handle.includes("e")) width = Math.max(minSize, state.originWidth + dx);
    if (state.handle.includes("s")) height = Math.max(minSize, state.originHeight + dy);
    if (state.handle.includes("w")) {
      const right = state.originLeft + state.originWidth;
      left = Math.min(right - minSize, state.originLeft + dx);
      width = right - left;
    }
    if (state.handle.includes("n")) {
      const bottom = state.originTop + state.originHeight;
      top = Math.min(bottom - minSize, state.originTop + dy);
      height = bottom - top;
    }

    return {
      left,
      top,
      width,
      height,
      sourceLeft: left,
      sourceTop: top,
      sourceWidth: width,
      sourceHeight: height,
    };
  }

  function resizedAssetBox(state: Extract<DragState, { kind: "asset-resize" }>, x: number, y: number) {
    const minSize = 16;
    const dx = x - state.startX;
    const dy = y - state.startY;
    let left = state.originLeft;
    let top = state.originTop;
    let width = state.originWidth;
    let height = state.originHeight;

    if (state.handle.includes("e")) width = Math.max(minSize, state.originWidth + dx);
    if (state.handle.includes("s")) height = Math.max(minSize, state.originHeight + dy);
    if (state.handle.includes("w")) {
      const right = state.originLeft + state.originWidth;
      left = Math.min(right - minSize, state.originLeft + dx);
      width = right - left;
    }
    if (state.handle.includes("n")) {
      const bottom = state.originTop + state.originHeight;
      top = Math.min(bottom - minSize, state.originTop + dy);
      height = bottom - top;
    }

    return {
      left,
      top,
      width,
      height,
      sourceLeft: left,
      sourceTop: top,
      sourceWidth: width,
      sourceHeight: height,
    };
  }

  async function stopPointer() {
    setDragState(null);

    if (!activeJob || !cropState) return;
    const left = Math.max(0, Math.min(cropState.startX, cropState.currentX));
    const top = Math.max(0, Math.min(cropState.startY, cropState.currentY));
    const width = Math.min(activeJob.width - left, Math.abs(cropState.currentX - cropState.startX));
    const height = Math.min(activeJob.height - top, Math.abs(cropState.currentY - cropState.startY));
    setCropState(null);

    if (width < 16 || height < 16) return;
    if (toolMode === "text") {
      const text = await createPendingTextMark(activeJob, left, top, width, height);
      setJobs((current) =>
        current.map((job) => (job.id === activeJob.id ? { ...job, texts: [...job.texts, text] } : job)),
      );
      setSelectedTextId(text.id);
      setSelectedAssetId("");
      return;
    }

    const asset = await cropAsset(activeJob, left, top, width, height);
    setJobs((current) =>
      current.map((job) =>
        job.id === activeJob.id ? { ...job, assets: [...job.assets, asset] } : job,
      ),
    );
    setSelectedAssetId(asset.id);
    setSelectedTextId("");
  }

  async function createPendingTextMark(job: ImageJob, left: number, top: number, width: number, height: number): Promise<TextLayer> {
    const image = await loadImage(job.dataUrl);
    const sampler = createImageSampler(image);
    return {
      id: crypto.randomUUID(),
      text: "",
      confidence: 0,
      sourceLeft: left,
      sourceTop: top,
      sourceWidth: width,
      sourceHeight: height,
      left,
      top,
      width,
      height,
      fontSize: Math.max(8, Math.round(height * 0.68)),
      color: sampler.textColor(left, top, width, height),
      maskColor: sampler.backgroundColor(left, top, width, height),
      visible: true,
      dirty: false,
      manual: true,
      pending: true,
    };
  }

  function selectMode(nextMode: Mode) {
    setMode(nextMode);
    clearSelection();
  }

  function selectTool(nextTool: ToolMode) {
    setToolMode(nextTool);
    clearSelection();
    setCropState(null);
  }

  async function downloadActiveSvg() {
    if (!activeJob || activeJob.status !== "ready") return;
    if (!activeJob.vectorSvg) return;
    downloadBlob(buildVectorSvgBlob(activeJob), `${stripExtension(activeJob.name)}-vector.svg`);
  }

  async function downloadActivePng() {
    if (!activeJob || activeJob.status !== "ready") return;
    const svg = await buildSvgMarkup(activeJob);
    const image = await loadImage(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`);
    const canvas = document.createElement("canvas");
    canvas.width = activeJob.width;
    canvas.height = activeJob.height;
    const context = canvas.getContext("2d");
    if (!context) return;
    context.drawImage(image, 0, 0);
    canvas.toBlob((blob) => {
      if (blob) downloadBlob(blob, `${stripExtension(activeJob.name)}.png`);
    }, "image/png");
  }

  async function downloadActivePptx() {
    if (!activeJob || activeJob.status !== "ready") return;
    await writePptx(activeJob, "visual");
  }

  async function downloadActiveLayeredPptx() {
    if (!activeJob || activeJob.status !== "ready") return;
    await writePptx(activeJob, "layered");
  }

  async function downloadSelectedPptx() {
    for (const job of selectedJobs.filter((item) => item.status === "ready")) {
      await writePptx(job, "visual");
    }
  }

  async function downloadSelectedSvg() {
    for (const job of selectedJobs.filter((item) => item.status === "ready")) {
      downloadBlob(buildVectorSvgBlob(job), `${stripExtension(job.name)}-vector.svg`);
    }
  }

  function downloadSelectedAsset() {
    if (!selectedAsset) return;
    downloadDataUrl(selectedAsset.dataUrl, `${selectedAsset.name}.png`);
  }

  const cropBox = cropState ? normalizeBox(cropState.startX, cropState.startY, cropState.currentX, cropState.currentY) : null;
  const showTextRegions = toolMode !== "asset";
  const showAssetRegions = toolMode !== "text";

  return (
    <div className="app" tabIndex={0} onKeyDown={handleKeyDown}>
      <header className="topbar">
        <div>
          <h1>钳工的美术箱</h1>
          <p>简单图可用 SVG（PPT 第三页），复杂图使用 PNG（PPT 第四页）。</p>
        </div>

        <div className="modeSwitch" aria-label="模式切换">
          <button className={mode === "single" ? "active" : ""} onClick={() => selectMode("single")}>
            单图
          </button>
          <button className={mode === "multi" ? "active" : ""} onClick={() => selectMode("multi")}>
            多图
          </button>
        </div>
      </header>

      <main className="layout">
        <aside className="sidebar">
          <section className="panel">
            <div className="panelTitle">图片选择</div>
            <div className="sourceHint">样例目录：{SAMPLE_DIR}</div>

            <div className="uploadButtons">
              <button onClick={() => fileInputRef.current?.click()}>
                <ImagePlus size={17} />
                选择图片
              </button>
              {mode === "multi" && (
                <button onClick={() => folderInputRef.current?.click()}>
                  <FolderOpen size={17} />
                  选择文件夹
                </button>
              )}
            </div>

            <input
              ref={fileInputRef}
              type="file"
              accept="image/*"
              multiple={mode === "multi"}
              onChange={(event) => handleFiles(event.target.files, true)}
              hidden
            />
            <input
              ref={folderInputRef}
              type="file"
              accept="image/*"
              multiple
              onChange={(event) => handleFiles(event.target.files, false)}
              hidden
              {...folderInputProps}
            />

            {jobs.length > 0 && (
              <div className="queueActions">
                <button onClick={convertSelectedJobs} disabled={!canConvertSelected}>
                  <Play size={16} />
                  转换选中
                </button>
                <button className="ghost" onClick={clearJobs}>
                  清空
                </button>
              </div>
            )}
          </section>

          {jobs.length > 0 && (
            <section className="panel jobList">
              <div className="panelTitle">图片队列</div>
              {jobs.map((job) => (
                <button
                  key={job.id}
                  className={`jobItem ${activeJob?.id === job.id ? "active" : ""}`}
                  onClick={() => {
                    setActiveJobId(job.id);
                    clearSelection();
                  }}
                >
                  <span
                    className="check"
                    onClick={(event) => {
                      event.stopPropagation();
                      toggleJob(job.id);
                    }}
                  >
                    {job.selected ? <CheckSquare size={16} /> : <Square size={16} />}
                  </span>
                  <FileImage size={17} />
                  <span className="jobName">{job.name}</span>
                  <span className={`badge ${job.status}`}>{job.status}</span>
                </button>
              ))}
            </section>
          )}

          <section className="panel">
            <div className="panelTitle">工具</div>
            <div className="toolGrid">
              <button className={toolMode === "select" ? "activeTool" : ""} onClick={() => selectTool("select")}>
                <MousePointer2 size={16} />
                选择编辑
              </button>
              <button className={toolMode === "text" ? "activeTool" : ""} onClick={() => selectTool("text")}>
                <Square size={16} />
                连续文字标记
              </button>
              <button className={toolMode === "asset" ? "activeTool" : ""} onClick={() => selectTool("asset")}>
                <Crop size={16} />
                连续素材标记
              </button>
            </div>
            <button className="wideButton" onClick={showAllTextMarks} disabled={!activeJob}>
              显示全部文字标记
            </button>
            <div className="toolHint">
              {toolMode === "text"
                ? "只显示文字识别区。拖红框标记文字，导出 PPTX 时会重新识别仍保留的红框。"
                : toolMode === "asset"
                  ? "只显示素材识别区。拖蓝框补完整素材，可连续框选多个素材。"
                  : "选择编辑模式下可点选已有文字框或素材框，然后移动、缩放或删除。"}
            </div>
          </section>

          <section className="panel">
            <div className="panelTitle">文字编辑</div>
            {selectedText ? (
              <div className="editorForm">
                {selectedText.pending ? (
                  <div className="emptyState">当前红框只代表识别范围。可拖动、缩放，确认后重新识别文字。</div>
                ) : (
                  <>
                    <label>
                      内容
                      <textarea
                        value={selectedText.text}
                        onChange={(event) => updateText(selectedText.id, { text: event.target.value }, true)}
                      />
                    </label>
                    <div className="fieldRow">
                      <label>
                        字号
                        <input
                          type="number"
                          min="6"
                          max="180"
                          value={Math.round(selectedText.fontSize)}
                          onChange={(event) =>
                            updateText(selectedText.id, { fontSize: Number(event.target.value) || 12 }, true)
                          }
                        />
                      </label>
                      <label>
                        颜色
                        <input
                          type="color"
                          value={selectedText.color}
                          onChange={(event) => updateText(selectedText.id, { color: event.target.value }, true)}
                        />
                      </label>
                    </div>
                  </>
                )}
                <button className="danger" onClick={removeText}>
                  <Trash2 size={16} />
                  {selectedText.pending ? "删除识别框" : "删除文字"}
                </button>
                {!selectedText.pending && (
                  <button className="ghost" onClick={ignoreText}>
                    忽略误识别
                  </button>
                )}
              </div>
            ) : (
              <div className="emptyState">
                <MousePointer2 size={18} />
                在画布上点选文字后编辑。
              </div>
            )}
          </section>

          <section className="panel">
            <div className="panelTitle">素材编辑</div>
            {selectedAsset ? (
              <div className="editorForm">
                <div className="assetPreview">
                  <img src={selectedAsset.dataUrl} alt={selectedAsset.name} />
                </div>
                <button onClick={downloadSelectedAsset}>
                  <Download size={16} />
                  导出素材 PNG
                </button>
                <button className="danger" onClick={removeAsset}>
                  <Trash2 size={16} />
                  删除素材
                </button>
              </div>
            ) : (
              <div className="emptyState">
                <Crop size={18} />
                框选或点选素材后编辑。
              </div>
            )}
          </section>

          <section className="panel">
            <div className="panelTitle">导出</div>
            <div className="fieldRow">
              <label>
                SVG colors
                <input
                  type="number"
                  min="2"
                  max="96"
                  value={traceColorCount}
                  onChange={(event) => setTraceColorCount(clampTraceColorCount(Number(event.target.value)))}
                />
              </label>
            </div>
            <div className="exportGrid">
              <button onClick={downloadActiveSvg} disabled={!activeJob || activeJob.status !== "ready" || !activeJob.vectorSvg}>
                <Download size={16} />
                SVG
              </button>
              <button onClick={downloadActivePng} disabled={!activeJob || activeJob.status !== "ready"}>
                <Download size={16} />
                PNG
              </button>
              <button onClick={downloadActivePptx} disabled={!activeJob || activeJob.status !== "ready"}>
                <Download size={16} />
                PPTX
              </button>
            </div>
            <button
              className="wideButton"
              onClick={downloadActiveLayeredPptx}
              disabled={!activeJob || activeJob.status !== "ready"}
            >
              <Download size={16} />
              实验PPTX分层
            </button>
            <button
              className="wideButton"
              onClick={traceActiveWithIllustrator}
              disabled={!activeJob || activeJob.status === "ocr" || activeJob.status === "reading"}
            >
              <ImagePlus size={16} />
              Illustrator描摹SVG
            </button>
            <button
              className="wideButton"
              onClick={downloadIllustratorLayeredPptx}
              disabled={!activeJob || activeJob.status === "ocr" || activeJob.status === "reading"}
            >
              <Download size={16} />
              Illustrator分层PPTX
            </button>
            {mode === "multi" && (
              <div className="batchExport">
                <button onClick={downloadSelectedSvg}>导出选中 SVG</button>
                <button onClick={downloadSelectedPptx}>导出选中 PPTX</button>
              </div>
            )}
          </section>
        </aside>

        <section className="workspace">
          <div className="statusBar">
            <span>{activeJob?.name ?? "未选择图片"}</span>
            <strong>{statusText}</strong>
          </div>

          <div className="canvasShell">
            {activeJob ? (
              <svg
                ref={svgRef}
                className={`canvas ${toolMode === "asset" || toolMode === "text" ? "cropMode" : ""}`}
                viewBox={`0 0 ${activeJob.width} ${activeJob.height}`}
                onPointerDown={handlePointerDown}
                onPointerMove={handlePointerMove}
                onPointerUp={stopPointer}
                onPointerLeave={stopPointer}
              >
                <image href={activeJob.url} width={activeJob.width} height={activeJob.height} />

                {showTextRegions &&
                  activeJob.texts.filter(needsMask).map((text) => (
                    <rect
                      key={`${text.id}-mask`}
                      x={text.sourceLeft - 2}
                      y={text.sourceTop - 2}
                      width={text.sourceWidth + 4}
                      height={text.sourceHeight + 4}
                      fill={text.maskColor}
                    />
                  ))}

                {showAssetRegions && activeJob.assets
                  .filter((asset) => asset.visible)
                  .map((asset) => (
                    <g
                      key={asset.id}
                      className={`assetLayer ${selectedAssetId === asset.id ? "selected" : ""}`}
                      onPointerDown={(event) => startAssetDrag(event, asset)}
                    >
                      {toolMode !== "asset" && (
                        <image href={asset.dataUrl} x={asset.left} y={asset.top} width={asset.width} height={asset.height} />
                      )}
                      <rect
                        x={asset.left - 4}
                        y={asset.top - 4}
                        width={asset.width + 8}
                        height={asset.height + 8}
                        fill="rgb(14 165 233 / 6%)"
                        stroke={selectedAssetId === asset.id ? "#0369a1" : "#0ea5e9"}
                        strokeDasharray="8 5"
                        strokeWidth={selectedAssetId === asset.id ? 2.4 : 1.4}
                      />
                      {selectedAssetId === asset.id &&
                        (["nw", "ne", "sw", "se"] as ResizeHandle[]).map((handle) => {
                          const size = Math.max(7, Math.min(16, Math.min(asset.width, asset.height) * 0.12));
                          const x = handle.includes("w") ? asset.left - size / 2 : asset.left + asset.width - size / 2;
                          const y = handle.includes("n") ? asset.top - size / 2 : asset.top + asset.height - size / 2;
                          return (
                            <rect
                              key={handle}
                              className={`resizeHandle ${handle}`}
                              x={x}
                              y={y}
                              width={size}
                              height={size}
                              rx={1.5}
                              fill="#ffffff"
                              stroke="#0369a1"
                              strokeWidth={1.5}
                              onPointerDown={(event) => startAssetResize(event, asset, handle)}
                            />
                          );
                        })}
                    </g>
                  ))}

                {showTextRegions && activeJob.texts
                  .filter((text) => text.visible)
                  .map((text) => (
                    <g
                      key={text.id}
                      className={`textLayer ${selectedTextId === text.id ? "selected" : ""}`}
                      onPointerDown={(event) => startTextDrag(event, text)}
                    >
                      <rect
                        x={text.left}
                        y={text.top}
                        width={Math.max(4, text.width)}
                        height={Math.max(4, text.height)}
                        fill={text.pending ? "rgb(220 38 38 / 7%)" : text.manual ? "rgb(249 115 22 / 5%)" : "rgb(249 115 22 / 5%)"}
                        stroke={selectedTextId === text.id ? "#2563eb" : text.pending ? "#dc2626" : "#f97316"}
                        strokeDasharray="8 5"
                        strokeWidth={selectedTextId === text.id ? 2 : text.pending ? 1.8 : 1.2}
                      />
                      {selectedTextId === text.id &&
                        (["nw", "ne", "sw", "se"] as ResizeHandle[]).map((handle) => {
                          const size = Math.max(6, Math.min(14, Math.min(text.width, text.height) * 0.18));
                          const x = handle.includes("w") ? text.left - size / 2 : text.left + text.width - size / 2;
                          const y = handle.includes("n") ? text.top - size / 2 : text.top + text.height - size / 2;
                          return (
                            <rect
                              key={handle}
                              className={`resizeHandle ${handle}`}
                              x={x}
                              y={y}
                              width={size}
                              height={size}
                              rx={1.5}
                              fill="#ffffff"
                              stroke="#2563eb"
                              strokeWidth={1.5}
                              onPointerDown={(event) => startTextResize(event, text, handle)}
                            />
                          );
                        })}
                      {text.dirty && text.text.trim().length > 0 && (
                        <text
                          x={text.left}
                          y={text.top + text.fontSize}
                          fontSize={text.fontSize}
                          fontFamily="'Times New Roman', Arial, sans-serif"
                          fill={text.color}
                        >
                          {text.text}
                        </text>
                      )}
                    </g>
                  ))}

                {cropBox && (
                  <rect
                    x={cropBox.left}
                    y={cropBox.top}
                    width={cropBox.width}
                    height={cropBox.height}
                    fill={toolMode === "text" ? "rgb(220 38 38 / 10%)" : "rgb(37 99 235 / 10%)"}
                    stroke={toolMode === "text" ? "#dc2626" : "#2563eb"}
                    strokeDasharray="10 6"
                    strokeWidth={3}
                  />
                )}
              </svg>
            ) : (
              <div className="dropHint">
                <FileImage size={34} />
                <span>先选择一张或多张图片</span>
              </div>
            )}
          </div>
        </section>
      </main>
    </div>
  );
}

function needsMask(text: TextLayer) {
  return text.dirty && !text.pending;
}

function shouldExportText(text: TextLayer) {
  return !text.pending && text.dirty && text.visible && text.text.trim().length > 0;
}

function shouldExportLayeredText(text: TextLayer) {
  if (text.pending || !text.visible || text.text.trim().length === 0) return false;
  if (text.dirty) return true;
  return (text.confidence ?? 100) >= 50 && looksLikeReliableText(text) && !looksLikeChartAxisText(text);
}

function shouldRemoveFromImage(text: TextLayer) {
  return text.maskOnly || (text.visible && (text.manual || shouldExportLayeredText(text)));
}

function shouldBlockForAssetExtraction(text: TextLayer) {
  if (text.maskOnly) return true;
  if (text.pending) return true;
  if (text.dirty) return true;
  return shouldExportLayeredText(text);
}

function clampTraceColorCount(value: number) {
  return Math.max(2, Math.min(96, Math.round(value) || 20));
}

function looksLikeReliableText(text: TextLayer) {
  const value = text.text.trim();
  if (/^\(?[a-z]\)$/i.test(value)) return true;
  if (/^[\\|/_\-=<>[\]{}()@#~^`'".,:;!?]+$/.test(value)) return false;
  if (/^[a-z]?[0-9]+[a-z]?$/i.test(value)) return false;
  if (/[|_{}[\]<>]/.test(value)) return false;

  const letters = value.match(/[a-z]/gi)?.length ?? 0;
  const alphaNumeric = value.match(/[a-z0-9]/gi)?.length ?? 0;
  const symbols = value.match(/[^a-z0-9\s+\-.%/]/gi)?.length ?? 0;
  const words = value.split(/\s+/).filter(Boolean);
  if (letters < 2 || alphaNumeric < 3) return false;
  if (symbols > letters) return false;
  if (words.length === 1 && value.length < 4 && !/^\(?[a-z]\)$/i.test(value)) return false;
  if (text.width < text.fontSize * 1.2 && value.length > 2) return false;
  if (text.height > text.fontSize * 2.4) return false;
  return true;
}

function looksLikeChartAxisText(text: TextLayer) {
  const value = text.text.trim();
  const normalized = value.toLowerCase().replace(/\s+/g, " ");
  if (/^[+\-]?(?:\d+(?:\.\d+)?|\.\d+)$/.test(normalized)) return true;
  if (/^[+\-]?\d+(?:\.\d+)?\s*(?:s|sec|ms|%|m\/s|rad\/s)?$/i.test(value)) return true;
  if (/^(?:[+\-]?(?:\d+(?:\.\d+)?|\.\d+)\s*){2,}$/.test(normalized)) return true;

  const chartTerms = [
    "time",
    "torque",
    "velocity",
    "reward",
    "rate",
    "work",
    "metabolic",
    "transition",
    "fall",
    "angular",
    "command",
    "reduction",
    "sample size",
    "subject",
    "seed",
    "uncertainty",
    "transition time",
    "statistical",
    "tests",
    "mechanical",
    "encoder",
  ];
  return text.fontSize <= 30 && chartTerms.some((term) => normalized.includes(term));
}

async function recognizeManualText(
  job: ImageJob,
  left: number,
  top: number,
  width: number,
  height: number,
) {
  const cropDataUrl = await cropImageRegionDataUrl(job.dataUrl, left, top, width, height);
  const ocrDataUrl = await buildOcrDataUrl(cropDataUrl);

  for (const language of ["eng+chi_sim", "eng"]) {
    let worker: Awaited<ReturnType<typeof createWorker>> | null = null;
    try {
      worker = await createWorker(language, 1);
      const result = await worker.recognize(ocrDataUrl.dataUrl);
      const rawItems: Array<{
        text: string;
        confidence: number;
        bbox?: { x0: number; y0: number; x1: number; y1: number };
      }> = collectOcrTextItems(result.data);
      const items = rawItems
        .filter((item): item is { text: string; confidence: number; bbox: { x0: number; y0: number; x1: number; y1: number } } =>
          Boolean(item.bbox),
        )
        .map((item) => ({
          text: item.text,
          confidence: item.confidence,
          bbox: {
            x0: item.bbox.x0 / ocrDataUrl.scale,
            y0: item.bbox.y0 / ocrDataUrl.scale,
            x1: item.bbox.x1 / ocrDataUrl.scale,
            y1: item.bbox.y1 / ocrDataUrl.scale,
          },
        }));
      const text =
        items.map((item: { text: string }) => String(item.text ?? "").trim()).filter(Boolean).join(" ") ||
        String(result.data.text ?? "").replace(/\s+/g, " ").trim();
      const confidences = items.map((item: { confidence: number }) => Number(item.confidence)).filter(Number.isFinite);
      const confidence =
        confidences.length > 0
          ? confidences.reduce((sum: number, value: number) => sum + value, 0) / confidences.length
          : Number(result.data.confidence ?? 100);
      const boxes = items.map((item) => item.bbox);
      const box =
        boxes.length > 0
          ? {
              left: Math.max(0, Math.min(...boxes.map((item) => item.x0))),
              top: Math.max(0, Math.min(...boxes.map((item) => item.y0))),
              width: Math.max(1, Math.max(...boxes.map((item) => item.x1)) - Math.min(...boxes.map((item) => item.x0))),
              height: Math.max(1, Math.max(...boxes.map((item) => item.y1)) - Math.min(...boxes.map((item) => item.y0))),
            }
          : null;
      await worker.terminate();
      return { text, confidence: Number.isFinite(confidence) ? confidence : 100, box, items };
    } catch (error) {
      if (worker) await worker.terminate();
      console.warn(`手动文字区域 OCR 失败: ${language}`, error);
    }
  }

  return { text: "", confidence: 0, box: null, items: [] };
}

async function cropImageRegionDataUrl(
  dataUrl: string,
  left: number,
  top: number,
  width: number,
  height: number,
) {
  const image = await loadImage(dataUrl);
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(width));
  canvas.height = Math.max(1, Math.round(height));
  const context = canvas.getContext("2d");
  if (!context) return dataUrl;

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(
    image,
    Math.max(0, Math.round(left)),
    Math.max(0, Math.round(top)),
    canvas.width,
    canvas.height,
    0,
    0,
    canvas.width,
    canvas.height,
  );
  return canvas.toDataURL("image/png");
}

async function buildOcrDataUrl(dataUrl: string) {
  const image = await loadImage(dataUrl);
  const scale = image.naturalWidth < 2600 ? 2 : 1.4;
  const width = Math.round(image.naturalWidth * scale);
  const height = Math.round(image.naturalHeight * scale);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) return { dataUrl, scale: 1 };

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.drawImage(image, 0, 0, width, height);

  const imageData = context.getImageData(0, 0, width, height);
  const pixels = imageData.data;
  for (let offset = 0; offset < pixels.length; offset += 4) {
    const red = pixels[offset];
    const green = pixels[offset + 1];
    const blue = pixels[offset + 2];
    const max = Math.max(red, green, blue);
    const min = Math.min(red, green, blue);
    const luminance = red * 0.299 + green * 0.587 + blue * 0.114;
    const coloredText = max - min > 34 && luminance < 235;
    const darkText = luminance < 205;
    if (coloredText || darkText) {
      pixels[offset] = Math.max(0, red - 35);
      pixels[offset + 1] = Math.max(0, green - 35);
      pixels[offset + 2] = Math.max(0, blue - 35);
    } else if (luminance > 225) {
      pixels[offset] = 255;
      pixels[offset + 1] = 255;
      pixels[offset + 2] = 255;
    }
  }
  context.putImageData(imageData, 0, 0);
  return { dataUrl: canvas.toDataURL("image/png"), scale };
}

async function buildTextLayers(dataUrl: string, data: any, ocrScale = 1): Promise<TextLayer[]> {
  const image = await loadImage(dataUrl);
  const sampler = createImageSampler(image);
  const items = collectOcrTextItems(data);

  const layers: TextLayer[] = [];
  for (const item of items) {
    const text = String(item.text ?? "").replace(/\s+/g, " ").trim();
    const box = item.bbox;
    if (!text || !box) continue;

    const sourceLeft = box.x0 / ocrScale;
    const sourceTop = box.y0 / ocrScale;
    const width = Math.max(1, (box.x1 - box.x0) / ocrScale);
    const height = Math.max(1, (box.y1 - box.y0) / ocrScale);
    if (width < 4 || height < 4) continue;
    if (width > image.naturalWidth * 0.35 && text.length < 30) continue;
    if (height > image.naturalHeight * 0.08) continue;

    layers.push({
        id: crypto.randomUUID(),
        text,
        confidence: Number(item.confidence ?? 100),
        sourceLeft,
        sourceTop,
        sourceWidth: width,
        sourceHeight: height,
        left: sourceLeft,
        top: sourceTop,
        width,
        height,
        fontSize: Math.max(8, Math.round(height * 0.82)),
        color: sampler.textColor(sourceLeft, sourceTop, width, height),
        maskColor: sampler.backgroundColor(sourceLeft, sourceTop, width, height),
        visible: true,
        dirty: false,
        pending: true,
      });
  }

  return layers.slice(0, 700);
}

function collectOcrTextItems(data: any) {
  const words = Array.isArray(data.words) ? data.words : [];
  if (words.length === 0) return Array.isArray(data.lines) ? data.lines : [];

  const cleanWords = words
    .map((word: any) => ({
      text: String(word.text ?? "").replace(/\s+/g, " ").trim(),
      bbox: word.bbox,
      confidence: Number(word.confidence ?? 100),
    }))
    .filter((word: any) => word.text.length > 0 && word.bbox && word.confidence >= 35)
    .sort((a: any, b: any) => wordCenterY(a) - wordCenterY(b) || a.bbox.x0 - b.bbox.x0);

  const rows: any[][] = [];
  for (const word of cleanWords) {
    const target = rows.find((row) => Math.abs(wordCenterY(row[0]) - wordCenterY(word)) <= Math.max(8, wordHeight(word) * 0.55));
    if (target) {
      target.push(word);
    } else {
      rows.push([word]);
    }
  }

  const items: Array<{ text: string; confidence: number; bbox: { x0: number; y0: number; x1: number; y1: number } }> = [];
  for (const row of rows) {
    const sorted = row.sort((a, b) => a.bbox.x0 - b.bbox.x0);
    let segment: any[] = [];
    for (const word of sorted) {
      const previous = segment[segment.length - 1];
      const gap = previous ? word.bbox.x0 - previous.bbox.x1 : 0;
      if (previous && gap > Math.max(48, wordHeight(word) * 3.2)) {
        items.push(wordsToTextItem(segment));
        segment = [];
      }
      segment.push(word);
    }
    if (segment.length > 0) items.push(wordsToTextItem(segment));
  }

  return items.filter((item) => item.text.length > 1);
}

function wordsToTextItem(words: any[]) {
  const bbox = words.reduce(
    (box, word) => ({
      x0: Math.min(box.x0, word.bbox.x0),
      y0: Math.min(box.y0, word.bbox.y0),
      x1: Math.max(box.x1, word.bbox.x1),
      y1: Math.max(box.y1, word.bbox.y1),
    }),
    { ...words[0].bbox },
  );
  return {
    text: words.map((word) => word.text).join(" "),
    confidence: words.reduce((sum, word) => sum + Number(word.confidence ?? 0), 0) / words.length,
    bbox,
  };
}

function wordCenterY(word: any) {
  return (word.bbox.y0 + word.bbox.y1) / 2;
}

function wordHeight(word: any) {
  return Math.max(1, word.bbox.y1 - word.bbox.y0);
}

function estimateManualTextFontSize(width: number, height: number, text: string) {
  const cleanText = text.trim();
  if (!cleanText) return Math.max(8, Math.round(height * 0.9));

  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");
  const fontFamily = /[\u3400-\u9fff]/.test(cleanText) ? "SimSun, serif" : "Times New Roman, serif";
  if (context) context.font = `100px ${fontFamily}`;
  const metrics = context?.measureText(cleanText);
  const measuredWidth = Math.max(1, metrics?.width ?? cleanText.length * 50);
  const measuredHeight = Math.max(
    1,
    (metrics?.actualBoundingBoxAscent ?? 72) + (metrics?.actualBoundingBoxDescent ?? 20),
  );
  const byHeight = Math.round((height * 0.96 * 100) / measuredHeight);
  const byWidth = Math.round((width * 0.96 * 100) / measuredWidth);
  return Math.max(8, Math.min(byHeight, byWidth || byHeight));
}

function textItemToLayer(
  region: TextLayer,
  item: { text: string; confidence: number; bbox: { x0: number; y0: number; x1: number; y1: number } },
  sampler: ReturnType<typeof createImageSampler>,
): TextLayer {
  const text = String(item.text ?? "").replace(/\s+/g, " ").trim();
  const left = region.left + item.bbox.x0;
  const top = region.top + item.bbox.y0;
  const width = Math.max(1, item.bbox.x1 - item.bbox.x0);
  const height = Math.max(1, item.bbox.y1 - item.bbox.y0);
  return {
    id: crypto.randomUUID(),
    sourceRegionId: region.id,
    text,
    confidence: Number(item.confidence ?? 100),
    sourceLeft: region.left,
    sourceTop: region.top,
    sourceWidth: region.width,
    sourceHeight: region.height,
    left,
    top,
    width,
    height,
    fontSize: estimateManualTextFontSize(width, height, text),
    color: sampler.textColor(left, top, width, height),
    maskColor: region.maskColor,
    visible: true,
    dirty: true,
    manual: true,
  };
}

function createImageSampler(image: HTMLImageElement) {
  const canvas = document.createElement("canvas");
  canvas.width = image.naturalWidth;
  canvas.height = image.naturalHeight;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  context?.drawImage(image, 0, 0);

  function pixelAt(x: number, y: number) {
    if (!context) return [255, 255, 255, 255];
    const safeX = Math.max(0, Math.min(canvas.width - 1, Math.round(x)));
    const safeY = Math.max(0, Math.min(canvas.height - 1, Math.round(y)));
    return Array.from(context.getImageData(safeX, safeY, 1, 1).data);
  }

  return {
    textColor(left: number, top: number, width: number, height: number) {
      if (!context) return "#111827";
      const data = context.getImageData(left, top, width, height).data;
      let r = 0;
      let g = 0;
      let b = 0;
      let count = 0;

      for (let index = 0; index < data.length; index += 16) {
        const red = data[index];
        const green = data[index + 1];
        const blue = data[index + 2];
        const max = Math.max(red, green, blue);
        const min = Math.min(red, green, blue);
        const luminance = red * 0.299 + green * 0.587 + blue * 0.114;
        if (luminance < 210 || (max - min > 34 && luminance < 245)) {
          r += red;
          g += green;
          b += blue;
          count += 1;
        }
      }

      return count > 0 ? rgbToHex(r / count, g / count, b / count) : "#111827";
    },
    backgroundColor(left: number, top: number, width: number, height: number) {
      const points = [
        pixelAt(left - 4, top - 4),
        pixelAt(left + width + 4, top - 4),
        pixelAt(left - 4, top + height + 4),
        pixelAt(left + width + 4, top + height + 4),
        pixelAt(left + width / 2, top - 5),
        pixelAt(left + width / 2, top + height + 5),
      ];
      const avg = points.reduce(
        (sum, color) => [sum[0] + color[0], sum[1] + color[1], sum[2] + color[2]],
        [0, 0, 0],
      );
      return rgbToHex(avg[0] / points.length, avg[1] / points.length, avg[2] / points.length);
    },
  };
}

async function cropAsset(job: ImageJob, left: number, top: number, width: number, height: number): Promise<AssetLayer> {
  const image = await loadImage(job.dataUrl);
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(width);
  canvas.height = Math.round(height);
  const context = canvas.getContext("2d");
  if (!context) throw new Error("无法创建素材画布");

  context.drawImage(
    image,
    Math.round(left),
    Math.round(top),
    Math.round(width),
    Math.round(height),
    0,
    0,
    Math.round(width),
    Math.round(height),
  );
  await refineTransparentAsset(context, canvas.width, canvas.height);
  const dataUrl = canvas.toDataURL("image/png");
  const vectorSvg = "";

  const name = `${stripExtension(job.name)}-素材-${job.assets.length + 1}`;
  return {
    id: crypto.randomUUID(),
    name,
    dataUrl,
    vectorSvg,
    maskColor: averageEdgeColor(context, canvas.width, canvas.height),
    sourceLeft: left,
    sourceTop: top,
    sourceWidth: width,
    sourceHeight: height,
    left,
    top,
    width,
    height,
    visible: true,
  };
}

async function autoExtractAssets(job: ImageJob, texts: TextLayer[]): Promise<AssetLayer[]> {
  try {
    return await autoExtractAssetsWithOpenCv(job, texts);
  } catch (error) {
    console.warn("OpenCV 自动抠图失败，回退到基础算法", error);
  }

  const image = await loadImage(job.dataUrl);
  const scale = Math.min(1, 1200 / image.naturalWidth);
  const width = Math.round(image.naturalWidth * scale);
  const height = Math.round(image.naturalHeight * scale);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) return [];

  context.drawImage(image, 0, 0, width, height);
  const imageData = context.getImageData(0, 0, width, height);
  const data = imageData.data;
  const blocked = new Uint8Array(width * height);

  texts.forEach((text) => {
    const pad = 3;
    const x0 = Math.max(0, Math.floor((text.sourceLeft - pad) * scale));
    const y0 = Math.max(0, Math.floor((text.sourceTop - pad) * scale));
    const x1 = Math.min(width, Math.ceil((text.sourceLeft + text.sourceWidth + pad) * scale));
    const y1 = Math.min(height, Math.ceil((text.sourceTop + text.sourceHeight + pad) * scale));
    for (let y = y0; y < y1; y += 1) {
      for (let x = x0; x < x1; x += 1) {
        blocked[y * width + x] = 1;
      }
    }
  });

  const foreground = new Uint8Array(width * height);
  for (let index = 0; index < width * height; index += 1) {
    if (blocked[index]) continue;
    const offset = index * 4;
    const red = data[offset];
    const green = data[offset + 1];
    const blue = data[offset + 2];
    const alpha = data[offset + 3];
    if (alpha < 20) continue;
    const max = Math.max(red, green, blue);
    const min = Math.min(red, green, blue);
    const dark = red * 0.299 + green * 0.587 + blue * 0.114 < 225;
    const saturated = max - min > 24;
    if (dark || saturated) foreground[index] = 1;
  }

  const visited = new Uint8Array(width * height);
  const boxes: Array<{ left: number; top: number; right: number; bottom: number; count: number }> = [];
  const queue: number[] = [];

  for (let start = 0; start < foreground.length; start += 1) {
    if (!foreground[start] || visited[start]) continue;
    visited[start] = 1;
    queue.length = 0;
    queue.push(start);

    let left = start % width;
    let right = left;
    let top = Math.floor(start / width);
    let bottom = top;
    let count = 0;

    for (let head = 0; head < queue.length; head += 1) {
      const current = queue[head];
      const x = current % width;
      const y = Math.floor(current / width);
      count += 1;
      left = Math.min(left, x);
      right = Math.max(right, x);
      top = Math.min(top, y);
      bottom = Math.max(bottom, y);

      const neighbors = [current - 1, current + 1, current - width, current + width];
      for (const next of neighbors) {
        if (next < 0 || next >= foreground.length || visited[next] || !foreground[next]) continue;
        const nx = next % width;
        if (Math.abs(nx - x) > 1) continue;
        visited[next] = 1;
        queue.push(next);
      }
    }

    const boxWidth = right - left + 1;
    const boxHeight = bottom - top + 1;
    if (count > 180 && boxWidth > 22 && boxHeight > 22) {
      boxes.push({ left, top, right, bottom, count });
    }
  }

  const merged = mergeAssetBoxes(boxes, width, height)
    .map((box) => ({
      left: Math.max(0, Math.round(box.left / scale) - 4),
      top: Math.max(0, Math.round(box.top / scale) - 4),
      right: Math.min(job.width, Math.round((box.right + 1) / scale) + 4),
      bottom: Math.min(job.height, Math.round((box.bottom + 1) / scale) + 4),
      count: box.count,
    }))
    .filter((box) => box.right - box.left > 36 && box.bottom - box.top > 36)
    .filter((box) => {
      const boxWidth = box.right - box.left;
      const boxHeight = box.bottom - box.top;
      const ratio = boxWidth / boxHeight;
      return ratio < 8 && ratio > 0.12;
    })
    .filter((box) => (box.right - box.left) * (box.bottom - box.top) < job.width * job.height * 0.12)
    .sort((a, b) => b.count - a.count)
    .slice(0, 80);

  const assets: AssetLayer[] = [];
  for (const box of merged) {
    const asset = await cropTransparentAsset(
      job,
      box.left,
      box.top,
      box.right - box.left,
      box.bottom - box.top,
      assets.length + 1,
    );
    assets.push(asset);
  }
  return assets;
}

async function vectorizeImage(job: ImageJob) {
  return vectorizeDataUrl(job.dataUrl, job.width, job.height);
}

async function vectorizeDataUrl(dataUrl: string, sourceWidth: number, sourceHeight: number, colorCount = 20) {
  const image = await loadImage(dataUrl);
  const maxWidth = 1800;
  const scale = Math.min(1, maxWidth / image.naturalWidth);
  const width = Math.round(image.naturalWidth * scale);
  const height = Math.round(image.naturalHeight * scale);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) return "";

  context.drawImage(image, 0, 0, width, height);
  const imageData = context.getImageData(0, 0, width, height);
  const tracer = await getImageTracer();
  const svg = tracer.imagedataToSVG(imageData, {
    ltres: 0.25,
    qtres: 0.25,
    pathomit: 2,
    rightangleenhance: true,
    colorsampling: 2,
    numberofcolors: clampTraceColorCount(colorCount),
    mincolorratio: 0,
    colorquantcycles: 5,
    layering: 0,
    strokewidth: 0,
    linefilter: false,
    scale: 1 / scale,
    roundcoords: 3,
    viewbox: true,
    desc: false,
    blurradius: 0,
    blurdelta: 20,
  });

  return svg
    .replace(/width="[^"]+"/, `width="${sourceWidth}"`)
    .replace(/height="[^"]+"/, `height="${sourceHeight}"`)
    .replace(/viewBox="[^"]+"/, `viewBox="0 0 ${sourceWidth} ${sourceHeight}"`);
}

async function prepareLayeredAssets(job: ImageJob, colorCount: number, cleanSourceDataUrl?: string) {
  const result: AssetLayer[] = [];
  const cleanSourceImage = await loadImage(cleanSourceDataUrl ?? job.dataUrl);
  for (const asset of job.assets) {
    if (!asset.visible) {
      result.push(asset);
      continue;
    }
    const exportAsset = expandAssetForExport(asset, job.width, job.height);
    const dataUrl = await cropAssetFromCleanSource(cleanSourceImage, exportAsset, "png");
    const traceDataUrl = await cropAssetFromCleanSource(cleanSourceImage, exportAsset, "trace");
    try {
      result.push({
        ...exportAsset,
        dataUrl,
        vectorSvg: await vectorizeAssetDataUrl(traceDataUrl, exportAsset.width, exportAsset.height, colorCount),
      });
    } catch (error) {
      console.warn("素材矢量化失败，保留 PNG 素材", error);
      result.push({ ...exportAsset, dataUrl, vectorSvg: "" });
    }
  }
  return result;
}

function expandAssetForExport(asset: AssetLayer, imageWidth: number, imageHeight: number): AssetLayer {
  const padding = Math.max(4, Math.min(18, Math.round(Math.min(asset.width, asset.height) * 0.06)));
  const left = Math.max(0, Math.round(asset.left) - padding);
  const top = Math.max(0, Math.round(asset.top) - padding);
  const right = Math.min(imageWidth, Math.round(asset.left + asset.width) + padding);
  const bottom = Math.min(imageHeight, Math.round(asset.top + asset.height) + padding);
  const width = Math.max(1, right - left);
  const height = Math.max(1, bottom - top);
  return {
    ...asset,
    left,
    top,
    width,
    height,
    sourceLeft: left,
    sourceTop: top,
    sourceWidth: width,
    sourceHeight: height,
  };
}

async function cropAssetFromCleanSource(
  image: HTMLImageElement,
  asset: AssetLayer,
  transparency: "none" | "png" | "trace",
) {
  const width = Math.max(1, Math.round(asset.width));
  const height = Math.max(1, Math.round(asset.height));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) return asset.dataUrl;

  context.drawImage(
    image,
    Math.round(asset.left),
    Math.round(asset.top),
    width,
    height,
    0,
    0,
    width,
    height,
  );
  if (transparency !== "none") makeBackgroundTransparent(context, width, height, transparency);
  return canvas.toDataURL("image/png");
}

async function vectorizeAssetDataUrl(dataUrl: string, sourceWidth: number, sourceHeight: number, colorCount = 20) {
  const image = await loadImage(dataUrl);
  const minTraceWidth = 1200;
  const maxTraceWidth = 2400;
  const upscale = image.naturalWidth < minTraceWidth ? Math.min(4, minTraceWidth / image.naturalWidth) : 1;
  const scale = Math.min(upscale, maxTraceWidth / image.naturalWidth);
  const width = Math.max(1, Math.round(image.naturalWidth * scale));
  const height = Math.max(1, Math.round(image.naturalHeight * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) return "";

  context.drawImage(image, 0, 0, width, height);
  cleanupAssetPixelsForTrace(context, width, height, colorCount);
  const tracer = await getImageTracer();
  const imageData = context.getImageData(0, 0, width, height);
  const primarySvg = cleanTracedAssetSvg(
    traceAssetImageData(tracer, imageData, scale, {
      colorCount: Math.max(colorCount, 48),
      ltres: 0.04,
      qtres: 0.04,
      pathomit: 0,
      colorquantcycles: 6,
      roundcoords: 3,
    }),
  );
  let cleanedSvg = primarySvg;
  if (isUnsafeTracedAssetSvg(cleanedSvg, width, height)) {
    cleanedSvg = cleanTracedAssetSvg(
      traceAssetImageData(tracer, imageData, scale, {
        colorCount: Math.max(48, colorCount),
        ltres: 0.08,
        qtres: 0.08,
        pathomit: 1,
        colorquantcycles: 4,
        roundcoords: 3,
      }),
    );
  }
  if (isSeverelyUnsafeTracedAssetSvg(cleanedSvg, width, height)) return "";
  if (!hasVisibleTracedAssetSvg(cleanedSvg)) return "";

  return cleanedSvg
    .replace(/width="[^"]+"/, `width="${sourceWidth}"`)
    .replace(/height="[^"]+"/, `height="${sourceHeight}"`)
    .replace(/viewBox="[^"]+"/, `viewBox="0 0 ${sourceWidth} ${sourceHeight}"`);
}

function traceAssetImageData(
  tracer: any,
  imageData: ImageData,
  scale: number,
  options: {
    colorCount: number;
    ltres: number;
    qtres: number;
    pathomit: number;
    colorquantcycles: number;
    roundcoords: number;
  },
) {
  return tracer.imagedataToSVG(imageData, {
    ltres: options.ltres,
    qtres: options.qtres,
    pathomit: options.pathomit,
    rightangleenhance: true,
    colorsampling: 2,
    numberofcolors: clampTraceColorCount(options.colorCount),
    mincolorratio: 0,
    colorquantcycles: options.colorquantcycles,
    layering: 0,
    strokewidth: 0,
    linefilter: false,
    scale: 1 / scale,
    roundcoords: options.roundcoords,
    viewbox: true,
    desc: false,
    blurradius: 0,
    blurdelta: 20,
  });
}

function cleanupAssetPixelsForTrace(context: CanvasRenderingContext2D, width: number, height: number, colorCount: number) {
  const imageData = context.getImageData(0, 0, width, height);
  const data = imageData.data;

  for (let offset = 0; offset < data.length; offset += 4) {
    const alpha = data[offset + 3];

    if (alpha < 96) {
      data[offset] = 255;
      data[offset + 1] = 255;
      data[offset + 2] = 255;
      data[offset + 3] = 0;
    } else if (alpha < 220) {
      data[offset + 3] = Math.min(255, Math.round(alpha * 1.18));
    }
  }

  context.putImageData(imageData, 0, 0);
  makeBackgroundTransparent(context, width, height, "trace");

  const cleaned = context.getImageData(0, 0, width, height);
  snapNearbyTraceColors(cleaned.data, colorCount);
  limitTraceColors(cleaned.data, colorCount);
  context.putImageData(cleaned, 0, 0);
}

function snapNearbyTraceColors(data: Uint8ClampedArray, colorCount: number) {
  const step = colorCount <= 12 ? 28 : colorCount <= 20 ? 22 : colorCount <= 48 ? 16 : 12;
  const buckets = new Map<string, { red: number; green: number; blue: number; count: number }>();

  for (let offset = 0; offset < data.length; offset += 4) {
    if (data[offset + 3] === 0) continue;
    const key = [
      Math.round(data[offset] / step),
      Math.round(data[offset + 1] / step),
      Math.round(data[offset + 2] / step),
    ].join(",");
    const bucket = buckets.get(key) ?? { red: 0, green: 0, blue: 0, count: 0 };
    bucket.red += data[offset];
    bucket.green += data[offset + 1];
    bucket.blue += data[offset + 2];
    bucket.count += 1;
    buckets.set(key, bucket);
  }

  for (const bucket of buckets.values()) {
    bucket.red = Math.round(bucket.red / bucket.count);
    bucket.green = Math.round(bucket.green / bucket.count);
    bucket.blue = Math.round(bucket.blue / bucket.count);
  }

  for (let offset = 0; offset < data.length; offset += 4) {
    if (data[offset + 3] === 0) continue;
    const key = [
      Math.round(data[offset] / step),
      Math.round(data[offset + 1] / step),
      Math.round(data[offset + 2] / step),
    ].join(",");
    const bucket = buckets.get(key);
    if (!bucket) continue;
    data[offset] = bucket.red;
    data[offset + 1] = bucket.green;
    data[offset + 2] = bucket.blue;
  }
}

function limitTraceColors(data: Uint8ClampedArray, colorCount: number) {
  const target = clampTraceColorCount(colorCount);
  const buckets = new Map<string, { red: number; green: number; blue: number; count: number }>();

  for (let offset = 0; offset < data.length; offset += 4) {
    if (data[offset + 3] === 0) continue;
    const key = `${data[offset]},${data[offset + 1]},${data[offset + 2]}`;
    const bucket = buckets.get(key) ?? { red: data[offset], green: data[offset + 1], blue: data[offset + 2], count: 0 };
    bucket.count += 1;
    buckets.set(key, bucket);
  }

  const palette = pickTracePalette(Array.from(buckets.values()), target);
  if (palette.length === 0 || buckets.size <= palette.length) return;

  for (let offset = 0; offset < data.length; offset += 4) {
    if (data[offset + 3] === 0) continue;
    const nearest = nearestPaletteColor(data[offset], data[offset + 1], data[offset + 2], palette);
    data[offset] = nearest.red;
    data[offset + 1] = nearest.green;
    data[offset + 2] = nearest.blue;
  }
}

function pickTracePalette(
  colors: Array<{ red: number; green: number; blue: number; count: number }>,
  target: number,
) {
  const sorted = [...colors].sort((a, b) => b.count - a.count);
  const palette: Array<{ red: number; green: number; blue: number; count: number }> = [];
  const used = new Set<string>();

  while (palette.length < target && palette.length < sorted.length) {
    let best = sorted[0];
    let bestScore = -1;
    for (const color of sorted) {
      const key = `${color.red},${color.green},${color.blue}`;
      if (used.has(key)) continue;
      const distance = palette.length === 0 ? 1 : nearestPaletteDistance(color, palette);
      const score = distance * Math.sqrt(color.count);
      if (score > bestScore) {
        best = color;
        bestScore = score;
      }
    }
    const key = `${best.red},${best.green},${best.blue}`;
    used.add(key);
    palette.push(best);
  }

  return palette;
}

function nearestPaletteDistance(
  color: { red: number; green: number; blue: number },
  palette: Array<{ red: number; green: number; blue: number }>,
) {
  let nearest = Number.POSITIVE_INFINITY;
  for (const item of palette) {
    const distance =
      (color.red - item.red) * (color.red - item.red) +
      (color.green - item.green) * (color.green - item.green) +
      (color.blue - item.blue) * (color.blue - item.blue);
    nearest = Math.min(nearest, distance);
  }
  return nearest;
}

function nearestPaletteColor(
  red: number,
  green: number,
  blue: number,
  palette: Array<{ red: number; green: number; blue: number }>,
) {
  let nearest = palette[0];
  let nearestDistance = Number.POSITIVE_INFINITY;
  for (const color of palette) {
    const distance =
      (red - color.red) * (red - color.red) +
      (green - color.green) * (green - color.green) +
      (blue - color.blue) * (blue - color.blue);
    if (distance < nearestDistance) {
      nearest = color;
      nearestDistance = distance;
    }
  }
  return nearest;
}

function cleanTracedAssetSvg(svg: string) {
  const viewBox = svg.match(/viewBox="([^"]+)"/i)?.[1];
  const values = viewBox?.split(/[\s,]+/).map(Number).filter(Number.isFinite) ?? [];
  const fullArea = values.length >= 4 ? Math.max(1, values[2] * values[3]) : 1;
  const edgeColors = tracedSvgEdgeColors(svg, values, fullArea);

  return svg.replace(/<path\b[^>]*>/gi, (path) => {
    if (isInvisibleSvgPath(path)) return "";
    const color = svgPathColor(path);
    const box = svgPathBox(path);
    const area = box ? Math.max(0, (box.maxX - box.minX) * (box.maxY - box.minY)) : 0;
    const strokeWidth = svgPathStrokeWidth(path);
    if (strokeWidth > Math.max(12, Math.sqrt(fullArea) * 0.12)) return "";
    if (!color) return path;

    const max = Math.max(color.r, color.g, color.b);
    const min = Math.min(color.r, color.g, color.b);
    const luminance = color.r * 0.299 + color.g * 0.587 + color.b * 0.114;
    const nearWhite = luminance > 238 && max - min < 24;
    const tiny = area > 0 && area < fullArea * 0.0025;
    const largeEdgeBackground =
      box !== null &&
      area > fullArea * 0.018 &&
      touchesSvgEdge(box, values) &&
      edgeColors.some((edgeColor) => colorDistance(color, edgeColor) < 38);
    const largeInteriorBackground =
      box !== null &&
      area > fullArea * 0.08 &&
      edgeColors.some((edgeColor) => colorDistance(color, edgeColor) < 30);
    const largePaleEdge =
      box !== null && area > fullArea * 0.18 && luminance > 218 && max - min < 44 && touchesSvgEdge(box, values);
    if (largeEdgeBackground || largeInteriorBackground) return "";
    if (largePaleEdge) return "";
    if (nearWhite && tiny) return "";
    return path;
  });
}

function tracedSvgEdgeColors(svg: string, viewBoxValues: number[], fullArea: number) {
  const buckets = new Map<string, { r: number; g: number; b: number; count: number }>();
  if (viewBoxValues.length < 4) return [];

  for (const match of svg.matchAll(/<path\b[^>]*>/gi)) {
    const path = match[0];
    const color = svgPathColor(path);
    const box = svgPathBox(path);
    if (!color || !box || !touchesSvgEdge(box, viewBoxValues)) continue;
    const area = Math.max(0, (box.maxX - box.minX) * (box.maxY - box.minY));
    if (area < fullArea * 0.001) continue;
    const key = `${Math.round(color.r / 12)},${Math.round(color.g / 12)},${Math.round(color.b / 12)}`;
    const bucket = buckets.get(key) ?? { r: 0, g: 0, b: 0, count: 0 };
    bucket.r += color.r;
    bucket.g += color.g;
    bucket.b += color.b;
    bucket.count += 1;
    buckets.set(key, bucket);
  }

  return Array.from(buckets.values())
    .filter((bucket) => bucket.count >= 2)
    .sort((a, b) => b.count - a.count)
    .slice(0, 6)
    .map((bucket) => ({
      r: Math.round(bucket.r / bucket.count),
      g: Math.round(bucket.g / bucket.count),
      b: Math.round(bucket.b / bucket.count),
    }));
}

function colorDistance(a: { r: number; g: number; b: number }, b: { r: number; g: number; b: number }) {
  return Math.hypot(a.r - b.r, a.g - b.g, a.b - b.b);
}

function isInvisibleSvgPath(path: string) {
  return /\sopacity="0(?:\.0*)?"/i.test(path) || /opacity\s*:\s*0(?:\.0*)?/i.test(path);
}

function svgPathStrokeWidth(path: string) {
  const value =
    path.match(/stroke-width:\s*([0-9.]+)/i)?.[1] ??
    path.match(/\sstroke-width="([0-9.]+)"/i)?.[1];
  const number = Number(value ?? 0);
  return Number.isFinite(number) ? number : 0;
}

function touchesSvgEdge(box: { minX: number; minY: number; maxX: number; maxY: number }, viewBoxValues: number[]) {
  if (viewBoxValues.length < 4) return false;
  const [x, y, width, height] = viewBoxValues;
  const pad = Math.max(1, Math.min(width, height) * 0.01);
  return box.minX <= x + pad || box.minY <= y + pad || box.maxX >= x + width - pad || box.maxY >= y + height - pad;
}

function hasVisibleTracedAssetSvg(svg: string) {
  return /<path\b[^>]*>/i.test(svg);
}

function isUnsafeTracedAssetSvg(svg: string, width: number, height: number) {
  const paths = Array.from(svg.matchAll(/<path\b[^>]*>/gi), (match) => match[0]);
  if (paths.length === 0) return true;

  const viewBox = svg.match(/viewBox="([^"]+)"/i)?.[1];
  const values = viewBox?.split(/[\s,]+/).map(Number).filter(Number.isFinite) ?? [];
  const fullArea = values.length >= 4 ? Math.max(1, values[2] * values[3]) : Math.max(1, width * height);
  let lightPaths = 0;

  for (const path of paths) {
    const color = svgPathColor(path);
    const box = svgPathBox(path);
    if (!color || !box) continue;

    const area = Math.max(0, (box.maxX - box.minX) * (box.maxY - box.minY));
    const max = Math.max(color.r, color.g, color.b);
    const min = Math.min(color.r, color.g, color.b);
    const luminance = color.r * 0.299 + color.g * 0.587 + color.b * 0.114;
    if (luminance > 232 && max - min < 28) lightPaths += 1;
    if (area > fullArea * 0.55 && luminance < 48) return true;
    if (area > fullArea * 0.7 && max - min < 18 && luminance < 210) return true;
  }

  return svg.length > 24_000_000 || (paths.length > 120_000 && lightPaths / paths.length > 0.8);
}

function isSeverelyUnsafeTracedAssetSvg(svg: string, width: number, height: number) {
  const paths = Array.from(svg.matchAll(/<path\b[^>]*>/gi), (match) => match[0]);
  if (paths.length === 0) return true;
  if (svg.length > 36_000_000) return true;

  const viewBox = svg.match(/viewBox="([^"]+)"/i)?.[1];
  const values = viewBox?.split(/[\s,]+/).map(Number).filter(Number.isFinite) ?? [];
  const fullArea = values.length >= 4 ? Math.max(1, values[2] * values[3]) : Math.max(1, width * height);

  for (const path of paths) {
    const color = svgPathColor(path);
    const box = svgPathBox(path);
    if (!color || !box) continue;
    const area = Math.max(0, (box.maxX - box.minX) * (box.maxY - box.minY));
    const max = Math.max(color.r, color.g, color.b);
    const min = Math.min(color.r, color.g, color.b);
    const luminance = color.r * 0.299 + color.g * 0.587 + color.b * 0.114;
    if (area > fullArea * 0.85 && luminance < 40) return true;
    if (area > fullArea * 0.9 && max - min < 18 && luminance < 190) return true;
  }

  return false;
}

function svgPathColor(path: string) {
  const rgb =
    path.match(/fill:\s*rgb\((\d+),\s*(\d+),\s*(\d+)\)/i) ??
    path.match(/\sfill="rgb\((\d+),\s*(\d+),\s*(\d+)\)"/i);
  if (rgb) return { r: Number(rgb[1]), g: Number(rgb[2]), b: Number(rgb[3]) };

  const hex =
    path.match(/fill:\s*(#[0-9a-fA-F]{6})/i)?.[1] ??
    path.match(/\sfill="(#[0-9a-fA-F]{6})"/i)?.[1];
  if (!hex) return null;
  const clean = hex.replace("#", "");
  return {
    r: parseInt(clean.slice(0, 2), 16),
    g: parseInt(clean.slice(2, 4), 16),
    b: parseInt(clean.slice(4, 6), 16),
  };
}

function svgPathBox(path: string) {
  const d = path.match(/\sd="([^"]+)"/i)?.[1];
  if (!d) return null;
  const values = Array.from(d.matchAll(/[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?/g), (match) => Number(match[0]));
  if (values.length < 2) return null;
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (let index = 0; index < values.length - 1; index += 2) {
    const x = values[index];
    const y = values[index + 1];
    minX = Math.min(minX, x);
    minY = Math.min(minY, y);
    maxX = Math.max(maxX, x);
    maxY = Math.max(maxY, y);
  }
  if (!Number.isFinite(minX) || !Number.isFinite(minY) || !Number.isFinite(maxX) || !Number.isFinite(maxY)) return null;
  return { minX, minY, maxX, maxY };
}

async function autoExtractAssetsWithOpenCv(job: ImageJob, texts: TextLayer[]): Promise<AssetLayer[]> {
  const cv = await getOpenCv();
  const image = await loadImage(job.dataUrl);
  const scale = Math.min(1, 1400 / image.naturalWidth);
  const width = Math.round(image.naturalWidth * scale);
  const height = Math.round(image.naturalHeight * scale);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) return [];

  context.drawImage(image, 0, 0, width, height);
  const source = cv.imread(canvas);
  const gray = new cv.Mat();
  const blurred = new cv.Mat();
  const binary = new cv.Mat();
  const contours = new cv.MatVector();
  const hierarchy = new cv.Mat();
  const kernel = cv.Mat.ones(3, 3, cv.CV_8U);
  const closeKernel = cv.Mat.ones(7, 7, cv.CV_8U);

  try {
    cv.cvtColor(source, gray, cv.COLOR_RGBA2GRAY);
    cv.GaussianBlur(gray, blurred, new cv.Size(3, 3), 0);
    cv.threshold(blurred, binary, 238, 255, cv.THRESH_BINARY_INV);

    texts.forEach((text) => {
      const pad = 4;
      const x = Math.max(0, Math.floor((text.sourceLeft - pad) * scale));
      const y = Math.max(0, Math.floor((text.sourceTop - pad) * scale));
      const w = Math.min(width - x, Math.ceil((text.sourceWidth + pad * 2) * scale));
      const h = Math.min(height - y, Math.ceil((text.sourceHeight + pad * 2) * scale));
      cv.rectangle(binary, new cv.Point(x, y), new cv.Point(x + w, y + h), new cv.Scalar(0), -1);
    });

    cv.morphologyEx(binary, binary, cv.MORPH_CLOSE, kernel);
    cv.findContours(binary, contours, hierarchy, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);

    const boxes: Array<{ left: number; top: number; right: number; bottom: number; count: number }> = [];
    for (let i = 0; i < contours.size(); i += 1) {
      const contour = contours.get(i);
      const rect = cv.boundingRect(contour);
      const area = cv.contourArea(contour);
      contour.delete();

      if (area < 180) continue;
      if (rect.width < 18 || rect.height < 18) continue;
      const ratio = rect.width / rect.height;
      if (ratio > 9 || ratio < 0.1) continue;
      if (rect.width * rect.height > width * height * 0.08) continue;

      boxes.push({
        left: rect.x,
        top: rect.y,
        right: rect.x + rect.width,
        bottom: rect.y + rect.height,
        count: area,
      });
    }

    const merged = mergeAssetBoxes(boxes, width, height)
      .map((box) => ({
        left: Math.max(0, Math.round(box.left / scale) - 3),
        top: Math.max(0, Math.round(box.top / scale) - 3),
        right: Math.min(job.width, Math.round(box.right / scale) + 3),
        bottom: Math.min(job.height, Math.round(box.bottom / scale) + 3),
        count: box.count,
      }))
      .filter((box) => box.right - box.left > 34 && box.bottom - box.top > 34)
      .filter((box) => {
        const boxWidth = box.right - box.left;
        const boxHeight = box.bottom - box.top;
        return boxWidth * boxHeight < job.width * job.height * 0.08;
      })
      .sort((a, b) => b.count - a.count)
      .slice(0, 160);

    const assets: AssetLayer[] = [];
    for (const box of merged) {
      const asset = await cropTransparentAsset(
        job,
        box.left,
        box.top,
        box.right - box.left,
        box.bottom - box.top,
        assets.length + 1,
      );
      assets.push(asset);
    }
    return assets;
  } finally {
    source.delete();
    gray.delete();
    blurred.delete();
    binary.delete();
    contours.delete();
    hierarchy.delete();
    kernel.delete();
    closeKernel.delete();
  }
}

async function cropTransparentAsset(
  job: ImageJob,
  left: number,
  top: number,
  width: number,
  height: number,
  index: number,
): Promise<AssetLayer> {
  const image = await loadImage(job.dataUrl);
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(width);
  canvas.height = Math.round(height);
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) throw new Error("无法创建素材画布");

  context.drawImage(image, Math.round(left), Math.round(top), Math.round(width), Math.round(height), 0, 0, Math.round(width), Math.round(height));
  const maskColor = averageEdgeColor(context, canvas.width, canvas.height);
  await refineTransparentAsset(context, canvas.width, canvas.height);
  const dataUrl = canvas.toDataURL("image/png");
  const vectorSvg = "";

  return {
    id: crypto.randomUUID(),
    name: `${stripExtension(job.name)}-自动素材-${index}`,
    dataUrl,
    vectorSvg,
    maskColor,
    sourceLeft: left,
    sourceTop: top,
    sourceWidth: width,
    sourceHeight: height,
    left,
    top,
    width,
    height,
    visible: true,
  };
}

async function refineTransparentAsset(context: CanvasRenderingContext2D, width: number, height: number) {
  makeBackgroundTransparent(context, width, height);
  return;

  try {
    const cv = await getOpenCv();
    grabCutTransparent(context, width, height, cv);
  } catch (error) {
    console.warn("GrabCut 失败，使用边缘背景透明化", error);
    makeBackgroundTransparent(context, width, height);
  }
}

function grabCutTransparent(context: CanvasRenderingContext2D, width: number, height: number, cv: any) {
  if (width < 10 || height < 10) {
    makeBackgroundTransparent(context, width, height);
    return;
  }

  const canvas = context.canvas;
  const source = cv.imread(canvas);
  const rgb = new cv.Mat();
  const mask = new cv.Mat();
  const bgdModel = new cv.Mat();
  const fgdModel = new cv.Mat();
  const resultMask = new cv.Mat();
  const result = new cv.Mat();

  try {
    cv.cvtColor(source, rgb, cv.COLOR_RGBA2RGB);
    mask.create(height, width, cv.CV_8UC1);
    mask.setTo(new cv.Scalar(cv.GC_PR_BGD));

    const marginX = Math.max(2, Math.round(width * 0.06));
    const marginY = Math.max(2, Math.round(height * 0.06));
    const rect = new cv.Rect(marginX, marginY, Math.max(1, width - marginX * 2), Math.max(1, height - marginY * 2));
    cv.grabCut(rgb, mask, rect, bgdModel, fgdModel, 2, cv.GC_INIT_WITH_RECT);

    const probableForeground = new cv.Mat(mask.rows, mask.cols, mask.type(), new cv.Scalar(cv.GC_PR_FGD));
    const sureForeground = new cv.Mat(mask.rows, mask.cols, mask.type(), new cv.Scalar(cv.GC_FGD));
    const fg1 = new cv.Mat();
    const fg2 = new cv.Mat();
    cv.compare(mask, probableForeground, fg1, cv.CMP_EQ);
    cv.compare(mask, sureForeground, fg2, cv.CMP_EQ);
    cv.bitwise_or(fg1, fg2, resultMask);
    source.copyTo(result, resultMask);

    cv.imshow(canvas, result);
    makeBackgroundTransparent(context, width, height);

    probableForeground.delete();
    sureForeground.delete();
    fg1.delete();
    fg2.delete();
  } finally {
    source.delete();
    rgb.delete();
    mask.delete();
    bgdModel.delete();
    fgdModel.delete();
    resultMask.delete();
    result.delete();
  }
}

async function buildSvgMarkup(job: ImageJob) {
  const cleanBackground = await buildCleanBackgroundDataUrl(job);
  const assets = job.assets
    .filter((asset) => asset.visible)
    .map(
      (asset) =>
        asset.vectorSvg
          ? `<svg x="${asset.left}" y="${asset.top}" width="${asset.width}" height="${asset.height}" viewBox="0 0 ${asset.width} ${asset.height}">${extractSvgInner(asset.vectorSvg)}</svg>`
          : `<image href="${asset.dataUrl}" x="${asset.left}" y="${asset.top}" width="${asset.width}" height="${asset.height}" />`,
    )
    .join("");

  const textNodes = job.texts
    .filter(shouldExportText)
    .map(
      (text) =>
        `<text x="${text.left}" y="${text.top + text.fontSize}" font-size="${text.fontSize}" font-family="Times New Roman, Arial, sans-serif" fill="${text.color}">${escapeXml(text.text)}</text>`,
    )
    .join("");

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${job.width}" height="${job.height}" viewBox="0 0 ${job.width} ${job.height}">
<image href="${cleanBackground}" width="${job.width}" height="${job.height}" />
${assets}
${textNodes}
</svg>`;
}

async function buildSvgBlob(job: ImageJob) {
  return new Blob([await buildSvgMarkup(job)], { type: "image/svg+xml;charset=utf-8" });
}

function buildVectorSvgBlob(job: ImageJob) {
  return new Blob([job.vectorSvg || ""], { type: "image/svg+xml;charset=utf-8" });
}

function normalizeSvgSize(svg: string, width: number, height: number) {
  const hasViewBox = /viewBox="[^"]+"/i.test(svg);
  const sized = svg
    .replace(/<svg\b([^>]*)>/i, (match) => {
      let next = match
        .replace(/\swidth="[^"]*"/i, "")
        .replace(/\sheight="[^"]*"/i, "");
      if (!hasViewBox) next = next.replace(/<svg\b/i, `<svg viewBox="0 0 ${width} ${height}"`);
      return next.replace(/<svg\b/i, `<svg width="${width}" height="${height}"`);
    });
  return sized;
}

function svgToDataUri(svg: string) {
  return `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}`;
}

function extractSvgInner(svg: string) {
  const match = svg.match(/<svg[^>]*>([\s\S]*)<\/svg>/i);
  return match ? match[1] : "";
}

async function writePptx(job: ImageJob, mode: "visual" | "layered") {
  const pptx = new PptxGenJS();
  const slideWidth = 13.333;
  const slideHeight = slideWidth * (job.height / job.width);
  const layoutName = "IMAGE_EDIT_LAYOUT";
  pptx.defineLayout({ name: layoutName, width: slideWidth, height: slideHeight });
  pptx.layout = layoutName;
  pptx.author = "钳工的美术箱";

  const slide = pptx.addSlide();
  const scaleX = slideWidth / job.width;
  const scaleY = slideHeight / job.height;
  if (mode === "visual" && job.vectorSvg) {
    slide.addImage({
      data: svgToDataUri(job.vectorSvg),
      x: 0,
      y: 0,
      w: slideWidth,
      h: slideHeight,
    });
    slide.addImage({ data: job.dataUrl, x: 0, y: 0, w: slideWidth, h: slideHeight });
  } else {
    const background = mode === "layered" ? await buildCleanBackgroundDataUrl(job) : job.dataUrl;
    slide.addImage({ data: background, x: 0, y: 0, w: slideWidth, h: slideHeight });
  }

  if (mode === "layered") {
    splitSvgPaths(job.vectorSvg, 260).forEach((svg) => {
      slide.addImage({
        data: svgToDataUri(svg),
        x: 0,
        y: 0,
        w: slideWidth,
        h: slideHeight,
      });
    });

    job.assets
      .filter((asset) => asset.visible)
      .forEach((asset) => {
        slide.addImage({
          data: asset.vectorSvg ? svgToDataUri(asset.vectorSvg) : asset.dataUrl,
          x: asset.left * scaleX,
          y: asset.top * scaleY,
          w: asset.width * scaleX,
          h: asset.height * scaleY,
        });
      });

    job.texts
      .filter(shouldExportText)
      .forEach((text) => {
        slide.addText(text.text, {
          x: text.left * scaleX,
          y: text.top * scaleY,
          w: Math.max(text.width, text.text.length * text.fontSize * 0.5) * scaleX,
          h: Math.max(text.height, text.fontSize * 1.2) * scaleY,
          margin: 0,
          fontFace: "Times New Roman",
          fontSize: Math.max(5, text.fontSize * scaleY * 72),
          color: text.color.replace("#", ""),
          breakLine: false,
          fit: "shrink",
        });
      });
  }

  await pptx.writeFile({ fileName: `${stripExtension(job.name)}.pptx` });
}

function splitSvgPaths(svg: string, batchSize: number) {
  const paths = Array.from(svg.matchAll(/<path\b[^>]*>/gi), (match) => match[0]);
  if (paths.length <= batchSize) return svg ? [svg] : [];
  const viewBox = svg.match(/viewBox="([^"]+)"/i)?.[1] ?? "0 0 100 100";
  const defs = Array.from(svg.matchAll(/<defs\b[\s\S]*?<\/defs>/gi), (match) => match[0]).join("");
  const styles = Array.from(svg.matchAll(/<style\b[\s\S]*?<\/style>/gi), (match) => match[0]).join("");
  const chunks: string[] = [];
  for (let index = 0; index < paths.length; index += batchSize) {
    chunks.push(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="${viewBox}">
${defs}
${styles}
${paths.slice(index, index + batchSize).join("\n")}
</svg>`);
  }
  return chunks.slice(0, 80);
}

async function buildCleanBackgroundDataUrl(job: ImageJob) {
  const image = await loadImage(job.dataUrl);
  const canvas = document.createElement("canvas");
  canvas.width = job.width;
  canvas.height = job.height;
  const context = canvas.getContext("2d");
  if (!context) return job.dataUrl;

  context.drawImage(image, 0, 0, job.width, job.height);

  job.assets
    .filter((asset) => asset.visible)
    .forEach((asset) => {
      fillRectWithLocalBackground(
        context,
        asset.sourceLeft,
        asset.sourceTop,
        asset.sourceWidth,
        asset.sourceHeight,
        averageAreaEdgeColor(context, asset.sourceLeft, asset.sourceTop, asset.sourceWidth, asset.sourceHeight),
      );
    });

  job.texts
    .filter(needsMask)
    .forEach((text) => {
      fillRectWithLocalBackground(
        context,
        text.sourceLeft - 2,
        text.sourceTop - 2,
        text.sourceWidth + 4,
        text.sourceHeight + 4,
        text.maskColor,
      );
    });

  return canvas.toDataURL("image/png");
}

async function buildTextReplacementBackgroundDataUrl(job: ImageJob, texts: TextLayer[]) {
  const image = await loadImage(job.dataUrl);
  const canvas = document.createElement("canvas");
  canvas.width = job.width;
  canvas.height = job.height;
  const context = canvas.getContext("2d");
  if (!context) return job.dataUrl;

  context.drawImage(image, 0, 0, job.width, job.height);
  try {
    await inpaintTextStrokes(context, job.width, job.height, texts);
  } catch (error) {
    texts.forEach((text) => eraseTextStrokesWithLocalColor(context, text, job.width, job.height));
    console.warn("OpenCV 鏂囧瓧鑳屾櫙淇澶辫触锛屼娇鐢ㄥ眬閮ㄥ钩鍧囪壊", error);
  }
  return canvas.toDataURL("image/png");
}

function eraseTextStrokesWithLocalColor(
  context: CanvasRenderingContext2D,
  text: TextLayer,
  imageWidth: number,
  imageHeight: number,
) {
  const pad = Math.max(2, Math.round(text.fontSize * 0.18));
  const left = Math.max(0, Math.floor(text.sourceLeft - pad));
  const top = Math.max(0, Math.floor(text.sourceTop - pad));
  const right = Math.min(imageWidth, Math.ceil(text.sourceLeft + text.sourceWidth + pad));
  const bottom = Math.min(imageHeight, Math.ceil(text.sourceTop + text.sourceHeight + pad));
  const width = right - left;
  const height = bottom - top;
  if (width <= 0 || height <= 0) return;

  const imageData = context.getImageData(left, top, width, height);
  const data = imageData.data;
  const bg = hexToRgb(averageAreaEdgeColor(context, left, top, width, height));

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const offset = (y * width + x) * 4;
      const red = data[offset];
      const green = data[offset + 1];
      const blue = data[offset + 2];
      const alpha = data[offset + 3];
      if (alpha < 20) continue;

      const luminance = red * 0.299 + green * 0.587 + blue * 0.114;
      const bgDiff = Math.abs(red - bg.r) + Math.abs(green - bg.g) + Math.abs(blue - bg.b);
      const coloredText = Math.max(red, green, blue) - Math.min(red, green, blue) > 45 && bgDiff > 70;
      if ((luminance < 205 && bgDiff > 48) || coloredText) {
        data[offset] = bg.r;
        data[offset + 1] = bg.g;
        data[offset + 2] = bg.b;
      }
    }
  }

  context.putImageData(imageData, left, top);
}

async function inpaintTextStrokes(
  context: CanvasRenderingContext2D,
  imageWidth: number,
  imageHeight: number,
  texts: TextLayer[],
) {
  if (texts.length === 0) return;
  const cv = await getOpenCv();
  const source = cv.imread(context.canvas);
  const rgb = new cv.Mat();
  const mask = cv.Mat.zeros(imageHeight, imageWidth, cv.CV_8UC1);
  const result = new cv.Mat();
  const kernel = cv.Mat.ones(3, 3, cv.CV_8U);

  try {
    texts.forEach((text) => addTextStrokeMask(context, mask, text, imageWidth, imageHeight));
    cv.dilate(mask, mask, kernel);
    cv.cvtColor(source, rgb, cv.COLOR_RGBA2RGB);
    cv.inpaint(rgb, mask, result, 3, cv.INPAINT_TELEA);
    cv.cvtColor(result, source, cv.COLOR_RGB2RGBA);
    cv.imshow(context.canvas, source);
  } finally {
    source.delete();
    rgb.delete();
    mask.delete();
    result.delete();
    kernel.delete();
  }
}

function addTextStrokeMask(
  context: CanvasRenderingContext2D,
  mask: any,
  text: TextLayer,
  imageWidth: number,
  imageHeight: number,
) {
  const pad = Math.max(2, Math.round(text.fontSize * 0.18));
  const left = Math.max(0, Math.floor(text.sourceLeft - pad));
  const top = Math.max(0, Math.floor(text.sourceTop - pad));
  const right = Math.min(imageWidth, Math.ceil(text.sourceLeft + text.sourceWidth + pad));
  const bottom = Math.min(imageHeight, Math.ceil(text.sourceTop + text.sourceHeight + pad));
  const width = right - left;
  const height = bottom - top;
  if (width <= 0 || height <= 0) return;

  const imageData = context.getImageData(left, top, width, height).data;
  const bg = hexToRgb(averageAreaEdgeColor(context, left, top, width, height));

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const offset = (y * width + x) * 4;
      const red = imageData[offset];
      const green = imageData[offset + 1];
      const blue = imageData[offset + 2];
      const alpha = imageData[offset + 3];
      if (alpha < 20) continue;

      const luminance = red * 0.299 + green * 0.587 + blue * 0.114;
      const bgDiff = Math.abs(red - bg.r) + Math.abs(green - bg.g) + Math.abs(blue - bg.b);
      const coloredText = Math.max(red, green, blue) - Math.min(red, green, blue) > 45 && bgDiff > 70;
      if ((luminance < 205 && bgDiff > 48) || coloredText) {
        mask.data[(top + y) * imageWidth + left + x] = 255;
      }
    }
  }
}

function normalizeBox(x1: number, y1: number, x2: number, y2: number) {
  return {
    left: Math.min(x1, x2),
    top: Math.min(y1, y2),
    width: Math.abs(x2 - x1),
    height: Math.abs(y2 - y1),
  };
}

function mergeBoxes(boxes: Array<{ left: number; top: number; right: number; bottom: number; count: number }>) {
  const result: Array<{ left: number; top: number; right: number; bottom: number; count: number }> = [];

  boxes.forEach((box) => {
    const existing = result.find((item) => boxesClose(item, box));
    if (!existing) {
      result.push({ ...box });
      return;
    }

    existing.left = Math.min(existing.left, box.left);
    existing.top = Math.min(existing.top, box.top);
    existing.right = Math.max(existing.right, box.right);
    existing.bottom = Math.max(existing.bottom, box.bottom);
    existing.count += box.count;
  });

  return result;
}

function boxesClose(
  a: { left: number; top: number; right: number; bottom: number },
  b: { left: number; top: number; right: number; bottom: number },
) {
  const gap = 10;
  return !(a.right + gap < b.left || b.right + gap < a.left || a.bottom + gap < b.top || b.bottom + gap < a.top);
}

function mergeAssetBoxes(
  boxes: Array<{ left: number; top: number; right: number; bottom: number; count: number }>,
  imageWidth: number,
  imageHeight: number,
) {
  const result: Array<{ left: number; top: number; right: number; bottom: number; count: number }> = [];
  const maxArea = imageWidth * imageHeight * 0.055;
  const maxWidth = imageWidth * 0.34;
  const maxHeight = imageHeight * 0.34;

  boxes
    .sort((a, b) => b.count - a.count)
    .forEach((box) => {
      let target = -1;
      let targetGap = Number.POSITIVE_INFINITY;

      for (let index = 0; index < result.length; index += 1) {
        const item = result[index];
        const gap = boxGapPx(item, box);
        const merged = {
          left: Math.min(item.left, box.left),
          top: Math.min(item.top, box.top),
          right: Math.max(item.right, box.right),
          bottom: Math.max(item.bottom, box.bottom),
          count: item.count + box.count,
        };
        const width = merged.right - merged.left;
        const height = merged.bottom - merged.top;
        const area = width * height;
        const sameRegion = gap <= 3 || (gap <= 8 && area <= maxArea * 0.45);
        if (sameRegion && area <= maxArea && width <= maxWidth && height <= maxHeight && gap < targetGap) {
          target = index;
          targetGap = gap;
        }
      }

      if (target < 0) {
        result.push({ ...box });
        return;
      }

      const item = result[target];
      item.left = Math.min(item.left, box.left);
      item.top = Math.min(item.top, box.top);
      item.right = Math.max(item.right, box.right);
      item.bottom = Math.max(item.bottom, box.bottom);
      item.count += box.count;
    });

  return result;
}

function boxGapPx(
  a: { left: number; top: number; right: number; bottom: number },
  b: { left: number; top: number; right: number; bottom: number },
) {
  const dx = Math.max(0, Math.max(a.left, b.left) - Math.min(a.right, b.right));
  const dy = Math.max(0, Math.max(a.top, b.top) - Math.min(a.bottom, b.bottom));
  return Math.hypot(dx, dy);
}

function makeBackgroundTransparent(
  context: CanvasRenderingContext2D,
  width: number,
  height: number,
  mode: "png" | "trace" = "trace",
) {
  const imageData = context.getImageData(0, 0, width, height);
  const data = imageData.data;
  const visited = new Uint8Array(width * height);
  const queue: number[] = [];
  const edgeColors = sampleEdgeColors(context, width, height);

  function isBackground(index: number) {
    const offset = index * 4;
    const red = data[offset];
    const green = data[offset + 1];
    const blue = data[offset + 2];
    const alpha = data[offset + 3];
    if (alpha < 20) return true;
    const max = Math.max(red, green, blue);
    const min = Math.min(red, green, blue);
    const luminance = red * 0.299 + green * 0.587 + blue * 0.114;
    const chroma = max - min;
    if (mode === "png") {
      if (luminance < 232) return false;
      return edgeColors.some((edgeColor) => {
        const edgeLuminance = edgeColor.r * 0.299 + edgeColor.g * 0.587 + edgeColor.b * 0.114;
        const diff = Math.abs(red - edgeColor.r) + Math.abs(green - edgeColor.g) + Math.abs(blue - edgeColor.b);
        return edgeLuminance > 220 && diff < 42 && (chroma < 34 || luminance > 246);
      });
    }
    if (luminance < 190) return false;
    if (chroma > 54 && luminance < 244) return false;
    return edgeColors.some((edgeColor) => {
      const diff = Math.abs(red - edgeColor.r) + Math.abs(green - edgeColor.g) + Math.abs(blue - edgeColor.b);
      return diff < 48 || (diff < 70 && luminance > 230 && chroma < 36);
    });
  }

  function push(index: number) {
    if (index < 0 || index >= visited.length || visited[index] || !isBackground(index)) return;
    visited[index] = 1;
    queue.push(index);
  }

  for (let x = 0; x < width; x += 1) {
    push(x);
    push((height - 1) * width + x);
  }
  for (let y = 0; y < height; y += 1) {
    push(y * width);
    push(y * width + width - 1);
  }

  for (let head = 0; head < queue.length; head += 1) {
    const current = queue[head];
    const x = current % width;
    const y = Math.floor(current / width);
    if (x > 0) push(current - 1);
    if (x < width - 1) push(current + 1);
    if (y > 0) push(current - width);
    if (y < height - 1) push(current + width);
  }

  for (let index = 0; index < visited.length; index += 1) {
    if (visited[index]) data[index * 4 + 3] = 0;
  }
  context.putImageData(imageData, 0, 0);
}

function sampleEdgeColors(context: CanvasRenderingContext2D, width: number, height: number) {
  const samples: Array<{ r: number; g: number; b: number }> = [];
  const stepX = Math.max(1, Math.floor(width / 10));
  const stepY = Math.max(1, Math.floor(height / 10));
  const add = (x: number, y: number) => {
    const pixel = context.getImageData(Math.max(0, Math.min(width - 1, x)), Math.max(0, Math.min(height - 1, y)), 1, 1).data;
    const luminance = pixel[0] * 0.299 + pixel[1] * 0.587 + pixel[2] * 0.114;
    if (pixel[3] >= 20 && luminance >= 172) samples.push({ r: pixel[0], g: pixel[1], b: pixel[2] });
  };

  for (let x = 0; x < width; x += stepX) {
    add(x, 0);
    add(x, height - 1);
  }
  for (let y = 0; y < height; y += stepY) {
    add(0, y);
    add(width - 1, y);
  }

  return samples.length > 0 ? samples : [hexToRgb(averageEdgeColor(context, width, height))];
}

function averageEdgeColor(context: CanvasRenderingContext2D, width: number, height: number) {
  const points: Array<[number, number]> = [];
  const stepX = Math.max(1, Math.floor(width / 8));
  const stepY = Math.max(1, Math.floor(height / 8));

  for (let x = 0; x < width; x += stepX) {
    points.push([x, 0], [x, height - 1]);
  }
  for (let y = 0; y < height; y += stepY) {
    points.push([0, y], [width - 1, y]);
  }

  let r = 0;
  let g = 0;
  let b = 0;
  points.forEach(([x, y]) => {
    const pixel = context.getImageData(x, y, 1, 1).data;
    r += pixel[0];
    g += pixel[1];
    b += pixel[2];
  });

  return rgbToHex(r / points.length, g / points.length, b / points.length);
}

function averageAreaEdgeColor(
  context: CanvasRenderingContext2D,
  left: number,
  top: number,
  width: number,
  height: number,
) {
  const canvasWidth = context.canvas.width;
  const canvasHeight = context.canvas.height;
  const points: Array<[number, number]> = [];
  const x0 = Math.max(0, Math.round(left));
  const y0 = Math.max(0, Math.round(top));
  const x1 = Math.min(canvasWidth - 1, Math.round(left + width));
  const y1 = Math.min(canvasHeight - 1, Math.round(top + height));
  const stepX = Math.max(1, Math.floor((x1 - x0 + 1) / 8));
  const stepY = Math.max(1, Math.floor((y1 - y0 + 1) / 8));

  for (let x = x0; x <= x1; x += stepX) {
    points.push([x, y0], [x, y1]);
  }
  for (let y = y0; y <= y1; y += stepY) {
    points.push([x0, y], [x1, y]);
  }

  let r = 0;
  let g = 0;
  let b = 0;
  points.forEach(([x, y]) => {
    const pixel = context.getImageData(x, y, 1, 1).data;
    r += pixel[0];
    g += pixel[1];
    b += pixel[2];
  });

  return rgbToHex(r / points.length, g / points.length, b / points.length);
}

function fillRectWithLocalBackground(
  context: CanvasRenderingContext2D,
  left: number,
  top: number,
  width: number,
  height: number,
  color: string,
) {
  context.save();
  context.fillStyle = color;
  context.fillRect(Math.round(left), Math.round(top), Math.round(width), Math.round(height));
  context.restore();
}

function rgbToHex(r: number, g: number, b: number) {
  return `#${[r, g, b]
    .map((value) => Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, "0"))
    .join("")}`;
}

function hexToRgb(hex: string) {
  const clean = hex.replace("#", "");
  return {
    r: parseInt(clean.slice(0, 2), 16),
    g: parseInt(clean.slice(2, 4), 16),
    b: parseInt(clean.slice(4, 6), 16),
  };
}

async function getOpenCv() {
  if (!openCvPromise) {
    openCvPromise = (async () => {
      const imported = await import("@techstark/opencv-js");
      const maybeModule: any = imported.default ?? imported;
      const cv = typeof maybeModule?.then === "function" ? await maybeModule : maybeModule;
      if (cv.Mat) return cv;
      await new Promise<void>((resolve) => {
        cv.onRuntimeInitialized = () => resolve();
      });
      return cv;
    })();
  }
  return openCvPromise;
}

async function getImageTracer() {
  if (!imageTracerPromise) {
    imageTracerPromise = import("imagetracerjs").then((module) => module.default ?? module);
  }
  return imageTracerPromise;
}

function readAsDataUrl(file: File) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function readImageSize(dataUrl: string) {
  return loadImage(dataUrl).then((image) => ({
    width: image.naturalWidth,
    height: image.naturalHeight,
  }));
}

function loadImage(src: string) {
  return new Promise<HTMLImageElement>((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("图片加载失败"));
    image.src = src;
  });
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function downloadDataUrl(dataUrl: string, fileName: string) {
  const link = document.createElement("a");
  link.href = dataUrl;
  link.download = fileName;
  link.click();
}

function stripExtension(fileName: string) {
  return fileName.replace(/\.[^.]+$/, "");
}

function escapeXml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export default App;
