# AI-Figure-Toolbox

中文名：钳工的美术箱

AI-Figure-Toolbox is a local web tool for converting AI-generated or raster scientific figures into editable research assets.

AI-Figure-Toolbox 是一个本地网页工具，用于把 AI 生成图、PNG/JPG 科研图转换为更方便编辑的论文素材，包括文本层、图片素材层、SVG 层、PNG 层和 PPTX。

## Features / 功能

- Import PNG/JPG scientific figures.
- 导入 PNG/JPG 科研图片。
- Mark, move, resize, and delete text recognition regions.
- 标记、移动、缩放、删除文字识别区域。
- Mark image asset regions separately from text regions.
- 将插图素材区域和文字区域分开标记。
- Export layered PPTX:
- 导出分层 PPTX：
  - Page 1: visual preview.
  - 第 1 页：原始视觉预览。
  - Page 2: editable text layer and cleaned background.
  - 第 2 页：可编辑文本层和清理后的背景。
  - Page 3: SVG assets, recommended for simple flat-color figures.
  - 第 3 页：SVG 素材，适合颜色简单、边界清晰的图。
  - Page 4: PNG assets, recommended for complex colorful figures.
  - 第 4 页：PNG 素材，适合复杂彩色插图。
- Optional Adobe Illustrator tracing/export workflow on Windows.
- Windows 下可选使用 Adobe Illustrator 进行描摹/导出。

## Important Usage Tip / 重要使用提醒

Text recognition regions and image asset regions should overlap as little as possible.

文字识别区和素材识别区尽量不要重合。

If the automatically detected regions are not accurate, adjust them manually before exporting. You can move, resize, or delete text regions and asset regions.

如果自动识别区域不准确，请在导出前手动调整。文本识别区和素材识别区都可以移动、缩放或删除。

Why this matters: overlapping regions may cause text to be erased from image assets, or image details to be incorrectly treated as text.

原因：区域重合可能导致文字从素材中擦除，或者插图细节被误当成文字处理。

## Network / Offline Notes / 联网与离线说明

The app now includes local Tesseract.js worker/core files and the English OCR language file. Basic English OCR and PPTX export can work without internet after dependencies are installed.

本项目已内置 Tesseract.js 的本地 worker/core 文件和英文 OCR 语言包。依赖安装完成后，基础英文 OCR 和 PPTX 导出可以在断网时使用。

Chinese OCR language data is not bundled by default to keep the repository smaller. If Chinese text recognition is required, internet access is recommended unless you manually add the corresponding local language data.

为了控制仓库体积，默认没有内置中文 OCR 语言包。如果需要中文识别，建议保持联网，或者手动加入对应的本地语言包。

First-time setup still requires internet because `pnpm install` needs to download npm dependencies.

首次安装仍然需要联网，因为 `pnpm install` 需要下载 npm 依赖。

Adobe Illustrator itself does not provide the web app runtime. Node.js and pnpm are still required.

Adobe Illustrator 只负责可选的描摹/导出流程，不提供网页应用运行环境；Node.js 和 pnpm 仍然是必需的。

## Requirements / 环境要求

Required / 必需：

- Node.js 18 or newer / Node.js 18 或更高版本
- npm, included with Node.js / npm，安装 Node.js 后自带
- pnpm

Optional / 可选：

- Adobe Illustrator, only needed for Illustrator tracing/export features.
- Adobe Illustrator，仅 Illustrator 描摹/导出功能需要。

Important / 重要：

Installing Adobe Illustrator alone is not enough. This project is a local web app. You must install Node.js and pnpm before starting it.

只安装 Adobe Illustrator 不够。本项目是本地网页应用，必须先安装 Node.js 和 pnpm 才能启动。

## First-Time Setup / 首次安装

1. Install Node.js LTS.

1. 安装 Node.js LTS 版本。

Download / 下载地址：

```text
https://nodejs.org/
```

During installation, keep the default options and make sure Node.js is added to PATH.

安装时保持默认选项，并确保 Node.js 被加入 PATH。

2. Close the current terminal and open a new PowerShell.

2. 关闭当前终端，重新打开一个新的 PowerShell。

3. Check Node.js and npm.

3. 检查 Node.js 和 npm 是否安装成功。

```powershell
node -v
npm -v
```

If both commands print version numbers, continue.

如果两个命令都能显示版本号，继续下一步。

4. Install pnpm.

4. 安装 pnpm。

```powershell
npm install -g pnpm
```

5. Check pnpm.

5. 检查 pnpm。

```powershell
pnpm -v
```

6. Install project dependencies.

6. 安装项目依赖。

```powershell
pnpm install
```

## Start the App / 启动软件

Option 1 / 方式一：double-click the startup script / 双击启动脚本：

```text
启动软件.bat
```

Option 2 / 方式二：run manually / 手动启动：

```powershell
pnpm dev
```

Then open / 然后打开：

```text
http://127.0.0.1:5173
```

## Build / 构建

```powershell
pnpm build
```

## Troubleshooting / 常见问题

### `npm` is not recognized / 无法识别 npm

If PowerShell shows:

如果 PowerShell 显示：

```text
npm is not recognized
无法将“npm”项识别为 cmdlet、函数、脚本文件或可运行程序的名称
```

It means Node.js is not installed correctly, or Node.js was not added to PATH.

这说明 Node.js 没有正确安装，或者 Node.js 没有加入 PATH。

Fix / 解决方法：

1. Install Node.js LTS from `https://nodejs.org/`.
2. Keep default installation options.
3. Close PowerShell.
4. Open a new PowerShell.
5. Run:

1. 从 `https://nodejs.org/` 安装 Node.js LTS。
2. 安装时保持默认选项。
3. 关闭 PowerShell。
4. 重新打开 PowerShell。
5. 执行：

```powershell
node -v
npm -v
```

Then install pnpm:

然后安装 pnpm：

```powershell
npm install -g pnpm
```

### `pnpm` is not recognized / 无法识别 pnpm

Install pnpm:

安装 pnpm：

```powershell
npm install -g pnpm
```

Then close and reopen PowerShell.

然后关闭并重新打开 PowerShell。

### `http://127.0.0.1:5173` cannot be opened / 无法打开网页

Run these commands in the project directory:

在项目目录中执行：

```powershell
pnpm install
pnpm dev
```

Then open:

然后打开：

```text
http://127.0.0.1:5173
```

### Illustrator export fails / Illustrator 导出失败

- Make sure Adobe Illustrator is installed.
- 确认已安装 Adobe Illustrator。
- Make sure Illustrator can be opened normally.
- 确认 Illustrator 可以正常打开。
- Try opening Illustrator once before exporting from the web app.
- 导出前可以先手动打开一次 Illustrator。

## Notes / 说明

SVG tracing works best for simple figures with limited colors and sharp boundaries.

SVG 更适合颜色少、边界清晰的简单图。

For complex colorful illustrations, PNG assets usually preserve visual fidelity better than aggressive SVG vectorization. The exported PPTX keeps SVG and PNG on separate pages so users can choose the better result manually.

对于复杂彩色插图，PNG 通常比强行 SVG 矢量化更能保留视觉效果。导出的 PPTX 会把 SVG 和 PNG 放在不同页面，方便用户自行选择更好的结果。
