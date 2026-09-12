<div align="center">

# Acacia Flow

**Offline file conversion for Windows**

[![Latest Release](https://img.shields.io/github/v/release/Chengyi7888/Acacia-flow?display_name=tag&label=latest%20release)](https://github.com/Chengyi7888/Acacia-flow/releases)
[![License](https://img.shields.io/github/license/Chengyi7888/Acacia-flow)](LICENSE)

<a href="#english">English</a> · <a href="#中文">中文</a>

</div>

## English

<a id="english"></a>

Acacia Flow is a Windows desktop application for converting documents, images, PDFs, audio, and video files locally. It provides drag-and-drop conversion, a conversion queue, duplicate-name handling, output-folder selection, and a set of offline utility tools.

### Download

Download the installer from the [latest GitHub Release](https://github.com/Chengyi7888/Acacia-flow/releases/latest).

The Windows installer includes the required conversion runtimes. Users do not need to install Node.js, FFmpeg, LibreOffice, developer plugins, or a cloud service separately.

### Main Features

- Image conversion: PNG, JPG, WEBP, GIF, AVIF, TIFF, BMP, and PDF.
- Audio conversion: MP3, WAV, FLAC, M4A, AAC, OGG, OPUS, and WMA.
- Video conversion: MP4, MOV, MKV, WEBM, GIF, audio extraction, and frame snapshots.
- Document and table conversion: TXT, MD, HTML, CSV, RTF, DOC, DOCX, XLS, XLSX, PPT, PPTX, PDF, and related formats.
- PDF text extraction and page rendering to image formats.
- Offline conversion of legacy Office files through the bundled LibreOffice runtime.
- Conversion queue, output folder selection, duplicate-name handling, and tray behavior.
- Built-in regular-expression, JSON, and Base64/URL utility tools.

### Offline Runtime

The installer bundles FFmpeg and a portable LibreOffice runtime. File processing is performed locally; normal conversion workflows do not upload user files.

### Build From Source

```text
npm install
npm run start
npm run test:smoke
npm run dist
```

The source tree contains `vendor/libreoffice`, which is copied into packaged application resources by the Electron Builder configuration.

### Project Layout

```text
assets/             Application icons, backgrounds, and assistant images
scripts/            Development, packaging, and smoke-test scripts
src/main/           Electron main process, preload layer, and conversion engine
web/                User interface and bundled browser libraries
vendor/libreoffice/ Portable Office conversion runtime
```

### License

The project is released under the MIT License. Bundled third-party components remain subject to their respective licenses.

## 中文

<a id="中文"></a>

Acacia Flow 是一款面向 Windows 的本地文件格式转换软件，支持文档、图片、PDF、音频和视频的拖拽转换，并提供转换队列、输出目录选择、同名文件处理和离线工具。

### 下载

请前往 [GitHub 最新 Release](https://github.com/Chengyi7888/Acacia-flow/releases/latest) 下载 Windows 安装包。

安装包已经包含运行所需的转换引擎，用户不需要另外安装 Node.js、FFmpeg、LibreOffice、本地插件或云服务。

### 主要功能

- 图片转换：PNG、JPG、WEBP、GIF、AVIF、TIFF、BMP、PDF。
- 音频转换：MP3、WAV、FLAC、M4A、AAC、OGG、OPUS、WMA。
- 视频转换：MP4、MOV、MKV、WEBM、GIF、音频提取和视频截图。
- 文档与表格转换：TXT、MD、HTML、CSV、RTF、DOC、DOCX、XLS、XLSX、PPT、PPTX、PDF 等。
- PDF 文本提取和页面渲染。
- 通过内置 LibreOffice 引擎离线处理老式 Office 文件。
- 转换队列、输出目录、同名文件处理和托盘行为设置。
- 内置正则表达式、JSON、Base64/URL 编解码工具。

### 离线运行

安装包内置 FFmpeg 和便携版 LibreOffice，文件转换在本机完成，正常使用不会上传用户文件。

### 开源与许可

本项目使用 MIT License。安装包中包含的第三方组件仍遵循其各自的许可协议。
