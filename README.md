# Acacia Flow

Acacia Flow 是一个面向 Windows 的本地离线文件格式转换工具。它基于 Electron 构建，提供图片、音频、视频、PDF、文本和常见文档格式的转换入口，并带有一个会随转换状态变化的小助手界面。

## 特性

- 本地离线运行，不依赖 Codex、ChatGPT、本机插件或云端服务。
- 支持常见图片格式互转，并支持图片生成 PDF。
- 支持音频和视频格式转换，内置 FFmpeg 二进制依赖。
- 支持 TXT、MD、HTML、CSV、RTF、DOCX、XLSX、PDF 等文本/表格类转换。
- 支持 PDF 文本提取和 PDF 到 TXT/MD/HTML/DOCX 的常用转换流程。
- 支持转换队列、右键删除、输出目录选择、同名文件处理和托盘关闭行为。
- 安装包会创建桌面快捷方式和开始菜单快捷方式。

## 下载使用

普通用户建议下载 Release 页面中的 Windows 安装包：

```text
Acacia Flow Setup 0.1.0.exe
```

双击安装后即可使用。应用安装后会在本机离线运行。

## 开发环境

需要安装：

- Node.js 20 或更高版本
- npm
- Windows 10/11

安装依赖：

```bash
npm install
```

本地启动：

```bash
npm start
```

运行核心转换烟测：

```bash
npm run test:smoke
```

生成 Windows 安装包：

```bash
npm run dist
```

构建完成后，安装包会输出到项目根目录：

```text
Acacia Flow Setup 0.1.0.exe
```

构建过程中产生的临时目录会自动清理。

## 项目结构

```text
assets/                 应用图标、背景图、小助手状态图
scripts/                启动、构建、烟测脚本
src/main/               Electron 主进程、预加载脚本、转换引擎
web/                    前端界面
web/vendor/             离线浏览器端依赖
package.json            项目配置和打包配置
```

## 转换能力说明

内置转换能力覆盖常见场景，但部分 Office 专有格式转换依赖 LibreOffice。若目标机器安装了 LibreOffice，Acacia Flow 会自动检测并调用它完成 DOC/DOCX/PPT/PPTX/XLS/XLSX/ODT/ODS/ODP 等更完整的 Office 转换。

不依赖 LibreOffice 的常用路径包括：

- TXT/MD/HTML/CSV/RTF 到 TXT、MD、HTML、CSV、RTF
- TXT/MD/HTML/CSV/RTF 到 DOCX、XLSX、PDF
- PDF 到 TXT、MD、HTML、DOCX
- 图片到 PNG、JPG、WEBP、GIF、AVIF、TIFF、BMP、PDF
- 音频到 MP3、WAV、FLAC、M4A、AAC、OGG、OPUS、WMA
- 视频到 MP4、MOV、MKV、WEBM、GIF，以及提取音频到常见音频格式

## 隐私和离线

Acacia Flow 的转换逻辑在本地执行，不会上传用户文件。项目源码和安装包不包含开发者本机的 Codex 配置、`.codex`、`.agents`、桌面路径或本地插件路径。

## 许可证

本项目使用 MIT License 开源。第三方依赖遵循其各自许可证。
