const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const { execFile, execFileSync } = require("child_process");
const os = require("os");
const { BrowserWindow } = require("electron");
const JSZip = require("../../web/vendor/jszip.min.js");
const XLSX = require("../../web/vendor/xlsx.full.min.js");
const docx = require("../../web/vendor/docx.umd.js");
const sharp = require("sharp");
const { PDFDocument } = require("pdf-lib");

const textTargets = new Set(["txt", "md", "html", "csv", "rtf"]);
const imageInputs = new Set(["jpg", "jpeg", "png", "webp", "gif", "avif", "tif", "tiff", "bmp"]);
const imageTargets = new Set(["png", "jpg", "webp", "gif", "avif", "tiff", "bmp", "pdf"]);
const audioInputs = new Set(["mp3", "wav", "flac", "m4a", "aac", "ogg", "opus", "wma"]);
const videoInputs = new Set(["mp4", "mov", "mkv", "webm", "avi", "m4v", "wmv", "flv"]);
const audioTargets = new Set(["mp3", "wav", "flac", "m4a", "aac", "ogg", "opus", "wma"]);
const videoTargets = new Set(["mp4", "mov", "mkv", "webm", "avi", "m4v", "wmv", "flv", "gif"]);
const officeInputs = new Set(["doc", "docx", "ppt", "pptx", "xls", "xlsx", "odt", "ods", "odp", "rtf"]);
const officeTargets = new Set(["pdf", "doc", "docx", "pptx", "xlsx", "xls", "odt", "ods", "odp", "txt", "html", "rtf"]);
const nativelyReadableOfficeInputs = new Set(["docx", "pptx", "xls", "xlsx", "rtf"]);
const libreOfficeOnlyTargets = new Set(["doc", "pptx", "odt", "ods", "odp"]);
const readable = new Set([
  "txt", "md", "html", "htm", "csv", "tsv", "rtf", "doc", "docx", "ppt", "pptx", "xlsx", "xls", "odt", "ods", "odp", "pdf",
  ...imageInputs, ...audioInputs, ...videoInputs
]);
const writable = new Set([
  ...textTargets, "doc", "docx", "xls", "xlsx", "pdf",
  ...imageTargets, ...audioTargets, ...videoTargets
]);

function extname(filePath) {
  const ext = path.extname(filePath).replace(/^\./, "").toLowerCase();
  if (ext === "jpeg") return "jpg";
  if (ext === "tif") return "tiff";
  if (ext === "htm") return "html";
  return ext;
}

function commandPath(name, fallback) {
  const candidates = [
    process.env[`ACACIA_${name.toUpperCase()}_PATH`],
    fallback,
    name
  ].filter(Boolean);
  return candidates.find(candidate => candidate === name || fs.existsSync(candidate)) || name;
}

function run(command, args, timeout = 1000 * 60 * 30) {
  return new Promise((resolve, reject) => {
    execFile(command, args, { timeout, windowsHide: true }, (error, stdout, stderr) => {
      if (error) {
        const detail = String(stderr || stdout || error.message).trim();
        reject(new Error(detail || error.message));
        return;
      }
      resolve({ stdout, stderr });
    });
  });
}

function ffmpegPath() {
  let bundled = "";
  try {
    bundled = require("@ffmpeg-installer/ffmpeg").path;
  } catch (_error) {
    // The optional binary may be absent in a source-only installation.
  }
  if (bundled.includes(".asar")) bundled = bundled.replace(/\.asar([\\/])/, ".asar.unpacked$1");
  return commandPath("ffmpeg", bundled);
}

function libreOfficePath() {
  const candidates = [
    process.env.ACACIA_LIBREOFFICE_PATH,
    process.resourcesPath && path.join(process.resourcesPath, "libreoffice", "program", "soffice.exe"),
    process.resourcesPath && path.join(process.resourcesPath, "libreoffice", "program", "soffice.com"),
    process.resourcesPath && path.join(process.resourcesPath, "vendor", "libreoffice", "program", "soffice.exe"),
    process.resourcesPath && path.join(process.resourcesPath, "vendor", "libreoffice", "program", "soffice.com"),
    process.resourcesPath && path.join(process.resourcesPath, "app.asar.unpacked", "vendor", "libreoffice", "program", "soffice.exe"),
    process.resourcesPath && path.join(process.resourcesPath, "app.asar.unpacked", "vendor", "libreoffice", "program", "soffice.com"),
    path.join(__dirname, "..", "..", "vendor", "libreoffice", "program", "soffice.exe"),
    path.join(__dirname, "..", "..", "vendor", "libreoffice", "program", "soffice.com"),
    path.join(__dirname, "..", "..", "bin", "libreoffice", "program", "soffice.exe"),
    path.join(__dirname, "..", "..", "bin", "libreoffice", "program", "soffice.com"),
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe"
  ].filter(Boolean);
  const local = candidates.find(candidate => candidate !== "soffice" && fs.existsSync(candidate));
  if (local) return local;
  try {
    const command = process.platform === "win32" ? "where.exe" : "which";
    const output = execFileSync(command, [process.platform === "win32" ? "soffice.exe" : "soffice"], {
      encoding: "utf8",
      windowsHide: true
    }).trim();
    return output.split(/\r?\n/).find(Boolean) || "";
  } catch (_error) {
    return "";
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function decodeXml(value) {
  return String(value)
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function cleanText(value) {
  return String(value)
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function decodeTextBuffer(buffer) {
  if (buffer.length >= 3 && buffer.subarray(0, 3).equals(Buffer.from([0xef, 0xbb, 0xbf]))) {
    return buffer.subarray(3).toString("utf8");
  }
  if (buffer.length >= 2 && buffer.subarray(0, 2).equals(Buffer.from([0xff, 0xfe]))) {
    return buffer.subarray(2).toString("utf16le");
  }
  const utf8 = buffer.toString("utf8");
  if (!utf8.includes("\ufffd")) return utf8;
  return new TextDecoder("gb18030").decode(buffer);
}

function stripHtml(html) {
  return String(html)
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|tr|h[1-6]|li|table)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function stripRtf(rtf) {
  return String(rtf)
    .replace(/\\u(-?\d+)\??/g, (_match, code) => String.fromCharCode(Number(code)))
    .replace(/\\par[d]?/g, "\n")
    .replace(/\\'[0-9a-fA-F]{2}/g, "")
    .replace(/\\[a-zA-Z]+-?\d* ?/g, "")
    .replace(/[{}]/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function parseCsv(text) {
  return String(text).split(/\r?\n/).filter(line => line.length).map(line => {
    const cells = [];
    let cell = "";
    let quoted = false;
    for (let i = 0; i < line.length; i += 1) {
      const char = line[i];
      if (char === '"' && line[i + 1] === '"') {
        cell += '"';
        i += 1;
      } else if (char === '"') {
        quoted = !quoted;
      } else if (char === "," && !quoted) {
        cells.push(cell.trim());
        cell = "";
      } else {
        cell += char;
      }
    }
    cells.push(cell.trim());
    return cells;
  });
}

function textToRows(text) {
  return String(text).split(/\r?\n/).filter(Boolean).map(line => line.includes("\t") ? line.split("\t") : [line]);
}

function tableToCsv(rows) {
  return rows.map(row => row.map(cell => {
    const value = cell == null ? "" : String(cell);
    return /[",\n]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value;
  }).join(",")).join("\n");
}

function sourceToText(data) {
  if (data.kind === "table") return String(data.rows.map(row => row.join("\t")).join("\n"))
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");
  if (data.kind === "html") return String(stripHtml(data.html))
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");
  return String(data.text || "").replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");
}

function tableToMarkdown(rows) {
  if (!rows.length) return "";
  const escaped = rows.map(row => row.map(cell => String(cell == null ? "" : cell).replaceAll("|", "\\|").replace(/\r?\n/g, "<br>")));
  const columns = Math.max(...escaped.map(row => row.length), 1);
  const fill = row => Array.from({ length: columns }, (_value, index) => row[index] || "");
  const [header, ...body] = escaped.map(fill);
  return [`| ${header.join(" | ")} |`, `| ${header.map(() => "---").join(" | ")} |`, ...body.map(row => `| ${row.join(" | ")} |`)].join("\n");
}

function tableToHtml(rows) {
  const body = rows.map(row => `<tr>${row.map(cell => `<td>${escapeHtml(cell == null ? "" : cell)}</td>`).join("")}</tr>`).join("\n");
  return `<!doctype html><html><head><meta charset="utf-8"><style>table{border-collapse:collapse;font-family:"Microsoft YaHei",Arial,sans-serif}td{border:1px solid #aeb6c7;padding:6px 9px;vertical-align:top}</style></head><body><table>${body}</table></body></html>`;
}

function sourceToRows(data) {
  return data.kind === "table" ? data.rows : textToRows(sourceToText(data));
}

function resolveOutputPath(folder, base, target, conflictMode = "rename") {
  const firstPath = path.join(folder, `${base}.${target}`);
  const pathTaken = candidate => {
    const candidateName = path.basename(candidate).toLowerCase();
    try {
      return fs.readdirSync(folder).some(name => name.toLowerCase() === candidateName);
    } catch (_error) {
      return fs.existsSync(candidate);
    }
  };
  const pagePathTaken = candidate => pathTaken(candidate.replace(/\.[^.]+$/, `_page_1.${target}`));
  if (conflictMode === "overwrite" || (!pathTaken(firstPath) && !pagePathTaken(firstPath))) {
    return { outPath: firstPath, skipped: false };
  }
  if (conflictMode === "skip") return { outPath: firstPath, skipped: true };
  for (let index = 1; index < 10000; index += 1) {
    const candidate = path.join(folder, `${base} (${index}).${target}`);
    if (!pathTaken(candidate) && !pagePathTaken(candidate)) {
      return { outPath: candidate, skipped: false };
    }
  }
  throw new Error("同名文件太多，无法自动重命名。");
}

async function extractDocx(filePath) {
  const zip = await JSZip.loadAsync(await fsp.readFile(filePath));
  const names = Object.keys(zip.files)
    .filter(name => /^word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/.test(name))
    .sort((a, b) => a === "word/document.xml" ? -1 : b === "word/document.xml" ? 1 : a.localeCompare(b));
  const blocks = [];
  for (const name of names) {
    const xml = await zip.files[name].async("text");
    const paragraphs = xml.split(/<\/w:p>/).map(part => cleanText(
      [...part.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
        .map(match => decodeXml(match[1])).join("")
    )).filter(Boolean);
    blocks.push(...paragraphs);
  }
  return { kind: "text", text: cleanText(blocks.join("\n")) };
}

async function extractXlsx(filePath) {
  const workbook = XLSX.read(await fsp.readFile(filePath), { type: "buffer", cellDates: false });
  const rows = [];
  workbook.SheetNames.forEach((sheetName, index) => {
    if (index > 0) rows.push([]);
    rows.push([sheetName]);
    rows.push(...XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, blankrows: false }));
  });
  return { kind: "table", rows };
}

async function extractPptx(filePath) {
  const zip = await JSZip.loadAsync(await fsp.readFile(filePath));
  const slideNames = Object.keys(zip.files)
    .filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => Number(a.match(/slide(\d+)/)[1]) - Number(b.match(/slide(\d+)/)[1]));
  const lines = [];
  for (const name of slideNames) {
    const xml = await zip.files[name].async("text");
    const text = cleanText([...xml.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)]
      .map(match => decodeXml(match[1])).join(" "));
    if (text) lines.push(text);
  }
  return { kind: "text", text: cleanText(lines.join("\n\n")) };
}

async function extractPdf(filePath) {
  try {
    const vendorDir = path.join(__dirname, "..", "..", "web", "vendor");
    const pdfjs = await import(pathToFileUrl(path.join(vendorDir, "pdf.min.mjs")));
    pdfjs.GlobalWorkerOptions.workerSrc = pathToFileUrl(path.join(vendorDir, "pdf.worker.min.mjs"));
    const pdfBytes = await fsp.readFile(filePath);
    const pdf = await pdfjs.getDocument({ data: new Uint8Array(pdfBytes) }).promise;
    const pages = [];
    for (let index = 1; index <= pdf.numPages; index += 1) {
      const page = await pdf.getPage(index);
      const content = await page.getTextContent();
      pages.push(content.items.map(item => item.str).join(" "));
    }
    return { kind: "text", text: cleanText(pages.join("\n\n")) };
  } catch (error) {
    throw new Error(`PDF 文本解析失败：${error.message}`);
  }
}

function pathToFileUrl(filePath) {
  return require("url").pathToFileURL(filePath).href;
}

async function extractSource(filePath) {
  const ext = extname(filePath);
  if (officeInputs.has(ext) && ext !== "docx" && ext !== "xlsx" && ext !== "xls" && ext !== "pptx" && ext !== "rtf") {
    const tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), "acacia-office-text-"));
    const tempDocx = path.join(tempDir, "source.docx");
    try {
      await convertWithLibreOffice(filePath, tempDocx, "docx");
      return await extractDocx(tempDocx);
    } finally {
      await fsp.rm(tempDir, { recursive: true, force: true });
    }
  }
  if (ext === "docx") return extractDocx(filePath);
  if (ext === "xlsx" || ext === "xls") return await extractXlsx(filePath);
  if (ext === "pptx") return extractPptx(filePath);
  if (ext === "pdf") return extractPdf(filePath);
  const text = decodeTextBuffer(await fsp.readFile(filePath));
  if (ext === "csv") return { kind: "table", rows: parseCsv(text) };
  if (ext === "tsv") return { kind: "table", rows: text.split(/\r?\n/).filter(Boolean).map(line => line.split("\t")) };
  if (ext === "rtf") return { kind: "text", text: stripRtf(text) };
  if (ext === "html") return { kind: "html", html: text, text: stripHtml(text) };
  return { kind: "text", text };
}

function renderTextLike(data, target) {
  const text = sourceToText(data);
  if (target === "txt") return text;
  if (target === "md") return data.kind === "table" ? tableToMarkdown(data.rows) : text;
  if (target === "html") return data.kind === "table" ? tableToHtml(data.rows) : `<!doctype html><html><head><meta charset="utf-8"></head><body><pre>${escapeHtml(text)}</pre></body></html>`;
  if (target === "csv") return tableToCsv(sourceToRows(data));
  if (target === "rtf") return `{\\rtf1\\ansi\\ansicpg65001\\deff0\n${text.replace(/[\\{}]/g, match => `\\${match}`).replace(/\r?\n/g, "\\par\n")}\n}`;
  throw new Error(`暂不支持导出为 ${target}`);
}

function encodeBmp(rawData, info) {
  const rowBytes = Math.ceil(info.width * 3 / 4) * 4;
  const pixels = Buffer.alloc(rowBytes * info.height);
  for (let y = 0; y < info.height; y += 1) {
    const sourceY = info.height - y - 1;
    for (let x = 0; x < info.width; x += 1) {
      const sourceOffset = (sourceY * info.width + x) * info.channels;
      const targetOffset = y * rowBytes + x * 3;
      pixels[targetOffset] = rawData[sourceOffset + 2] || 0;
      pixels[targetOffset + 1] = rawData[sourceOffset + 1] || 0;
      pixels[targetOffset + 2] = rawData[sourceOffset] || 0;
    }
  }
  const header = Buffer.alloc(54);
  header.write("BM", 0, 2, "ascii");
  header.writeUInt32LE(54 + pixels.length, 2);
  header.writeUInt32LE(54, 10);
  header.writeUInt32LE(40, 14);
  header.writeInt32LE(info.width, 18);
  header.writeInt32LE(info.height, 22);
  header.writeUInt16LE(1, 26);
  header.writeUInt16LE(24, 28);
  header.writeUInt32LE(pixels.length, 34);
  return Buffer.concat([header, pixels]);
}

async function encodeBmpFromImage(input) {
  const { data, info } = await sharp(input)
    .flatten({ background: "#ffffff" })
    .removeAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });
  return encodeBmp(data, {
    width: info.width,
    height: info.height,
    channels: info.channels
  });
}

async function writeDocx(data, outPath) {
  const { Document, Packer, Paragraph, TextRun } = docx;
  const children = sourceToText(data).split(/\r?\n/)
    .map(line => new Paragraph({ children: [new TextRun(line || " ")] }));
  const document = new Document({ sections: [{ children }] });
  await fsp.writeFile(outPath, await Packer.toBuffer(document));
}

async function convertWithLibreOffice(sourcePath, outPath, target) {
  const executable = libreOfficePath();
  if (!executable) {
    throw new Error("此转换需要 LibreOffice，但当前电脑未检测到 LibreOffice。请安装 LibreOffice 后重试。");
  }
  const tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), "acacia-office-"));
  const filter = target === "doc" ? "doc:MS Word 97" :
    target === "xls" ? "xls:MS Excel 97" :
    target === "xlsx" ? "xlsx:Calc MS Excel 2007 XML" :
    target === "docx" ? "docx:Office Open XML Text" :
    target === "pptx" ? "pptx:Impress MS PowerPoint 2007 XML" :
    target === "html" ? "html:XHTML Writer File" :
    target === "rtf" ? "rtf:Rich Text Format" :
    target;
  const base = path.basename(sourcePath).replace(/\.[^.]+$/, "");
  const generated = path.join(tempDir, `${base}.${target}`);
  try {
    await run(executable, [
      "--headless",
      "--convert-to", filter,
      "--outdir", tempDir,
      sourcePath
    ], 1000 * 60 * 5);
    if (!fs.existsSync(generated)) {
      throw new Error(`LibreOffice 未生成预期的 ${target.toUpperCase()} 文件。`);
    }
    await fsp.copyFile(generated, outPath);
  } finally {
    await fsp.rm(tempDir, { recursive: true, force: true });
  }
}

async function writeDoc(data, outPath) {
  const tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), "acacia-doc-"));
  const tempDocx = path.join(tempDir, "source.docx");
  try {
    await writeDocx(data, tempDocx);
    await convertWithLibreOffice(tempDocx, outPath, "doc");
  } finally {
    await fsp.rm(tempDir, { recursive: true, force: true });
  }
}

function writeXlsx(data, outPath, bookType) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(sourceToRows(data)), "Sheet1");
  fs.writeFileSync(outPath, XLSX.write(workbook, { bookType, type: "buffer" }));
}

async function writePdf(data, outPath) {
  const win = new BrowserWindow({ show: false, webPreferences: { offscreen: true } });
  try {
    const html = `<!doctype html><meta charset="utf-8"><style>body{font-family:"Microsoft YaHei",Arial,sans-serif;white-space:pre-wrap;line-height:1.6;margin:40px}</style><body>${escapeHtml(sourceToText(data))}</body>`;
    await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    await fsp.writeFile(outPath, await win.webContents.printToPDF({ printBackground: true, pageSize: "A4" }));
  } finally {
    win.destroy();
  }
}

async function writeSourcePdf(sourcePath, outPath) {
  const ext = extname(sourcePath);
  if (officeInputs.has(ext) && ext !== "docx" && ext !== "xlsx" && ext !== "xls" && ext !== "pptx" && ext !== "rtf") {
    await convertWithLibreOffice(sourcePath, outPath, "pdf");
    return;
  }
  if (["doc", "docx", "ppt", "pptx", "xls", "xlsx", "odt", "ods", "odp"].includes(ext)) {
    await convertWithLibreOffice(sourcePath, outPath, "pdf");
    return;
  }
  const source = await fsp.readFile(sourcePath);
  const win = new BrowserWindow({ show: false, webPreferences: { offscreen: true } });
  try {
    if (ext === "html") {
      const tempHtml = path.join(os.tmpdir(), `acacia-source-${Date.now()}-${Math.random().toString(16).slice(2)}.html`);
      await fsp.writeFile(tempHtml, source);
      await win.loadFile(tempHtml);
      await fsp.writeFile(outPath, await win.webContents.printToPDF({ printBackground: true, pageSize: "A4" }));
      await fsp.rm(tempHtml, { force: true }).catch(() => {});
      return;
    }
    const text = decodeTextBuffer(source);
    const html = `<!doctype html><meta charset="utf-8"><style>
      @page{size:A4;margin:18mm}body{font-family:"Microsoft YaHei",Arial,sans-serif;
      white-space:pre-wrap;word-break:break-word;line-height:1.6;margin:0}
    </style><body>${escapeHtml(text)}</body>`;
    await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    await fsp.writeFile(outPath, await win.webContents.printToPDF({ printBackground: true, pageSize: "A4" }));
  } finally {
    win.destroy();
  }
}

function dataUrlToBuffer(dataUrl) {
  return Buffer.from(String(dataUrl).split(",")[1] || "", "base64");
}

async function renderPdfPagesToImages(pdfPath, target, scale = 1.5) {
  const vendorDir = path.join(__dirname, "..", "..", "web", "vendor");
  const pdfModuleUrl = pathToFileUrl(path.join(vendorDir, "pdf.min.mjs"));
  const workerModuleUrl = pathToFileUrl(path.join(vendorDir, "pdf.worker.min.mjs"));
  const pdfBytes = (await fsp.readFile(pdfPath)).toString("base64");
  const win = new BrowserWindow({
    show: false,
    webPreferences: {
      contextIsolation: false,
      nodeIntegration: false,
      sandbox: false
    }
  });
  const html = `<!doctype html>
    <meta charset="utf-8">
    <style>html,body{margin:0;background:#fff}canvas{display:block}</style>
    <script type="module">
      import * as pdfjsLib from ${JSON.stringify(pdfModuleUrl)};
      pdfjsLib.GlobalWorkerOptions.workerSrc = ${JSON.stringify(workerModuleUrl)};
      window.__acaciaRenderPdf = async (base64, renderScale) => {
        const binary = atob(base64);
        const bytes = Uint8Array.from(binary, char => char.charCodeAt(0));
        const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
        const pages = [];
        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
          const page = await pdf.getPage(pageNumber);
          const viewport = page.getViewport({ scale: renderScale });
          const canvas = document.createElement("canvas");
          canvas.width = Math.ceil(viewport.width);
          canvas.height = Math.ceil(viewport.height);
          await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
          pages.push(canvas.toDataURL("image/png"));
        }
        return pages;
      };
    </script>`;
  const tempHtml = path.join(os.tmpdir(), `acacia-pdf-render-${Date.now()}-${Math.random().toString(16).slice(2)}.html`);
  await fsp.writeFile(tempHtml, html, "utf8");
  try {
    await win.loadFile(tempHtml);
    await win.webContents.executeJavaScript("new Promise(resolve => { const wait = () => window.__acaciaRenderPdf ? resolve(true) : setTimeout(wait, 20); wait(); })");
    const pageDataUrls = await win.webContents.executeJavaScript(
      `window.__acaciaRenderPdf(${JSON.stringify(pdfBytes)}, ${Number(scale) || 1.5})`
    );
    const extension = target === "jpg" ? "jpg" : target;
    const pages = [];
    for (const dataUrl of pageDataUrls) {
      const pngBuffer = dataUrlToBuffer(dataUrl);
      const image = sharp(pngBuffer);
      let buffer = pngBuffer;
      if (target === "jpg") buffer = await image.jpeg({ quality: 92, background: "#ffffff" }).flatten().toBuffer();
      if (target === "webp") buffer = await image.webp({ quality: 92 }).toBuffer();
      if (target === "tiff") buffer = await image.tiff({ compression: "lzw" }).toBuffer();
      if (target === "bmp") {
        buffer = await encodeBmpFromImage(pngBuffer);
      }
      pages.push({
        buffer,
        extension
      });
    }
    return pages;
  } finally {
    win.destroy();
    await fsp.rm(tempHtml, { force: true }).catch(() => {});
  }
}

async function writePdfAsDocx(pdfPath, outPath) {
  const { Document, Packer, Paragraph, ImageRun, PageBreak } = docx;
  const pages = await renderPdfPagesToImages(pdfPath, "png", 2);
  if (!pages.length) throw new Error("没有渲染出可写入 Word 的页面。");

  const children = [];
  for (let index = 0; index < pages.length; index += 1) {
    const metadata = await sharp(pages[index].buffer).metadata();
    const width = metadata.width || 1;
    const height = metadata.height || 1;
    const fit = Math.min(560 / width, 740 / height);
    children.push(new Paragraph({
      children: [new ImageRun({
        data: pages[index].buffer,
        type: "png",
        transformation: {
          width: Math.max(1, Math.round(width * fit)),
          height: Math.max(1, Math.round(height * fit))
        }
      })]
    }));
    if (index < pages.length - 1) children.push(new Paragraph({ children: [new PageBreak()] }));
  }
  const document = new Document({ sections: [{ children }] });
  await fsp.writeFile(outPath, await Packer.toBuffer(document));
}

async function writePdfAsHtml(pdfPath, outPath) {
  const pages = await renderPdfPagesToImages(pdfPath, "png", 1.5);
  if (!pages.length) throw new Error("没有渲染出可写入网页的页面。");
  const content = pages.map((page, index) => `<img alt="第 ${index + 1} 页" src="data:image/png;base64,${page.buffer.toString("base64")}">`).join("\n");
  const html = `<!doctype html><html><head><meta charset="utf-8"><style>
    body{margin:0;background:#edf0f5;padding:24px;font-family:"Microsoft YaHei",Arial,sans-serif}
    img{display:block;width:min(100%,900px);height:auto;margin:0 auto 24px;background:#fff;box-shadow:0 8px 28px rgba(0,0,0,.16)}
  </style></head><body>${content}</body></html>`;
  await fsp.writeFile(outPath, html, "utf8");
}

async function writeVisualDocxFromSource(sourcePath, outPath) {
  const tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), "acacia-visual-docx-"));
  const tempPdf = path.join(tempDir, "source.pdf");
  try {
    await writeSourcePdf(sourcePath, tempPdf);
    await writePdfAsDocx(tempPdf, outPath);
  } finally {
    await fsp.rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}

async function writeVisualDocFromSource(sourcePath, outPath) {
  const tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), "acacia-visual-doc-"));
  const tempDocx = path.join(tempDir, "source.docx");
  try {
    if (extname(sourcePath) === "pdf") await writePdfAsDocx(sourcePath, tempDocx);
    else await writeVisualDocxFromSource(sourcePath, tempDocx);
    await convertWithLibreOffice(tempDocx, outPath, "doc");
  } finally {
    await fsp.rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}

async function writeRenderedImages(data, outPath, target) {
  const tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), "acacia-render-image-"));
  const tempPdf = path.join(tempDir, "render.pdf");
  try {
    await writePdf(data, tempPdf);
    const pages = await renderPdfPagesToImages(tempPdf, target);
    if (!pages.length) throw new Error("没有渲染出可用的页面。");
    return writeImagePages(pages, outPath, target);
  } finally {
    await fsp.rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}

async function writeImagePages(pages, outPath, target) {
  if (!pages.length) throw new Error("没有渲染出可用的页面。");
  const outputExt = target === "jpg" ? "jpg" : target;
  const base = outPath.replace(/\.[^.]+$/, "");
  const outputs = [];
  for (let index = 0; index < pages.length; index += 1) {
    const pagePath = `${base}_page_${index + 1}.${outputExt}`;
    await fsp.writeFile(pagePath, pages[index].buffer);
    outputs.push(pagePath);
  }
  return outputs;
}

async function convertImage(sourcePath, outPath, target) {
  const imageInput = extname(sourcePath) === "bmp" ? await readBmpAsPng(sourcePath) : sourcePath;
  if (target === "pdf") {
    const pdf = await PDFDocument.create();
    const pngBuffer = await sharp(imageInput).rotate().png().toBuffer();
    const image = await pdf.embedPng(pngBuffer);
    const page = pdf.addPage([image.width, image.height]);
    page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height });
    await fsp.writeFile(outPath, await pdf.save());
    return;
  }
  const format = target === "jpg" ? "jpeg" : target;
  let pipeline = sharp(imageInput, { animated: target === "gif" }).rotate();
  if (format === "jpeg") pipeline = pipeline.jpeg({ quality: 92 });
  else if (format === "png") pipeline = pipeline.png();
  else if (format === "webp") pipeline = pipeline.webp({ quality: 92 });
  else if (format === "gif") pipeline = pipeline.gif();
  else if (format === "avif") pipeline = pipeline.avif({ quality: 85 });
  else if (format === "tiff") pipeline = pipeline.tiff({ compression: "lzw" });
  else if (format === "bmp") {
    await fsp.writeFile(outPath, await encodeBmpFromImage(imageInput));
    return;
  }
  await fsp.writeFile(outPath, await pipeline.toBuffer());
}

async function readBmpAsPng(sourcePath) {
  const tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), "acacia-bmp-input-"));
  const pngPath = path.join(tempDir, "source.png");
  try {
    await run(ffmpegPath(), [
      "-hide_banner",
      "-loglevel", "error",
      "-y",
      "-i", sourcePath,
      "-frames:v", "1",
      pngPath
    ]);
    return await fsp.readFile(pngPath);
  } finally {
    await fsp.rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}

async function copyPdf(sourcePath, outPath) {
  const pdf = await PDFDocument.load(await fsp.readFile(sourcePath));
  await fsp.writeFile(outPath, await pdf.save());
}

async function convertMedia(sourcePath, outPath, target) {
  const args = ["-hide_banner", "-y", "-i", sourcePath];
  if (audioTargets.has(target)) {
    args.push("-vn");
    if (target === "mp3") args.push("-codec:a", "libmp3lame", "-q:a", "2");
    if (target === "m4a") args.push("-codec:a", "aac", "-b:a", "192k");
    if (target === "ogg" || target === "opus") args.push("-codec:a", "libopus", "-b:a", "160k");
    if (target === "aac") args.push("-codec:a", "aac", "-b:a", "192k");
    if (target === "wma") args.push("-codec:a", "wmav2", "-b:a", "192k");
  } else if (target === "gif") {
    args.push("-vf", "fps=12,scale='min(720,iw)':-2:flags=lanczos", "-loop", "0");
  } else if (target === "webm") {
    args.push("-codec:v", "libvpx-vp9", "-crf", "32", "-b:v", "0", "-codec:a", "libopus");
  } else if (target === "avi") {
    args.push("-codec:v", "mpeg4", "-q:v", "4", "-codec:a", "libmp3lame", "-q:a", "3");
  } else if (target === "m4v") {
    args.push("-codec:v", "libx264", "-preset", "medium", "-crf", "23", "-codec:a", "aac", "-movflags", "+faststart");
  } else if (target === "wmv") {
    args.push("-codec:v", "wmv2", "-q:v", "3", "-codec:a", "wmav2", "-b:a", "192k");
  } else if (target === "flv") {
    args.push("-codec:v", "flv", "-q:v", "4", "-codec:a", "libmp3lame", "-q:a", "3");
  } else {
    args.push("-codec:v", "libx264", "-preset", "medium", "-crf", "23", "-codec:a", "aac", "-movflags", "+faststart");
  }
  args.push(outPath);
  try {
    await run(ffmpegPath(), args);
  } catch (error) {
    const message = String(error?.message || error);
    if (/does not contain any stream|Stream map.*matches no streams|audio.*stream/i.test(message) && audioTargets.has(target)) {
      throw new Error("源视频没有可提取的音频轨，不能转换为音频文件。");
    }
    throw error;
  }
}

async function convertVideoFrame(sourcePath, outPath, target) {
  const args = [
    "-hide_banner", "-loglevel", "error", "-y", "-i", sourcePath,
    "-vf", "thumbnail,scale='min(1280,iw)':-2",
    "-frames:v", "1", outPath
  ];
  await run(ffmpegPath(), args);
}

function categoryForExt(ext) {
  if (imageInputs.has(ext)) return "image";
  if (audioInputs.has(ext)) return "audio";
  if (videoInputs.has(ext)) return "video";
  if (ext === "pdf") return "pdf";
  if (["doc", "docx", "ppt", "pptx", "odt", "odp"].includes(ext)) return "document";
  if (["xlsx", "xls", "ods", "csv", "tsv"].includes(ext)) return "spreadsheet";
  if (["txt", "md", "html", "rtf"].includes(ext)) return "text";
  return "unknown";
}

function isSupportedRoute(sourceExt, target) {
  const category = categoryForExt(sourceExt);
  if (category === "image") return imageTargets.has(target);
  if (category === "audio") return audioTargets.has(target);
  if (category === "video") return audioTargets.has(target) || videoTargets.has(target) || ["jpg", "png"].includes(target);
  if (category === "pdf") return ["txt", "md", "html", "doc", "docx", "pdf", "jpg", "png", "webp", "tiff", "bmp"].includes(target);
  if (category === "spreadsheet") return ["txt", "md", "html", "csv", "xlsx", "xls", "pdf", "jpg", "png", "webp", "tiff", "bmp"].includes(target);
  if (category === "document") return ["txt", "md", "html", "doc", "docx", "pdf", "rtf", "jpg", "png", "webp", "tiff", "bmp"].includes(target);
  if (category === "text") return [...textTargets, "docx", "doc", "xlsx", "xls", "pdf", "jpg", "png", "webp", "tiff", "bmp"].includes(target);
  return false;
}

function needsLibreOffice(sourceExt, target) {
  if (!officeInputs.has(sourceExt) || !officeTargets.has(target)) return false;
  if (sourceExt === target) return false;
  return true;
}

async function convertFile({ sourcePath, target, outputDir, conflictMode = "rename", onlyText = false }) {
  if (!sourcePath || !fs.existsSync(sourcePath)) throw new Error("找不到源文件。");
  const sourceExt = extname(sourcePath);
  const normalizedTarget = String(target || "").toLowerCase();
  if (!readable.has(sourceExt) || !writable.has(normalizedTarget) || !isSupportedRoute(sourceExt, normalizedTarget)) {
    throw new Error(`当前版本暂不支持 ${sourceExt} -> ${normalizedTarget}`);
  }
  const base = path.basename(sourcePath).replace(/\.[^.]+$/, "");
  const folder = outputDir || path.dirname(sourcePath);
  await fsp.mkdir(folder, { recursive: true });
  const { outPath, skipped } = resolveOutputPath(folder, base, normalizedTarget, conflictMode);
  if (skipped) return { outPath, characters: 0, skipped: true };

  const category = categoryForExt(sourceExt);
  if (category === "image") {
    await convertImage(sourcePath, outPath, normalizedTarget);
    return { outPath, characters: 0, skipped: false };
  }
  if (category === "audio" || category === "video") {
    if (category === "video" && ["jpg", "png"].includes(normalizedTarget)) {
      await convertVideoFrame(sourcePath, outPath, normalizedTarget);
      return { outPath, characters: 0, skipped: false };
    }
    await convertMedia(sourcePath, outPath, normalizedTarget);
    return { outPath, characters: 0, skipped: false };
  }
  if (category === "pdf" && normalizedTarget === "pdf") {
    await copyPdf(sourcePath, outPath);
    return { outPath, characters: 0, skipped: false };
  }

  if (!onlyText && category === "pdf" && normalizedTarget === "docx") {
    await writePdfAsDocx(sourcePath, outPath);
    return { outPath, characters: 0, skipped: false };
  }

  if (!onlyText && category === "pdf" && normalizedTarget === "doc") {
    await writeVisualDocFromSource(sourcePath, outPath);
    return { outPath, characters: 0, skipped: false };
  }

  if (!onlyText && category === "pdf" && normalizedTarget === "html") {
    await writePdfAsHtml(sourcePath, outPath);
    return { outPath, characters: 0, skipped: false };
  }

  if (["jpg", "png", "webp", "tiff", "bmp"].includes(normalizedTarget) && ["pdf", "document", "spreadsheet", "text"].includes(category)) {
    const outputSeed = outPath;
    if (category === "pdf") {
      const pages = await renderPdfPagesToImages(sourcePath, normalizedTarget);
      const pagePaths = await writeImagePages(pages, outputSeed, normalizedTarget);
      return { outPath: pagePaths[0], outPaths: pagePaths, characters: 0, skipped: false };
    }
    if (!onlyText && ["document", "spreadsheet", "text"].includes(category)) {
      const tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), "acacia-source-render-"));
      const tempPdf = path.join(tempDir, "source.pdf");
      try {
        await writeSourcePdf(sourcePath, tempPdf);
        const pages = await renderPdfPagesToImages(tempPdf, normalizedTarget);
        const pagePaths = await writeImagePages(pages, outputSeed, normalizedTarget);
        return { outPath: pagePaths[0], outPaths: pagePaths, characters: 0, skipped: false };
      } finally {
        await fsp.rm(tempDir, { recursive: true, force: true }).catch(() => {});
      }
    }
    const data = await extractSource(sourcePath);
    const pagePaths = await writeRenderedImages(data, outputSeed, normalizedTarget);
    return { outPath: pagePaths[0], outPaths: pagePaths, characters: sourceToText(data).length, skipped: false };
  }

  if (!onlyText && officeInputs.has(sourceExt) && officeTargets.has(normalizedTarget)) {
    if (["spreadsheet", "text"].includes(category) && textTargets.has(normalizedTarget)) {
      const data = await extractSource(sourcePath);
      const bom = ["txt", "md", "csv"].includes(normalizedTarget) ? "\ufeff" : "";
      await fsp.writeFile(outPath, bom + renderTextLike(data, normalizedTarget), "utf8");
      return { outPath, characters: sourceToText(data).length, skipped: false };
    }
    if (["spreadsheet", "text"].includes(category) && ["xlsx", "xls"].includes(normalizedTarget)) {
      const data = await extractSource(sourcePath);
      writeXlsx(data, outPath, normalizedTarget);
      return { outPath, characters: sourceToText(data).length, skipped: false };
    }
    if (sourceExt === normalizedTarget) {
      await fsp.copyFile(sourcePath, outPath);
    } else {
      await convertWithLibreOffice(sourcePath, outPath, normalizedTarget);
    }
    return { outPath, characters: 0, skipped: false };
  }

  if (!onlyText && ["text"].includes(category) && ["docx", "doc"].includes(normalizedTarget)) {
    if (normalizedTarget === "docx") await writeVisualDocxFromSource(sourcePath, outPath);
    else await writeVisualDocFromSource(sourcePath, outPath);
    return { outPath, characters: 0, skipped: false };
  }

  if (needsLibreOffice(sourceExt, normalizedTarget)) {
    await convertWithLibreOffice(sourcePath, outPath, normalizedTarget);
    return { outPath, characters: 0, skipped: false };
  }

  const data = await extractSource(sourcePath);
  if (textTargets.has(normalizedTarget)) {
    const bom = ["txt", "md", "csv"].includes(normalizedTarget) ? "\ufeff" : "";
    await fsp.writeFile(outPath, bom + renderTextLike(data, normalizedTarget), "utf8");
  } else if (normalizedTarget === "docx") {
    await writeDocx(data, outPath);
  } else if (normalizedTarget === "doc") {
    await writeDoc(data, outPath);
  } else if (normalizedTarget === "xlsx") {
    writeXlsx(data, outPath, "xlsx");
  } else if (normalizedTarget === "xls") {
    writeXlsx(data, outPath, "xls");
  } else if (normalizedTarget === "pdf") {
    await writePdf(data, outPath);
  }
  return { outPath, characters: sourceToText(data).length, skipped: false };
}

module.exports = {
  convertFile,
  readable,
  writable,
  isSupportedRoute,
  categoryForExt
};
