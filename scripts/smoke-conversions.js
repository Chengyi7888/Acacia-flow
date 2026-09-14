const fs = require("fs/promises");
const path = require("path");
const os = require("os");
const { execFile } = require("child_process");
const { app } = require("electron");
const sharp = require("sharp");
const converterModule = process.env.ACACIA_CONVERTER_MODULE || "../src/main/converter";
const { convertFile, readable, writable, isSupportedRoute } = require(converterModule);

function run(command, args) {
  return new Promise((resolve, reject) => {
    execFile(command, args, { windowsHide: true }, (error, stdout, stderr) => {
      if (error) {
        reject(new Error(String(stderr || stdout || error.message).trim()));
        return;
      }
      resolve();
    });
  });
}

async function exists(filePath) {
  try {
    const stat = await fs.stat(filePath);
    return stat.size > 0;
  } catch (_error) {
    return false;
  }
}

function normalizedExt(filePath) {
  const extension = path.extname(filePath).slice(1).toLowerCase();
  if (extension === "jpeg") return "jpg";
  if (extension === "tif") return "tiff";
  if (extension === "htm") return "html";
  return extension;
}

function isOutputPresent(result) {
  const paths = result.outPaths?.length ? result.outPaths : [result.outPath];
  return Promise.all(paths.map(exists)).then(values => values.every(Boolean));
}

async function convertAndRecord(results, name, sourcePath, target, outputDir, onlyText = false) {
  try {
    const result = await convertFile({
      sourcePath,
      target,
      outputDir,
      conflictMode: "rename",
      onlyText
    });
    results.push({
      name,
      source: normalizedExt(sourcePath),
      target,
      path: result.outPath,
      paths: result.outPaths,
      ok: await isOutputPresent(result)
    });
    return result;
  } catch (error) {
    results.push({
      name,
      source: normalizedExt(sourcePath),
      target,
      ok: false,
      error: String(error?.message || error)
    });
    return null;
  }
}

async function main() {
  app.on("window-all-closed", () => {});
  await app.whenReady();
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "acacia-smoke-"));
  const inputDir = path.join(root, "inputs");
  const fixtureDir = path.join(root, "fixtures");
  const outputDir = path.join(root, "outputs");
  const reportDir = path.join(__dirname, "..", "build", "test-results");
  const reportPath = path.join(reportDir, "smoke-result.json");
  await fs.mkdir(inputDir, { recursive: true });
  await fs.mkdir(outputDir, { recursive: true });
  await fs.mkdir(fixtureDir, { recursive: true });
  await fs.mkdir(reportDir, { recursive: true });

  const ffmpeg = require("@ffmpeg-installer/ffmpeg").path.replace(/\.asar([\\/])/, ".asar.unpacked$1");
  const txt = path.join(inputDir, "sample.txt");
  const md = path.join(inputDir, "sample.md");
  const html = path.join(inputDir, "sample.html");
  const csv = path.join(inputDir, "sample.csv");
  const png = path.join(inputDir, "sample.png");
  const wav = path.join(inputDir, "sample.wav");
  const mp4 = path.join(inputDir, "sample.mp4");
  const sampleDoc = process.env.ACACIA_SAMPLE_DOC || path.join(inputDir, "sample.doc");
  const samplePdf = process.env.ACACIA_SAMPLE_PDF || "";

  await fs.writeFile(txt, [
    "Acacia Flow 全量转换测试",
    "中文 English 12345",
    "符号：☆ ♡ © ® → ← ± × ÷ <tag> & \"quotes\"",
    "",
    "不同字号由富文本样例覆盖；此处验证纯文本的换行与空行。",
    "    保留四个空格缩进"
  ].join("\n"), "utf8");
  await fs.writeFile(md, [
    "# Acacia Flow 标题",
    "",
    "正文含有 **粗体**、*斜体*、`代码` 和 ☆ ♡。",
    "",
    "## 二级标题",
    "",
    "- 第一项",
    "- 第二项",
    "",
    "| 名称 | 数值 |",
    "| --- | ---: |",
    "| 中文 | 123 |"
  ].join("\n"), "utf8");
  await fs.writeFile(html, `<!doctype html>
<meta charset="utf-8">
<style>
  body{font-family:"Microsoft YaHei",Arial,sans-serif;line-height:1.6}
  h1{font-size:30px;color:#3a4573} h2{font-size:22px}
  .small{font-size:12px}.large{font-size:26px}.panel{padding:12px;background:#eef3ff;border:1px solid #aab8e8}
  table{border-collapse:collapse}td,th{border:1px solid #777;padding:6px}
</style>
<h1>Acacia Flow ☆</h1>
<p class="small">小字号：中文 English 123 &amp; &lt;符号&gt;</p>
<p>常规文字 <strong>粗体</strong> <em>斜体</em> ♡</p>
<p class="large">大字号文本：格式转换测试</p>
<div class="panel">带背景色的图片区块：<img alt="测试图片" src="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='160' height='70'%3E%3Crect width='160' height='70' fill='%236f82c7'/%3E%3Ctext x='12' y='42' fill='white' font-size='26'%3EAcacia %E2%99%A1%3C/text%3E%3C/svg%3E"></div>
<h2>表格</h2><table><tr><th>名称</th><th>数值</th></tr><tr><td>中文</td><td>123</td></tr></table>`, "utf8");
  const hasExternalSampleDoc = Boolean(process.env.ACACIA_SAMPLE_DOC) && await exists(process.env.ACACIA_SAMPLE_DOC);
  const hasExternalSamplePdf = Boolean(samplePdf) && await exists(samplePdf);
  await fs.writeFile(csv, "\ufeff名称,数值,备注\n中文,123,☆ ♡\nEnglish,456,\"<tag> & symbol\"\n", "utf8");
  await sharp(Buffer.from(`<svg xmlns="http://www.w3.org/2000/svg" width="720" height="360">
    <rect width="720" height="360" fill="#f0f4ff"/>
    <rect x="32" y="28" width="656" height="304" rx="18" fill="#6e7fc4"/>
    <circle cx="130" cy="180" r="68" fill="#ffdaa6"/>
    <text x="230" y="125" font-size="34" fill="white" font-family="Microsoft YaHei, Arial">Acacia Flow ☆</text>
    <text x="230" y="182" font-size="24" fill="white" font-family="Microsoft YaHei, Arial">中文 English 123</text>
    <text x="230" y="240" font-size="16" fill="white" font-family="Microsoft YaHei, Arial">♡ &amp; &lt;format&gt; → PDF</text>
  </svg>`)).png().toFile(png);
  await run(ffmpeg, ["-hide_banner", "-y", "-f", "lavfi", "-i", "sine=frequency=880:duration=0.5", wav]);
  await run(ffmpeg, [
    "-hide_banner", "-y",
    "-f", "lavfi", "-i", "testsrc2=size=320x180:rate=12:duration=0.6",
    "-f", "lavfi", "-i", "sine=frequency=440:duration=0.6",
    "-shortest", "-pix_fmt", "yuv420p", mp4
  ]);

  const results = [];
  const fixtures = new Map([
    ["txt", txt],
    ["md", md],
    ["html", html],
    ["csv", csv],
    ["png", png],
    ["wav", wav],
    ["mp4", mp4]
  ]);

  const createFixture = async (sourcePath, target) => {
    const result = await convertAndRecord(results, `fixture ${normalizedExt(sourcePath)} to ${target}`, sourcePath, target, inputDir);
    if (result?.outPath) {
      const stablePath = path.join(fixtureDir, `fixture.${target}`);
      await fs.copyFile(result.outPath, stablePath);
      fixtures.set(target, stablePath);
    }
  };

  for (const target of ["jpg", "webp", "gif", "avif", "tiff", "bmp"]) await createFixture(png, target);
  for (const target of ["mp3", "flac", "m4a", "aac", "ogg", "opus", "wma"]) await createFixture(wav, target);
  for (const target of ["mov", "mkv", "webm", "avi", "m4v", "wmv", "flv"]) await createFixture(mp4, target);
  for (const target of ["xlsx", "xls"]) await createFixture(csv, target);
  const stableCsv = path.join(fixtureDir, "source.csv");
  await fs.copyFile(csv, stableCsv);
  const stableTable = new Map([
    ["csv", stableCsv]
  ]);
  for (const target of ["xlsx", "xls"]) {
    const result = await convertAndRecord(results, `stable csv to ${target}`, stableCsv, target, fixtureDir);
    if (result?.outPath) {
      const stablePath = path.join(fixtureDir, `source-${target}.${target}`);
      const sourceOutput = result.outPaths?.[0] || result.outPath;
      await fs.copyFile(sourceOutput, stablePath);
      stableTable.set(target, stablePath);
    }
  }
  for (const target of ["docx", "doc", "rtf"]) await createFixture(html, target);
  const pdfFixture = await convertAndRecord(results, "fixture html to pdf", html, "pdf", inputDir);
  if (pdfFixture?.outPath) fixtures.set("pdf", pdfFixture.outPath);

  for (const [sourceExt, sourcePath] of fixtures) {
    let sourceForConversion = sourcePath;
    if (sourceExt === "xlsx" || sourceExt === "xls") {
      const stablePath = stableTable.get(sourceExt);
      if (stablePath) sourceForConversion = stablePath;
    }
    for (const target of writable) {
      if (!isSupportedRoute(sourceExt, target)) continue;
      await convertAndRecord(results, `${sourceExt} to ${target}`, sourceForConversion, target, outputDir);
    }
  }

  await convertAndRecord(results, "txt to docx text-only", txt, "docx", outputDir, true);
  await convertAndRecord(results, "pdf to docx text-only", fixtures.get("pdf"), "docx", outputDir, true);

  if (hasExternalSampleDoc) {
    for (const target of writable) {
      if (isSupportedRoute(normalizedExt(sampleDoc), target)) {
        await convertAndRecord(results, `external ${normalizedExt(sampleDoc)} to ${target}`, sampleDoc, target, outputDir);
      }
    }
  }
  if (hasExternalSamplePdf) {
    for (const target of writable) {
      if (isSupportedRoute("pdf", target)) {
        await convertAndRecord(results, `external pdf to ${target}`, samplePdf, target, outputDir);
      }
    }
  }

  const report = {
    root,
    supportedInputFormats: [...readable].sort(),
    supportedOutputFormats: [...writable].sort(),
    total: results.length,
    passed: results.filter(result => result.ok).length,
    failed: results.filter(result => !result.ok).length,
    results
  };
  await fs.writeFile(reportPath, JSON.stringify(report, null, 2), "utf8");
  console.log(JSON.stringify(report, null, 2));
  app.exit(results.every(result => result.ok) ? 0 : 1);
}

main().catch(error => {
  console.error(error);
  if (app?.exit) app.exit(1);
  else process.exit(1);
});
