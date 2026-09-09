const fs = require("fs/promises");
const path = require("path");
const os = require("os");
const { execFile } = require("child_process");
const { app } = require("electron");
const sharp = require("sharp");
const converterModule = process.env.ACACIA_CONVERTER_MODULE || "../src/main/converter";
const { convertFile } = require(converterModule);

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

async function main() {
  app.on("window-all-closed", () => {});
  await app.whenReady();
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "acacia-smoke-"));
  const inputDir = path.join(root, "inputs");
  const outputDir = path.join(root, "outputs");
  const reportDir = path.join(__dirname, "..", "build", "test-results");
  const reportPath = path.join(reportDir, "smoke-result.json");
  await fs.mkdir(inputDir, { recursive: true });
  await fs.mkdir(outputDir, { recursive: true });
  await fs.mkdir(reportDir, { recursive: true });

  const ffmpeg = require("@ffmpeg-installer/ffmpeg").path.replace(/\.asar([\\/])/, ".asar.unpacked$1");
  const txt = path.join(inputDir, "sample.txt");
  const csv = path.join(inputDir, "sample.csv");
  const png = path.join(inputDir, "sample.png");
  const wav = path.join(inputDir, "sample.wav");
  const mp4 = path.join(inputDir, "sample.mp4");

  await fs.writeFile(txt, "Acacia Flow smoke test\nTXT to PDF works.", "utf8");
  await fs.writeFile(csv, "name,value\nalpha,1\nbeta,2\n", "utf8");
  await sharp({
    create: {
      width: 64,
      height: 64,
      channels: 4,
      background: { r: 170, g: 161, b: 255, alpha: 1 }
    }
  }).png().toFile(png);
  await run(ffmpeg, ["-hide_banner", "-y", "-f", "lavfi", "-i", "sine=frequency=880:duration=0.25", wav]);
  await run(ffmpeg, [
    "-hide_banner", "-y",
    "-f", "lavfi", "-i", "testsrc=size=64x64:rate=10:duration=0.3",
    "-f", "lavfi", "-i", "sine=frequency=440:duration=0.3",
    "-shortest", "-pix_fmt", "yuv420p", mp4
  ]);

  const checks = [
    ["text to pdf", txt, "pdf"],
    ["text to docx", txt, "docx"],
    ["text to xlsx", txt, "xlsx"],
    ["csv to xlsx", csv, "xlsx"],
    ["image to pdf", png, "pdf"],
    ["image to jpg", png, "jpg"],
    ["audio to mp3", wav, "mp3"],
    ["video to mp4", mp4, "mp4"],
    ["video to mp3", mp4, "mp3"]
  ];

  const results = [];
  for (const [name, sourcePath, target] of checks) {
    try {
      const result = await convertFile({ sourcePath, target, outputDir, conflictMode: "rename" });
      results.push({ name, target, path: result.outPath, ok: await exists(result.outPath) });
    } catch (error) {
      results.push({ name, target, ok: false, error: String(error?.message || error) });
    }
  }

  const pdfResult = results.find(result => result.name === "text to pdf");
  const docxResult = results.find(result => result.name === "text to docx");
  if (docxResult?.path) {
    try {
      const result = await convertFile({ sourcePath: docxResult.path, target: "txt", outputDir, conflictMode: "rename" });
      results.push({ name: "docx to txt", target: "txt", path: result.outPath, ok: await exists(result.outPath) });
    } catch (error) {
      results.push({ name: "docx to txt", target: "txt", ok: false, error: String(error?.message || error) });
    }
  }

  if (pdfResult?.path) {
    try {
      const result = await convertFile({ sourcePath: pdfResult.path, target: "txt", outputDir, conflictMode: "rename" });
      results.push({ name: "pdf to txt", target: "txt", path: result.outPath, ok: await exists(result.outPath) });
    } catch (error) {
      results.push({ name: "pdf to txt", target: "txt", ok: false, error: String(error?.message || error) });
    }
  }

  const report = { root, results };
  await fs.writeFile(reportPath, JSON.stringify(report, null, 2), "utf8");
  console.log(JSON.stringify(report, null, 2));
  app.exit(results.every(result => result.ok) ? 0 : 1);
}

main().catch(error => {
  console.error(error);
  if (app?.exit) app.exit(1);
  else process.exit(1);
});
