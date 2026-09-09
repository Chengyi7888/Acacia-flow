const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");
const packageJson = require("../package.json");

const root = path.resolve(__dirname, "..");
const outDir = path.join(root, "build", "installer");
const productName = packageJson.build?.productName || packageJson.name;
const unpackedExe = path.join(outDir, "win-unpacked", `${productName}.exe`);
const installerName = `${productName} Setup ${packageJson.version}.exe`;
const builtInstaller = path.join(outDir, installerName);
const finalInstaller = path.join(root, installerName);
const iconPath = path.join(root, "assets", "acacia_flow.ico");

function run(command, args) {
  const result = spawnSync(command, args, {
    cwd: root,
    stdio: "inherit",
    windowsHide: true
  });
  if (result.status !== 0) {
    process.exit(result.status || 1);
  }
}

function findRcedit() {
  const localRcedit = path.join(root, "node_modules", "electron-winstaller", "vendor", "rcedit.exe");
  if (fs.existsSync(localRcedit)) return localRcedit;

  const cacheRoot = path.join(process.env.LOCALAPPDATA || "", "electron-builder", "Cache", "winCodeSign");
  const stack = fs.existsSync(cacheRoot) ? [cacheRoot] : [];
  while (stack.length) {
    const current = stack.pop();
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      const fullPath = path.join(current, entry.name);
      if (entry.isDirectory()) {
        stack.push(fullPath);
      } else if (entry.name.toLowerCase() === "rcedit-x64.exe") {
        return fullPath;
      }
    }
  }
  throw new Error("rcedit.exe not found. Run npm install before building the installer.");
}

fs.rmSync(outDir, { recursive: true, force: true });
fs.rmSync(finalInstaller, { force: true });
const electronBuilderCli = require.resolve("electron-builder/out/cli/cli.js");

run(process.execPath, [electronBuilderCli, "--win", "nsis"]);

if (!fs.existsSync(unpackedExe)) {
  throw new Error(`Packaged executable not found: ${unpackedExe}`);
}

run(findRcedit(), [unpackedExe, "--set-icon", iconPath]);
run(process.execPath, [
  electronBuilderCli,
  "--win",
  "nsis",
  "--prepackaged",
  path.join(outDir, "win-unpacked")
]);

for (const relativePath of [
  path.join("win-unpacked"),
  "builder-debug.yml",
  `${installerName}.blockmap`
]) {
  fs.rmSync(path.join(outDir, relativePath), { recursive: true, force: true });
}

if (!fs.existsSync(builtInstaller)) {
  throw new Error(`Installer not found: ${builtInstaller}`);
}

fs.copyFileSync(builtInstaller, finalInstaller);
fs.rmSync(outDir, { recursive: true, force: true });
