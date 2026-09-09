const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");

const electronPackageDir = path.dirname(require.resolve("electron/package.json"));
const executableName = fs.readFileSync(path.join(electronPackageDir, "path.txt"), "utf8").trim();
const executablePath = path.join(electronPackageDir, "dist", executableName);

if (!fs.existsSync(executablePath)) {
  throw new Error(`Electron executable not found: ${executablePath}`);
}

const child = spawn(executablePath, [path.resolve(__dirname, "..")], {
  stdio: "inherit",
  windowsHide: false
});

child.on("error", error => {
  console.error(error);
  process.exitCode = 1;
});

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
    return;
  }
  process.exitCode = code ?? 0;
});
