const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");

const electronPackageDir = path.dirname(require.resolve("electron/package.json"));
const executableName = fs.readFileSync(path.join(electronPackageDir, "path.txt"), "utf8").trim();
const executablePath = path.join(electronPackageDir, "dist", executableName);
const smokeAppPath = path.join(__dirname, "smoke-electron-app");

const child = spawn(executablePath, [smokeAppPath], {
  cwd: path.resolve(__dirname, ".."),
  stdio: "inherit",
  windowsHide: true
});

child.on("error", error => {
  console.error(error);
  process.exitCode = 1;
});

child.on("exit", code => {
  process.exitCode = code ?? 0;
});
