const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const { execFile } = require("child_process");

const protectedRoots = [
  process.env.SystemRoot || "C:\\Windows",
  process.env.ProgramData || "C:\\ProgramData"
].map(item => path.resolve(item).toLowerCase());

const elevatedRoots = [
  process.env.ProgramFiles,
  process.env["ProgramFiles(x86)"]
].filter(Boolean).map(item => path.resolve(item).toLowerCase());

function normalizePath(value) {
  return path.resolve(String(value || "").trim().replace(/[\\/]+$/, ""));
}

function isSameOrInside(parent, child) {
  const normalizedParent = normalizePath(parent).toLowerCase();
  const normalizedChild = normalizePath(child).toLowerCase();
  return normalizedChild === normalizedParent || normalizedChild.startsWith(`${normalizedParent}${path.sep}`);
}

function run(command, args) {
  return new Promise((resolve, reject) => {
    execFile(command, args, { windowsHide: true }, (error, stdout, stderr) => {
      if (error) {
        reject(new Error(String(stderr || stdout || error.message).trim()));
        return;
      }
      resolve({ stdout, stderr });
    });
  });
}

async function isElevated() {
  if (process.platform !== "win32") return true;
  try {
    await run("fltmc.exe", []);
    return true;
  } catch (_error) {
    return false;
  }
}

async function pathExists(targetPath) {
  try {
    await fsp.lstat(targetPath);
    return true;
  } catch (_error) {
    return false;
  }
}

function validateMove(source, destinationRoot) {
  if (!source || !destinationRoot) throw new Error("请选择源文件夹和目标位置。");
  if (!path.isAbsolute(source) || !path.isAbsolute(destinationRoot)) {
    throw new Error("文件夹路径无效。");
  }
  if (source === destinationRoot) throw new Error("源文件夹和目标位置不能相同。");
  if (protectedRoots.some(root => source.toLowerCase() === root || isSameOrInside(root, source))) {
    throw new Error("为避免破坏系统，不能移动 Windows 或 ProgramData 等系统目录。");
  }
  if (isSameOrInside(source, destinationRoot)) {
    throw new Error("目标位置不能位于源文件夹内部。");
  }
}

function touchesElevatedRoot(...targets) {
  return targets.some(target => {
    const normalizedTarget = normalizePath(target).toLowerCase();
    return elevatedRoots.some(root => normalizedTarget === root || isSameOrInside(root, normalizedTarget));
  });
}

async function copyDirectory(source, destination, onProgress) {
  const entries = await fsp.readdir(source, { withFileTypes: true });
  await fsp.mkdir(destination, { recursive: true });
  for (const entry of entries) {
    const from = path.join(source, entry.name);
    const to = path.join(destination, entry.name);
    if (entry.isDirectory()) {
      await copyDirectory(from, to, onProgress);
    } else if (entry.isSymbolicLink()) {
      const linkTarget = await fsp.readlink(from);
      await fsp.symlink(linkTarget, to, "junction");
    } else {
      await fsp.copyFile(from, to);
    }
    onProgress?.();
  }
}

async function removeDirectory(targetPath) {
  await fsp.rm(targetPath, { recursive: true, force: true });
}

async function createJunction(linkPath, targetPath) {
  await run("cmd.exe", ["/d", "/c", "mklink", "/J", linkPath, targetPath]);
}

async function moveFolder({ sourcePath, destinationRoot, hideOriginal = false, onProgress } = {}) {
  const source = normalizePath(sourcePath);
  const root = normalizePath(destinationRoot);
  validateMove(source, root);

  const sourceStat = await fsp.lstat(source).catch(() => null);
  if (!sourceStat || !sourceStat.isDirectory()) throw new Error("源文件夹不存在或不是文件夹。");
  if (await pathExists(root) && !(await fsp.lstat(root)).isDirectory()) {
    throw new Error("目标位置不是文件夹。");
  }
  await fsp.mkdir(root, { recursive: true });

  const destination = path.join(root, path.basename(source));
  if (await pathExists(destination)) throw new Error("目标位置已经存在同名文件夹。");
  if (touchesElevatedRoot(source, destination) && !(await isElevated())) {
    throw new Error("当前迁移涉及 Program Files 等高权限目录。请先关闭正在使用该软件的程序，然后右键 Acacia Flow，选择“以管理员身份运行”后再迁移。");
  }

  const sameDrive = path.parse(source).root.toLowerCase() === path.parse(destination).root.toLowerCase();
  let sourceMoved = false;
  try {
    if (sameDrive) {
      await fsp.rename(source, destination);
      sourceMoved = true;
      onProgress?.({ phase: "move", progress: 70 });
    } else {
      await copyDirectory(source, destination, onProgress ? () => onProgress({ phase: "copy", progress: 45 }) : undefined);
      await removeDirectory(source);
      sourceMoved = true;
      onProgress?.({ phase: "move", progress: 70 });
    }

    await createJunction(source, destination);
    onProgress?.({ phase: "link", progress: 100 });

    if (hideOriginal) {
      await run("attrib.exe", ["+H", source]);
    }

    return {
      source,
      destination,
      junction: source,
      sameDrive
    };
  } catch (error) {
    if (await pathExists(source) && sourceMoved) {
      await removeDirectory(source).catch(() => {});
    }
    if (await pathExists(destination)) {
      if (sourceMoved && sameDrive) {
        await fsp.rename(destination, source).catch(() => {});
      } else if (sourceMoved) {
        await copyDirectory(destination, source).catch(() => {});
        await removeDirectory(destination).catch(() => {});
      } else {
        await removeDirectory(destination).catch(() => {});
      }
    }
    const message = String(error.message || "");
    if (/EPERM|EACCES|operation not permitted|access is denied|拒绝访问/i.test(message)) {
      throw new Error("移动目录失败：Windows 拒绝了本次文件操作。请关闭正在占用该目录的软件、终端或后台服务，并以管理员身份运行 Acacia Flow 后重试。");
    }
    throw new Error(`移动目录失败：${message}`);
  }
}

module.exports = { moveFolder };
