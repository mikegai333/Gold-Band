import { execSync } from 'node:child_process';
import { copyFileSync, cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, '..');

function parseArgs(args) {
  const options = {
    exe: null,
    installer: null,
    version: null,
    runtime: null,
    output: null,
    zip: true,
  };

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === '--exe' && i + 1 < args.length) {
      options.exe = resolve(args[++i]);
    } else if (arg === '--installer' && i + 1 < args.length) {
      options.installer = resolve(args[++i]);
    } else if (arg === '--version' && i + 1 < args.length) {
      options.version = args[++i];
    } else if (arg === '--runtime' && i + 1 < args.length) {
      options.runtime = resolve(args[++i]);
    } else if (arg === '--output' && i + 1 < args.length) {
      options.output = resolve(args[++i]);
    } else if (arg === '--no-zip') {
      options.zip = false;
    }
  }

  return options;
}

function find7z() {
  const candidates = ['7zz', '7z', '/opt/homebrew/bin/7zz', '/usr/local/bin/7zz'];
  for (const cmd of candidates) {
    try {
      execSync(`which ${cmd}`, { stdio: 'ignore' });
      return cmd;
    } catch {
      // ignore
    }
  }
  return null;
}

function main() {
  const options = parseArgs(process.argv.slice(2));

  // Determine version
  let version = options.version;
  if (!version) {
    try {
      const pkg = JSON.parse(readFileSync(join(repoRoot, 'package.json'), 'utf8'));
      version = pkg.version;
    } catch {
      version = '0.19.1';
    }
  }
  if (version.startsWith('v')) {
    version = version.slice(1);
  }

  // Determine target directory
  const folderName = `Gold-Band-v${version}-windows-x64-portable`;
  const distDir = join(repoRoot, 'dist');
  const targetDir = options.output || join(distDir, folderName);

  console.log(`Packaging Gold Band Windows x64 Portable v${version}...`);
  console.log(`Target directory: ${targetDir}`);

  mkdirSync(targetDir, { recursive: true });

  // Locate or extract Gold Band.exe
  let exeSource = options.exe;
  if (!exeSource && options.installer) {
    if (!existsSync(options.installer)) {
      console.error(`Installer file not found: ${options.installer}`);
      process.exit(1);
    }
    const sevenZip = find7z();
    if (!sevenZip) {
      console.error('7-Zip (7z or 7zz) is required to extract the installer.');
      process.exit(1);
    }

    const tempExtractDir = join(targetDir, '__installer_temp');
    mkdirSync(tempExtractDir, { recursive: true });
    try {
      console.log(`Extracting executable from installer: ${options.installer}`);
      execSync(`${sevenZip} x -y -o"${tempExtractDir}" "${options.installer}"`, { stdio: 'inherit' });
      const extractedExe = join(tempExtractDir, 'gold-band-desktop.exe');
      const fallbackExe = join(tempExtractDir, 'Gold Band.exe');
      if (existsSync(extractedExe)) {
        exeSource = extractedExe;
      } else if (existsSync(fallbackExe)) {
        exeSource = fallbackExe;
      } else {
        console.error('Could not locate executable inside the installer archive.');
        process.exit(1);
      }
    } catch (err) {
      console.error('Failed to extract installer:', err);
      process.exit(1);
    }
  }

  if (!exeSource) {
    const candidatePaths = [
      join(repoRoot, 'src-tauri', 'target', 'release', 'Gold Band.exe'),
      join(repoRoot, 'src-tauri', 'target', 'release', 'gold-band-desktop.exe'),
      join(repoRoot, 'src-tauri', 'target', 'x86_64-pc-windows-msvc', 'release', 'Gold Band.exe'),
      join(repoRoot, 'src-tauri', 'target', 'x86_64-pc-windows-msvc', 'release', 'gold-band-desktop.exe'),
      join(distDir, 'Gold-Band-v0.18.0-windows-x64-portable', 'Gold Band.exe'),
    ];
    for (const p of candidatePaths) {
      if (existsSync(p)) {
        exeSource = p;
        break;
      }
    }
  }

  if (!exeSource || !existsSync(exeSource)) {
    console.error('No Gold Band.exe binary found. Please specify --exe <path> or --installer <path>.');
    process.exit(1);
  }

  const destExe = join(targetDir, 'Gold Band.exe');
  console.log(`Copying executable from ${exeSource} -> ${destExe}`);
  copyFileSync(exeSource, destExe);

  // Clean temp extraction if used
  const tempExtractDir = join(targetDir, '__installer_temp');
  if (existsSync(tempExtractDir)) {
    rmSync(tempExtractDir, { recursive: true, force: true });
  }

  // Locate and copy WebView2 runtime
  let runtimeSource = options.runtime;
  if (!runtimeSource) {
    const candidateRuntimes = [
      join(distDir, 'Gold-Band-v0.18.0-windows-x64-portable', 'runtime', 'webview2'),
      join(repoRoot, 'runtime', 'webview2'),
    ];
    for (const p of candidateRuntimes) {
      if (existsSync(p) && existsSync(join(p, 'msedgewebview2.exe'))) {
        runtimeSource = p;
        break;
      }
    }
  }

  if (runtimeSource && existsSync(runtimeSource)) {
    const destRuntime = join(targetDir, 'runtime', 'webview2');
    if (!existsSync(destRuntime)) {
      console.log(`Copying WebView2 runtime from ${runtimeSource}...`);
      mkdirSync(join(targetDir, 'runtime'), { recursive: true });
      cpSync(runtimeSource, destRuntime, { recursive: true });
    } else {
      console.log(`WebView2 runtime already present in ${destRuntime}`);
    }
  } else {
    console.warn('Warning: WebView2 Fixed Version Runtime not found. The portable package will rely on system WebView2.');
  }

  // Write Gold Band.vbs
  const vbsContent = `Option Explicit
Dim WshShell, fso, currentDir, wv2Path

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
currentDir = fso.GetParentFolderName(WScript.ScriptFullName)
wv2Path = currentDir & "\\runtime\\webview2"

If fso.FileExists(wv2Path & "\\msedgewebview2.exe") Then
    WshShell.Environment("PROCESS")("WEBVIEW2_BROWSER_EXECUTABLE_FOLDER") = wv2Path
End If

WshShell.Run """" & currentDir & "\\Gold Band.exe""", 1, False
Set WshShell = Nothing
Set fso = Nothing
`;
  writeFileSync(join(targetDir, 'Gold Band.vbs'), vbsContent, 'utf8');

  // Write 启动 Gold Band.bat
  const batContent = `@echo off
chcp 65001 >nul
title Gold Band (离线便携版)
cd /d "%~dp0"

:: 设置 WebView2 固定版本运行时路径（优先使用自带离线内核）
if exist "%~dp0runtime\\webview2\\msedgewebview2.exe" (
    set "WEBVIEW2_BROWSER_EXECUTABLE_FOLDER=%~dp0runtime\\webview2"
)

:: 启动应用程序
start "" "%~dp0Gold Band.exe" %*
exit
`;
  writeFileSync(join(targetDir, '启动 Gold Band.bat'), batContent, 'utf8');

  // Write 使用说明-离线免安装版.txt
  const readmeContent = `【Gold Band 离线免安装绿色版】使用说明

1. 简介：
   本安装包为 Gold Band 的 Windows x64 离线绿色便携版本，已完整内置 Microsoft Edge WebView2 固定版本运行时内核（Fixed Version Runtime）。
   完全无需联网，无需宿主系统预装任何环境，无需 Administrator 管理员权限，解压即用。

2. 启动方式：
   - 方式一（推荐）：双击运行【Gold Band.vbs】，静默启动应用（无控制台黑框弹出）。
   - 方式二：双击运行【启动 Gold Band.bat】，自动加载自带 WebView2 运行时并启动。
   - 方式三：若您的 Windows 系统中已安装有 WebView2 运行时，也可直接双击运行【Gold Band.exe】。

3. 目录结构：
   ${folderName}/
   ├── Gold Band.exe              # 主程序可执行文件
   ├── Gold Band.vbs              # 免黑框静默启动器（推荐快捷方式目标）
   ├── 启动 Gold Band.bat          # 批处理启动脚本
   ├── runtime/
   │   └── webview2/              # 内置的 WebView2 离线固定运行时组件
   └── 使用说明-离线免安装版.txt     # 本说明文件

4. 提示与技巧：
   - 请勿随意删除或重命名 runtime/ 文件夹，否则在未预装 WebView2 的断网机器上将无法正常显示界面。
   - 如需在桌面创建快捷方式，建议右键【Gold Band.vbs】选择“发送到 -> 桌面快捷方式”，并可在属性中为其更换主程序 Gold Band.exe 的图标。
`;
  writeFileSync(join(targetDir, '使用说明-离线免安装版.txt'), readmeContent, 'utf8');

  // Create ZIP archive if requested
  if (options.zip) {
    const zipPath = join(distDir, `${folderName}.zip`);
    console.log(`Creating ZIP archive: ${zipPath}`);
    const sevenZip = find7z();
    if (sevenZip) {
      execSync(`${sevenZip} a -tzip -mx=9 "${zipPath}" "${folderName}"`, {
        cwd: distDir,
        stdio: 'inherit',
      });
    } else {
      execSync(`zip -r -9 "${zipPath}" "${folderName}"`, {
        cwd: distDir,
        stdio: 'inherit',
      });
    }
    console.log(`Portable package successfully created: ${zipPath}`);
  } else {
    console.log(`Portable package ready at: ${targetDir}`);
  }
}

main();
