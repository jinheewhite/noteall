/* noteall — note(판서노트)의 화면·기능은 그대로 두고, 여기에 파워포인트·
   한글(한쇼) 불러오기를 더합니다. 그 PC에 실제 설치된 파워포인트·한글을
   COM 자동화로 띄워 PDF로 저장시키고, note가 이미 가진 "PDF 열기" 경로에
   그대로 넘깁니다 — note의 화면·필기·저장 로직은 한 줄도 안 건드립니다. */
const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const { execFile } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { autoUpdater } = require('electron-updater');

/* ProgID가 레지스트리(HKEY_CLASSES_ROOT)에 등록되어 있으면 설치된 것으로 봅니다.
   64비트 reg.exe는 기본적으로 64비트 보기만 보므로, 32비트로 설치된 프로그램을
   놓치지 않도록 /reg:32 보기도 함께 확인합니다. */
function regKeyExists(progId) {
  return new Promise(resolve => {
    const tryView = view => new Promise(res => {
      const args = ['query', `HKCR\\${progId}`];
      if (view) args.push('/reg:' + view);
      execFile('reg.exe', args, err => res(!err));
    });
    tryView().then(found => {
      if (found) return resolve(true);
      tryView('32').then(resolve);
    });
  });
}

async function detectOffice() {
  const [hasPowerPoint, hasHangul] = await Promise.all([
    regKeyExists('PowerPoint.Application'),
    regKeyExists('HWPFrame.HwpObject'),
  ]);
  return { hasPowerPoint, hasHangul };
}

function warnMissing({ hasPowerPoint, hasHangul }) {
  const missing = [];
  if (!hasPowerPoint) missing.push('파워포인트');
  if (!hasHangul) missing.push('한글');
  if (!missing.length) return;
  dialog.showMessageBox({
    type: 'warning',
    title: '설치된 프로그램 확인',
    message: `이 PC에서 ${missing.join('·')}을(를) 찾지 못했습니다.`,
    detail: `${missing.join('·')} 파일을 불러오려면 그 프로그램이 이 PC에 설치되어 있어야 합니다.\n` +
      'PDF·그림 판서는 지금처럼 그대로 쓰실 수 있습니다.',
    buttons: ['확인'],
  });
}

const OFFICE_EXT_KIND = {
  '.ppt': 'ppt', '.pptx': 'ppt',
  '.hwp': 'hwp', '.hwpx': 'hwp',
};

function convertToPdf(inputPath) {
  return new Promise((resolve, reject) => {
    const kind = OFFICE_EXT_KIND[path.extname(inputPath).toLowerCase()];
    if (!kind) return reject(new Error('지원하지 않는 파일 형식입니다 (파워포인트·한글만 됩니다).'));

    const outputPath = path.join(os.tmpdir(), `noteall-${Date.now()}-${Math.random().toString(36).slice(2)}.pdf`);
    const scriptPath = path.join(__dirname, 'convert.ps1');
    execFile('powershell.exe', [
      '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
      '-File', scriptPath,
      '-Kind', kind, '-InputPath', inputPath, '-OutputPath', outputPath,
    ], { timeout: 120000, windowsHide: true }, (err, stdout, stderr) => {
      if (err) return reject(new Error(`변환 중 문제가 생겼습니다: ${(stderr || err.message).trim()}`));
      if (!fs.existsSync(outputPath)) return reject(new Error('변환된 PDF를 찾을 수 없습니다.'));
      resolve(outputPath);
    });
  });
}

ipcMain.handle('pick-office-file', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const { canceled, filePaths } = await dialog.showOpenDialog(win, {
    title: '파워포인트·한글 파일 열기',
    filters: [
      { name: '파워포인트·한글', extensions: ['pptx', 'ppt', 'hwp', 'hwpx'] },
    ],
    properties: ['openFile'],
  });
  if (canceled || !filePaths.length) return null;
  return filePaths[0];
});

ipcMain.handle('import-office', async (event, inputPath) => {
  const pdfPath = await convertToPdf(inputPath);
  try {
    const bytes = fs.readFileSync(pdfPath);
    const name = path.basename(inputPath, path.extname(inputPath)) + '.pdf';
    return { bytes, name };
  } finally {
    fs.unlink(pdfPath, () => {});
  }
});

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, 'preload.js'),
    },
  });
  win.loadFile(path.join(__dirname, 'index.html'));
  return win;
}

app.whenReady().then(async () => {
  createWindow();
  const found = await detectOffice();
  warnMissing(found);

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });

  /* note의 PWA 판과 같은 철학 — 조용히 새 판을 받아 두고, 다음에 켤 때
     자동으로 적용합니다. 업데이트를 확인하는 순간에만 통신이 생기고,
     그 외에는 완전히 오프라인입니다. */
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.checkForUpdates().catch(() => {});
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
