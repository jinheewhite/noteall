/* noteall — note(판서노트)의 화면·기능은 그대로 두고, 여기에 파워포인트·
   한글(한쇼) 불러오기를 더합니다. 그 PC에 실제 설치된 파워포인트·한글을
   COM 자동화로 띄워 PDF로 저장시키고, note가 이미 가진 "PDF 열기" 경로에
   그대로 넘깁니다 — note의 화면·필기·저장 로직은 한 줄도 안 건드립니다. */
const { app, BrowserWindow, dialog, ipcMain, Menu, shell } = require('electron');
const { execFile } = require('child_process');
const path = require('path');
const fs = require('fs');
const fsp = fs.promises;
const os = require('os');
const { autoUpdater } = require('electron-updater');

/* 저장 폴더 — 브라우저용 File System Access API(showDirectoryPicker) 대신
   Node fs를 직접 씁니다. 이 앱은 늘 윈도우 PC에서만 도니, 매번 사용자가
   고르게 하지 않고 "문서\기전판서앱"에 처음부터 자동으로 연결해 둡니다.
   다른 폴더로 바꾸고 싶을 때만 네이티브 폴더 선택창(dialog)을 씁니다 —
   showDirectoryPicker는 일렉트론 안에서 가끔 창이 먹통처럼 보이는 문제가 있었습니다. */
const BACKUP_CFG_PATH = path.join(app.getPath('userData'), 'backup-dir.json');
function readBackupDirCfg() {
  try { return JSON.parse(fs.readFileSync(BACKUP_CFG_PATH, 'utf8')).dir || null; } catch (e) { return null; }
}
function writeBackupDirCfg(dir) {
  try { fs.writeFileSync(BACKUP_CFG_PATH, JSON.stringify({ dir })); } catch (e) {}
}
async function resolveBackupDir() {
  let dir = readBackupDirCfg();
  if (!dir || !fs.existsSync(dir)) {
    dir = path.join(app.getPath('documents'), '기전판서앱');
    await fsp.mkdir(dir, { recursive: true });
    writeBackupDirCfg(dir);
  }
  return dir;
}

ipcMain.handle('fsop-get-dir', async () => {
  const dir = await resolveBackupDir();
  return { path: dir, name: path.basename(dir) };
});
ipcMain.handle('fsop-pick-dir', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const { canceled, filePaths } = await dialog.showOpenDialog(win, {
    title: '판서 저장 폴더 고르기',
    properties: ['openDirectory', 'createDirectory'],
  });
  if (canceled || !filePaths.length) return null;
  await fsp.mkdir(filePaths[0], { recursive: true });
  writeBackupDirCfg(filePaths[0]);
  return { path: filePaths[0], name: path.basename(filePaths[0]) };
});
ipcMain.handle('fsop-read-dir', async (_event, dirPath) => {
  const ents = await fsp.readdir(dirPath, { withFileTypes: true });
  return ents.map(d => ({ name: d.name, path: path.join(dirPath, d.name), isDirectory: d.isDirectory() }));
});
ipcMain.handle('fsop-mkdir', async (_event, dirPath) => { await fsp.mkdir(dirPath, { recursive: true }); });
ipcMain.handle('fsop-read-file', async (_event, filePath) => fsp.readFile(filePath));
ipcMain.handle('fsop-write-file', async (_event, { path: p, data }) => {
  await fsp.writeFile(p, Buffer.isBuffer(data) ? data : Buffer.from(data));
});
ipcMain.handle('fsop-remove', async (_event, { path: p, recursive }) => {
  await fsp.rm(p, { recursive: !!recursive, force: true });
});
/* USB 같은 이동식 드라이브에 저장 폴더를 정해 둔 경우, 뽑혀 있는 동안에도
   서재가 "파일이 있다"고 속이지 않도록 — 지금 실제로 보이는지만 확인합니다. */
ipcMain.handle('fsop-exists', (_event, p) => fs.existsSync(p));

/* 화면 오른쪽 위에 늘 보이는 최소화·종료 단추용 — 메뉴(창 ▸ 최소화/닫기)와
   같은 동작이지만, 터치로 드롭다운을 여는 것보다 한 번에 눌러 쓰기 쉽게 */
ipcMain.handle('get-webview-preload-path', () =>
  'file://' + path.join(__dirname, 'webview-preload.js').replace(/\\/g, '/'));

ipcMain.handle('win-minimize', (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  if (win) win.minimize();
});

/* 서재의 "네이버 열기" — 네이버는 일렉트론 내장 화면에서 접속을 막아서,
   네이버가 직접 만든 웨일 브라우저로 따로 띄웁니다(진짜 브라우저라
   100% 정상 동작합니다). "start"의 빈 제목(" ")은 경로에 공백이 있을 때
   그 첫 인자를 제목으로 착각하는 cmd.exe의 오랜 버릇을 피하기 위함입니다. */
ipcMain.handle('open-naver-whale', () => {
  execFile('cmd.exe', ['/c', 'start', '', 'whale', 'https://www.naver.com', '--start-maximized'], { windowsHide: true });
});


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
  const [hasPowerPoint, hasHangul, hasExcel] = await Promise.all([
    regKeyExists('PowerPoint.Application'),
    regKeyExists('HWPFrame.HwpObject'),
    regKeyExists('Excel.Application'),
  ]);
  return { hasPowerPoint, hasHangul, hasExcel };
}

function warnMissing({ hasPowerPoint, hasHangul, hasExcel }) {
  const missing = [];
  if (!hasPowerPoint) missing.push('파워포인트');
  if (!hasHangul) missing.push('한글');
  if (!hasExcel) missing.push('엑셀');
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
  '.xls': 'excel', '.xlsx': 'excel',
};

/* 패키지된 앱은 convert.ps1 을 app.asar 밖(app.asar.unpacked)에 둡니다 —
   asar 안의 경로는 Electron 안에서만 보이는 가상 경로라, powershell.exe
   같은 바깥 프로그램은 그 안의 파일을 읽을 수 없기 때문입니다. */
const scriptPath = app.isPackaged
  ? path.join(process.resourcesPath, 'app.asar.unpacked', 'convert.ps1')
  : path.join(__dirname, 'convert.ps1');

function convertToPdf(inputPath) {
  return new Promise((resolve, reject) => {
    const kind = OFFICE_EXT_KIND[path.extname(inputPath).toLowerCase()];
    if (!kind) return reject(new Error('지원하지 않는 파일 형식입니다 (파워포인트·한글·엑셀만 됩니다).'));

    const outputPath = path.join(os.tmpdir(), `noteall-${Date.now()}-${Math.random().toString(36).slice(2)}.pdf`);
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

const OFFICE_PICK_FILTERS = {
  ppt: { title: '파워포인트 파일 열기', filters: [{ name: '파워포인트', extensions: ['pptx', 'ppt'] }] },
  hwp: { title: '한글 파일 열기', filters: [{ name: '한글', extensions: ['hwpx', 'hwp'] }] },
  excel: { title: '엑셀 파일 열기', filters: [{ name: '엑셀', extensions: ['xlsx', 'xls'] }] },
  all: { title: '파워포인트·한글·엑셀 파일 열기', filters: [{ name: '파워포인트·한글·엑셀', extensions: ['pptx', 'ppt', 'hwp', 'hwpx', 'xlsx', 'xls'] }] },
};

ipcMain.handle('pick-office-file', async (event, kind) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const picked = OFFICE_PICK_FILTERS[kind] || OFFICE_PICK_FILTERS.all;
  const { canceled, filePaths } = await dialog.showOpenDialog(win, {
    title: picked.title,
    filters: picked.filters,
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

/* 수동 업데이트 확인 — 시작할 때 조용히 하는 확인과 달리, 눌러서 하는
   확인은 결과를 알려줘야 "눌렀는데 아무 반응 없네" 가 안 됩니다. */
function checkForUpdatesManually(win) {
  const cleanup = () => {
    autoUpdater.removeListener('update-available', onAvailable);
    autoUpdater.removeListener('update-not-available', onNotAvailable);
    autoUpdater.removeListener('error', onError);
  };
  const onAvailable = (info) => {
    cleanup();
    dialog.showMessageBox(win, {
      type: 'info', title: '업데이트',
      message: `새 버전(${info.version})을 받고 있습니다.`,
      detail: '다 받으면 다음에 프로그램을 켤 때 자동으로 적용됩니다.',
      buttons: ['확인'],
    });
  };
  const onNotAvailable = () => {
    cleanup();
    dialog.showMessageBox(win, {
      type: 'info', title: '업데이트', message: '이미 최신 버전입니다.', buttons: ['확인'],
    });
  };
  const onError = (err) => {
    cleanup();
    dialog.showMessageBox(win, {
      type: 'warning', title: '업데이트 확인 실패',
      message: '업데이트를 확인하지 못했습니다.',
      detail: String((err && err.message) || err),
      buttons: ['확인'],
    });
  };
  autoUpdater.once('update-available', onAvailable);
  autoUpdater.once('update-not-available', onNotAvailable);
  autoUpdater.once('error', onError);
  autoUpdater.checkForUpdates().catch(onError);
}

function buildMenu() {
  return Menu.buildFromTemplate([
    {
      label: '파일',
      submenu: [
        {
          label: '열기 (PDF·그림)',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'open'),
        },
        {
          label: '파워포인트·한글·엑셀 열기',
          click: (_mi, win) => win && win.webContents.send('menu-open-office', 'all'),
        },
        { type: 'separator' },
        {
          label: '저장 폴더 연결하기',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'connect-dir'),
        },
        {
          label: '이 쪽 저장',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'save-page'),
        },
        { type: 'separator' },
        {
          label: '닫기 (서재로)',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'close-file'),
        },
        { label: '종료', click: () => app.quit() },
      ],
    },
    {
      label: '도구',
      submenu: [
        {
          label: '도구상자 보이기·감추기',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'toggle-palette'),
        },
        {
          label: '폴더 설정',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'folder-settings'),
        },
        { type: 'separator' },
        {
          label: '전체화면',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'fullscreen'),
        },
        {
          label: '칠판 열기·닫기',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'board'),
        },
        { type: 'separator' },
        {
          label: '확대',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'zoom-in'),
        },
        {
          label: '축소',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'zoom-out'),
        },
        {
          label: '폭 맞춤',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'fit-width'),
        },
        {
          label: '쪽 맞춤',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'fit-page'),
        },
        { type: 'separator' },
        {
          label: '되돌리기',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'undo'),
        },
        {
          label: '쪽 비우기',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'clear-page'),
        },
        {
          label: '전체 쪽 비우기',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'clear-all-pages'),
        },
      ],
    },
    {
      label: '창',
      submenu: [
        { role: 'minimize', label: '최소화' },
        { role: 'close', label: '닫기' },
      ],
    },
    {
      label: '도움말',
      submenu: [
        {
          label: '사용법',
          click: (_mi, win) => win && win.webContents.send('menu-action', 'guide'),
        },
        { type: 'separator' },
        {
          label: '전자칠판 문제해결 매뉴얼',
          click: () => shell.openExternal('https://botem.com/kijun'),
        },
        {
          label: '업데이트 확인',
          click: (_mi, win) => checkForUpdatesManually(win),
        },
        { type: 'separator' },
        {
          label: `버전 정보 (${app.getVersion()})`,
          click: (_mi, win) => dialog.showMessageBox(win, {
            type: 'info', title: '버전 정보',
            message: '기전판서앱',
            detail: `버전 ${app.getVersion()}`,
            buttons: ['확인'],
          }),
        },
      ],
    },
  ]);
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    show: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, 'preload.js'),
      webviewTag: true,   /* 서재의 "크롬 열기" — 앱 창 안에 내장된 웹 화면 */
    },
  });
  /* 전자칠판 같은 큰 화면에서, note의 "화면이 좁으면 도구 막대를 두 줄로"
     규칙이 괜히 걸리지 않도록 — 창을 그 화면 크기에 맞춰 최대화합니다. */
  win.once('ready-to-show', () => {
    win.maximize();
    win.show();
  });
  win.loadFile(path.join(__dirname, 'index.html'));
  win.webContents.on('console-message', (_e, _level, message) => console.log('[renderer]', message));
  return win;
}

app.whenReady().then(async () => {
  Menu.setApplicationMenu(buildMenu());
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
