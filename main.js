/* noteall — note(판서노트)의 화면·기능은 그대로 두고, 여기에 파워포인트·
   한글(한쇼) 불러오기를 더합니다. 1단계: 슬라이드를 이미지로 뽑아 note
   화면에 페이지처럼 띄웁니다. 그 전 단계로, COM 자동화가 되려면 그 PC에
   프로그램이 실제로 설치되어 있어야 하므로, 없으면 켤 때 바로 알려 줍니다. */
const { app, BrowserWindow, dialog } = require('electron');
const { execFile } = require('child_process');
const path = require('path');

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

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
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
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
