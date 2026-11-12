const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('noteall', {
  // 파일 선택 → COM 자동화로 PDF 변환 → {bytes, name} 또는 취소 시 null.
  // 실패하면 사람이 읽을 수 있는 메시지로 reject 합니다.
  importOfficeFile: async (kind) => {
    const filePath = await ipcRenderer.invoke('pick-office-file', kind);
    if (!filePath) return null;
    return ipcRenderer.invoke('import-office', filePath);
  },
  // 메뉴 "파일 ▸ 파워포인트 열기 / 한글 열기"를 눌렀을 때 — 화면의 버튼을
  // 그대로 눌러주는 것과 같습니다(같은 코드 경로를 씁니다). kind: 'ppt'|'hwp'.
  onMenuOpenOffice: (cb) => ipcRenderer.on('menu-open-office', (_e, kind) => cb(kind)),
  // 메뉴 "도구" 아래 항목들(도구상자 보이기·전체화면·칠판 열기·되돌리기) —
  // action 이름만 넘겨주고, 어떤 버튼을 누를지는 index.html 쪽에서 정합니다.
  onMenuAction: (cb) => ipcRenderer.on('menu-action', (_e, action) => cb(action)),
  // 저장 폴더 — Node fs를 직접 쓰는 쪽입니다(브라우저 showDirectoryPicker 대신).
  fsop: {
    getDir: () => ipcRenderer.invoke('fsop-get-dir'),
    pickDir: () => ipcRenderer.invoke('fsop-pick-dir'),
    readDir: (p) => ipcRenderer.invoke('fsop-read-dir', p),
    mkdir: (p) => ipcRenderer.invoke('fsop-mkdir', p),
    readFile: (p) => ipcRenderer.invoke('fsop-read-file', p),
    writeFile: (p, data) => ipcRenderer.invoke('fsop-write-file', { path: p, data }),
    remove: (p, recursive) => ipcRenderer.invoke('fsop-remove', { path: p, recursive }),
    exists: (p) => ipcRenderer.invoke('fsop-exists', p),
  },
  winMinimize: () => ipcRenderer.invoke('win-minimize'),
  openWhale: () => ipcRenderer.invoke('open-naver-whale'),
  getAppVersion: () => ipcRenderer.invoke('get-app-version'),
  isUpdateReady: () => ipcRenderer.invoke('is-update-ready'),
  installUpdateNow: () => ipcRenderer.invoke('install-update-now'),
  // 서재 "웨일 열기"의 <webview>에 심을 전용 preload 경로(window.open 가짜 탭 보강용)
  // (샌드박스 preload에서는 require('path')도 __dirname도 못 써서, 메인 프로세스에 물어봅니다)
  getWebviewPreloadPath: () => ipcRenderer.invoke('get-webview-preload-path'),
});
