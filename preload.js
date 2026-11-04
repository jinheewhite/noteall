const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('noteall', {
  // 파일 선택 → COM 자동화로 PDF 변환 → {bytes, name} 또는 취소 시 null.
  // 실패하면 사람이 읽을 수 있는 메시지로 reject 합니다.
  importOfficeFile: async () => {
    const filePath = await ipcRenderer.invoke('pick-office-file');
    if (!filePath) return null;
    return ipcRenderer.invoke('import-office', filePath);
  },
});
