/* 서재 "크롬 열기" 안의 <webview> 전용 — 네이버 같은 사이트는 링크를
   window.open()으로 "새 탭"처럼 열고, 그 반환값(새 창 객체)을 계속
   붙잡고 씁니다(포커스 주기, 열려 있는지 확인 등). 저희는 새 창을 진짜로
   만들지 않고 지금 화면에 그대로 이어서 열기 때문에, window.open()이
   아무것도 못 돌려주면 그 뒤 코드가 멈춰 버립니다 — 그래서 가짜 창
   객체를 하나 돌려줘서, 사이트 쪽 코드가 계속 이어지게 둡니다. */
window.open = function (url) {
  try {
    if (url) window.location.href = url;
  } catch (e) {}
  return { closed: false, focus() {}, close() {}, postMessage() {}, location: { href: url || '' } };
};
