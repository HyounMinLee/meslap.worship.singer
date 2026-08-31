/**
 * ============================================================
 *  예배 준비 도구 — 설정 파일
 *
 *  이 파일은 index.html 과 같은 폴더에 두세요.
 *  index.html 을 새 것으로 바꿔도 이 파일은 그대로 두면 되므로,
 *  중계 서버 주소를 매번 다시 넣을 필요가 없습니다.
 *
 *  ※ 이 파일에는 GitHub 토큰이나 접속 암호를 절대 넣지 마세요.
 *    화면을 여는 사람 누구나 볼 수 있습니다. 비밀은 중계 서버에만 둡니다.
 * ============================================================
 */
window.WORSHIP_CONFIG = {

  /* ---------- 반드시 채워야 하는 값 ---------- */

  // Apps Script 웹 앱 주소 (배포 → 배포 관리 에서 복사)
  PROXY_URL: "https://script.google.com/macros/s/AKfycbzp4G2CZ8zB4dN8nJ0BHrJCzkeZY5Y4pOu0UBIOJ0esnZ9eVruZopp9BSZsXawamSpoMQ/exec",


  /* ---------- 여기부터는 없어도 됩니다 (지우면 기본값 사용) ---------- */

  // PPT 만들기 화면이 처음 열릴 때의 기본 설정
  DEFAULT_FONT: "DX시인과나",   // "DX시인과나" 또는 "맑은 고딕"
  DEFAULT_DECK_SIZE: 33,        // 글자 크기 (20~60, 짝수)
  DEFAULT_RATIO: "4x3",        // "16x9" 또는 "4x3"
  DEFAULT_LINE_SPACING: "1.5",  // "1", "1.5", "2"
  DEFAULT_BG: "light",           // "dark"(검정 바탕) 또는 "light"(흰 바탕)
  DEFAULT_LANG: "ko",           // "ko"(한글만) / "en"(영어만) / "both"(한글+영어)

  // PDF 만들기 화면의 콘티 글자 크기 기본값 (18~30)
  DEFAULT_CONTE_SIZE: 22
};
