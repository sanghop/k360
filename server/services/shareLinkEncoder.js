// shareLinkEncoder.js
// 쉐어포인트 공유 링크(사람이 브라우저에서 여는 그 주소)를, Graph API의 /shares/ 엔드포인트가
// 요구하는 형식으로 변환합니다. (Microsoft 공식 문서에 정의된 인코딩 규칙을 그대로 구현)
//
// 규칙: 1) URL을 base64로 인코딩  2) 끝의 '=' 제거, '/'는 '_'로, '+'는 '-'로 치환  3) 앞에 "u!" 붙이기

function encodeShareUrl(url) {
  const base64Value = Buffer.from(url, "utf-8").toString("base64");
  const unpadded = base64Value.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  return `u!${unpadded}`;
}

module.exports = { encodeShareUrl };
