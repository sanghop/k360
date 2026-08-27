// graphClient.js
// Graph API에 앱 전용 인증(client credentials)으로 접속해서, 쉐어포인트 폴더 내용을
// 재귀적으로 조회하고, 각 파일의 "누구나 볼 수 있는 링크"를 만드는 기능을 제공합니다.
//
// 필요한 환경변수: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET

const { encodeShareUrl } = require("./shareLinkEncoder");

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/**
 * client credentials 방식으로 앱 전용 액세스 토큰을 발급받습니다.
 */
async function getAccessToken() {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET 환경변수가 모두 필요합니다.");
  }

  const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    }),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`토큰 발급 실패 (${res.status}): ${text}`);
  }
  const data = await res.json();
  return data.access_token;
}

async function graphGet(token, path) {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API 요청 실패 (${path}) (${res.status}): ${text}`);
  }
  return res.json();
}

/**
 * 쉐어포인트 공유 링크(폴더)를 실제 driveItem(폴더) 정보로 변환합니다.
 */
async function resolveShareLink(token, shareUrl) {
  const encoded = encodeShareUrl(shareUrl);
  return graphGet(token, `/shares/${encoded}/driveItem`);
}

/**
 * 특정 폴더(driveId + itemId)의 하위 항목(파일/폴더) 목록을 가져옵니다.
 */
async function listChildren(token, driveId, itemId) {
  const data = await graphGet(token, `/drives/${driveId}/items/${itemId}/children`);
  return data.value || [];
}

/**
 * 특정 파일에 대해 "링크가 있으면 누구나 볼 수 있는" 영구 공유 링크를 만듭니다.
 * (이미 같은 설정으로 만들어진 링크가 있으면 Graph가 그걸 재사용해서 돌려줍니다)
 */
async function createAnonymousViewLink(token, driveId, itemId) {
  const res = await fetch(`${GRAPH_BASE}/drives/${driveId}/items/${itemId}/createLink`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ type: "view", scope: "anonymous" }),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`공유 링크 생성 실패 (${res.status}): ${text}`);
  }
  const data = await res.json();
  return data.link?.webUrl;
}

/**
 * 파일 내용을 텍스트로 직접 읽어옵니다 (trajectory.json 같은 작은 파일용).
 */
async function getFileContent(token, driveId, itemId) {
  const res = await fetch(`${GRAPH_BASE}/drives/${driveId}/items/${itemId}/content`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`파일 내용 조회 실패 (${res.status})`);
  return res.text();
}

module.exports = {
  getAccessToken,
  resolveShareLink,
  listChildren,
  createAnonymousViewLink,
  getFileContent,
};
