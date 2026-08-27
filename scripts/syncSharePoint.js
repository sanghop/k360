#!/usr/bin/env node
// syncSharePoint.js
// GitHub Actions가 하루 한 번 실행하는 스크립트입니다.
// 쉐어포인트의 "프로젝트 폴더 > 층 폴더 > 날짜 폴더(영상+궤적)" 구조를 읽어서,
// index.json 파일 하나로 정리합니다. 이 파일을 GitHub Pages 뷰어가 읽어서 목록을 보여줍니다.
//
// 폴더 구조 가정:
//   [공유 폴더 최상위]
//     └ 프로젝트명 폴더
//         └ 층 폴더 (예: "1F")
//             ├ floorplan.png (또는 .jpg)  <- 그 층의 평면도, 선택사항
//             └ 날짜 폴더 (예: "2026-08-19")
//                 ├ sample_360.mp4 (영상)
//                 └ trajectory.json (궤적)
//
// 필요한 환경변수: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, SHAREPOINT_FOLDER_URL

const fs = require("fs");
const path = require("path");
const {
  getAccessToken,
  resolveShareLink,
  listChildren,
  createAnonymousViewLink,
  getFileContent,
} = require("../server/services/graphClient");

const OUTPUT_PATH = path.join(__dirname, "..", "docs", "index.json");

function isVideoFile(name) {
  return /\.(mp4|mov)$/i.test(name);
}
function isImageFile(name) {
  return /\.(png|jpg|jpeg)$/i.test(name);
}
function isDateFolderName(name) {
  return /^\d{4}-\d{2}-\d{2}/.test(name); // "2026-08-19" 또는 "2026-08-19_2" 형태 허용
}

async function processDateFolder(token, driveId, folder) {
  const children = await listChildren(token, driveId, folder.id);
  const videoFile = children.find((c) => c.file && isVideoFile(c.name));
  const trajectoryFile = children.find((c) => c.file && c.name.toLowerCase() === "trajectory.json");

  if (!videoFile || !trajectoryFile) {
    console.warn(`  경고: "${folder.name}" 폴더에 영상 또는 trajectory.json이 없어서 건너뜁니다.`);
    return null;
  }

  const videoUrl = await createAnonymousViewLink(token, driveId, videoFile.id);
  const trajectoryText = await getFileContent(token, driveId, trajectoryFile.id);
  let points;
  try {
    points = JSON.parse(trajectoryText);
  } catch (err) {
    console.warn(`  경고: "${folder.name}"의 trajectory.json 파싱 실패:`, err.message);
    return null;
  }

  return { date: folder.name, videoUrl, points };
}

async function processFloorFolder(token, driveId, floorFolder) {
  const children = await listChildren(token, driveId, floorFolder.id);
  const floorplanFile = children.find((c) => c.file && isImageFile(c.name));
  const dateFolders = children.filter((c) => c.folder && isDateFolderName(c.name));

  let floorplanUrl = null;
  if (floorplanFile) {
    floorplanUrl = await createAnonymousViewLink(token, driveId, floorplanFile.id);
  }

  const runs = [];
  for (const dateFolder of dateFolders) {
    console.log(`  - 날짜 폴더 처리 중: ${dateFolder.name}`);
    const run = await processDateFolder(token, driveId, dateFolder);
    if (run) runs.push(run);
  }

  return { floor: floorFolder.name, floorplanUrl, runs };
}

async function processProjectFolder(token, driveId, projectFolder) {
  const children = await listChildren(token, driveId, projectFolder.id);
  const floorFolders = children.filter((c) => c.folder);

  const floors = [];
  for (const floorFolder of floorFolders) {
    console.log(` 층 폴더 처리 중: ${floorFolder.name}`);
    floors.push(await processFloorFolder(token, driveId, floorFolder));
  }

  return { projectName: projectFolder.name, floors };
}

async function main() {
  const shareUrl = process.env.SHAREPOINT_FOLDER_URL;
  if (!shareUrl) throw new Error("SHAREPOINT_FOLDER_URL 환경변수가 필요합니다.");

  console.log("Graph API 인증 중...");
  const token = await getAccessToken();

  console.log("공유 폴더 정보 확인 중...");
  const rootFolder = await resolveShareLink(token, shareUrl);
  const driveId = rootFolder.parentReference.driveId;

  console.log("최상위 폴더 목록 조회 중...");
  const projectFolders = (await listChildren(token, driveId, rootFolder.id)).filter((c) => c.folder);

  const projects = [];
  for (const projectFolder of projectFolders) {
    console.log(`프로젝트 폴더 처리 중: ${projectFolder.name}`);
    projects.push(await processProjectFolder(token, driveId, projectFolder));
  }

  const index = { generatedAt: new Date().toISOString(), projects };

  fs.mkdirSync(path.dirname(OUTPUT_PATH), { recursive: true });
  fs.writeFileSync(OUTPUT_PATH, JSON.stringify(index, null, 2), "utf-8");
  console.log(`완료: ${OUTPUT_PATH} (${projects.length}개 프로젝트)`);
}

main().catch((err) => {
  console.error("동기화 실패:", err);
  process.exit(1);
});
