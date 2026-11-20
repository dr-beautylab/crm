// Code.gs(원본 Apps Script) → 클라우드플레어용 모듈로 감싸는 생성기
// 사용: node tools/gen.mjs <Code.gs 경로>
import fs from 'node:fs';
import path from 'node:path';

const src = fs.readFileSync(process.argv[2], 'utf8');
const here = path.dirname(new URL(import.meta.url).pathname);
const overrides = fs.readFileSync(path.join(here, '../src/overrides.js'), 'utf8');

const apiNames = [...new Set([...src.matchAll(/^function (api_[A-Za-z0-9_]+)\s*\(/gm), ...overrides.matchAll(/^function (api_[A-Za-z0-9_]+)\s*\(/gm)].map(m => m[1]))];

const out = `/* 자동 생성 파일 — 직접 고치지 말고 tools/gen.mjs 로 다시 만드세요 */
import { KDate, makeServices } from '../src/shim.js';

export function makeGas(book, io) {
  const Date = KDate;
  const { SpreadsheetApp, Utilities, UrlFetchApp, GmailApp, Logger, PropertiesService, CacheService,
          ScriptApp, CalendarApp, DriveApp, MailApp, HtmlService, Session, LockService } = makeServices(book, io);
  const globalThis = {};

/* ===== 원본 Code.gs ===== */
${src}

/* ===== 클라우드플레어용 덮어쓰기 ===== */
${overrides}

  const __API = { ${apiNames.map(n => `${n}: ${n}`).join(', ')} };
  return { __API, api, processBookingMails, autoSettle, sendReminderAdvance, sendReminderTomorrow,
           sendReminderToday, sendRetouchYear, dailyMaintenance, setup, publicForm, publicSubmit, authOf, addPhotoRow, expiredPhotos, CUR: () => CUR };
}
`;
fs.writeFileSync(path.join(here, '../functions/_gas.js'), out);
console.log('생성 완료: api', apiNames.length, '개');
