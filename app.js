/* global XLSX */

const STORAGE_KEY = "tuyenDungWorkflow_v1";

/** Google Sheet — link chia sẻ /edit hoặc xuất bản pubhtml */
const DEFAULT_GOOGLE_SHEET_PUB_URL =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vQs96Y38iag14gR2tVg72YahtcmbfOFyuZbwUNFlzp7qA2juTWHxon6SDe6hRdLDBpx_sDkcMGvV8MR/pubhtml";

/** Google AI Studio (Gemini) — mặc định trong app, không hiển thị trên giao diện */
const DEFAULT_GEMINI_API_KEY = "";
const DEFAULT_GEMINI_MODEL = "gemini-2.5-flash";

const STATUS_LABELS = {
  new: "Chưa tạo câu hỏi",
  generating: "Đang tạo AI",
  pending_hr: "Chờ HR duyệt",
  approved: "Đã duyệt",
  sent_hrm: "Đã gửi HRM",
  done: "Hoàn tất",
  rejected: "Từ chối",
};

const AUTO_GEN_CONCURRENCY = 2;

const HEADER_ALIASES = {
  name: ["họ và tên", "ho ten", "họ tên", "hoten", "ten", "tên", "fullname", "full name", "name", "ung vien", "ứng viên"],
  email: ["email", "mail", "e-mail"],
  phone: ["sdt", "điện thoại", "dien thoai", "phone", "mobile", "số điện thoại"],
  industry: ["ngành", "nganh", "chuyên ngành", "chuyen nganh", "industry", "lĩnh vực", "linh vuc", "sector", "nhóm ngành"],
  position: ["vị trí", "vi tri", "position", "job", "title", "chức danh", "chuc danh", "vị trí ứng tuyển"],
  experience: ["kinh nghiệm", "kinh nghiem", "experience", "số năm", "so nam"],
  education: ["học vấn", "hoc van", "education", "bằng cấp", "bang cap", "trình độ", "trinh do"],
  skills: ["kỹ năng", "ky nang", "skills", "skill", "competency"],
  note: ["ghi chú", "ghi chu", "note", "mô tả", "mo ta", "description"],
};

function norm(s) {
  return String(s ?? "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function pickMappedRow(row, headerMap) {
  const out = {};
  for (const [key, colIdx] of Object.entries(headerMap)) {
    if (colIdx >= 0) out[key] = row[colIdx];
  }
  return out;
}

function detectHeaderMap(headers) {
  const normalized = headers.map((h, i) => ({ raw: h, i, n: norm(h) }));
  const map = {
    name: -1,
    email: -1,
    phone: -1,
    industry: -1,
    position: -1,
    experience: -1,
    education: -1,
    skills: -1,
    note: -1,
  };

  for (const key of Object.keys(map)) {
    const aliases = HEADER_ALIASES[key];
    for (const { i, n } of normalized) {
      if (!n) continue;
      if (aliases.some((a) => n === a || n.includes(a))) {
        map[key] = i;
        break;
      }
    }
  }

  if (map.name < 0 && headers.length) {
    map.name = 0;
  }

  return map;
}

function rowToCandidate(rowArr, headers, headerMap, sheetName, rowIndex) {
  const cells = rowArr.map((c) => (c === undefined || c === null ? "" : c));
  const mapped = pickMappedRow(cells, headerMap);
  const displayName =
    String(mapped.name || "").trim() ||
    `Ứng viên dòng ${rowIndex + 2}`;

  const raw = {};
  headers.forEach((h, idx) => {
    if (h != null && String(h).trim() !== "") raw[String(h).trim()] = cells[idx];
  });

  return {
    id: `${sheetName}::${rowIndex}`,
    sheetName,
    rowIndex,
    displayName,
    industry: String(mapped.industry || "").trim() || "—",
    position: String(mapped.position || "").trim() || "—",
    email: String(mapped.email || "").trim(),
    phone: String(mapped.phone || "").trim(),
    experience: String(mapped.experience || "").trim(),
    education: String(mapped.education || "").trim(),
    skills: String(mapped.skills || "").trim(),
    note: String(mapped.note || "").trim(),
    raw,
  };
}

/**
 * Chuyển link Google Sheet thành URL CSV:
 * - /edit, /view, /d/{id} → export?format=csv (cần quyền «Bất kỳ ai có link» xem được)
 * - /pubhtml, /pub → pub?output=csv (xuất bản lên web)
 */
function toGoogleSheetCsvUrl(pubUrl) {
  const raw = String(pubUrl ?? "").trim();
  if (!raw) return "";
  let u;
  try {
    u = new URL(raw);
  } catch {
    return "";
  }
  if (!u.hostname.includes("google.com") || !u.pathname.includes("/spreadsheets/")) return "";

  const path = u.pathname.replace(/\/+$/u, "");

  const docMatch = path.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (docMatch && !path.includes("/e/")) {
    const id = docMatch[1];
    const gid = u.searchParams.get("gid") || "0";
    return `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${encodeURIComponent(gid)}`;
  }

  if (path.includes("/e/")) {
    const pubPath = path.replace(/\/pubhtml\/?$/i, "/pub");
    if (pubPath.endsWith("/pub")) {
      const base = `${u.protocol}//${u.host}${pubPath}`;
      return `${base}?output=csv`;
    }
  }

  return "";
}

/** Parse CSV (hỗ trợ dấu ngoặc kép, xuống dòng trong ô). */
function parseCsv(text) {
  const t = String(text ?? "").replace(/^\uFEFF/, "");
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    const next = t[i + 1];
    if (inQuotes) {
      if (c === '"') {
        if (next === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ",") {
      row.push(field);
      field = "";
    } else if (c === "\n" || (c === "\r" && next === "\n")) {
      if (c === "\r") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (c !== "\r") {
      field += c;
    }
  }
  row.push(field);
  if (row.some((cell) => String(cell).trim() !== "") || rows.length === 0) {
    rows.push(row);
  }
  return rows.map((r) => r.map((cell) => (cell === undefined || cell === null ? "" : cell)));
}

function candidatesFromMatrix(rows, sheetName) {
  if (!rows.length) return [];

  let headerRowIdx = 0;
  let headers = (rows[0] || []).map((h) => String(h).trim());

  const nonEmpty = headers.filter(Boolean).length;
  if (nonEmpty < 2 && rows.length > 1) {
    headerRowIdx = 1;
    headers = (rows[1] || []).map((h) => String(h).trim());
  }

  const headerMap = detectHeaderMap(headers);
  const dataStart = headerRowIdx + 1;
  const out = [];

  for (let r = dataStart; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row.some((c) => String(c).trim() !== "")) continue;
    out.push(rowToCandidate(row, headers, headerMap, sheetName, r));
  }
  return out;
}

function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const candidates = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!rows.length) continue;
    candidates.push(...candidatesFromMatrix(rows, sheetName));
  }

  return candidates;
}

/**
 * Tải nội dung CSV từ URL Google (pub).
 * Trình duyệt thường báo "Failed to fetch" khi mở file qua file:// hoặc khi Google không cho CORS —
 * khi đó thử lại qua api.allorigins.win (chỉ dùng cho sheet đã công khai).
 */
async function fetchPublishedCsvText(csvUrl) {
  const direct = async () => {
    const res = await fetch(csvUrl, { credentials: "omit", cache: "no-store" });
    if (!res.ok) throw new Error(`Tải sheet lỗi HTTP ${res.status}`);
    return res.text();
  };

  const viaAllOrigins = async () => {
    const proxy = `https://api.allorigins.win/get?url=${encodeURIComponent(csvUrl)}`;
    const res = await fetch(proxy, { credentials: "omit", cache: "no-store" });
    if (!res.ok) throw new Error(`Proxy AllOrigins HTTP ${res.status}`);
    const data = await res.json();
    const raw = data?.contents;
    if (typeof raw !== "string") throw new Error("Proxy không trả về nội dung.");
    return raw;
  };

  let text;
  let usedProxy = false;
  try {
    text = await direct();
  } catch (e1) {
    const m = String(e1?.message ?? e1).toLowerCase();
    const looksLikeBlockedFetch =
      e1 instanceof TypeError || m.includes("failed to fetch") || m.includes("networkerror");
    if (!looksLikeBlockedFetch) throw e1;
    try {
      text = await viaAllOrigins();
      usedProxy = true;
    } catch (e2) {
      throw new Error(
        `Không tải được sheet (thường do mở app bằng file:// — trình duyệt chặn CORS). ` +
          `Hãy mở qua địa chỉ http://localhost (VS Code «Live Server», hoặc chạy \`npx serve .\` trong thư mục app). ` +
          `Lỗi gốc: ${e1?.message || e1}; proxy: ${e2?.message || e2}`
      );
    }
  }

  const t = text.trim();
  if (t.startsWith("<!DOCTYPE") || (t.includes("<html") && t.length < 8000)) {
    throw new Error(
      "Google trả về trang HTML thay vì CSV — kiểm tra Sheet vẫn bật «Xuất bản lên web» và link pubhtml đúng."
    );
  }

  return { text, usedProxy };
}

async function loadCandidatesFromGoogleSheet(pubUrl) {
  const csvUrl = toGoogleSheetCsvUrl(pubUrl);
  if (!csvUrl) {
    throw new Error("Không chuyển được link Google Sheet sang CSV. Dùng link dạng .../pubhtml hoặc .../pub.");
  }
  const { text, usedProxy } = await fetchPublishedCsvText(csvUrl);
  const rows = parseCsv(text);
  const list = candidatesFromMatrix(rows, "Danh sách (Google Sheet)");
  return { candidates: list, usedProxy };
}

function defaultWorkflowEntry() {
  return {
    status: "new",
    aiSuggestions: [],
    questions: [],
    questionsText: "",
    generating: false,
    hrComment: "",
    approvedAt: null,
    rejectedAt: null,
    hrmSentAt: null,
    hrmResponse: null,
    reportAt: null,
    log: [],
  };
}

function migrateEntry(entry) {
  if (!Array.isArray(entry.aiSuggestions)) entry.aiSuggestions = [];
  if (!Array.isArray(entry.questions)) entry.questions = [];
  if (entry.generating === undefined) entry.generating = false;
  if (
    entry.status === "pending_hr" &&
    entry.questions.length > 0 &&
    entry.aiSuggestions.length === 0
  ) {
    entry.aiSuggestions = [...entry.questions];
    entry.questions = [];
  }
}

function getEffectiveStatus(entry) {
  if (entry.generating) return "generating";
  return entry.status;
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { byId: {}, settings: {} };
    return JSON.parse(raw);
  } catch {
    return { byId: {}, settings: {} };
  }
}

function saveState(state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function pushLog(entry, message) {
  entry.log = entry.log || [];
  entry.log.push({ t: new Date().toISOString(), message });
}

function templateQuestions(candidate) {
  const pos = candidate.position || "vị trí ứng tuyển";
  const ind = candidate.industry || "ngành";
  const skills = candidate.skills || "kỹ năng được nêu trong hồ sơ";
  const edu = candidate.education || "trình độ của bạn";
  const exp = candidate.experience || "kinh nghiệm làm việc";

  return [
    `Với vị trí ${pos} trong lĩnh vực ${ind}, anh/chị mô tả ngắn gọn một dự án hoặc công việc tiêu biểu nhất liên quan trực tiếp đến vai trò này?`,
    `Anh/chị đánh giá thế mạnh chính của mình so với các ứng viên khác cho vị trí ${pos} là gì?`,
    `Kỹ năng/kiến thức nào trong nhóm "${skills}" anh/chị đã áp dụng thực tế và kết quả đạt được?`,
    `Liên quan ${edu} và ${exp}, anh/chị đã học được điều gì quan trọng nhất cho công việc hiện tại?`,
    `Trong môi trường ${ind}, anh/chị xử lý thế nào khi deadline gấp nhưng chất lượng vẫn phải đảm bảo?`,
    `Một tình huống khó với đồng nghiệp/khách hàng nội bộ: anh/chị chọn hướng xử lý nào và vì sao?`,
    `Trong 90 ngày đầu nếu được nhận, anh/chị ưu tiên 3 mục tiêu cụ thể nào và cách đo lường?`,
    `Anh/chị có câu hỏi nào cho chúng tôi về vị trí, đội ngũ hoặc văn hóa làm việc không?`,
  ];
}

function parseQuestionsJsonFromModelText(text) {
  const raw = String(text ?? "").trim();
  if (!raw) throw new Error("Gemini không trả về nội dung.");
  try {
    return JSON.parse(raw);
  } catch {
    const m = raw.match(/\{[\s\S]*\}/);
    if (!m) throw new Error("Không parse được JSON từ Gemini.");
    return JSON.parse(m[0]);
  }
}

async function generateQuestionsWithGemini(candidate, apiKey, modelId) {
  const model = String(modelId || DEFAULT_GEMINI_MODEL).replace(/^models\//, "");
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(
    model
  )}:generateContent?key=${encodeURIComponent(apiKey)}`;

  const systemText =
    'Bạn là chuyên gia tuyển dụng. Chỉ trả về một JSON hợp lệ dạng {"questions": ["...", ...]} với đúng 8 chuỗi: mỗi chuỗi là một câu hỏi phỏng vấn tiếng Việt, súc tích, không trùng lặp, phù hợp ứng viên và vị trí.';

  const userText = JSON.stringify({
    candidate: {
      name: candidate.displayName,
      industry: candidate.industry,
      position: candidate.position,
      education: candidate.education,
      experience: candidate.experience,
      skills: candidate.skills,
      note: candidate.note,
    },
  });

  const body = {
    systemInstruction: { parts: [{ text: systemText }] },
    contents: [{ role: "user", parts: [{ text: userText }] }],
    generationConfig: {
      temperature: 0.5,
      responseMimeType: "application/json",
    },
  };

  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  const textRaw = await res.text();
  let data = {};
  try {
    data = textRaw ? JSON.parse(textRaw) : {};
  } catch {
    if (!res.ok) throw new Error(textRaw || res.statusText);
    throw new Error("Phản hồi Gemini không phải JSON.");
  }

  if (!res.ok) {
    const msg = data?.error?.message || textRaw || res.statusText;
    throw new Error(typeof msg === "string" ? msg : JSON.stringify(msg));
  }

  if (data?.promptFeedback?.blockReason) {
    throw new Error(`Yêu cầu bị chặn: ${data.promptFeedback.blockReason}`);
  }

  const cand = data?.candidates?.[0];
  const text = cand?.content?.parts?.map((p) => p.text).join("") || "";
  if (!text && cand?.finishReason) {
    throw new Error(`Gemini kết thúc sớm: ${cand.finishReason}`);
  }

  const parsed = parseQuestionsJsonFromModelText(text);
  const qs = Array.isArray(parsed.questions) ? parsed.questions.map(String) : [];
  if (qs.length < 3) throw new Error("Gemini trả về quá ít câu hỏi.");
  return qs.slice(0, 12);
}

function downloadBlob(filename, blob) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

/** --- App --- */

let candidates = [];
let state = loadState();
let drawerCandidateId = null;
let sheetFetchLoading = false;
let lastSheetError = "";
let autoGenActive = 0;
let autoGenQueued = 0;

const els = {
  fileInput: document.getElementById("fileInput"),
  btnReloadSheet: document.getElementById("btnReloadSheet"),
  sheetHint: document.getElementById("sheetHint"),
  autoGenHint: document.getElementById("autoGenHint"),
  filterIndustry: document.getElementById("filterIndustry"),
  filterStatus: document.getElementById("filterStatus"),
  search: document.getElementById("search"),
  stats: document.getElementById("stats"),
  candidateTableBody: document.getElementById("candidateTableBody"),
  tableEmpty: document.getElementById("tableEmpty"),
  hrmUrl: document.getElementById("hrmUrl"),
  geminiApiKey: document.getElementById("geminiApiKey"),
  geminiModel: document.getElementById("geminiModel"),
  btnExportState: document.getElementById("btnExportState"),
  btnImportState: document.getElementById("btnImportState"),
  btnExportCsv: document.getElementById("btnExportCsv"),
  stateImport: document.getElementById("stateImport"),
  drawer: document.getElementById("drawer"),
  drawerBackdrop: document.getElementById("drawerBackdrop"),
  drawerTitle: document.getElementById("drawerTitle"),
  drawerMeta: document.getElementById("drawerMeta"),
  drawerBody: document.getElementById("drawerBody"),
  btnCloseDrawer: document.getElementById("btnCloseDrawer"),
};

function getGeminiApiKey() {
  const fromUi = els.geminiApiKey?.value?.trim();
  if (fromUi) return fromUi;
  return (state.settings?.geminiApiKey || DEFAULT_GEMINI_API_KEY || "").trim();
}

function getGeminiModel() {
  const fromUi = els.geminiModel?.value?.trim();
  if (fromUi) return fromUi;
  return (state.settings?.geminiModel || DEFAULT_GEMINI_MODEL).trim();
}

function persistSettings() {
  state.settings = {
    ...(state.settings || {}),
    hrmUrl: els.hrmUrl?.value?.trim() || "",
    geminiApiKey: els.geminiApiKey?.value?.trim() || "",
    geminiModel: els.geminiModel?.value?.trim() || DEFAULT_GEMINI_MODEL,
  };
  saveState(state);
}

function hydrateSettings() {
  const s = state.settings || {};
  if (s.hrmUrl && els.hrmUrl) els.hrmUrl.value = s.hrmUrl;
  if (els.geminiApiKey) {
    els.geminiApiKey.value = s.geminiApiKey || DEFAULT_GEMINI_API_KEY || "";
  }
  if (els.geminiModel) {
    els.geminiModel.value = s.geminiModel || DEFAULT_GEMINI_MODEL;
  }
}

function ensureEntry(id) {
  if (!state.byId[id]) state.byId[id] = defaultWorkflowEntry();
  const entry = state.byId[id];
  migrateEntry(entry);
  return entry;
}

function updateAutoGenHint() {
  if (!els.autoGenHint) return;
  const busy = autoGenActive + autoGenQueued;
  if (busy > 0) {
    els.autoGenHint.classList.remove("hidden");
    els.autoGenHint.textContent = `Đang tạo câu hỏi AI: ${autoGenActive} đang chạy, ${autoGenQueued} chờ...`;
  } else {
    els.autoGenHint.classList.add("hidden");
    els.autoGenHint.textContent = "";
  }
}

async function generateForCandidate(candidate) {
  const entry = ensureEntry(candidate.id);
  const apiKey = getGeminiApiKey();
  if (!apiKey) {
    throw new Error("Chưa nhập Gemini API Key.");
  }
  entry.generating = true;
  renderTable();
  if (drawerCandidateId === candidate.id) renderDrawer();
  try {
    const qs = await generateQuestionsWithGemini(candidate, apiKey, getGeminiModel());
    entry.aiSuggestions = qs;
    entry.status = "pending_hr";
    entry.generating = false;
    pushLog(entry, "AI đã tạo câu hỏi gợi ý (Gemini).");
    saveState(state);
    return qs;
  } catch (e) {
    entry.generating = false;
    if (entry.status === "generating") entry.status = "new";
    saveState(state);
    throw e;
  } finally {
    renderTable();
    if (drawerCandidateId === candidate.id) renderDrawer();
  }
}

function scheduleAutoGenerate() {
  const apiKey = getGeminiApiKey();
  if (!apiKey) {
    if (els.autoGenHint) {
      els.autoGenHint.classList.remove("hidden");
      els.autoGenHint.textContent = "Nhập API Key để bật tự động tạo câu hỏi.";
    }
    return;
  }

  const queue = candidates.filter((c) => {
    const e = ensureEntry(c.id);
    if (e.generating) return false;
    if (e.aiSuggestions.length > 0) return false;
    return e.status === "new" || e.status === "rejected";
  });

  autoGenQueued = queue.length;
  updateAutoGenHint();

  const pump = async () => {
    while (autoGenActive < AUTO_GEN_CONCURRENCY && queue.length) {
      const cand = queue.shift();
      autoGenQueued = queue.length;
      autoGenActive++;
      updateAutoGenHint();
      generateForCandidate(cand)
        .catch((err) => {
          console.warn("Auto-gen failed:", cand.displayName, err);
          pushLog(ensureEntry(cand.id), `Lỗi tạo AI: ${err.message || err}`);
          saveState(state);
        })
        .finally(() => {
          autoGenActive--;
          updateAutoGenHint();
          renderStats();
          renderTable();
          pump();
        });
    }
  };
  pump();
}

function statusCounts() {
  const c = {
    new: 0,
    generating: 0,
    pending_hr: 0,
    approved: 0,
    sent_hrm: 0,
    done: 0,
    rejected: 0,
  };
  for (const cand of candidates) {
    const st = getEffectiveStatus(ensureEntry(cand.id));
    if (c[st] !== undefined) c[st]++;
  }
  return c;
}

function filteredCandidates() {
  const ind = els.filterIndustry.value;
  const st = els.filterStatus.value;
  const q = norm(els.search.value);

  return candidates.filter((c) => {
    if (ind && c.industry !== ind) return false;
    const entry = ensureEntry(c.id);
    const effective = getEffectiveStatus(entry);
    if (st && effective !== st) return false;
    if (q) {
      const hay = norm(
        [c.displayName, c.position, c.industry, c.email, c.phone, c.skills, c.experience].join(" ")
      );
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

function renderIndustryFilter() {
  const set = new Set(candidates.map((c) => c.industry).filter(Boolean));
  const prev = els.filterIndustry.value;
  els.filterIndustry.innerHTML = '<option value="">Tất cả ngành</option>';
  [...set].sort().forEach((v) => {
    const o = document.createElement("option");
    o.value = v;
    o.textContent = v;
    els.filterIndustry.appendChild(o);
  });
  if ([...set].includes(prev)) els.filterIndustry.value = prev;
}

function renderStats() {
  const c = statusCounts();
  els.stats.textContent =
    `Tổng ${candidates.length} — Mới: ${c.new} | AI: ${c.generating} | Chờ duyệt: ${c.pending_hr} | Đã duyệt: ${c.approved} | HRM: ${c.sent_hrm} | Xong: ${c.done} | Từ chối: ${c.rejected}`;
}

function badgeHtml(statusKey) {
  const label = STATUS_LABELS[statusKey] || statusKey;
  return `<span class="badge ${statusKey}">${label}</span>`;
}

function openDrawer(id) {
  drawerCandidateId = id;
  els.drawer.classList.remove("hidden");
  els.drawerBackdrop.classList.remove("hidden");
  requestAnimationFrame(() => {
    els.drawer.classList.add("is-open");
    els.drawerBackdrop.classList.add("is-open");
  });
  renderDrawer();
  renderTable();
}

function closeDrawer() {
  els.drawer.classList.remove("is-open");
  els.drawerBackdrop.classList.remove("is-open");
  setTimeout(() => {
    els.drawer.classList.add("hidden");
    els.drawerBackdrop.classList.add("hidden");
    drawerCandidateId = null;
    renderTable();
  }, 280);
}

function syncOfficialQuestions(entry) {
  entry.questionsText = (entry.questions || []).join("\n");
}

function addAiToOfficial(entry, text) {
  const t = String(text || "").trim();
  if (!t) return;
  if (!entry.questions) entry.questions = [];
  if (entry.questions.some((q) => norm(q) === norm(t))) return;
  entry.questions.push(t);
  syncOfficialQuestions(entry);
  saveState(state);
}

function approveCandidate(cand, entry) {
  const official = (entry.questions || []).map((q) => String(q).trim()).filter(Boolean);
  if (!official.length) {
    alert("Cần ít nhất một câu hỏi chính thức (bấm + trên câu AI hoặc nhập tay).");
    return false;
  }
  entry.questions = official;
  syncOfficialQuestions(entry);
  entry.status = "approved";
  entry.approvedAt = new Date().toISOString();
  pushLog(entry, "HR đã duyệt câu hỏi.");
  saveState(state);
  return true;
}

function rejectCandidate(entry, candidateId) {
  entry.status = "rejected";
  entry.rejectedAt = new Date().toISOString();
  entry.aiSuggestions = [];
  pushLog(entry, "HR từ chối — sẽ tạo lại câu hỏi AI.");
  saveState(state);
  const cand = candidates.find((x) => x.id === candidateId);
  if (cand) scheduleAutoGenerate();
}

async function sendToHrm(cand, entry) {
  const payload = {
    candidate: cand,
    questions: entry.questions,
    aiSuggestions: entry.aiSuggestions,
    hrComment: entry.hrComment,
    approvedAt: entry.approvedAt,
    clientTs: new Date().toISOString(),
  };
  const url = els.hrmUrl?.value?.trim() || "";
  try {
    if (url) {
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const txt = await res.text();
      entry.hrmResponse = { ok: res.ok, status: res.status, body: txt.slice(0, 4000) };
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
    } else {
      entry.hrmResponse = { mock: true, message: "Không cấu hình URL — mô phỏng thành công." };
    }
    entry.status = "sent_hrm";
    entry.hrmSentAt = new Date().toISOString();
    pushLog(entry, "Đã gửi sang hệ thống HR.");
    saveState(state);
    persistSettings();
    return true;
  } catch (e) {
    alert(`Gửi HRM lỗi: ${e.message || e}`);
    entry.hrmResponse = { error: String(e.message || e) };
    saveState(state);
    return false;
  }
}

function exportAllCsv() {
  const headers = [
    "id",
    "name",
    "industry",
    "email",
    "phone",
    "status",
    "ai_count",
    "official_count",
    "questions",
    "approvedAt",
    "hrmSentAt",
  ];
  const lines = [headers.join(",")];
  for (const cand of candidates) {
    const e = ensureEntry(cand.id);
    const row = [
      cand.id,
      cand.displayName,
      cand.industry,
      cand.email,
      cand.phone,
      getEffectiveStatus(e),
      (e.aiSuggestions || []).length,
      (e.questions || []).length,
      JSON.stringify((e.questions || []).join(" | ")),
      e.approvedAt || "",
      e.hrmSentAt || "",
    ].map((cell) => {
      const s = String(cell ?? "").replace(/"/g, '""');
      return `"${s}"`;
    });
    lines.push(row.join(","));
  }
  downloadBlob(
    `tom_tat_ung_vien_${new Date().toISOString().slice(0, 10)}.csv`,
    new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" })
  );
}

function renderTable() {
  if (!els.candidateTableBody) return;
  const list = filteredCandidates();
  els.candidateTableBody.innerHTML = "";

  if (!list.length) {
    els.tableEmpty?.classList.remove("hidden");
    if (sheetFetchLoading && !candidates.length) {
      els.tableEmpty.textContent = "Đang tải danh sách từ Google Sheet...";
    } else if (!candidates.length && lastSheetError) {
      els.tableEmpty.textContent = `Không tải được sheet: ${lastSheetError}`;
    } else if (!candidates.length) {
      els.tableEmpty.textContent = "Chưa có dữ liệu. Tải Google Sheet hoặc nhập Excel.";
    } else {
      els.tableEmpty.textContent = "Không có ứng viên khớp bộ lọc.";
    }
    return;
  }

  els.tableEmpty?.classList.add("hidden");

  for (const c of list) {
    const entry = ensureEntry(c.id);
    const st = getEffectiveStatus(entry);
    const aiN = entry.aiSuggestions?.length || 0;
    const offN = entry.questions?.length || 0;
    const tr = document.createElement("tr");
    if (c.id === drawerCandidateId) tr.classList.add("is-active");

    const canApprove = st === "pending_hr" && offN > 0;
    const canReject = st === "pending_hr" || st === "approved";
    const canHrm = st === "approved" && offN > 0;

    tr.innerHTML = `
      <td>
        <span class="cell-name" data-action="open"></span>
        <span class="cell-sub"></span>
      </td>
      <td class="cell-muted"></td>
      <td class="cell-muted"></td>
      <td class="cell-muted"></td>
      <td></td>
      <td class="q-count"></td>
      <td>
        <div class="row-actions">
          <button type="button" class="btn-icon btn-icon--view" data-action="open" title="Xem / sửa câu hỏi">👁</button>
          <button type="button" class="btn-icon btn-icon--approve" data-action="approve" title="Duyệt nhanh" ${canApprove ? "" : "disabled"}>✓</button>
          <button type="button" class="btn-icon btn-icon--reject" data-action="reject" title="Từ chối" ${canReject ? "" : "disabled"}>✕</button>
          <button type="button" class="btn-icon btn-icon--hrm" data-action="hrm" title="Gửi HRM" ${canHrm ? "" : "disabled"}>➤</button>
        </div>
      </td>
    `;

    tr.querySelector(".cell-name").textContent = c.displayName;
    tr.querySelector(".cell-sub").textContent = c.experience || c.position || "—";
    tr.cells[1].textContent = c.industry;
    tr.cells[2].textContent = c.email || "—";
    tr.cells[3].textContent = c.phone || "—";
    tr.cells[4].innerHTML = badgeHtml(st);
    tr.querySelector(".q-count").innerHTML = `AI <strong>${aiN}</strong> · Duyệt <strong>${offN}</strong>`;

    tr.querySelectorAll('[data-action="open"]').forEach((el) => {
      el.addEventListener("click", (ev) => {
        ev.stopPropagation();
        openDrawer(c.id);
      });
    });

    tr.querySelector('[data-action="approve"]')?.addEventListener("click", (ev) => {
      ev.stopPropagation();
      if (approveCandidate(c, entry)) {
        renderStats();
        renderTable();
        if (drawerCandidateId === c.id) renderDrawer();
      }
    });

    tr.querySelector('[data-action="reject"]')?.addEventListener("click", (ev) => {
      ev.stopPropagation();
      rejectCandidate(entry, c.id);
      renderStats();
      renderTable();
      if (drawerCandidateId === c.id) renderDrawer();
    });

    tr.querySelector('[data-action="hrm"]')?.addEventListener("click", async (ev) => {
      ev.stopPropagation();
      const ok = await sendToHrm(c, entry);
      if (ok) {
        renderStats();
        renderTable();
        if (drawerCandidateId === c.id) renderDrawer();
      }
    });

    els.candidateTableBody.appendChild(tr);
  }
}

function renderDrawer() {
  const c = candidates.find((x) => x.id === drawerCandidateId);
  if (!c || !els.drawerBody) return;

  const entry = ensureEntry(c.id);
  const st = getEffectiveStatus(entry);
  const aiList = entry.aiSuggestions?.length
    ? entry.aiSuggestions
    : entry.generating
      ? []
      : [];

  els.drawerTitle.textContent = c.displayName;
  els.drawerMeta.textContent = [c.industry, c.email, c.phone].filter(Boolean).join(" · ");

  const canApprove = st === "pending_hr" && (entry.questions?.length || 0) > 0;
  const canSend = st === "approved" && (entry.questions?.length || 0) > 0;
  const canFinalize = st === "sent_hrm";
  const canRegen = !entry.generating && (st === "new" || st === "rejected" || st === "pending_hr");

  els.drawerBody.innerHTML = `
    <div class="questions-split">
      <div class="split-col split-col--ai">
        <h3>Gợi ý từ AI</h3>
        <p class="col-desc">Bấm + để thêm vào danh sách chính thức bên phải.</p>
        <div class="q-cards" id="aiCards"></div>
      </div>
      <div class="split-col split-col--official">
        <h3>Câu hỏi chính thức</h3>
        <p class="col-desc">HR chỉnh sửa trực tiếp — dùng khi duyệt &amp; gửi HRM.</p>
        <div class="official-editor" id="officialEditor"></div>
        <button type="button" class="btn btn-ghost btn-sm" id="btnAddOfficial" style="margin-top:0.5rem">+ Thêm câu trống</button>
      </div>
    </div>

    <div class="drawer-section">
      <h4>HR duyệt &amp; workflow</h4>
      <label>Ghi chú HR (tùy chọn)
        <textarea id="hrComment" rows="2"></textarea>
      </label>
      <div class="drawer-actions" style="margin-top:0.75rem">
        <button type="button" class="btn btn-secondary" id="btnRegen" ${canRegen ? "" : "disabled"}>Tạo lại AI</button>
        <button type="button" class="btn btn-success" id="btnApprove" ${canApprove ? "" : "disabled"}>Duyệt câu hỏi</button>
        <button type="button" class="btn btn-danger" id="btnReject" ${st === "pending_hr" || st === "approved" ? "" : "disabled"}>Từ chối</button>
        <button type="button" class="btn btn-primary" id="btnSendHrm" ${canSend ? "" : "disabled"}>Gửi HRM</button>
        <button type="button" class="btn btn-secondary" id="btnReport" ${canFinalize ? "" : "disabled"}>Hoàn tất &amp; tải JSON</button>
      </div>
      <pre id="hrmOut" class="hint hidden" style="white-space:pre-wrap;max-height:120px;overflow:auto;margin-top:0.75rem"></pre>
      <ul class="timeline" id="timeline"></ul>
    </div>

    <details class="drawer-section" style="border-top:none;padding-top:0">
      <summary class="hint" style="cursor:pointer;font-weight:600">Dữ liệu gốc (Sheet / Excel)</summary>
      <div class="raw-fields" style="margin-top:0.75rem"></div>
    </details>
  `;

  const aiMount = els.drawerBody.querySelector("#aiCards");
  if (entry.generating) {
    aiMount.innerHTML = '<p class="empty-cards">Đang tạo câu hỏi AI...</p>';
  } else if (!aiList.length) {
    aiMount.innerHTML = '<p class="empty-cards">Chưa có gợi ý — đợi tự động tạo hoặc bấm «Tạo lại AI».</p>';
  } else {
    aiList.forEach((text, idx) => {
      const card = document.createElement("article");
      card.className = "q-card q-card--ai";
      card.innerHTML = `
        <p class="q-card-text"></p>
        <div class="q-card-actions">
          <button type="button" class="btn-add-q" title="Thêm vào danh sách chính thức">+</button>
        </div>
      `;
      card.querySelector(".q-card-text").textContent = text;
      card.querySelector(".btn-add-q").addEventListener("click", () => {
        addAiToOfficial(entry, text);
        renderDrawer();
        renderTable();
      });
      aiMount.appendChild(card);
    });
  }

  const offMount = els.drawerBody.querySelector("#officialEditor");
  const renderOfficialFields = () => {
    offMount.innerHTML = "";
    const qs = entry.questions?.length ? entry.questions : [""];
    qs.forEach((text, idx) => {
      const wrap = document.createElement("div");
      wrap.className = "official-item";
      wrap.innerHTML = `
        <div style="display:flex;gap:0.5rem;align-items:flex-start">
          <textarea data-off="${idx}" placeholder="Câu hỏi ${idx + 1}..."></textarea>
          <button type="button" class="btn-remove-q" data-rm="${idx}" title="Xóa">×</button>
        </div>
      `;
      wrap.querySelector("textarea").value = text;
      wrap.querySelector("textarea").addEventListener("change", (e) => {
        entry.questions[idx] = e.target.value.trim();
        syncOfficialQuestions(entry);
        saveState(state);
        renderTable();
      });
      wrap.querySelector("[data-rm]").addEventListener("click", () => {
        entry.questions.splice(idx, 1);
        syncOfficialQuestions(entry);
        saveState(state);
        renderOfficialFields();
        renderTable();
      });
      offMount.appendChild(wrap);
    });
  };
  renderOfficialFields();

  els.drawerBody.querySelector("#btnAddOfficial").addEventListener("click", () => {
    if (!entry.questions) entry.questions = [];
    entry.questions.push("");
    syncOfficialQuestions(entry);
    saveState(state);
    renderOfficialFields();
  });

  const hrTa = els.drawerBody.querySelector("#hrComment");
  hrTa.value = entry.hrComment || "";
  hrTa.addEventListener("change", () => {
    entry.hrComment = hrTa.value;
    saveState(state);
  });

  const hrmOut = els.drawerBody.querySelector("#hrmOut");
  if (entry.hrmResponse) {
    hrmOut.classList.remove("hidden");
    hrmOut.textContent =
      typeof entry.hrmResponse === "string"
        ? entry.hrmResponse
        : JSON.stringify(entry.hrmResponse, null, 2);
  }

  const rawBox = els.drawerBody.querySelector(".raw-fields");
  for (const [k, v] of Object.entries(c.raw)) {
    const d = document.createElement("div");
    d.innerHTML = `<strong></strong><span></span>`;
    d.querySelector("strong").textContent = k;
    d.querySelector("span").textContent = String(v ?? "");
    rawBox.appendChild(d);
  }

  const tl = els.drawerBody.querySelector("#timeline");
  (entry.log || []).slice(-10).reverse().forEach((l) => {
    const li = document.createElement("li");
    li.textContent = `${l.t} — ${l.message}`;
    tl.appendChild(li);
  });

  els.drawerBody.querySelector("#btnRegen").addEventListener("click", async () => {
    const btn = els.drawerBody.querySelector("#btnRegen");
    btn.disabled = true;
    btn.textContent = "Đang tạo...";
    try {
      await generateForCandidate(c);
    } catch (e) {
      alert(String(e.message || e));
    }
    renderDrawer();
    renderStats();
    renderTable();
  });

  els.drawerBody.querySelector("#btnApprove").addEventListener("click", () => {
    entry.hrComment = hrTa.value;
    if (approveCandidate(c, entry)) {
      renderStats();
      renderTable();
      renderDrawer();
    }
  });

  els.drawerBody.querySelector("#btnReject").addEventListener("click", () => {
    rejectCandidate(entry, c.id);
    renderStats();
    renderTable();
    renderDrawer();
  });

  els.drawerBody.querySelector("#btnSendHrm").addEventListener("click", async () => {
    const ok = await sendToHrm(c, entry);
    if (ok) {
      renderStats();
      renderTable();
      renderDrawer();
    }
  });

  els.drawerBody.querySelector("#btnReport").addEventListener("click", () => {
    const report = {
      generatedAt: new Date().toISOString(),
      candidate: c,
      workflow: entry,
    };
    entry.status = "done";
    entry.reportAt = new Date().toISOString();
    pushLog(entry, "Đã hoàn tất và xuất báo cáo.");
    saveState(state);
    downloadBlob(
      `bao_cao_${norm(c.displayName).replace(/\s+/g, "_")}.json`,
      new Blob([JSON.stringify(report, null, 2)], { type: "application/json" })
    );
    renderStats();
    renderTable();
    renderDrawer();
  });
}

els.fileInput.addEventListener("change", async (ev) => {
  const file = ev.target.files?.[0];
  if (!file) return;
  const buf = await file.arrayBuffer();
  try {
    candidates = parseWorkbook(buf);
    lastSheetError = "";
    if (els.sheetHint) {
      els.sheetHint.textContent = candidates.length
        ? `Đang dùng file Excel (${candidates.length} dòng).`
        : "";
    }
    if (!candidates.length) {
      alert("Không đọc được dòng dữ liệu nào. Kiểm tra dòng tiêu đề cột trong Excel.");
    }
    renderIndustryFilter();
    renderStats();
    renderTable();
    scheduleAutoGenerate();
  } catch (e) {
    alert("Lỗi đọc Excel: " + (e.message || e));
  }
  ev.target.value = "";
});

async function refreshCandidatesFromGoogleSheet() {
  sheetFetchLoading = true;
  lastSheetError = "";
  if (els.sheetHint) els.sheetHint.textContent = "Đang tải...";
  renderTable();
  const btn = els.btnReloadSheet;
  if (btn) {
    btn.disabled = true;
    btn.textContent = "Đang tải...";
  }
  try {
    const { candidates: list, usedProxy } = await loadCandidatesFromGoogleSheet(
      DEFAULT_GOOGLE_SHEET_PUB_URL
    );
    candidates = list;
    if (!candidates.length) {
      lastSheetError = "Không có dòng dữ liệu (kiểm tra dòng tiêu đề cột).";
    }
    if (els.sheetHint) {
      const proxyNote = usedProxy ? " (qua proxy vì trình duyệt chặn tải trực tiếp)." : "";
      els.sheetHint.textContent = lastSheetError
        ? lastSheetError
        : `Đã tải ${candidates.length} ứng viên từ Google Sheet.${proxyNote}`;
    }
  } catch (e) {
    lastSheetError = String(e.message || e);
    candidates = [];
    if (els.sheetHint) {
      els.sheetHint.textContent = `Lỗi: ${lastSheetError}`;
    }
  }
  sheetFetchLoading = false;
  if (btn) {
    btn.disabled = false;
    btn.textContent = "Tải lại từ Google Sheet";
  }
  renderIndustryFilter();
  renderStats();
  renderTable();
  scheduleAutoGenerate();
}

if (els.btnReloadSheet) {
  els.btnReloadSheet.addEventListener("click", () => {
    refreshCandidatesFromGoogleSheet();
  });
}

["change", "input"].forEach((evt) => {
  els.filterIndustry.addEventListener(evt, () => {
    renderTable();
  });
  els.filterStatus.addEventListener(evt, () => {
    renderTable();
  });
  els.search.addEventListener(evt, () => {
    renderTable();
  });
});

["change", "blur"].forEach((evt) => {
  els.hrmUrl?.addEventListener(evt, persistSettings);
  els.geminiApiKey?.addEventListener(evt, persistSettings);
  els.geminiModel?.addEventListener(evt, persistSettings);
});

if (els.btnCloseDrawer) {
  els.btnCloseDrawer.addEventListener("click", closeDrawer);
}
if (els.drawerBackdrop) {
  els.drawerBackdrop.addEventListener("click", closeDrawer);
}
if (els.btnExportCsv) {
  els.btnExportCsv.addEventListener("click", exportAllCsv);
}

els.btnExportState.addEventListener("click", () => {
  const blob = new Blob([JSON.stringify({ ...state, exportedAt: new Date().toISOString() }, null, 2)], {
    type: "application/json",
  });
  downloadBlob("tuyen_dung_trang_thai.json", blob);
});

els.btnImportState.addEventListener("click", () => els.stateImport.click());

els.stateImport.addEventListener("change", async (ev) => {
  const f = ev.target.files?.[0];
  if (!f) return;
  try {
    const text = await f.text();
    const parsed = JSON.parse(text);
    if (!parsed.byId) throw new Error("File không hợp lệ.");
    state = parsed;
    saveState(state);
    hydrateSettings();
    renderStats();
    renderTable();
  } catch (e) {
    alert("Nhập trạng thái lỗi: " + (e.message || e));
  }
  ev.target.value = "";
});

hydrateSettings();
renderIndustryFilter();
renderStats();
renderTable();
refreshCandidatesFromGoogleSheet();
