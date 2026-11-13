/* =========================================================
   BIỂU MẪU A4 – BẢNG ĐIỀU KHIỂN SOẠN HỒ SƠ
   - Dữ liệu hồ sơ nằm trong `state`; mọi phần tử data-field trên trang được vẽ lại từ đó.
   - Người ký cố định, quy tắc tự động, mẫu câu chữ lưu ở localStorage (a4Settings).
   - Trang cha (index.html) dùng window.A4Editor để lấy thông tin khi in / lưu lịch sử.
   ========================================================= */

"use strict";

/* =========================
   CẤU HÌNH
   ========================= */

const STORAGE_KEYS = {
  selectedPrinter: "selectedPrinter",
  printerData: "printerData",
  settings: "a4Settings",
  ui: "a4EditorUi",
};

// Các trường lấy từ file Excel, có thể lưu ngược vào danh sách
const EXCEL_FIELDS = [
  "department",
  "shortDepartment",
  "requester",
  "officer",
  "representative",
  "representativeStaff",
  "daysW4",
  "monthW4",
];

// Giá trị không có trong Excel. Sửa trên bảng điều khiển rồi bấm "Lưu làm mặc định".
const FACTORY_DEFAULTS = {
  planOffice: "Phan Nhật Mạnh",
  planOfficeHidden: true,
  techBefore: "Nguyễn Văn Toàn",
  // Đơn vị là Phòng BĐKT: người ký cột "Đơn vị sử dụng" ở BB trước sửa chữa (thay tên trong Excel)
  techUnitBefore: "Nguyễn Đức Dũng",
  techAfter: "Nguyễn Văn Toàn",
  functional1: "Lê Tuấn Cường",
  functional1Position: "NVKH",
  functional2: "Vũ Hải Hà",
  functional2Position: "NVKT",
  receiver: "Đào Thị Thanh",
  accountant: "Dương Thị Lan Thu",
  substituteGiver: "Đặng Thị Ngọc Ngà",
  docSymbol: "TTr-",
};

// Danh sách tên, mỗi dòng một người
const FACTORY_RULES = {
  // Lãnh đạo đơn vị vẫn ký nhưng không in tên
  hiddenOfficerNames: "Nguyễn Thị Thu",
  // Người đề nghị thuộc Phòng Kế hoạch: tờ trình không cần thêm cột ký của đơn vị
  requestersWithoutUnitColumn: "Lê Tuấn Cường",
};

// Tuỳ chọn được ghi nhớ ngay khi bật/tắt
const FACTORY_OPTIONS = {
  linkBroken: true,
  linkRecovered: true,
};

// Tuỳ chọn tự tính theo quy tắc, trừ khi người dùng tự bật/tắt cho hồ sơ đang soạn
const AUTO_TOGGLES = {
  hideOfficer: (s) => nameInList(s.officer, s.hiddenOfficerNames),
  showUnitColumn: (s) =>
    !nameInList(s.requester, s.requestersWithoutUnitColumn),
};

const DATE_ROWS = [
  { label: "Tờ trình", day: "toTrinhDay", month: "toTrinhMonth" },
  { label: "Phiếu báo hỏng", day: "baoHongDay", month: "baoHongMonth" },
  { label: "BB trước sửa chữa", day: "truocDay", month: "truocMonth" },
  { label: "BB sau sửa chữa", day: "sauDay", month: "sauMonth" },
  { label: "BB thu hồi", day: "daysW4", month: "monthW4" },
  { label: "BB căn cứ (thu hồi)", day: "canCuDay", month: "canCuMonth" },
];

// Cột của bảng thu hồi tự điền từ bảng vật tư tờ trình
const LINKED_COLUMNS = ["name", "unit", "qty"];

// Bấm chỗ hiển thị trên trang thì nhảy tới ô nhập tương ứng
const FIELD_INPUT_ALIAS = {
  giver: () => "giverOverride",
  unitSignBefore: () =>
    usesTechUnitSigner() ? "techUnitBefore" : "representative",
};

// Nhận ra đơn vị là Phòng Bảo đảm kỹ thuật (viết đủ, viết tắt BĐKT, có/không dấu)
const TECH_DEPARTMENT_PATTERN = /\b(bao dam|dam bao) ky thuat\b|\bbdkt\b/;

// Tăng khi cần sửa giá trị cũ đã lưu trong trình duyệt (xem migrateSettings)
const SETTINGS_VERSION = 2;

const PANEL_GROUPS = [
  {
    id: "unit",
    title: "Thông tin đơn vị",
    open: true,
    hint: "Lấy từ file Excel. Sửa ở đây chỉ áp dụng cho hồ sơ này; bấm “Lưu vào danh sách” để giữ cho các lần sau.",
    fields: [
      { key: "department", label: "Phòng ban", required: true },
      { key: "shortDepartment", label: "Tên đơn vị ở chỗ ký", required: true },
      {
        key: "requester",
        label: "Người đề nghị",
        people: true,
        required: true,
      },
      {
        key: "officer",
        label: "Lãnh đạo đơn vị ký",
        people: true,
        required: true,
        toggle: { key: "hideOfficer", label: "Ẩn tên khi in" },
      },
      {
        key: "representative",
        label: "Đại diện đơn vị sử dụng (ký biên bản)",
        people: true,
        required: true,
      },
      {
        key: "representativeStaff",
        label: "Nhân viên đại diện đơn vị",
        people: true,
        required: true,
      },
    ],
    actions: [
      { action: "save-to-list", label: "Lưu vào danh sách", primary: true },
    ],
  },
  {
    id: "equipment",
    title: "Thiết bị sửa chữa",
    open: true,
    hint: "Nhập một lần, tự hiện ở cả hai biên bản kiểm tra và phiếu báo hỏng.",
    fields: [
      {
        key: "equipmentName",
        label: "Tên vật tư thiết bị",
        multiline: true,
        required: true,
      },
      { key: "serialNumber", label: "Số S/N (nếu có)" },
    ],
  },
  {
    id: "dates",
    title: "Ngày tháng & số văn bản",
    open: true,
    build: buildDatesGroup,
  },
  {
    id: "signers",
    title: "Người ký cố định",
    hint: "Sửa ở đây áp dụng cho hồ sơ này. Bấm “Lưu làm mặc định” để dùng cho các hồ sơ sau.",
    fields: [
      {
        key: "planOffice",
        label: "Phòng Kế hoạch",
        people: true,
        required: true,
        toggle: { key: "planOfficeHidden", label: "Ẩn tên khi in" },
      },
      {
        key: "techBefore",
        label: "Phòng BĐKT – BB trước sửa chữa",
        people: true,
        required: true,
      },
      {
        key: "techUnitBefore",
        label: "Đơn vị sử dụng – BB trước sửa chữa, khi đơn vị là Phòng BĐKT",
        people: true,
        placeholder: "Để trống = dùng tên trong Excel",
      },
      {
        key: "techAfter",
        label: "Phòng BĐKT – BB sau sửa chữa",
        people: true,
        required: true,
      },
      {
        key: "functional1",
        label: "Đại diện đơn vị chức năng 1",
        people: true,
        required: true,
        position: "functional1Position",
      },
      {
        key: "functional2",
        label: "Đại diện đơn vị chức năng 2",
        people: true,
        required: true,
        position: "functional2Position",
      },
      {
        key: "receiver",
        label: "Bên nhận – Kho trung tâm",
        people: true,
        required: true,
      },
      {
        key: "accountant",
        label: "Kế toán tài sản",
        people: true,
        required: true,
      },
      {
        key: "substituteGiver",
        label: "Bên giao thay khi người đề nghị là kế toán",
        people: true,
      },
      {
        key: "giverOverride",
        label: "Bên giao BB thu hồi (để trống = tự động)",
        people: true,
      },
    ],
    actions: [
      { action: "save-defaults", label: "Lưu làm mặc định", primary: true },
      { action: "reset-defaults", label: "Khôi phục gốc" },
    ],
  },
  {
    id: "options",
    title: "Tuỳ chọn trình bày",
    toggles: [
      { key: "showUnitColumn", label: "Tờ trình có cột ký của đơn vị" },
      { key: "linkBroken", label: "Phiếu báo hỏng: tự điền tên thiết bị" },
      { key: "linkRecovered", label: "BB thu hồi: tự điền vật tư từ tờ trình" },
    ],
    actions: [
      { action: "relink-recovered", label: "Tự điền lại bảng thu hồi" },
    ],
  },
  {
    id: "rules",
    title: "Quy tắc tự động",
    hint: "Mỗi dòng một tên. Áp dụng khi mở hồ sơ; vẫn bật/tắt tay được cho từng hồ sơ.",
    fields: [
      {
        key: "hiddenOfficerNames",
        label: "Lãnh đạo đơn vị luôn ẩn tên khi in",
        textarea: true,
      },
      {
        key: "requestersWithoutUnitColumn",
        label: "Người đề nghị không cần cột ký đơn vị trên tờ trình",
        textarea: true,
      },
    ],
    actions: [
      { action: "save-defaults", label: "Lưu làm mặc định", primary: true },
    ],
  },
  { id: "pages", title: "Văn bản cần in", open: true, build: buildPagesGroup },
];

const PANEL_FIELDS = PANEL_GROUPS.flatMap((group) => group.fields || []);
const REQUIRED_FIELDS = new Set([
  ...PANEL_FIELDS.filter((f) => f.required).map((f) => f.key),
  "unitSignBefore",
]);
const PEOPLE_FIELDS = PANEL_FIELDS.filter((f) => f.people).map((f) => f.key);

const FIELD_LABELS = (() => {
  const labels = {
    giver: "Bên giao",
    unitSignBefore: "Đơn vị sử dụng ký – BB trước sửa chữa",
    year: "Năm",
    docNumber: "Số",
    docSymbol: "Ký hiệu",
    functional1Position: "Chức vụ",
    functional2Position: "Chức vụ",
  };
  DATE_ROWS.forEach(({ label, day, month }) => {
    labels[day] = `Ngày – ${label}`;
    labels[month] = `Tháng – ${label}`;
  });
  PANEL_FIELDS.forEach((field) => {
    labels[field.key] = field.label;
  });
  return labels;
})();

/* =========================
   TIỆN ÍCH
   ========================= */

function onReady(callback) {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", callback);
  } else {
    callback();
  }
}

function loadJson(key, fallback = null) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (error) {
    console.error(`Không đọc được dữ liệu localStorage: ${key}`, error);
    return fallback;
  }
}

function saveJson(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch (error) {
    console.error(`Không lưu được dữ liệu localStorage: ${key}`, error);
    return false;
  }
}

const toText = (value) =>
  value === undefined || value === null ? "" : String(value);

const normalizeName = (value) =>
  toText(value).normalize("NFC").trim().toLowerCase();

const isSameName = (a, b) => normalizeName(a) === normalizeName(b);

const splitLines = (text) =>
  toText(text)
    .split(/[\n,;]/)
    .map((item) => item.trim())
    .filter(Boolean);

// Bỏ dấu tiếng Việt để so khớp: "Bảo đảm kỹ thuật" -> "bao dam ky thuat"
const plainText = (value) =>
  normalizeName(value)
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d");

function nameInList(name, listText) {
  if (!normalizeName(name)) return false;
  return splitLines(listText).some((item) => isSameName(item, name));
}

// Tạo phần tử DOM gọn: h("button", { class: "btn" }, "Nội dung")
function h(tag, attrs = {}, ...children) {
  const element = document.createElement(tag);

  Object.entries(attrs).forEach(([name, value]) => {
    if (value === null || value === undefined || value === false) return;
    if (name === "class") element.className = value;
    else if (name === "text") element.textContent = value;
    else element.setAttribute(name, value === true ? "" : String(value));
  });

  children.flat(Infinity).forEach((child) => {
    if (child === null || child === undefined || child === false) return;
    element.append(child instanceof Node ? child : String(child));
  });

  return element;
}

// Nghị định 30/2020/NĐ-CP: ngày dưới 10 và tháng 1, 2 phải thêm số 0 phía trước
function formatDatePart(value, part) {
  const text = toText(value).trim();
  if (!/^\d{1,2}$/.test(text)) return text;

  const number = Number(text);
  const needsZero = part === "day" ? number < 10 : number < 3;
  return needsZero ? `0${number}` : String(number);
}

function todayParts() {
  const today = new Date();
  return {
    day: String(today.getDate()),
    month: String(today.getMonth() + 1),
    year: String(today.getFullYear()),
  };
}

function flash(element) {
  element.classList.remove("flash");
  void element.offsetWidth; // chạy lại hiệu ứng khi bấm liên tiếp
  element.classList.add("flash");
  setTimeout(() => element.classList.remove("flash"), 1300);
}

function showToast(message, type = "info") {
  let container = document.querySelector(".toast-container");
  if (!container) {
    container = h("div", {
      class: "toast-container",
      "data-editor-ui": "",
      role: "status",
      "aria-live": "polite",
    });
    document.body.append(container);
  }

  const toast = h("div", { class: `toast toast-${type}`, text: message });
  container.append(toast);
  setTimeout(() => toast.classList.add("is-leaving"), 3500);
  setTimeout(() => toast.remove(), 3900);
}

/* =========================
   TRẠNG THÁI
   ========================= */

let settings = null;
let state = {};
let sourceRow = {};
let excludedDocs = new Set();
let dirty = false;
let panel = null;
let templateBar = null;
let activeTemplateBlock = null;
let currentIssues = [];

const manualToggles = new Set();
const factoryTemplates = new Map();
const rowTemplates = new Map();
const renderedText = new WeakMap();

const isEmbedded = window.parent && window.parent !== window;

// Sửa giá trị đã "Lưu làm mặc định" ở bản cũ mà nay không còn đúng
function migrateSettings(saved) {
  const version = Number(saved.version) || 1;
  const defaults = { ...saved.defaults };

  // v2: Phòng BĐKT ở BB trước sửa chữa đổi từ Nguyễn Đức Dũng sang Nguyễn Văn Toàn
  // (Nguyễn Đức Dũng chuyển sang ký cột Đơn vị sử dụng khi đơn vị là Phòng BĐKT)
  if (version < 2 && isSameName(defaults.techBefore, "Nguyễn Đức Dũng")) {
    delete defaults.techBefore;
  }

  return { ...saved, defaults };
}

function loadSettings() {
  const saved = migrateSettings(loadJson(STORAGE_KEYS.settings, {}) || {});
  return {
    version: SETTINGS_VERSION,
    defaults: { ...FACTORY_DEFAULTS, ...saved.defaults },
    rules: { ...FACTORY_RULES, ...saved.rules },
    options: { ...FACTORY_OPTIONS, ...saved.options },
    excludedDocs: Array.isArray(saved.excludedDocs) ? saved.excludedDocs : [],
    templates: { ...saved.templates },
  };
}

function saveSettings() {
  if (!saveJson(STORAGE_KEYS.settings, settings)) {
    showToast("Không lưu được cài đặt vào trình duyệt.", "error");
  }
}

function createInitialState(row) {
  const blankDates = Object.fromEntries(
    DATE_ROWS.flatMap(({ day, month }) => [
      [day, ""],
      [month, ""],
    ]),
  );
  const excelValues = Object.fromEntries(
    EXCEL_FIELDS.map((key) => [key, toText(row[key])]),
  );

  return {
    ...blankDates,
    ...excelValues,
    ...settings.defaults,
    ...settings.rules,
    ...settings.options,
    hideOfficer: false,
    showUnitColumn: true,
    giverOverride: "",
    equipmentName: "",
    serialNumber: "",
    docNumber: "",
    year: todayParts().year,
  };
}

// Người đề nghị là kế toán tài sản thì không thể vừa giao vừa ký kế toán → dùng người thay
function getAutoGiver() {
  return isSameName(state.requester, state.accountant)
    ? state.substituteGiver
    : state.requester;
}

const isTechDepartment = (department) =>
  TECH_DEPARTMENT_PATTERN.test(plainText(department));

// Đơn vị là Phòng BĐKT thì cột "Đơn vị sử dụng" ở BB trước sửa chữa dùng người ký riêng
function usesTechUnitSigner() {
  return (
    isTechDepartment(state.department) &&
    Boolean(toText(state.techUnitBefore).trim())
  );
}

// Giá trị cuối cùng để hiển thị (đã tính các trường tự động)
function buildView() {
  const view = {
    ...state,
    giver: toText(state.giverOverride).trim() || getAutoGiver(),
    unitSignBefore: usesTechUnitSigner()
      ? state.techUnitBefore
      : state.representative,
  };

  Object.entries(AUTO_TOGGLES).forEach(([key, compute]) => {
    view[key] = manualToggles.has(key) ? state[key] : compute(state);
  });

  return view;
}

function markDirty() {
  dirty = true;
}

function setField(key, value) {
  state[key] = value;

  if (key in AUTO_TOGGLES) manualToggles.add(key);
  if (key in FACTORY_OPTIONS) {
    settings.options[key] = value;
    saveSettings();
  }

  markDirty();
  render();
}

function setFields(patch) {
  Object.assign(state, patch);
  markDirty();
  refreshPanelInputs();
  render();
}

/* =========================
   VẼ TRANG
   ========================= */

function getPages() {
  return Array.from(document.querySelectorAll(".page"));
}

function render() {
  const view = buildView();
  renderDocument(view);
  renderPages();
  syncLinkedTables(view);
  renderPanelState(view);
  renderIssues(view);
}

function renderDocument(view) {
  document.querySelectorAll(".page [data-field]").forEach((element) => {
    renderField(element, view[element.dataset.field]);
  });

  document.querySelectorAll(".page [data-hidden-by]").forEach((element) => {
    element.classList.toggle(
      "is-name-hidden",
      Boolean(view[element.dataset.hiddenBy]),
    );
  });

  document.querySelectorAll(".page [data-shown-by]").forEach((element) => {
    element.classList.toggle("is-hidden", !view[element.dataset.shownBy]);
  });
}

function renderField(element, rawValue) {
  const { format, blank } = element.dataset;
  let text = toText(rawValue).trim();
  if (format === "day" || format === "month")
    text = formatDatePart(text, format);

  element.classList.toggle("is-empty", !text && blank === undefined);

  const content = text || blank || "";
  // Không ghi lại DOM nếu không đổi, tránh làm nhảy con trỏ khi đang gõ gần đó
  if (renderedText.get(element) === content) return;
  renderedText.set(element, content);

  if (format === "trung-tam") {
    const match = content.match(/^(.*?trung tâm)\s+(.+)$/i);
    if (match) {
      element.replaceChildren(match[1], document.createElement("br"), match[2]);
      return;
    }
  }

  element.textContent = content;
}

function renderPages() {
  let firstFound = false;

  getPages().forEach((page) => {
    const excluded = excludedDocs.has(page.dataset.doc);
    page.classList.toggle("is-excluded", excluded);
    page.classList.toggle("is-first-printed", !excluded && !firstFound);
    if (!excluded) firstFound = true;
  });

  document.querySelectorAll("[data-doc-toggle]").forEach((input) => {
    input.checked = !excludedDocs.has(input.dataset.docToggle);
  });
}

// Gắn nhãn, chặn gõ trực tiếp vào ô tự điền (bấm vào sẽ mở ô nhập bên trái)
function prepareFieldElements(root = document) {
  root.querySelectorAll(".page [data-field]").forEach((element) => {
    const key = element.dataset.field;
    const label = FIELD_LABELS[key] || key;

    element.contentEditable = "false";
    element.dataset.label = label;
    element.title = `Bấm để sửa: ${label}`;
    element.classList.toggle("is-optional", !REQUIRED_FIELDS.has(key));
  });
}

/* =========================
   KIỂM TRA TRƯỚC KHI IN
   ========================= */

function collectIssues(view) {
  const issues = [];
  const isPrinted = (element) =>
    !element.closest(".is-excluded, .is-hidden, .is-name-hidden");

  const missingInputs = new Set();
  REQUIRED_FIELDS.forEach((key) => {
    if (toText(view[key]).trim()) return;

    const used = Array.from(
      document.querySelectorAll(`.page [data-field="${key}"]`),
    ).some(isPrinted);
    // Nhiều chỗ trên trang cùng lấy từ một ô nhập thì chỉ nhắc một lần
    const inputKey = FIELD_INPUT_ALIAS[key]?.() || key;
    if (!used || missingInputs.has(inputKey)) return;

    missingInputs.add(inputKey);
    issues.push({ label: `Chưa nhập: ${FIELD_LABELS[key]}`, key });
  });

  getTables().forEach((table) => {
    if (!isPrinted(table)) return;

    const title = table.closest(".page").dataset.title;
    getDataRows(table).forEach((row, index) => {
      const nameCell = row.querySelector('[data-col="name"]');
      if (nameCell && !nameCell.textContent.trim()) {
        issues.push({
          label: `${title}: dòng ${index + 1} chưa có tên`,
          cell: nameCell,
        });
      }
    });
  });

  if (getPages().every((page) => excludedDocs.has(page.dataset.doc))) {
    issues.push({ label: "Chưa chọn văn bản nào để in" });
  }

  return issues;
}

function renderIssues(view) {
  currentIssues = collectIssues(view);
  const box = panel.querySelector("[data-issues]");

  if (!currentIssues.length) {
    box.className = "panel-issues is-ok";
    box.replaceChildren(
      h("strong", { text: "✓ Đã đủ thông tin, sẵn sàng in" }),
    );
    return;
  }

  box.className = "panel-issues has-issues";
  box.replaceChildren(
    h("strong", { text: `Cần kiểm tra (${currentIssues.length})` }),
    h(
      "ul",
      {},
      currentIssues.map((issue, index) =>
        h(
          "li",
          {},
          h(
            "button",
            {
              type: "button",
              class: "issue-link",
              "data-action": "focus-issue",
              "data-index": index,
            },
            issue.label,
          ),
        ),
      ),
    ),
  );
}

function focusIssue(issue) {
  if (!issue) return;

  if (issue.key) {
    focusPanelField(issue.key);
    return;
  }

  if (issue.cell) {
    issue.cell.scrollIntoView({ behavior: "smooth", block: "center" });
    issue.cell.focus({ preventScroll: true });
    flash(issue.cell);
  }
}

/* =========================
   BẢNG: THÊM / XOÁ DÒNG, TỰ ĐIỀN
   ========================= */

function getTables() {
  return Array.from(document.querySelectorAll(".page table[data-table]"));
}

function getTable(name) {
  return document.querySelector(`.page table[data-table="${name}"]`);
}

function getDataRows(table) {
  return table?.tBodies[0] ? Array.from(table.tBodies[0].rows) : [];
}

const rowText = (row) =>
  Array.from(row.cells)
    .slice(1)
    .map((cell) => cell.textContent.trim())
    .join("|");

function initTables() {
  getTables().forEach((table) => {
    const [firstRow] = getDataRows(table);
    if (!firstRow) return;

    // Dòng đầu trong HTML là mẫu cho các dòng thêm sau (giữ chữ mặc định như "Hỏng...")
    rowTemplates.set(table.dataset.table, firstRow.cloneNode(true));
    getDataRows(table).forEach(decorateRow);
    renumber(table);

    table.after(
      h(
        "button",
        {
          type: "button",
          class: "table-add-row",
          "data-editor-ui": "",
          "data-action": "row-append",
          "data-table-id": table.dataset.table,
        },
        "+ Thêm dòng",
      ),
    );
  });
}

function decorateRow(row) {
  const cell = row.cells[0];
  if (!cell) return;

  cell.classList.add("col-stt");
  cell.removeAttribute("contenteditable");
  cell.replaceChildren(
    h("span", { class: "row-no" }),
    h(
      "span",
      { class: "row-tools", "data-editor-ui": "", contenteditable: "false" },
      h(
        "button",
        {
          type: "button",
          tabindex: "-1",
          "data-action": "row-add",
          title: "Thêm dòng bên dưới (Ctrl+Enter)",
          "aria-label": "Thêm dòng",
        },
        "+",
      ),
      h(
        "button",
        {
          type: "button",
          tabindex: "-1",
          "data-action": "row-delete",
          title: "Xoá dòng này",
          "aria-label": "Xoá dòng",
        },
        "×",
      ),
    ),
  );
}

function createRow(table) {
  const row = rowTemplates.get(table.dataset.table).cloneNode(true);
  decorateRow(row);
  return row;
}

function renumber(table) {
  getDataRows(table).forEach((row, index) => {
    const number = row.querySelector(".row-no");
    if (number) number.textContent = String(index + 1);
  });
}

function addRowAfter(table, refRow, { manual = true, focus = true } = {}) {
  const row = createRow(table);
  // Dòng người dùng tự thêm không bị tự xoá khi đồng bộ bảng thu hồi
  if (manual) row.dataset.manualRow = "1";

  if (refRow) refRow.after(row);
  else table.tBodies[0].append(row);

  renumber(table);
  if (focus) row.querySelector('[contenteditable="true"]')?.focus();
  return row;
}

function removeRow(row) {
  const table = row.closest("table");
  const rows = getDataRows(table);
  const template = rowTemplates.get(table.dataset.table);
  const hasContent = rowText(row) !== rowText(template);

  if (hasContent && !confirm(`Xoá dòng ${rows.indexOf(row) + 1} của bảng?`)) {
    return;
  }

  if (rows.length === 1) {
    row.replaceWith(createRow(table));
    showToast("Bảng cần ít nhất một dòng nên đã làm trống dòng này.");
  } else {
    row.remove();
  }

  renumber(table);
  onTablesChanged();
}

function onTablesChanged() {
  markDirty();
  const view = buildView();
  syncLinkedTables(view);
  renderIssues(view);
}

function handleCellInput(cell) {
  const table = cell.closest("table[data-table]");

  // Ô đã gõ tay thì không bị tự điền đè nữa
  if (["recovered", "broken"].includes(table.dataset.table)) {
    cell.dataset.touched = "1";
    unmarkLinked(cell);
  }

  onTablesChanged();
}

function syncLinkedTables(view) {
  const supplies = getTable("supplies");
  const recovered = getTable("recovered");

  if (supplies && recovered) {
    if (view.linkRecovered) mirrorRows(supplies, recovered);
    else recovered.querySelectorAll("td.is-linked").forEach(unmarkLinked);
  }

  const brokenCell = getDataRows(getTable("broken"))[0]?.querySelector(
    '[data-col="name"]',
  );
  if (brokenCell) {
    if (view.linkBroken)
      linkCell(brokenCell, toText(view.equipmentName).trim());
    else unmarkLinked(brokenCell);
  }
}

// Bảng thu hồi = vật tư cũ tháo ra của các vật tư mua mới trong tờ trình
function mirrorRows(source, target) {
  const sourceRows = getDataRows(source);
  let targetRows = getDataRows(target);

  while (targetRows.length < sourceRows.length) {
    addRowAfter(target, targetRows.at(-1), { manual: false, focus: false });
    targetRows = getDataRows(target);
  }

  const isAutoRow = (row) =>
    !row.dataset.manualRow && !row.querySelector("[data-touched]");
  while (
    targetRows.length > Math.max(sourceRows.length, 1) &&
    isAutoRow(targetRows.at(-1))
  ) {
    targetRows.pop().remove();
  }
  renumber(target);

  sourceRows.forEach((row, index) => {
    LINKED_COLUMNS.forEach((column) => {
      const from = row.querySelector(`[data-col="${column}"]`);
      const to = targetRows[index]?.querySelector(`[data-col="${column}"]`);
      if (from && to) linkCell(to, from.textContent.trim());
    });
  });
}

function linkCell(cell, value) {
  if (cell.dataset.touched) {
    unmarkLinked(cell);
    return;
  }

  cell.classList.add("is-linked");
  cell.title = "Tự điền – gõ vào ô để sửa riêng";
  if (cell.textContent !== value) cell.textContent = value;
}

function unmarkLinked(cell) {
  cell.classList.remove("is-linked");
  cell.removeAttribute("title");
}

function relinkRecovered() {
  getTable("recovered")
    ?.querySelectorAll("[data-col][data-touched]")
    .forEach((cell) => delete cell.dataset.touched);

  setField("linkRecovered", true);
  showToast("Đã tự điền lại tên, ĐVT, số lượng từ bảng tờ trình.", "success");
}

/* =========================
   MẪU CÂU CHỮ (data-persist)
   ========================= */

const TEMPLATE_TAGS = new Set([
  "p",
  "div",
  "br",
  "b",
  "strong",
  "i",
  "em",
  "u",
]);
const FIELD_ATTRIBUTES = ["data-field", "data-format", "data-blank"];

// Chỉ giữ chữ, xuống dòng, định dạng cơ bản và chỗ tự điền (để trống)
function sanitizeTemplate(html) {
  const template = document.createElement("template");
  template.innerHTML = html;
  cleanTemplateNode(template.content);
  return template.innerHTML.replace(/\s+/g, " ").trim();
}

function cleanTemplateNode(parent) {
  Array.from(parent.childNodes).forEach((node) => {
    if (node.nodeType === Node.TEXT_NODE) return;
    if (node.nodeType !== Node.ELEMENT_NODE) {
      node.remove();
      return;
    }

    const tag = node.tagName.toLowerCase();

    if (tag === "span" && node.hasAttribute("data-field")) {
      Array.from(node.attributes).forEach(({ name }) => {
        if (!FIELD_ATTRIBUTES.includes(name)) node.removeAttribute(name);
      });
      node.textContent = "";
      return;
    }

    cleanTemplateNode(node);

    if (TEMPLATE_TAGS.has(tag)) {
      Array.from(node.attributes).forEach(({ name }) =>
        node.removeAttribute(name),
      );
    } else {
      node.replaceWith(...node.childNodes);
    }
  });
}

function initTemplates() {
  document.querySelectorAll(".page [data-persist]").forEach((block) => {
    const key = block.dataset.persist;
    factoryTemplates.set(key, sanitizeTemplate(block.innerHTML));

    const saved = settings.templates[key];
    if (typeof saved === "string") block.innerHTML = sanitizeTemplate(saved);
  });
}

function buildTemplateBar() {
  templateBar = h(
    "div",
    { class: "template-bar", "data-editor-ui": "", hidden: true },
    h("span", { class: "template-status" }),
    h(
      "button",
      {
        type: "button",
        class: "btn btn-small",
        "data-action": "template-save",
        title: "Các hồ sơ mở sau sẽ dùng nội dung này",
      },
      "Lưu làm mẫu",
    ),
    h(
      "button",
      {
        type: "button",
        class: "btn btn-small",
        "data-action": "template-reset",
        title: "Đưa đoạn này về nội dung ban đầu",
      },
      "Về mẫu gốc",
    ),
  );

  // Giữ con trỏ trong đoạn đang sửa khi bấm nút trên thanh
  templateBar.addEventListener("mousedown", (event) => event.preventDefault());
  document.body.append(templateBar);
}

function showTemplateBar(block) {
  activeTemplateBlock = block;
  templateBar.hidden = false;
  updateTemplateBar();
}

function hideTemplateBar() {
  activeTemplateBlock = null;
  if (templateBar) templateBar.hidden = true;
}

function updateTemplateBar() {
  if (!activeTemplateBlock) return;

  const key = activeTemplateBlock.dataset.persist;
  const current = sanitizeTemplate(activeTemplateBlock.innerHTML);
  const factory = factoryTemplates.get(key);
  const saved = settings.templates[key];
  const baseline = saved ?? factory;

  let status = saved === undefined ? "Mẫu gốc" : "Đang dùng mẫu đã lưu";
  if (current !== baseline) status = "Đã sửa – chỉ áp dụng hồ sơ này";

  templateBar.querySelector(".template-status").textContent = status;
  templateBar.querySelector('[data-action="template-save"]').disabled =
    current === baseline;
  templateBar.querySelector('[data-action="template-reset"]').disabled =
    current === factory && saved === undefined;

  positionTemplateBar();
}

function positionTemplateBar() {
  if (!activeTemplateBlock || templateBar.hidden) return;

  const rect = activeTemplateBlock.getBoundingClientRect();
  const top = rect.top + window.scrollY - templateBar.offsetHeight - 8;
  templateBar.style.top = `${Math.max(top, window.scrollY + 4)}px`;
  templateBar.style.left = `${rect.left + window.scrollX}px`;
}

function saveTemplate() {
  if (!activeTemplateBlock) return;

  const key = activeTemplateBlock.dataset.persist;
  settings.templates[key] = sanitizeTemplate(activeTemplateBlock.innerHTML);
  saveSettings();
  updateTemplateBar();
  showToast(
    "Đã lưu làm mẫu. Các hồ sơ mở sau sẽ dùng nội dung này.",
    "success",
  );
}

function resetTemplate() {
  const block = activeTemplateBlock;
  if (!block) return;

  const key = block.dataset.persist;
  block.innerHTML = factoryTemplates.get(key);
  delete settings.templates[key];
  saveSettings();

  prepareFieldElements(block);
  markDirty();
  render();
  updateTemplateBar();
  showToast("Đã đưa đoạn này về nội dung gốc.", "success");
}

/* =========================
   BẢNG ĐIỀU KHIỂN
   ========================= */

function buildPanel() {
  panel = h(
    "aside",
    {
      class: "editor-panel",
      "data-editor-ui": "",
      "aria-label": "Bảng điều khiển soạn hồ sơ",
    },
    h(
      "div",
      { class: "panel-header" },
      h(
        "div",
        {},
        h("div", { class: "panel-title", text: "Soạn hồ sơ" }),
        h("div", { class: "panel-subtitle", "data-panel-subtitle": "" }),
      ),
      h(
        "button",
        {
          type: "button",
          class: "icon-btn",
          "data-action": "collapse-panel",
          title: "Thu gọn bảng điều khiển",
          "aria-label": "Thu gọn",
        },
        "«",
      ),
    ),
    h("div", { class: "panel-issues", "data-issues": "", role: "status" }),
    PANEL_GROUPS.map(buildGroup),
    h(
      "div",
      { class: "shortcut-help" },
      h(
        "div",
        {},
        h("kbd", { text: "Ctrl" }),
        " + ",
        h("kbd", { text: "P" }),
        " In hồ sơ",
      ),
      h(
        "div",
        {},
        h("kbd", { text: "Ctrl" }),
        " + ",
        h("kbd", { text: "Enter" }),
        " Thêm dòng trong bảng",
      ),
      h("div", { text: "Bấm vào chữ tô xanh trên trang để sửa nhanh." }),
      h("div", { text: "Ô viền nét đứt: gõ trực tiếp trên trang." }),
    ),
    buildPeopleList(),
  );

  document.body.prepend(panel);
  document.body.append(
    h(
      "button",
      {
        type: "button",
        class: "panel-reopen",
        "data-editor-ui": "",
        "data-action": "expand-panel",
      },
      "☰ Soạn hồ sơ",
    ),
  );
}

function buildGroup(group) {
  return h(
    "details",
    { class: "panel-group", "data-group": group.id, open: Boolean(group.open) },
    h("summary", {}, group.title),
    h(
      "div",
      { class: "group-body" },
      group.hint ? h("p", { class: "group-hint", text: group.hint }) : null,
      (group.fields || []).map(buildField),
      (group.toggles || []).map(buildToggle),
      group.build ? group.build() : null,
      group.actions
        ? h(
            "div",
            { class: "group-actions" },
            group.actions.map(buildActionButton),
          )
        : null,
    ),
  );
}

function buildField(field) {
  const id = `field-${field.key}`;
  const control =
    field.multiline || field.textarea
      ? h("textarea", {
          id,
          rows: field.textarea ? 3 : 2,
          "data-model": field.key,
        })
      : h("input", {
          id,
          type: "text",
          autocomplete: "off",
          list: field.people ? "peopleList" : null,
          placeholder: field.placeholder,
          "data-model": field.key,
        });

  return h(
    "div",
    { class: "field" },
    h(
      "label",
      { for: id },
      field.label,
      field.required ? h("span", { class: "req", text: " *" }) : null,
    ),
    field.position
      ? h(
          "div",
          { class: "field-pair" },
          control,
          h("input", {
            type: "text",
            autocomplete: "off",
            placeholder: "Chức vụ",
            "aria-label": `${field.label} – chức vụ`,
            "data-model": field.position,
          }),
        )
      : control,
    field.toggle ? buildToggle(field.toggle) : null,
  );
}

function buildToggle({ key, label }) {
  return h(
    "label",
    { class: "toggle" },
    h("input", { type: "checkbox", "data-model": key }),
    h("span", { text: label }),
    key in AUTO_TOGGLES
      ? h("span", {
          class: "auto-badge",
          "data-auto-badge": key,
          title: "Đang theo quy tắc tự động",
          text: "tự động",
        })
      : null,
  );
}

function buildActionButton({ action, label, primary }) {
  return h(
    "button",
    {
      type: "button",
      class: primary ? "btn btn-primary" : "btn",
      "data-action": action,
    },
    label,
  );
}

function buildDatesGroup() {
  const smallField = (key, label) =>
    h(
      "div",
      { class: "field" },
      h("label", { for: `field-${key}`, text: label }),
      h("input", {
        id: `field-${key}`,
        type: "text",
        autocomplete: "off",
        "data-model": key,
      }),
    );

  const dateInput = (key, label) =>
    h("input", {
      type: "text",
      inputmode: "numeric",
      maxlength: 2,
      autocomplete: "off",
      "aria-label": label,
      "data-model": key,
    });

  return [
    h(
      "div",
      { class: "field-grid" },
      smallField("docNumber", "Số tờ trình"),
      smallField("docSymbol", "Ký hiệu"),
      smallField("year", "Năm"),
    ),
    h(
      "table",
      { class: "date-table" },
      h(
        "thead",
        {},
        h(
          "tr",
          {},
          h("th", { text: "Văn bản" }),
          h("th", { text: "Ngày" }),
          h("th", { text: "Tháng" }),
          h("th"),
        ),
      ),
      h(
        "tbody",
        {},
        DATE_ROWS.map((row, index) =>
          h(
            "tr",
            {},
            h("td", { text: row.label }),
            h("td", {}, dateInput(row.day, `${row.label}: ngày`)),
            h("td", {}, dateInput(row.month, `${row.label}: tháng`)),
            h(
              "td",
              {},
              h(
                "button",
                {
                  type: "button",
                  class: "link-btn",
                  "data-action": "date-today",
                  "data-row": index,
                  title: "Điền ngày hôm nay",
                },
                "Hôm nay",
              ),
            ),
          ),
        ),
      ),
    ),
    h("p", {
      class: "group-hint",
      text: "Để trống để chừa chỗ viết tay. Ký hiệu tờ trình được lưu bằng nút “Lưu làm mặc định” ở mục Người ký.",
    }),
    h(
      "div",
      { class: "group-actions" },
      buildActionButton({
        action: "date-all-today",
        label: "Hôm nay cho tất cả",
      }),
      buildActionButton({ action: "date-clear", label: "Để trống tất cả" }),
    ),
  ];
}

function buildPagesGroup() {
  return h(
    "ul",
    { class: "doc-list" },
    getPages().map((page, index) =>
      h(
        "li",
        {},
        h(
          "label",
          { class: "toggle" },
          h("input", { type: "checkbox", "data-doc-toggle": page.dataset.doc }),
          h("span", { text: `${index + 1}. ${page.dataset.title}` }),
        ),
        h(
          "button",
          {
            type: "button",
            class: "link-btn",
            "data-action": "goto-page",
            "data-doc": page.dataset.doc,
          },
          "Xem",
        ),
      ),
    ),
  );
}

function buildPageLabels() {
  getPages().forEach((page, index) => {
    page.before(
      h(
        "div",
        { class: "page-label", "data-editor-ui": "" },
        h("span", {
          class: "page-label-title",
          text: `${index + 1}. ${page.dataset.title}`,
        }),
        h(
          "label",
          { class: "toggle" },
          h("input", { type: "checkbox", "data-doc-toggle": page.dataset.doc }),
          h("span", { text: "In trang này" }),
        ),
      ),
    );
  });
}

// Gợi ý tên khi gõ: lấy từ danh sách Excel và người ký cố định
function buildPeopleList() {
  const names = new Set();
  const add = (value) => {
    const name = toText(value).trim();
    if (name) names.add(name);
  };

  (loadJson(STORAGE_KEYS.printerData, []) || []).forEach((row) => {
    ["requester", "officer", "representative", "representativeStaff"].forEach(
      (key) => add(row?.[key]),
    );
  });
  PEOPLE_FIELDS.forEach((key) => add(state[key]));

  return h(
    "datalist",
    { id: "peopleList" },
    [...names]
      .sort((a, b) => a.localeCompare(b, "vi"))
      .map((name) => h("option", { value: name })),
  );
}

// Ghi giá trị state vào ô nhập (bỏ qua ô đang gõ để không nhảy con trỏ)
function refreshPanelInputs() {
  panel.querySelectorAll("[data-model]").forEach((input) => {
    if (input.type === "checkbox" || input === document.activeElement) return;

    const value = toText(state[input.dataset.model]);
    if (input.value !== value) input.value = value;
  });
}

function getBaseline(key) {
  if (EXCEL_FIELDS.includes(key)) return toText(sourceRow[key]);
  if (key in FACTORY_DEFAULTS) return toText(settings.defaults[key]);
  if (key in FACTORY_RULES) return toText(settings.rules[key]);
  return undefined;
}

function renderPanelState(view) {
  panel.querySelectorAll("[data-model]").forEach((input) => {
    const key = input.dataset.model;

    if (input.type === "checkbox") {
      input.checked = Boolean(view[key]);
      return;
    }

    const baseline = getBaseline(key);
    input.classList.toggle(
      "is-modified",
      baseline !== undefined && toText(state[key]) !== baseline,
    );
  });

  panel.querySelectorAll("[data-auto-badge]").forEach((badge) => {
    badge.hidden = manualToggles.has(badge.dataset.autoBadge);
  });

  const giverInput = panel.querySelector('[data-model="giverOverride"]');
  if (giverInput)
    giverInput.placeholder = `Tự động: ${getAutoGiver() || "(trống)"}`;

  panel.querySelector("[data-panel-subtitle]").textContent =
    toText(state.department) || "Chưa chọn đơn vị";
}

function focusPanelField(fieldKey) {
  const key = FIELD_INPUT_ALIAS[fieldKey]?.() || fieldKey;
  const input = panel.querySelector(`[data-model="${key}"]`);
  if (!input) return;

  setPanelCollapsed(false);
  const group = input.closest("details");
  if (group && !group.open) group.open = true;

  input.scrollIntoView({ block: "center", behavior: "smooth" });
  input.focus({ preventScroll: true });
  if (input.type !== "checkbox") input.select?.();
  flash(input);
}

function setPanelCollapsed(collapsed) {
  if (document.body.classList.contains("panel-collapsed") === collapsed) return;

  document.body.classList.toggle("panel-collapsed", collapsed);
  hideTemplateBar();
  saveUiPrefs();
}

function applyUiPrefs() {
  const prefs = loadJson(STORAGE_KEYS.ui, {}) || {};
  const collapsed =
    typeof prefs.collapsed === "boolean"
      ? prefs.collapsed
      : window.innerWidth < 1100;
  document.body.classList.toggle("panel-collapsed", collapsed);

  if (Array.isArray(prefs.openGroups)) {
    panel.querySelectorAll("details[data-group]").forEach((details) => {
      details.open = prefs.openGroups.includes(details.dataset.group);
    });
  }
}

function saveUiPrefs() {
  saveJson(STORAGE_KEYS.ui, {
    collapsed: document.body.classList.contains("panel-collapsed"),
    openGroups: Array.from(panel.querySelectorAll("details[open]")).map(
      (details) => details.dataset.group,
    ),
  });
}

/* =========================
   THAO TÁC
   ========================= */

function saveToList() {
  const index = sourceRow._index;
  const rows = loadJson(STORAGE_KEYS.printerData, []) || [];

  if (!Number.isInteger(index) || !rows[index]) {
    showToast("Không tìm thấy dòng tương ứng trong danh sách.", "error");
    return;
  }

  EXCEL_FIELDS.forEach((key) => {
    rows[index][key] = toText(state[key]).trim();
  });

  if (!saveJson(STORAGE_KEYS.printerData, rows)) {
    showToast("Không lưu được danh sách.", "error");
    return;
  }

  sourceRow = { ...rows[index], _index: index };
  saveJson(STORAGE_KEYS.selectedPrinter, sourceRow);
  notifyParent("list-updated");
  render();
  showToast(
    "Đã lưu vào danh sách. Lưu ý: nhập lại file Excel sẽ thay thế thay đổi này.",
    "success",
  );
}

function saveDefaults() {
  Object.keys(FACTORY_DEFAULTS).forEach((key) => {
    settings.defaults[key] = state[key];
  });
  Object.keys(FACTORY_RULES).forEach((key) => {
    settings.rules[key] = toText(state[key]);
  });

  saveSettings();
  render();
  showToast("Đã lưu làm mặc định cho các hồ sơ sau.", "success");
}

function resetDefaults() {
  if (
    !confirm(
      "Khôi phục người ký, ký hiệu văn bản và quy tắc tự động về giá trị gốc ban đầu?",
    )
  ) {
    return;
  }

  settings.defaults = { ...FACTORY_DEFAULTS };
  settings.rules = { ...FACTORY_RULES };
  saveSettings();
  manualToggles.clear();
  setFields({ ...FACTORY_DEFAULTS, ...FACTORY_RULES });
  showToast("Đã khôi phục giá trị gốc.", "success");
}

function fillToday(rows) {
  const today = todayParts();
  const patch = { year: today.year };

  rows.forEach(({ day, month }) => {
    patch[day] = today.day;
    patch[month] = today.month;
  });

  setFields(patch);
}

function clearDates() {
  const patch = {};
  DATE_ROWS.forEach(({ day, month }) => {
    patch[day] = "";
    patch[month] = "";
  });
  setFields(patch);
}

function setDocIncluded(doc, included) {
  if (included) excludedDocs.delete(doc);
  else excludedDocs.add(doc);

  settings.excludedDocs = [...excludedDocs];
  saveSettings();
  render();
}

function scrollToDoc(doc) {
  const page = document.querySelector(`.page[data-doc="${doc}"]`);
  (page?.previousElementSibling || page)?.scrollIntoView({
    behavior: "smooth",
    block: "start",
  });
}

function notifyParent(type) {
  if (isEmbedded) window.parent.postMessage({ source: "a4-editor", type }, "*");
}

// In qua trang cha để hồ sơ được lưu vào lịch sử
function requestPrint() {
  if (isEmbedded) notifyParent("print");
  else window.print();
}

const ACTIONS = {
  "collapse-panel": () => setPanelCollapsed(true),
  "expand-panel": () => setPanelCollapsed(false),
  "save-to-list": saveToList,
  "save-defaults": saveDefaults,
  "reset-defaults": resetDefaults,
  "relink-recovered": relinkRecovered,
  "date-today": (button) => fillToday([DATE_ROWS[Number(button.dataset.row)]]),
  "date-all-today": () => fillToday(DATE_ROWS),
  "date-clear": clearDates,
  "goto-page": (button) => scrollToDoc(button.dataset.doc),
  "focus-issue": (button) =>
    focusIssue(currentIssues[Number(button.dataset.index)]),
  "row-add": (button) => {
    const row = button.closest("tr");
    addRowAfter(row.closest("table"), row);
    onTablesChanged();
  },
  "row-delete": (button) => removeRow(button.closest("tr")),
  "row-append": (button) => {
    const table = getTable(button.dataset.tableId);
    addRowAfter(table, getDataRows(table).at(-1));
    onTablesChanged();
  },
  "template-save": saveTemplate,
  "template-reset": resetTemplate,
};

/* =========================
   SỰ KIỆN
   ========================= */

function handleInput(event) {
  const target = event.target;

  if (target.dataset?.docToggle) {
    setDocIncluded(target.dataset.docToggle, target.checked);
    return;
  }

  if (target.dataset?.model) {
    setField(
      target.dataset.model,
      target.type === "checkbox" ? target.checked : target.value,
    );
    return;
  }

  // Gõ trực tiếp trên trang
  markDirty();
  const cell = target.closest?.(".page table[data-table] td");
  if (cell) handleCellInput(cell);
  if (activeTemplateBlock) updateTemplateBar();
}

function handleClick(event) {
  const actionElement = event.target.closest("[data-action]");
  if (actionElement) {
    ACTIONS[actionElement.dataset.action]?.(actionElement, event);
    return;
  }

  const fieldElement = event.target.closest(".page [data-field]");
  if (fieldElement) {
    focusPanelField(fieldElement.dataset.field);
  }
}

function handleKeydown(event) {
  const withCtrl = event.ctrlKey || event.metaKey;

  if (withCtrl && event.key.toLowerCase() === "p") {
    event.preventDefault();
    requestPrint();
    return;
  }

  if (withCtrl && event.key === "Enter") {
    const row = event.target.closest?.(".page table[data-table] tbody tr");
    if (row) {
      event.preventDefault();
      addRowAfter(row.closest("table"), row);
      onTablesChanged();
    }
    return;
  }

  if (event.key === "Enter" && event.target.closest?.("[data-single-line]")) {
    event.preventDefault();
  }
}

function pastePlainText(event) {
  const editable = event.target.closest?.('[contenteditable="true"]');
  if (!editable) return;

  event.preventDefault();
  const text = (event.clipboardData || window.clipboardData).getData(
    "text/plain",
  );

  // insertText giữ được Ctrl+Z và tự phát sự kiện input
  if (!document.execCommand("insertText", false, text)) {
    insertTextAtCursor(editable, text);
    editable.dispatchEvent(new Event("input", { bubbles: true }));
  }
}

function insertTextAtCursor(editable, text) {
  const selection = window.getSelection();
  if (!selection?.rangeCount) {
    editable.append(text);
    return;
  }

  const range = selection.getRangeAt(0);
  range.deleteContents();
  range.insertNode(document.createTextNode(text));
  range.collapse(false);
  selection.removeAllRanges();
  selection.addRange(range);
}

function bindEvents() {
  document.addEventListener("input", handleInput);
  document.addEventListener("click", handleClick);
  document.addEventListener("keydown", handleKeydown);
  document.addEventListener("paste", pastePlainText);

  document.addEventListener("focusin", (event) => {
    const block = event.target.closest?.("[data-persist]");
    if (block) showTemplateBar(block);
    else if (!templateBar.contains(event.target)) hideTemplateBar();
  });

  document.addEventListener("focusout", (event) => {
    if (!activeTemplateBlock) return;
    const next = event.relatedTarget;
    if (
      next &&
      (activeTemplateBlock.contains(next) || templateBar.contains(next))
    ) {
      return;
    }
    hideTemplateBar();
  });

  window.addEventListener("resize", positionTemplateBar);

  panel.querySelectorAll("details[data-group]").forEach((details) => {
    details.addEventListener("toggle", saveUiPrefs);
  });
}

/* =========================
   KHỞI TẠO
   ========================= */

function init() {
  settings = loadSettings();
  sourceRow = loadJson(STORAGE_KEYS.selectedPrinter, {}) || {};
  state = createInitialState(sourceRow);
  excludedDocs = new Set(settings.excludedDocs);

  document.body.classList.add("is-editing");
  initTemplates();
  initTables();
  prepareFieldElements();
  buildPageLabels();
  buildPanel();
  buildTemplateBar();
  applyUiPrefs();
  refreshPanelInputs();
  bindEvents();
  render();

  // Dùng bởi index.html khi in và lưu lịch sử
  window.A4Editor = {
    getMeta: () =>
      Object.fromEntries(
        EXCEL_FIELDS.map((key) => [key, toText(state[key]).trim()]),
      ),
    getIssues: () => collectIssues(buildView()).map((issue) => issue.label),
    isDirty: () => dirty,
    markSaved: () => {
      dirty = false;
    },
  };

  if (!sourceRow.department) {
    showToast(
      "Chưa có dữ liệu đơn vị – hãy mở hồ sơ từ danh sách hoặc nhập ở bảng bên trái.",
      "warning",
    );
  }
}

onReady(init);
