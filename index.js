const STORAGE_KEYS = {
  printerData: "printerData",
  selectedPrinter: "selectedPrinter",
};

// Số dòng tiêu đề ở đầu sheet Excel cần bỏ qua
const EXCEL_HEADER_ROWS = 3;

const EXCEL_COLUMN_INDEX = {
  department: 2,
  shortDepartment: 3,
  requester: 4,
  officer: 5,
  representative: 6,
  representativeStaff: 7,
  daysW4: 8,
  monthW4: 9,
};

const REQUIRED_FIELDS = [
  "department",
  "shortDepartment",
  "requester",
  "officer",
  "representative",
  "representativeStaff",
];

const TABLE_FIELDS = Object.keys(EXCEL_COLUMN_INDEX);

const HISTORY_SEARCH_FIELDS = [
  "department",
  "shortDepartment",
  "requester",
  "officer",
  "representative",
  "representativeStaff",
];

// Lịch sử được đọc một lần rồi giữ trong bộ nhớ, ô tìm kiếm chỉ lọc trên bản này
let historyCache = [];

const getElement = (id) => document.getElementById(id);

function saveJson(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
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

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function formatDate(iso) {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return "";

  const pad = (value) => String(value).padStart(2, "0");
  return `${pad(date.getDate())}/${pad(date.getMonth() + 1)}/${date.getFullYear()} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function formatBytes(bytes) {
  if (!bytes) return "0 KB";
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  if (bytes < 1024 ** 3) return `${(bytes / 1024 ** 2).toFixed(1)} MB`;
  return `${(bytes / 1024 ** 3).toFixed(1)} GB`;
}

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFC")
    .toLowerCase()
    .trim();
}

/* =========================
   ĐỌC EXCEL & BẢNG DỮ LIỆU
   ========================= */

function cleanExcelValue(value) {
  if (value === undefined || value === null) return "";

  const text = String(value);

  // Nếu trong Excel cố ý nhập &nbsp; thì chuyển thành khoảng trắng không ngắt dòng
  if (text.includes("&nbsp;")) {
    return text.replace(/&nbsp;/gi, " ");
  }

  return text
    .replace(/ /g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function mapExcelRow(row) {
  return Object.fromEntries(
    Object.entries(EXCEL_COLUMN_INDEX).map(([key, index]) => [
      key,
      cleanExcelValue(row[index]),
    ]),
  );
}

function hasRequiredData(row) {
  return REQUIRED_FIELDS.every((field) => Boolean(row[field]));
}

async function readExcelFile(file) {
  if (typeof XLSX === "undefined") {
    alert(
      "Chưa tải được thư viện đọc Excel. Kiểm tra kết nối mạng rồi tải lại trang.",
    );
    return;
  }

  try {
    const data = new Uint8Array(await file.arrayBuffer());
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const dataRows = rows
      .slice(EXCEL_HEADER_ROWS)
      .map(mapExcelRow)
      .filter(hasRequiredData);

    if (!dataRows.length) {
      alert("Không tìm thấy dòng dữ liệu hợp lệ nào trong file Excel.");
      return;
    }

    saveJson(STORAGE_KEYS.printerData, dataRows);
    renderTable(dataRows);
  } catch (error) {
    console.error("Lỗi đọc file Excel", error);
    alert("Không đọc được file Excel. Hãy kiểm tra lại file.");
  }
}

function renderTable(data) {
  const tbody = document.querySelector("#printerTable tbody");
  const fragment = document.createDocumentFragment();

  data.forEach((row, index) => {
    const tr = document.createElement("tr");

    TABLE_FIELDS.forEach((field) => {
      const td = document.createElement("td");
      td.textContent = row[field];
      tr.appendChild(td);
    });

    const actionTd = document.createElement("td");
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = "Xuất hồ sơ";
    button.addEventListener("click", () => openModal(row, index));
    actionTd.appendChild(button);
    tr.appendChild(actionTd);

    fragment.appendChild(tr);
  });

  tbody.replaceChildren(fragment);
}

function loadDataFromLocalStorage() {
  const data = loadJson(STORAGE_KEYS.printerData, []);
  if (data.length) renderTable(data);
}

/* =========================
   MODAL IN HỒ SƠ
   ========================= */

// _index giúp trang A4 lưu ngược thay đổi vào đúng dòng của danh sách
function openModal(selectedData, index) {
  saveJson(STORAGE_KEYS.selectedPrinter, { ...selectedData, _index: index });
  getElement("modalTitle").textContent = selectedData.department
    ? `Soạn hồ sơ – ${selectedData.department}`
    : "Soạn hồ sơ";
  getElement("modalIframe").src = "A4.html";
  getElement("modal").classList.add("active");
}

function closeModal() {
  if (
    getEditor()?.isDirty() &&
    !confirm("Hồ sơ đã sửa nhưng chưa in. Đóng và bỏ các thay đổi?")
  ) {
    return;
  }

  getElement("modal").classList.remove("active");
  getElement("modalIframe").src = "about:blank";
}

function getIframeDocument() {
  const iframe = getElement("modalIframe");
  return iframe?.contentDocument || iframe?.contentWindow?.document || null;
}

// API do A4.js cung cấp (null nếu trang A4 chưa tải xong)
function getEditor() {
  try {
    return getElement("modalIframe").contentWindow?.A4Editor || null;
  } catch {
    return null;
  }
}

// Chép giá trị người dùng đã nhập vào ô input/textarea/select sang bản sao
function copyFormValues(sourceRoot, cloneRoot) {
  const sourceFields = sourceRoot.querySelectorAll("input, textarea, select");
  const clonedFields = cloneRoot.querySelectorAll("input, textarea, select");

  sourceFields.forEach((source, index) => {
    const cloneField = clonedFields[index];
    if (!cloneField) return;

    if (source.tagName === "TEXTAREA") {
      cloneField.textContent = source.value;
      return;
    }

    if (source.tagName === "SELECT") {
      Array.from(cloneField.options).forEach((option, optionIndex) => {
        option.toggleAttribute(
          "selected",
          Boolean(source.options[optionIndex]?.selected),
        );
      });
      return;
    }

    if (source.type === "checkbox" || source.type === "radio") {
      cloneField.toggleAttribute("checked", source.checked);
    }

    cloneField.setAttribute("value", source.value);
  });
}

function readCssText(sheet) {
  try {
    return [...sheet.cssRules].map((rule) => rule.cssText).join("\n");
  } catch {
    return null; // trình duyệt không cho đọc (vd. mở bằng file://)
  }
}

// Nhúng thẳng CSS vào bản chụp: sau này A4.css có sửa thì hồ sơ đã lưu
// vẫn hiện và in lại đúng như lúc in. Không đọc được CSS thì giữ thẻ <link>.
function inlineStylesheets(doc, clone) {
  const sourceLinks = doc.querySelectorAll('link[rel~="stylesheet"]');
  const clonedLinks = clone.querySelectorAll('link[rel~="stylesheet"]');

  sourceLinks.forEach((link, index) => {
    const css = link.sheet ? readCssText(link.sheet) : null;
    if (css === null || !clonedLinks[index]) return;

    const style = doc.createElement("style");
    style.dataset.source = link.getAttribute("href") || "";
    style.textContent = css;
    clonedLinks[index].replaceWith(style);
  });
}

// Tạo bản HTML tĩnh của hồ sơ để lưu lịch sử:
// - bỏ <script> để khi xem lại không bị A4.js điền đè dữ liệu của hồ sơ đang chọn
// - bỏ bảng điều khiển, nút thêm/xoá dòng và các trang không in (giảm dung lượng)
// - nhúng CSS và thêm <base> để bản lưu tự hiển thị đúng, không phụ thuộc A4.css sau này
function createSnapshotHTML(doc) {
  const clone = doc.documentElement.cloneNode(true);
  copyFormValues(doc, clone);
  inlineStylesheets(doc, clone);

  clone
    .querySelectorAll("script, button, [data-editor-ui], .is-excluded")
    .forEach((element) => element.remove());
  clone
    .querySelectorAll("[contenteditable]")
    .forEach((element) => element.removeAttribute("contenteditable"));
  clone.querySelector("body")?.classList.remove("is-editing", "panel-collapsed");

  const head = clone.querySelector("head");
  if (head && !head.querySelector("base")) {
    const base = doc.createElement("base");
    base.href = doc.baseURI;
    head.prepend(base);
  }

  return `<!DOCTYPE html>\n${clone.outerHTML}`;
}

function getSelectedMeta() {
  // Ưu tiên thông tin đã sửa trên trang A4
  const editorMeta = getEditor()?.getMeta();
  if (editorMeta) return editorMeta;

  const selected = loadJson(STORAGE_KEYS.selectedPrinter, {});
  return Object.fromEntries(
    TABLE_FIELDS.map((field) => [field, selected[field] || ""]),
  );
}

async function saveSnapshotToHistory(html) {
  try {
    await HistoryStore.add({ meta: getSelectedMeta(), html });
  } catch (error) {
    console.error("Không lưu được lịch sử", error);
    alert("Không lưu được hồ sơ vào lịch sử: " + error.message);
  }
}

function printIframe() {
  const doc = getIframeDocument();
  if (doc?.readyState !== "complete" || !doc.body?.children.length) {
    alert("Hồ sơ chưa tải xong, vui lòng đợi một chút rồi bấm in lại.");
    return;
  }

  const editor = getEditor();
  const issues = editor?.getIssues() || [];
  if (issues.length) {
    const preview = issues.slice(0, 8).join("\n- ");
    const more = issues.length > 8 ? `\n... và ${issues.length - 8} mục khác` : "";
    if (!confirm(`Hồ sơ còn ${issues.length} mục cần kiểm tra:\n- ${preview}${more}\n\nVẫn in?`)) {
      return;
    }
  }

  // Chụp nội dung ngay trước khi in, việc nén/lưu chạy nền nên không làm chậm hộp thoại in
  saveSnapshotToHistory(createSnapshotHTML(doc));
  editor?.markSaved();

  const iframeWindow = getElement("modalIframe").contentWindow;
  iframeWindow.focus();
  iframeWindow.print();
}

/* =========================
   XUẤT WORD
   ========================= */

function formatLocalDate(date) {
  const pad = (value) => String(value).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

// Ví dụ: "Hồ sơ sửa chữa - Phòng Kỹ thuật - 2026-11-03.docx"
function buildWordFileName(meta, date = new Date()) {
  const name = ["Hồ sơ sửa chữa", meta?.department, formatLocalDate(date)]
    .filter(Boolean)
    .join(" - ")
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .trim();
  return `${name.slice(0, 150)}.docx`;
}

function buildWordTitle(meta) {
  return ["Hồ sơ sửa chữa", meta?.department].filter(Boolean).join(" - ");
}

async function withBusyButton(button, label, task) {
  const original = button.textContent;
  button.disabled = true;
  button.textContent = label;
  try {
    return await task();
  } finally {
    button.disabled = false;
    button.textContent = original;
  }
}

async function downloadWord(html, meta, date) {
  try {
    const blob = await DocxExport.fromHtml(html, { title: buildWordTitle(meta) });
    downloadFile(blob, buildWordFileName(meta, date), DocxExport.MIME);
    return true;
  } catch (error) {
    console.error("Lỗi xuất Word", error);
    alert("Không xuất được file Word: " + error.message);
    return false;
  }
}

// Xuất đúng những văn bản đang chọn in, với nội dung đang sửa trên trang A4
async function exportWord() {
  const doc = getIframeDocument();
  if (doc?.readyState !== "complete" || !doc.body?.children.length) {
    alert("Hồ sơ chưa tải xong, vui lòng đợi một chút rồi bấm xuất lại.");
    return;
  }
  if (!doc.querySelector(".page:not(.is-excluded)")) {
    alert('Chưa chọn văn bản nào. Hãy tích chọn ở mục "Văn bản cần in".');
    return;
  }

  const button = getElement("btnExportWord");
  await withBusyButton(button, "Đang xuất...", async () => {
    const saved = await downloadWord(createSnapshotHTML(doc), getSelectedMeta());
    // Nội dung đã nằm trong file Word -> đóng cửa sổ không cần hỏi lại
    if (saved) getEditor()?.markSaved();
  });
}

async function exportHistoryWord(id, button) {
  const item = findHistoryItem(id);
  await withBusyButton(button, "Đang xuất...", async () => {
    const html = await loadHistoryHtml(id);
    if (!html) {
      alert("Không tìm thấy hồ sơ.");
      return;
    }
    await downloadWord(
      prepareHistoryHtml(html),
      item?.meta,
      item ? new Date(item.createdAt) : new Date(),
    );
  });
}

/* =========================
   LỊCH SỬ HỒ SƠ ĐÃ IN
   ========================= */

async function openHistoryModal() {
  getElement("historyModal").classList.add("active");
  getElement("historySearch").focus();
  await refreshHistory();
}

function closeHistoryModal() {
  getElement("historyModal").classList.remove("active");
}

async function refreshHistory() {
  try {
    historyCache = await HistoryStore.list();
  } catch (error) {
    console.error("Không đọc được lịch sử", error);
    historyCache = [];
    getElement("historyList").innerHTML =
      `<p class="empty-state">Không đọc được lịch sử: ${escapeHtml(error.message)}</p>`;
    return;
  }

  renderHistoryList();
  renderHistorySummary();
}

async function renderHistorySummary() {
  const summary = getElement("historySummary");
  const parts = [`${historyCache.length} hồ sơ`];

  const estimate = await HistoryStore.estimateUsage().catch(() => null);
  if (estimate?.quota) {
    parts.push(
      `đã dùng ${formatBytes(estimate.usage)} / ${formatBytes(estimate.quota)}`,
    );
  }

  summary.textContent = parts.join(" · ");
}

function getFilteredHistory() {
  const keyword = normalizeText(getElement("historySearch").value);
  if (!keyword) return historyCache;

  return historyCache.filter((item) =>
    normalizeText(
      HISTORY_SEARCH_FIELDS.map((field) => item.meta?.[field]).join(" "),
    ).includes(keyword),
  );
}

function renderHistoryList() {
  const container = getElement("historyList");
  const history = getFilteredHistory();

  if (!history.length) {
    container.innerHTML = historyCache.length
      ? `<p class="empty-state">Không có hồ sơ nào khớp từ khoá.</p>`
      : `<p class="empty-state">Chưa có lịch sử nào.</p>`;
    return;
  }

  const rows = history
    .map((item, index) => {
      const printInfo =
        item.printCount > 1
          ? `<div class="muted">In ${item.printCount} lần, gần nhất ${formatDate(item.printedAt)}</div>`
          : "";

      return `
        <tr>
          <td>${index + 1}</td>
          <td>${formatDate(item.createdAt)}${printInfo}</td>
          <td>${escapeHtml(item.meta?.department)}</td>
          <td>${escapeHtml(item.meta?.requester)}</td>
          <td>${escapeHtml(item.meta?.officer)}</td>
          <td class="actions">
            <button type="button" data-action="view" data-id="${escapeHtml(item.id)}">Xem</button>
            <button type="button" data-action="reprint" data-id="${escapeHtml(item.id)}">In lại</button>
            <button type="button" class="btn-secondary" data-action="word" data-id="${escapeHtml(item.id)}" title="Tải file Word (.docx)">Word</button>
            <button type="button" class="btn-danger" data-action="delete" data-id="${escapeHtml(item.id)}">Xoá</button>
          </td>
        </tr>
      `;
    })
    .join("");

  container.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>STT</th>
          <th>Thời gian in</th>
          <th>Phòng ban</th>
          <th>Người đề nghị</th>
          <th>Cán bộ</th>
          <th>Thao tác</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

/* ---------- Xem / in lại hồ sơ đã lưu ---------- */

// CSS của biểu mẫu A4 trước bản nâng cấp 03/11/2026, dành cho hồ sơ lưu từ trước đó
const LEGACY_CSS_FILE = "history-legacy.css";
const PRINT_FRAME_TIMEOUT_MS = 15000;

// Thanh "In hồ sơ này" trong cửa sổ xem lại (không in ra giấy)
const HISTORY_VIEW_CSS = `
  .history-view-bar {
    position: fixed; right: 16px; bottom: 16px; z-index: 50;
    display: flex; align-items: center; gap: 10px;
    padding: 8px 8px 8px 14px; border-radius: 10px;
    color: #fff; background: #1e3a8a; box-shadow: 0 6px 16px rgba(15, 23, 42, 0.3);
    font: 13px/1.4 "Segoe UI", system-ui, Arial, sans-serif;
  }
  .history-view-bar button {
    margin: 0; padding: 7px 16px; border: 0; border-radius: 6px;
    color: #1e3a8a; background: #fff; cursor: pointer;
    font: 600 13px/1.2 "Segoe UI", system-ui, Arial, sans-serif;
  }
  .history-view-bar button:hover { background: #dbeafe; }
  @media print { .history-view-bar { display: none !important; } }
`;

function findHistoryItem(id) {
  return historyCache.find((entry) => entry.id === id);
}

function buildHistoryTitle(item) {
  return ["Hồ sơ", item?.meta?.department, formatDate(item?.createdAt)]
    .filter(Boolean)
    .join(" – ");
}

async function loadHistoryHtml(id) {
  try {
    return await HistoryStore.getHtml(id);
  } catch (error) {
    console.error("Không đọc được hồ sơ trong lịch sử", error);
    return null;
  }
}

// Hồ sơ lưu trước bản nâng cấp: các trang là div.WordSection…, chưa có lớp .page
function isLegacySnapshot(doc) {
  return !doc.querySelector(".page") && Boolean(doc.querySelector('[class*="WordSection"]'));
}

// Chuẩn bị bản lưu trước khi xem / in lại / xuất Word:
// - đường dẫn tương đối (A4.css...) luôn tính theo thư mục dự án đang mở,
//   kể cả khi bản lưu được tạo lúc mở bằng cổng / địa chỉ khác
// - hồ sơ cũ dùng lại CSS cũ: CSS hiện tại không khớp cấu trúc cũ nên vỡ bố cục, lộ tên cần ẩn
function prepareHistoryHtml(html, { title = "", toolbar = "" } = {}) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const head = doc.head;

  head.querySelectorAll("base").forEach((element) => element.remove());
  const base = doc.createElement("base");
  base.href = new URL("A4.html", location.href).href;
  head.prepend(base);

  if (isLegacySnapshot(doc)) {
    head.querySelectorAll('link[rel~="stylesheet"]').forEach((link) => link.remove());
    const link = doc.createElement("link");
    link.rel = "stylesheet";
    link.href = new URL(LEGACY_CSS_FILE, location.href).href;
    head.append(link);

    // Bản cũ còn nút "Thêm dòng / Xoá dòng" và vùng sửa trực tiếp
    doc.querySelectorAll("script, button").forEach((element) => element.remove());
    doc
      .querySelectorAll("[contenteditable]")
      .forEach((element) => element.removeAttribute("contenteditable"));
    // Đánh dấu từng trang để xuất Word nhận ra
    doc
      .querySelectorAll('body > [class*="WordSection"]')
      .forEach((section) => section.classList.add("page"));
  }

  if (title) doc.title = title;

  if (toolbar) {
    const style = doc.createElement("style");
    style.textContent = HISTORY_VIEW_CSS;
    head.append(style);

    const bar = doc.createElement("aside");
    bar.className = "history-view-bar";
    const label = doc.createElement("span");
    label.textContent = toolbar;
    const button = doc.createElement("button");
    button.type = "button";
    button.textContent = "In hồ sơ này";
    button.setAttribute("onclick", "window.print()");
    bar.append(label, button);
    doc.body.append(bar);
  }

  return `<!DOCTYPE html>\n${doc.documentElement.outerHTML}`;
}

async function viewHistory(id) {
  // Mở cửa sổ trước khi chờ dữ liệu để trình duyệt không coi là popup tự bật
  const win = window.open("", "_blank");
  if (!win) {
    alert("Trình duyệt đang chặn cửa sổ bật lên. Hãy cho phép popup cho trang này.");
    return;
  }

  const html = await loadHistoryHtml(id);
  if (!html) {
    win.close();
    alert("Không tìm thấy hồ sơ.");
    return;
  }

  const item = findHistoryItem(id);
  const page = prepareHistoryHtml(html, {
    title: buildHistoryTitle(item),
    toolbar: item ? `Bản lưu lúc ${formatDate(item.createdAt)}` : "Bản lưu",
  });
  // Mở bằng địa chỉ blob thay cho document.write: tải lại tab vẫn ra đúng hồ sơ
  win.location.replace(
    URL.createObjectURL(new Blob([page], { type: "text/html;charset=utf-8" })),
  );
}

// Dựng hồ sơ trong một iframe ẩn và chỉ trả về khi CSS, phông chữ đã tải xong
// (trước đây lệnh in chạy lúc CSS chưa tải nên bản in mất hết định dạng)
function loadPrintFrame(html) {
  getElement("historyPrintFrame")?.remove();

  const iframe = document.createElement("iframe");
  iframe.id = "historyPrintFrame";
  iframe.title = "Bản in lại hồ sơ";
  iframe.tabIndex = -1;
  iframe.setAttribute("aria-hidden", "true");
  Object.assign(iframe.style, {
    position: "fixed",
    right: "0",
    bottom: "0",
    width: "0",
    height: "0",
    border: "0",
    visibility: "hidden",
  });

  return new Promise((resolve, reject) => {
    const timer = setTimeout(
      () => reject(new Error("Quá thời gian tải hồ sơ.")),
      PRINT_FRAME_TIMEOUT_MS,
    );

    iframe.addEventListener(
      "load",
      async () => {
        clearTimeout(timer);
        await iframe.contentDocument?.fonts?.ready;
        resolve(iframe);
      },
      { once: true },
    );

    iframe.srcdoc = html;
    document.body.append(iframe);
  });
}

async function reprintHistory(id, button) {
  await withBusyButton(button, "Đang in...", async () => {
    const html = await loadHistoryHtml(id);
    if (!html) {
      alert("Không tìm thấy hồ sơ.");
      return;
    }

    try {
      const frame = await loadPrintFrame(
        prepareHistoryHtml(html, { title: buildHistoryTitle(findHistoryItem(id)) }),
      );
      frame.contentWindow.focus();
      frame.contentWindow.print();
    } catch (error) {
      console.error("Lỗi in lại hồ sơ", error);
      alert("Không in lại được hồ sơ: " + error.message);
    }
  });
}

async function deleteHistory(id) {
  if (!confirm("Bạn có chắc muốn xoá bản ghi này không?")) return;

  await HistoryStore.remove(id);
  await refreshHistory();
}

async function clearHistory() {
  if (!historyCache.length) return;
  if (
    !confirm(
      `Bạn có chắc muốn xoá toàn bộ ${historyCache.length} hồ sơ trong lịch sử không?\nNên bấm "Xuất sao lưu" trước khi xoá.`,
    )
  ) {
    return;
  }

  await HistoryStore.clear();
  await refreshHistory();
}

function downloadFile(content, fileName, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function exportHistoryJson() {
  const history = await HistoryStore.exportAll();
  if (!history.length) {
    alert("Chưa có dữ liệu lịch sử.");
    return;
  }

  const today = new Date().toISOString().slice(0, 10);
  downloadFile(
    JSON.stringify(history),
    `lich-su-ho-so-da-in-${today}.json`,
    "application/json;charset=utf-8",
  );
}

async function importHistoryJson(file) {
  try {
    const records = JSON.parse(await file.text());
    const imported = await HistoryStore.importAll(records);
    alert(`Đã nhập ${imported} hồ sơ (bỏ qua hồ sơ đã có).`);
    await refreshHistory();
  } catch (error) {
    console.error("Lỗi nhập file sao lưu", error);
    alert("File sao lưu không hợp lệ.");
  }
}

/* =========================
   KHỞI TẠO
   ========================= */

function closeTopModal() {
  const openModals = document.querySelectorAll(".modal.active");
  const top = openModals[openModals.length - 1];
  if (top?.id === "historyModal") closeHistoryModal();
  else if (top?.id === "modal") closeModal();
}

function bindEvents() {
  getElement("fileInput").addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (file) readExcelFile(file);
    event.target.value = ""; // cho phép chọn lại cùng một file
  });

  getElement("btnHeaderHistory").addEventListener("click", openHistoryModal);
  getElement("btnHistory").addEventListener("click", openHistoryModal);
  getElement("btnPrint").addEventListener("click", printIframe);
  getElement("btnExportWord").addEventListener("click", exportWord);
  getElement("btnCloseModal").addEventListener("click", closeModal);

  getElement("historySearch").addEventListener("input", renderHistoryList);
  getElement("btnExportHistory").addEventListener("click", exportHistoryJson);
  getElement("btnImportHistory").addEventListener("click", () =>
    getElement("historyImportInput").click(),
  );
  getElement("historyImportInput").addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (file) importHistoryJson(file);
    event.target.value = "";
  });
  getElement("btnClearHistory").addEventListener("click", clearHistory);
  getElement("btnCloseHistory").addEventListener("click", closeHistoryModal);

  getElement("historyList").addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) return;

    const { action, id } = button.dataset;
    if (action === "view") viewHistory(id);
    else if (action === "reprint") reprintHistory(id, button);
    else if (action === "word") exportHistoryWord(id, button);
    else if (action === "delete") deleteHistory(id);
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeTopModal();
  });

  // Tin nhắn từ trang A4 trong iframe (Ctrl+P, lưu thay đổi vào danh sách)
  window.addEventListener("message", (event) => {
    if (event.source !== getElement("modalIframe").contentWindow) return;

    const { source, type } = event.data || {};
    if (source !== "a4-editor") return;

    if (type === "print") printIframe();
    else if (type === "list-updated") loadDataFromLocalStorage();
  });
}

document.addEventListener("DOMContentLoaded", () => {
  bindEvents();
  loadDataFromLocalStorage();

  HistoryStore.requestPersistence();
  HistoryStore.migrateLegacy()
    .then((count) => {
      if (count) console.info(`Đã chuyển ${count} hồ sơ lịch sử cũ sang IndexedDB.`);
    })
    .catch((error) => console.error("Lỗi chuyển lịch sử cũ", error));
});
