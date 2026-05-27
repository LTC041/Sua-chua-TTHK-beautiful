const STORAGE_KEYS = {
  printerData: "printerData",
  selectedPrinter: "selectedPrinter",
  history: "repairHistoryRaw",
};

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

const REQUIRED_FIELDS = Object.keys(EXCEL_COLUMN_INDEX);

function getElement(id) {
  return document.getElementById(id);
}

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

function readExcelFile(file) {
  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const dataRows = rows
      .slice(3)
      .map(mapExcelRow)
      .filter(hasRequiredData);

    saveJson(STORAGE_KEYS.printerData, dataRows);
    renderTable(dataRows);
  };

  reader.readAsArrayBuffer(file);
}

function mapExcelRow(row) {
  return Object.fromEntries(
    Object.entries(EXCEL_COLUMN_INDEX).map(([key, index]) => [key, row[index] || ""]),
  );
}

function hasRequiredData(row) {
  return REQUIRED_FIELDS.every((field) => Boolean(row[field]));
}

function loadDataFromLocalStorage() {
  const data = loadJson(STORAGE_KEYS.printerData, []);
  if (data.length) renderTable(data);
}

function renderTable(data) {
  const tbody = document.querySelector("#printerTable tbody");
  tbody.innerHTML = "";

  data.forEach((row) => {
    const tr = document.createElement("tr");

    const visibleFields = [
      "department",
      "shortDepartment",
      "requester",
      "officer",
      "representative",
      "representativeStaff",
      "daysW4",
      "monthW4",
    ];

    visibleFields.forEach((field) => {
      const td = document.createElement("td");
      td.textContent = row[field];
      tr.appendChild(td);
    });

    const actionTd = document.createElement("td");
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = "Xuất hồ sơ";
    button.addEventListener("click", () => openModal(row));
    actionTd.appendChild(button);
    tr.appendChild(actionTd);

    tbody.appendChild(tr);
  });
}

function openModal(selectedData) {
  saveJson(STORAGE_KEYS.selectedPrinter, selectedData);

  const iframe = getElement("modalIframe");
  iframe.src = "A4.html";

  getElement("modal").classList.add("active");
}

function closeModal() {
  getElement("modal").classList.remove("active");
}

function printIframe() {
  saveCurrentSnapshotToHistory();

  const iframeWindow = getElement("modalIframe").contentWindow;
  iframeWindow.focus();
  iframeWindow.print();
}

function getHistory() {
  return loadJson(STORAGE_KEYS.history, []);
}

function setHistory(data) {
  saveJson(STORAGE_KEYS.history, data);
}

function escapeHtml(str) {
  return String(str || "")
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

function getIframeDocument() {
  const iframe = getElement("modalIframe");
  return iframe?.contentDocument || iframe?.contentWindow?.document || null;
}

function createSnapshotHTML(doc) {
  const clone = doc.documentElement.cloneNode(true);

  const sourceFields = doc.querySelectorAll("input, textarea, select");
  const clonedFields = clone.querySelectorAll("input, textarea, select");

  sourceFields.forEach((source, index) => {
    const cloneField = clonedFields[index];
    if (!cloneField) return;

    if (source.tagName === "TEXTAREA") {
      cloneField.textContent = source.value;
      return;
    }

    if (source.tagName === "SELECT") {
      Array.from(cloneField.options).forEach((option, optionIndex) => {
        option.selected = source.options[optionIndex]?.selected || false;
      });
      return;
    }

    if (source.type === "checkbox" || source.type === "radio") {
      if (source.checked) cloneField.setAttribute("checked", "checked");
      else cloneField.removeAttribute("checked");
      cloneField.setAttribute("value", source.value);
      return;
    }

    cloneField.setAttribute("value", source.value);
  });

  return `<!DOCTYPE html>\n${clone.outerHTML}`;
}

function saveCurrentSnapshotToHistory() {
  const doc = getIframeDocument();
  if (!doc?.documentElement) {
    alert("Chưa lấy được dữ liệu hồ sơ để lưu lịch sử.");
    return;
  }

  const selected = loadJson(STORAGE_KEYS.selectedPrinter, {});
  const record = {
    id: `HS${Date.now()}`,
    createdAt: new Date().toISOString(),
    meta: {
      department: selected.department || "",
      shortDepartment: selected.shortDepartment || "",
      requester: selected.requester || "",
      officer: selected.officer || "",
      representative: selected.representative || "",
      representativeStaff: selected.representativeStaff || "",
      daysW4: selected.daysW4 || "",
      monthW4: selected.monthW4 || "",
    },
    html: createSnapshotHTML(doc),
  };

  const history = getHistory();
  history.unshift(record);
  setHistory(history);
}

function openHistoryModal() {
  getElement("historyModal").classList.add("active");
  renderHistoryList();
}

function closeHistoryModal() {
  getElement("historyModal").classList.remove("active");
}

function getFilteredHistory() {
  const keyword = (getElement("historySearch").value || "").trim().toLowerCase();
  const history = getHistory();

  if (!keyword) return history;

  return history.filter((item) => {
    const text = [
      item.meta.department,
      item.meta.shortDepartment,
      item.meta.requester,
      item.meta.officer,
      item.meta.representative,
      item.meta.representativeStaff,
    ]
      .join(" ")
      .toLowerCase();

    return text.includes(keyword);
  });
}

function renderHistoryList() {
  const container = getElement("historyList");
  const history = getFilteredHistory();

  if (!history.length) {
    container.innerHTML = "<p>Chưa có lịch sử nào.</p>";
    return;
  }

  container.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>STT</th>
          <th>Thời gian lưu</th>
          <th>Phòng ban</th>
          <th>Người đề nghị</th>
          <th>Cán bộ</th>
          <th>Thao tác</th>
        </tr>
      </thead>
      <tbody>
        ${history
          .map(
            (item, index) => `
              <tr>
                <td>${index + 1}</td>
                <td>${formatDate(item.createdAt)}</td>
                <td>${escapeHtml(item.meta.department)}</td>
                <td>${escapeHtml(item.meta.requester)}</td>
                <td>${escapeHtml(item.meta.officer)}</td>
                <td>
                  <button type="button" onclick="viewHistory('${item.id}')">Xem</button>
                  <button type="button" onclick="deleteHistory('${item.id}')">Xoá</button>
                </td>
              </tr>
            `,
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function viewHistory(id) {
  const item = getHistory().find((record) => record.id === id);
  if (!item) {
    alert("Không tìm thấy hồ sơ.");
    return;
  }

  const win = window.open("", "_blank");
  if (!win) {
    alert("Trình duyệt đang chặn popup.");
    return;
  }

  win.document.open();
  win.document.write(item.html);
  win.document.close();
}

function deleteHistory(id) {
  if (!confirm("Bạn có chắc muốn xoá bản ghi này không?")) return;

  const history = getHistory().filter((record) => record.id !== id);
  setHistory(history);
  renderHistoryList();
}

function clearHistory() {
  if (!confirm("Bạn có chắc muốn xoá toàn bộ lịch sử không?")) return;

  localStorage.removeItem(STORAGE_KEYS.history);
  renderHistoryList();
}

function exportHistoryJson() {
  const history = getHistory();
  if (!history.length) {
    alert("Chưa có dữ liệu lịch sử.");
    return;
  }

  const blob = new Blob([JSON.stringify(history, null, 2)], {
    type: "application/json;charset=utf-8",
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "lich-su-ho-so-da-in.json";
  link.click();
  URL.revokeObjectURL(url);
}

function bindEvents() {
  getElement("fileInput").addEventListener("change", function (event) {
    const file = event.target.files[0];
    if (file) readExcelFile(file);
  });
}

document.addEventListener("DOMContentLoaded", function () {
  bindEvents();
  loadDataFromLocalStorage();
});

window.openModal = openModal;
window.closeModal = closeModal;
window.printIframe = printIframe;
window.openHistoryModal = openHistoryModal;
window.closeHistoryModal = closeHistoryModal;
window.renderHistoryList = renderHistoryList;
window.viewHistory = viewHistory;
window.deleteHistory = deleteHistory;
window.clearHistory = clearHistory;
window.exportHistoryJson = exportHistoryJson;
