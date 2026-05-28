function onReady(callback) {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", callback);
  } else {
    callback();
  }
}

function loadSelectedPrinter() {
  try {
    return JSON.parse(localStorage.getItem("selectedPrinter") || "null");
  } catch (error) {
    console.error("Dữ liệu selectedPrinter không hợp lệ.", error);
    return null;
  }
}

const setText = (className, value) => {
  document.querySelectorAll(`.${className}`).forEach((element) => {
    element.textContent = value === undefined || value === null ? "" : value;
  });
};

function applySelectedPrinterData() {
  const printerData = loadSelectedPrinter();

  if (!printerData) {
    console.error("Không tìm thấy dữ liệu trong localStorage!");
    return;
  }

  setText("department", printerData.department);
  setText("department-short", printerData.shortDepartment);
  setText("requester", printerData.requester);
  setText("officer", printerData.officer);
  setText("representative", printerData.representative);
  setText("representativeStaff", printerData.representativeStaff);
  setText("dayW4", printerData.daysW4);
  setText("monthW4", printerData.monthW4);
}

function breakAfterTrungTam() {
  document.querySelectorAll(".break-after-trung-tam").forEach((element) => {
    element.innerHTML = element.innerHTML.replace(/(Trung tâm)\s/, "$1<br>");
    element.style.textAlign = "center";
  });
}

function fixWidowText(selector) {
  document.querySelectorAll(selector).forEach((element) => {
    const words = element.textContent.trim().split(/\s+/).filter(Boolean);
    if (words.length <= 1) return;

    const lastWord = words.pop();
    words[words.length - 1] = `${words[words.length - 1]}\u00A0${lastWord}`;
    element.textContent = words.join(" ");
  });
}

function syncQuantityInputs() {
  const input1 = document.querySelector("#table1 input");
  const input2 = document.getElementById("input2");

  if (!input1 || !input2) return;

  input2.value = input1.value;
  input1.addEventListener("input", () => {
    input2.value = input1.value;
  });
}

function updateMainRequestSignature() {
  const section = document.querySelector(".WordSection3");
  if (!section) return;

  const requesterElement = section.querySelector(".requester");
  if (!requesterElement) return;

  const isLeTuanCuong =
    requesterElement.textContent.trim().toLowerCase() === "lê tuấn cường";
  if (!isLeTuanCuong) return;

  const signatureDiv = section.querySelector(".none-sign");
  if (signatureDiv) signatureDiv.style.display = "none";

  section.querySelectorAll(".footer-page-buy").forEach((element) => {
    if (window.getComputedStyle(element).justifyContent === "space-between") {
      element.style.justifyContent = "space-around";
    }
  });
}

function hideOfficerNguyenThiThu() {
  document.querySelectorAll(".officer").forEach((element) => {
    if (element.textContent.trim() === "Nguyễn Thị Thu") {
      element.style.color = "white";
    }
  });
}

function getDataRows(table) {
  return Array.from(table.rows).filter(
    (row) => row.querySelectorAll("td").length > 0,
  );
}

function resetCellContent(cell, cellIndex, nextIndex) {
  const formControls = cell.querySelectorAll("input, select, textarea");

  if (cellIndex === 0) {
    cell.textContent = String(nextIndex);
    return;
  }

  if (formControls.length) {
    formControls.forEach((control, controlIndex) => {
      resetFormControl(control);
      refreshControlIdentifier(control, nextIndex, controlIndex);
    });
    return;
  }

  cell.textContent = "";
}

function resetFormControl(control) {
  const tagName = control.tagName;

  if (tagName === "INPUT") {
    const type = (control.type || "text").toLowerCase();
    if (type === "checkbox" || type === "radio") {
      control.checked = false;
    } else {
      control.value = "";
    }
    return;
  }

  if (tagName === "SELECT") {
    control.selectedIndex = -1;
    return;
  }

  if (tagName === "TEXTAREA") {
    control.value = "";
  }
}

function refreshControlIdentifier(control, rowIndex, controlIndex) {
  const baseId =
    control.id ||
    control.name ||
    control.dataset.id ||
    control.tagName.toLowerCase();
  const suffix = `_${rowIndex}_${controlIndex}`;

  if (control.id) control.id = `${baseId}${suffix}`;
  if (control.name) control.name = `${control.name}${suffix}`;
}

function addRow(tableId) {
  const table = document.getElementById(tableId);
  if (!table) return;

  const dataRows = getDataRows(table);
  if (!dataRows.length) return;

  const lastRow = dataRows[dataRows.length - 1];
  const newRow = lastRow.cloneNode(true);
  const nextIndex = dataRows.length + 1;

  Array.from(newRow.cells).forEach((cell, cellIndex) => {
    resetCellContent(cell, cellIndex, nextIndex);
  });

  lastRow.parentNode.insertBefore(newRow, lastRow.nextSibling);
}

function deleteRow(tableId) {
  const table = document.getElementById(tableId);
  if (!table) return;

  const dataRows = getDataRows(table);
  if (dataRows.length <= 1) {
    alert("Phải có ít nhất một dòng!");
    return;
  }

  dataRows[dataRows.length - 1].remove();
}

function updateThuHoiSignature() {
  document.querySelectorAll(".requester").forEach((requesterElement) => {
    const requesterName = requesterElement.textContent.trim();
    const requesterBlock = requesterElement.closest(".ben-giao-requester");
    const ngocNgaBlock = findNextSiblingByClass(
      requesterBlock,
      "ben-giao-ngocnga",
    );

    if (requesterName === "Dương Thị Lan Thu") {
      requesterBlock?.classList.add("hidden");
      ngocNgaBlock?.classList.remove("hidden");
    } else {
      requesterBlock?.classList.remove("hidden");
      ngocNgaBlock?.classList.add("hidden");
    }
  });
}

function findNextSiblingByClass(element, className) {
  let sibling = element?.nextElementSibling || null;

  while (sibling) {
    if (sibling.classList.contains(className)) return sibling;
    sibling = sibling.nextElementSibling;
  }

  return null;
}

function pastePlainText(event) {
  const editable = event.target.closest('[contenteditable="true"]');
  if (!editable) return;

  event.preventDefault();
  const text = (event.clipboardData || window.clipboardData).getData(
    "text/plain",
  );
  insertTextAtCursor(editable, text);
}

function insertTextAtCursor(editable, text) {
  if (!window.getSelection || !document.createRange) {
    document.execCommand("insertText", false, text);
    return;
  }

  const selection = window.getSelection();

  if (!selection.rangeCount) {
    editable.focus();
    document.execCommand("insertText", false, text);
    return;
  }

  const range = selection.getRangeAt(0);
  range.deleteContents();
  range.insertNode(document.createTextNode(text));
  range.collapse(false);

  selection.removeAllRanges();
  selection.addRange(range);
}

onReady(() => {
  applySelectedPrinterData();
  breakAfterTrungTam();
  fixWidowText(".printerName");
  syncQuantityInputs();
  updateMainRequestSignature();
  hideOfficerNguyenThiThu();
  updateThuHoiSignature();
});

document.addEventListener("paste", pastePlainText);

window.addRow = addRow;
window.deleteRow = deleteRow;
