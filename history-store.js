// Kho lưu lịch sử hồ sơ đã in.
// - Dùng IndexedDB thay cho localStorage (localStorage chỉ ~5MB, IndexedDB cho phép hàng trăm MB trở lên).
// - Nén HTML bằng gzip trước khi lưu, mỗi bản ghi giảm từ ~35KB xuống còn vài KB.
// - In lại đúng hồ sơ giống hệt bản gần nhất thì chỉ tăng số lần in, không tạo bản ghi mới.
const HistoryStore = (() => {
  const DB_NAME = "hoSoSuaChua";
  const DB_VERSION = 1;
  const STORE = "history";
  const LEGACY_STORAGE_KEY = "repairHistoryRaw";

  const canCompress =
    typeof CompressionStream === "function" &&
    typeof DecompressionStream === "function";

  let dbPromise = null;

  function openDb() {
    if (!dbPromise) {
      dbPromise = new Promise((resolve, reject) => {
        if (!window.indexedDB) {
          reject(new Error("Trình duyệt không hỗ trợ IndexedDB."));
          return;
        }

        const request = indexedDB.open(DB_NAME, DB_VERSION);

        request.onupgradeneeded = () => {
          const db = request.result;
          if (!db.objectStoreNames.contains(STORE)) {
            const store = db.createObjectStore(STORE, { keyPath: "id" });
            store.createIndex("createdAt", "createdAt");
          }
        };

        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    }

    return dbPromise;
  }

  // Chạy một thao tác trong transaction, trả về kết quả khi transaction hoàn tất.
  async function run(mode, action) {
    const db = await openDb();

    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, mode);
      let result;

      const request = action(tx.objectStore(STORE), (value) => {
        result = value;
      });

      tx.oncomplete = () => resolve(request ? request.result : result);
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  }

  async function compress(text) {
    if (!canCompress) return text;

    const stream = new Blob([text])
      .stream()
      .pipeThrough(new CompressionStream("gzip"));
    return new Response(stream).arrayBuffer();
  }

  async function decompress(data) {
    if (typeof data === "string") return data;

    const stream = new Blob([data])
      .stream()
      .pipeThrough(new DecompressionStream("gzip"));
    return new Response(stream).text();
  }

  // Băm FNV-1a 32-bit: đủ nhanh để so sánh hai bản HTML có giống nhau không.
  function hashText(text) {
    let hash = 0x811c9dc5;
    for (let i = 0; i < text.length; i++) {
      hash ^= text.charCodeAt(i);
      hash = Math.imul(hash, 0x01000193);
    }
    return (hash >>> 0).toString(16);
  }

  function createId() {
    return `HS${Date.now()}${Math.random().toString(36).slice(2, 6)}`;
  }

  async function toStoredRecord(record) {
    const html = String(record.html || "");

    return {
      id: record.id || createId(),
      createdAt: record.createdAt || new Date().toISOString(),
      printedAt: record.printedAt || record.createdAt || new Date().toISOString(),
      printCount: record.printCount || 1,
      meta: record.meta || {},
      hash: hashText(html),
      size: html.length,
      html: await compress(html),
    };
  }

  // Thêm bản ghi mới. Nếu trùng nội dung với bản gần nhất thì chỉ cập nhật số lần in.
  async function add({ meta, html }) {
    const record = await toStoredRecord({ meta, html });

    return run("readwrite", (store, setResult) => {
      const cursorRequest = store.index("createdAt").openCursor(null, "prev");

      cursorRequest.onsuccess = () => {
        const cursor = cursorRequest.result;
        const latest = cursor?.value;

        if (latest && latest.hash === record.hash) {
          cursor.update({
            ...latest,
            printedAt: record.printedAt,
            printCount: (latest.printCount || 1) + 1,
          });
          setResult({ duplicated: true, id: latest.id });
          return;
        }

        store.put(record);
        setResult({ duplicated: false, id: record.id });
      };
    });
  }

  // Lấy danh sách (mới nhất trước), không giải nén HTML để hiển thị nhanh.
  function list() {
    return run("readonly", (store, setResult) => {
      const items = [];
      const cursorRequest = store.index("createdAt").openCursor(null, "prev");

      cursorRequest.onsuccess = () => {
        const cursor = cursorRequest.result;
        if (!cursor) {
          setResult(items);
          return;
        }

        const { html, ...summary } = cursor.value;
        items.push(summary);
        cursor.continue();
      };
    });
  }

  async function getHtml(id) {
    const record = await run("readonly", (store) => store.get(id));
    return record ? decompress(record.html) : null;
  }

  function remove(id) {
    return run("readwrite", (store) => store.delete(id));
  }

  function clear() {
    return run("readwrite", (store) => store.clear());
  }

  // Xuất toàn bộ ra dạng JSON thường (HTML đã giải nén) để sao lưu hoặc chuyển máy.
  async function exportAll() {
    const records = await run("readonly", (store) => store.getAll());
    records.sort((a, b) => b.createdAt.localeCompare(a.createdAt));

    return Promise.all(
      records.map(async ({ hash, size, html, ...rest }) => ({
        ...rest,
        html: await decompress(html),
      })),
    );
  }

  // Nhập từ file sao lưu (hỗ trợ cả định dạng cũ của localStorage). Bỏ qua bản ghi đã có.
  async function importAll(records) {
    const valid = (Array.isArray(records) ? records : []).filter(
      (item) => item && typeof item.html === "string",
    );
    // Bản ghi cũ còn thẻ <script>, khi xem lại A4.js sẽ điền đè dữ liệu sai nên bỏ đi
    const prepared = await Promise.all(
      valid.map((item) =>
        toStoredRecord({
          ...item,
          html: item.html.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, ""),
        }),
      ),
    );
    const existingIds = new Set(
      await run("readonly", (store) => store.getAllKeys()),
    );
    const fresh = prepared.filter((item) => !existingIds.has(item.id));

    if (fresh.length) {
      await run("readwrite", (store) => {
        fresh.forEach((item) => store.put(item));
      });
    }

    return fresh.length;
  }

  // Chuyển dữ liệu lịch sử cũ từ localStorage sang IndexedDB (chạy một lần).
  async function migrateLegacy() {
    const raw = localStorage.getItem(LEGACY_STORAGE_KEY);
    if (!raw) return 0;

    let records;
    try {
      records = JSON.parse(raw);
    } catch (error) {
      console.error("Lịch sử cũ trong localStorage bị lỗi, bỏ qua.", error);
      return 0;
    }

    const imported = await importAll(records);
    localStorage.removeItem(LEGACY_STORAGE_KEY);
    return imported;
  }

  async function estimateUsage() {
    if (!navigator.storage?.estimate) return null;
    return navigator.storage.estimate();
  }

  // Xin trình duyệt không tự xoá dữ liệu khi ổ đĩa sắp đầy.
  function requestPersistence() {
    navigator.storage?.persist?.().catch(() => {});
  }

  return {
    add,
    list,
    getHtml,
    remove,
    clear,
    exportAll,
    importAll,
    migrateLegacy,
    estimateUsage,
    requestPersistence,
  };
})();
