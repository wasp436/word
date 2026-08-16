// 優先載入同資料夾的 docx.js（可完全離線使用），不存在時退回 CDN
(function loadDocxLib() {
  const local = document.createElement("script");
  local.src = "docx.js";
  local.onerror = () => {
    const cdn = document.createElement("script");
    cdn.src = "https://cdn.jsdelivr.net/npm/docx@8/build/index.umd.js";
    cdn.onerror = () => {
      console.error(
        "docx 函式庫載入失敗：本機找不到 docx.js，CDN 也無法連線。",
      );
    };
    document.head.appendChild(cdn);
  };
  document.head.appendChild(local);
})();

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const cameraBtn = document.getElementById("cameraBtn");
const cameraInput = document.getElementById("cameraInput");
const thumbs = document.getElementById("thumbs");
const emptyNote = document.getElementById("emptyNote");
const exportBtn = document.getElementById("exportBtn");
const saveWarn = document.getElementById("saveWarn");
const clearBtn = document.getElementById("clearBtn");
const dateInput = document.getElementById("date");
const locationInput = document.getElementById("location");
const floorInput = document.getElementById("floor");
const itemInput = document.getElementById("item");

function locationWithFloor() {
  const loc = locationInput.value.trim();
  const floor = floorInput.value.trim();
  return floor ? `${loc}${floor}F` : loc;
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function todayIso() {
  const now = new Date();
  return `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())}`;
}

function formatDate(isoDate) {
  if (!isoDate) return "";
  const [y, m, d] = isoDate.split("-").map(Number);
  return `${y}/${pad2(m)}/${pad2(d)}`;
}

function formatDateCompact(isoDate) {
  if (!isoDate) return "";
  const [y, m, d] = isoDate.split("-").map(Number);
  return `${y}${pad2(m)}${pad2(d)}`;
}

let images = [];

const PHOTO_RATIO = 4 / 3;
const PHOTO_MAX_W = 1200;
const LEGACY_STORAGE_KEY = "siteRecordFormV1"; // 舊版單一儲存鍵，僅供遷移
const TEXT_KEY = "siteRecordTextV2";
const IMAGES_KEY = "siteRecordImagesV2";

function refreshEmptyNote() {
  emptyNote.style.display = images.length ? "none" : "block";
}

// 文字與照片分開儲存：打字時只寫入文字，避免每個字元都重新序列化全部照片
function saveText() {
  try {
    localStorage.setItem(
      TEXT_KEY,
      JSON.stringify({
        date: dateInput.value,
        location: locationInput.value,
        floor: floorInput.value,
        item: itemInput.value,
      }),
    );
  } catch (err) {
    console.warn("無法儲存文字欄位到 localStorage", err);
  }
}

function saveImages() {
  try {
    localStorage.setItem(
      IMAGES_KEY,
      JSON.stringify(
        images.map((i) => ({
          name: i.name,
          dataUrl: i.dataUrl,
          rot: i.rot,
        })),
      ),
    );
    saveWarn.hidden = true;
  } catch (err) {
    console.warn("無法儲存到 localStorage（可能空間不足）", err);
    saveWarn.hidden = false;
  }
}

function loadState() {
  const parse = (key) => {
    try {
      return JSON.parse(localStorage.getItem(key));
    } catch (err) {
      return null;
    }
  };
  let text = parse(TEXT_KEY);
  let imgs = parse(IMAGES_KEY);
  // 從舊版單一鍵遷移
  if (!text && !imgs) {
    const legacy = parse(LEGACY_STORAGE_KEY);
    if (legacy) {
      text = legacy;
      imgs = legacy.images;
      localStorage.removeItem(LEGACY_STORAGE_KEY);
    }
  }
  if (text) {
    if (text.date) dateInput.value = text.date;
    locationInput.value = text.location || "";
    floorInput.value = text.floor || "";
    itemInput.value = text.item || "";
  }
  (imgs || []).forEach((data) => {
    images.push({
      name: data.name,
      dataUrl: data.dataUrl,
      rot: data.rot || "0",
    });
  });
  renderAllThumbs();
}

function resizeImage(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      const outW = PHOTO_MAX_W;
      const outH = outW / PHOTO_RATIO;
      const canvas = document.createElement("canvas");
      canvas.width = outW;
      canvas.height = outH;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0, outW, outH);
      URL.revokeObjectURL(url);
      resolve(canvas.toDataURL("image/jpeg", 0.9));
    };
    img.onerror = reject;
    img.src = url;
  });
}

function addFiles(fileList) {
  const files = Array.from(fileList).filter((f) => f.type.startsWith("image/"));
  if (!files.length) return;
  Promise.all(
    files.map((file) =>
      resizeImage(file)
        .then((dataUrl) => ({
          name: file.name,
          dataUrl,
          rot: (Math.random() * 4 - 2).toFixed(2),
        }))
        .catch(() => null),
    ),
  ).then((items) => {
    const failed = files.filter((_, i) => !items[i]).map((f) => f.name);
    items.filter(Boolean).forEach((item) => {
      images.push(item);
    });
    renderAllThumbs();
    saveImages();
    if (failed.length) {
      alert(`以下檔案無法讀取，已略過：\n${failed.join("\n")}`);
    }
  });
}

let dragSrcIndex = null;

function moveImage(from, to) {
  if (to < 0 || to >= images.length || from === to) return;
  const [moved] = images.splice(from, 1);
  images.splice(to, 0, moved);
  renderAllThumbs();
  saveImages();
}

function renderAllThumbs() {
  thumbs.innerHTML = "";
  images.forEach((item, idx) => {
    thumbs.appendChild(createThumb(item, idx));
  });
  refreshEmptyNote();
}

function createThumb(item, idx) {
  const el = document.createElement("div");
  el.className = "thumb";
  el.style.setProperty("--r", item.rot + "deg");
  el.draggable = true;
  el.innerHTML = `
      <img alt="" draggable="false">
      <div class="cap"><\/div>
      <button type="button" class="mv mv-l" title="往前移">&lsaquo;<\/button>
      <button type="button" class="mv mv-r" title="往後移">&rsaquo;<\/button>
      <div class="rm" title="移除">&times;<\/div>
    `;
  el.querySelector("img").src = item.dataUrl;
  el.querySelector(".cap").textContent = item.name;
  el.querySelector(".rm").addEventListener("click", () => {
    images.splice(idx, 1);
    renderAllThumbs();
    saveImages();
  });
  el.querySelector(".mv-l").addEventListener("click", () => {
    moveImage(idx, idx - 1);
  });
  el.querySelector(".mv-r").addEventListener("click", () => {
    moveImage(idx, idx + 1);
  });
  el.addEventListener("dragstart", (e) => {
    dragSrcIndex = idx;
    el.classList.add("dragging");
    e.dataTransfer.effectAllowed = "move";
  });
  el.addEventListener("dragend", () => {
    dragSrcIndex = null;
    el.classList.remove("dragging");
  });
  el.addEventListener("dragover", (e) => {
    if (dragSrcIndex === null) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
  });
  el.addEventListener("drop", (e) => {
    if (dragSrcIndex === null) return;
    e.preventDefault();
    e.stopPropagation();
    moveImage(dragSrcIndex, idx);
  });
  return el;
}

dropzone.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", (e) => {
  addFiles(e.target.files);
  fileInput.value = "";
});

cameraBtn.addEventListener("click", (e) => {
  e.stopPropagation();
  cameraInput.click();
});
cameraInput.addEventListener("change", (e) => {
  addFiles(e.target.files);
  cameraInput.value = "";
});

[dateInput, locationInput, itemInput].forEach((el) => {
  el.addEventListener("input", saveText);
});

floorInput.addEventListener("input", () => {
  const raw = floorInput.value;
  const filtered = raw.replace(/[^a-zA-Z0-9]/g, "");
  // 只有出現非法字元才重寫 value，並把游標放回原本位置
  if (filtered !== raw) {
    const pos = floorInput.selectionStart;
    const posAfter = raw.slice(0, pos).replace(/[^a-zA-Z0-9]/g, "").length;
    floorInput.value = filtered;
    floorInput.setSelectionRange(posAfter, posAfter);
  }
  saveText();
});

clearBtn.addEventListener("click", () => {
  if (!confirm("確定要清除已儲存的表單資料與照片嗎？")) return;
  localStorage.removeItem(TEXT_KEY);
  localStorage.removeItem(IMAGES_KEY);
  localStorage.removeItem(LEGACY_STORAGE_KEY);
  dateInput.value = todayIso();
  locationInput.value = "";
  floorInput.value = "";
  itemInput.value = "";
  images = [];
  renderAllThumbs();
  saveWarn.hidden = true;
});

dateInput.value = todayIso();
loadState();

["dragenter", "dragover"].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.add("dragover");
  });
});
["dragleave", "drop"].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.remove("dragover");
  });
});
dropzone.addEventListener("drop", (e) => {
  if (e.dataTransfer.files && e.dataTransfer.files.length) {
    addFiles(e.dataTransfer.files);
  }
});

function sanitizeFilename(name) {
  return name.replace(/[\\/:*?"<>|]/g, "").trim();
}

function computeGrid() {
  // 固定兩欄、固定尺寸，不再壓縮成單頁；超出的列由 Word 自動換頁
  const w = 7.6; // cm
  return { cols: 2, w, h: Number(((w * 3) / 4).toFixed(2)) };
}

async function dataUrlToBytes(dataUrl) {
  const res = await fetch(dataUrl);
  return new Uint8Array(await res.arrayBuffer());
}

async function buildDocx({ item, dateStr, loc, grid, cols }) {
  const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    ImageRun,
    AlignmentType,
    WidthType,
    BorderStyle,
  } = docx;

  const line = { style: BorderStyle.SINGLE, size: 6, color: "808080" };
  const cellBorders = {
    top: line,
    bottom: line,
    left: line,
    right: line,
  };
  // 儲存格內距（twip），讓圖片與框線、相鄰圖片之間留出間隔
  const cellMargins = { top: 120, bottom: 120, left: 120, right: 120 };
  const pxW = Math.round(grid.w * 37.8);
  const pxH = Math.round(grid.h * 37.8);

  const allBytes = await Promise.all(
    images.map((img) => dataUrlToBytes(img.dataUrl)),
  );

  const rows = [];
  for (let i = 0; i < allBytes.length; i += cols) {
    rows.push(allBytes.slice(i, i + cols));
  }

  const tableRows = [];
  for (const row of rows) {
    const cells = [];
    for (const bytes of row) {
      cells.push(
        new TableCell({
          width: {
            size: Math.round(100 / cols),
            type: WidthType.PERCENTAGE,
          },
          borders: cellBorders,
          margins: cellMargins,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new ImageRun({
                  data: bytes,
                  type: "jpg",
                  transformation: { width: pxW, height: pxH },
                }),
              ],
            }),
          ],
        }),
      );
    }
    tableRows.push(new TableRow({ children: cells, cantSplit: true }));
  }

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 160 },
      children: [
        new TextRun({
          text: item || "領料記錄表",
          bold: true,
          size: 32,
          font: "微軟正黑體",
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 360 },
      border: {
        bottom: {
          style: BorderStyle.SINGLE,
          size: 6,
          color: "CCCCCC",
          space: 8,
        },
      },
      children: [
        new TextRun({
          text: `日期：${dateStr || "—"}  地點：${loc || "—"}`,
          size: 22,
          color: "666666",
          font: "微軟正黑體",
        }),
      ],
    }),
  ];

  if (tableRows.length) {
    children.push(
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top: line,
          bottom: line,
          left: line,
          right: line,
          insideHorizontal: line,
          insideVertical: line,
        },
        rows: tableRows,
      }),
    );
  } else {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "（無照片）", color: "999999" })],
      }),
    );
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1247, bottom: 1247, left: 1247, right: 1247 },
          },
        },
        children,
      },
    ],
  });

  return Packer.toBlob(doc);
}

// 必須在點擊當下立刻呼叫，不能等圖片處理完才呼叫，
// 否則瀏覽器可能已經判定使用者手勢逾期，導致視窗不會跳出、直接靜默下載
async function pickSaveHandle(suggestedName) {
  if (!window.showSaveFilePicker) return null;
  try {
    return await window.showSaveFilePicker({
      suggestedName,
      types: [
        {
          description: "Word 文件",
          accept: {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
              [".docx"],
          },
        },
      ],
    });
  } catch (err) {
    if (err.name === "AbortError") throw err; // 使用者取消存檔
    console.warn("另存新檔失敗，改用瀏覽器預設下載", err);
    return null;
  }
}

async function writeFile(handle, blob, suggestedName) {
  if (handle) {
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
    return;
  }
  // 不支援 File System Access API 時（如 Firefox），退回傳統下載方式
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = suggestedName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

exportBtn.addEventListener("click", async () => {
  if (typeof docx === "undefined") {
    console.error(
      "docx 全域變數不存在——本機 docx.js 不在同資料夾，且 CDN 腳本可能被封鎖或 404。請確認 docx.js 與本 HTML 檔放在一起，或開啟開發者工具的 Network 分頁檢查請求狀態。",
    );
    alert(
      "Word 匯出元件尚未載入。請確認 docx.js 檔案與本頁面放在同一個資料夾（或連上網路後重新整理頁面）再試一次。",
    );
    return;
  }

  exportBtn.disabled = true;
  try {
    const loc = locationWithFloor();
    const item = itemInput.value.trim();
    const dateStr = formatDate(dateInput.value);
    const dateCompact = formatDateCompact(dateInput.value);
    const grid = computeGrid();
    const cols = grid.cols;
    const baseName =
      sanitizeFilename([dateCompact, loc, item].filter(Boolean).join("_")) ||
      "領料記錄表";
    const suggestedName = baseName + ".docx";

    let handle;
    try {
      handle = await pickSaveHandle(suggestedName);
    } catch (err) {
      if (err.name === "AbortError") return; // 使用者取消存檔
      throw err;
    }

    const blob = await buildDocx({ item, dateStr, loc, grid, cols });
    await writeFile(handle, blob, suggestedName);
  } catch (err) {
    console.error("Word 匯出失敗：", err);
    alert("Word 匯出失敗，請再試一次。（詳細錯誤請見瀏覽器主控台）");
  } finally {
    exportBtn.disabled = false;
  }
});
