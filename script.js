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

// 優先載入同資料夾的 jszip.js（可完全離線使用），不存在時退回 CDN
(function loadJsZipLib() {
  const local = document.createElement("script");
  local.src = "jszip.js";
  local.onerror = () => {
    const cdn = document.createElement("script");
    cdn.src = "https://cdn.jsdelivr.net/npm/jszip@3/dist/jszip.min.js";
    cdn.onerror = () => {
      console.error(
        "JSZip 函式庫載入失敗：本機找不到 jszip.js，CDN 也無法連線。",
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
const exportFormat = document.getElementById("exportFormat");
const saveMode = document.getElementById("saveMode");
const autoExportOnDrop = document.getElementById("autoExportOnDrop");
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
        exportFormat: exportFormat.value,
        saveMode: saveMode.value,
        autoExportOnDrop: autoExportOnDrop.checked,
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
    if (text.exportFormat) {
      exportFormat.value = text.exportFormat;
    } else if (typeof text.zipEnabled === "boolean") {
      // 舊版設定遷移：以前只有「同時輸出 ZIP」一個開關
      exportFormat.value = text.zipEnabled ? "zip" : "docx";
    }
    if (text.saveMode) saveMode.value = text.saveMode;
    autoExportOnDrop.checked = !!text.autoExportOnDrop;
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
  if (!files.length) return Promise.resolve();
  return Promise.all(
    files.map((file) =>
      resizeImage(file)
        .then((dataUrl) => ({
          name: file.name,
          dataUrl,
          rot: (Math.random() * 4 - 2).toFixed(2),
          // 只留在記憶體給 ZIP 匯出用原圖，不會被存進 localStorage
          // （重新整理頁面後就會遺失，ZIP 匯出時會退回用壓縮版本）
          file,
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

exportFormat.addEventListener("change", saveText);
saveMode.addEventListener("change", saveText);
autoExportOnDrop.addEventListener("change", saveText);

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

function clearSavedData() {
  localStorage.removeItem(TEXT_KEY);
  localStorage.removeItem(IMAGES_KEY);
  localStorage.removeItem(LEGACY_STORAGE_KEY);
  dateInput.value = todayIso();
  locationInput.value = "";
  floorInput.value = "";
  itemInput.value = "";
  exportFormat.value = "docx";
  saveMode.value = "prompt";
  images = [];
  renderAllThumbs();
  saveWarn.hidden = true;
  // 重新寫回被清掉的 TEXT_KEY，讓「拖入資料夾時自動輸出」這類非表單資料的
  // 偏好設定不會因為清除表單而跟著遺失
  saveText();
}

clearBtn.addEventListener("click", () => {
  if (!confirm("確定要清除已儲存的表單資料與照片嗎？")) return;
  clearSavedData();
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
// 樓層樣式：「B1樓」「2F」「3樓」「11MF」等，允許字母前綴（地下室 B）
// 及數字後接單一字母（如 6M、11M）再接「樓/F」結尾標記
const FLOOR_TOKEN_RE = /([A-Za-z]{0,2}\d{1,3}[A-Za-z]?)\s*(?:樓|楼|[Ff])/;
// 結尾直接是數字、沒加「樓」或「F」的樣式，例如「科技3」
const FLOOR_TRAILING_RE = /([A-Za-z]{0,2}\d{1,3})$/;
// 緊鄰地點文字的樓層片段：「B1」（字母前綴＋數字）或「6M」（數字＋單一字母後綴），
// 字母後綴只抓 1 個字，避免把後面接著的項目文字（如 EF-3）一起吃進來
const FLOOR_NEIGHBOR_RE = /^([A-Za-z]{1,2}\d{1,3}|\d{1,3}[A-Za-z]?)/;

// 地點別名／俗稱對照表：資料夾常用簡稱不是正式地點名稱的縮寫或子字串時，
// 在這裡加一筆「別名: 正式地點」即可辨識
const LOCATION_ALIASES = {
  醫科: "科技大樓",
};

// 判斷 namePart 是否對得到 loc 這個地點，並拆出地點文字前後剩下的部分
// （分開記錄前後段，而不是合併成一段文字，樓層才能準確定位在地點的哪一側）
function evaluateLocation(namePart, loc) {
  if (loc.includes(namePart)) {
    // namePart 是地點的縮寫，且沒有多打其他字（例如「中正」對「中正樓」）
    return { score: namePart.length, before: "", after: "" };
  }
  const idx = namePart.indexOf(loc);
  if (idx !== -1) {
    // 完整地點文字出現在 namePart 中間或結尾，前後可能還夾雜其他文字
    return {
      score: loc.length,
      before: namePart.slice(0, idx),
      after: namePart.slice(idx + loc.length),
    };
  }
  // 地點縮寫沒打完整，且後面緊接項目文字、中間沒有分隔符
  // （例如「中正2FAH-22皮帶更換」拿掉樓層後變成「中正AH-22皮帶更換」；
  // 「一門診3換軸承」的「一門診」則是「第一門診」拿掉開頭「第」字的縮寫，
  // 不是從頭對齊的前綴，所以要在 loc 裡任意位置找 namePart 的開頭片段）
  // 從最長開始試，找 namePart 開頭多少字是 loc 的子字串
  for (let k = namePart.length; k >= 2; k--) {
    const prefix = namePart.slice(0, k);
    if (loc.includes(prefix)) {
      return { score: k, before: "", after: namePart.slice(k) };
    }
  }
  return null;
}

// candidates：{ match: 用來比對的文字, canonical: 比對成功後要填入的正式地點名稱 }
function findBestLocation(candidates, namePart) {
  if (!namePart) return null;
  let best = null;
  for (const { match, canonical } of candidates) {
    const result = evaluateLocation(namePart, match);
    if (result && (!best || result.score > best.score)) {
      best = {
        location: canonical,
        before: result.before,
        after: result.after,
        score: result.score,
      };
    }
  }
  return best;
}

function matchLocationAndFloor(folderName) {
  const options = Array.from(
    document.querySelectorAll("#locationOptions option"),
  ).map((o) => o.value);
  const candidates = options
    .map((loc) => ({ match: loc, canonical: loc }))
    .concat(
      Object.entries(LOCATION_ALIASES).map(([alias, canonical]) => ({
        match: alias,
        canonical,
      })),
    );

  const trim = (s) => s.replace(/^[\s_\-]+|[\s_\-]+$/g, "");

  // 依序嘗試：明確樓層樣式 → 結尾裸數字 → 整串當地點（樓層另外用舊法推算）
  const attempts = [];
  const floorMatch = folderName.match(FLOOR_TOKEN_RE);
  if (floorMatch) {
    attempts.push({
      namePart: trim(
        folderName.slice(0, floorMatch.index) +
          folderName.slice(floorMatch.index + floorMatch[0].length),
      ),
      floor: floorMatch[1],
    });
  }
  const trailMatch = folderName.match(FLOOR_TRAILING_RE);
  if (trailMatch) {
    attempts.push({
      namePart: trim(folderName.slice(0, trailMatch.index)),
      floor: trailMatch[1],
    });
  }
  attempts.push({ namePart: folderName, floor: "" });

  for (const { namePart, floor } of attempts) {
    const found = findBestLocation(candidates, namePart);
    if (!found) continue;
    const best = found.location;

    let f = floor;
    let item = "";
    if (f) {
      // namePart 已經把樓層樣式拿掉，扣除地點文字後剩下的部分就是項目名稱
      item = trim(found.before + found.after);
    } else {
      // 沒抓到明確樓層樣式：取地點緊鄰前後的英數字當樓層，其餘當項目名稱
      const { before, after } = found;
      const afterStripped = after.replace(/^[\s_-]+/, "");
      const afterNum = afterStripped.match(FLOOR_NEIGHBOR_RE);
      const beforeNum = before.match(
        /([A-Za-z]{1,2}\d{1,3}|\d{1,3}[A-Za-z]?)[\s_-]*$/,
      );
      if (afterNum) {
        f = afterNum[1];
        item = trim(before + afterStripped.slice(afterNum[0].length));
      } else if (beforeNum) {
        f = beforeNum[1];
        item = trim(
          before.slice(0, before.length - beforeNum[0].length) + after,
        );
      } else {
        item = trim(before + after);
      }
    }
    if (f.length > 1 && /f$/i.test(f)) f = f.slice(0, -1);

    return { location: best, floor: f.slice(0, 4), item };
  }
  return null;
}

function readAllFileEntries(reader) {
  return new Promise((resolve, reject) => {
    let entries = [];
    const readBatch = () => {
      reader.readEntries((results) => {
        if (!results.length) {
          resolve(entries);
        } else {
          entries = entries.concat(results);
          readBatch();
        }
      }, reject);
    };
    readBatch();
  });
}

async function getFilesFromDirectoryEntry(dirEntry) {
  const entries = await readAllFileEntries(dirEntry.createReader());
  const files = [];
  for (const entry of entries) {
    if (entry.isFile) {
      files.push(
        await new Promise((resolve, reject) => entry.file(resolve, reject)),
      );
    } else if (entry.isDirectory) {
      files.push(...(await getFilesFromDirectoryEntry(entry)));
    }
  }
  return files;
}

dropzone.addEventListener("drop", async (e) => {
  const items = e.dataTransfer.items;
  const entries = items
    ? Array.from(items)
        .map((it) => (it.webkitGetAsEntry ? it.webkitGetAsEntry() : null))
        .filter(Boolean)
    : [];

  if (entries.some((en) => en.isDirectory)) {
    const autoExport = autoExportOnDrop.checked;

    // 每個拖進來的資料夾都是獨立一筆記錄：依序清空表單→比對地點樓層項目→
    // 加入該資料夾的照片→（開了自動輸出的話）輸出，再處理下一個資料夾
    for (const entry of entries) {
      if (entry.isDirectory) {
        clearSavedData();

        // 自動輸出時不能再跳出另存新檔視窗（拖曳後經過圖片壓縮等非同步處理，
        // 使用者手勢已經逾期，瀏覽器會擋掉跳窗），固定用 ZIP＋直接下載
        exportFormat.value = "zip";
        saveMode.value = "download";
        saveText();

        const match = matchLocationAndFloor(entry.name);
        if (match) {
          locationInput.value = match.location;
          floorInput.value = match.floor;
          if (match.item) itemInput.value = match.item;
          saveText();
        }
        await addFiles(await getFilesFromDirectoryEntry(entry));

        if (autoExport) {
          await runExport();
        }
      } else if (entry.isFile) {
        await addFiles([
          await new Promise((resolve, reject) => entry.file(resolve, reject)),
        ]);
      }
    }
    return;
  }

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

const SAVE_TYPES = {
  docx: {
    description: "Word 文件",
    mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ext: ".docx",
  },
  zip: {
    description: "ZIP 壓縮檔",
    mime: "application/zip",
    ext: ".zip",
  },
};

// 必須在點擊當下立刻呼叫，不能等圖片處理完才呼叫，
// 否則瀏覽器可能已經判定使用者手勢逾期，導致視窗不會跳出、直接靜默下載
async function pickSaveHandle(suggestedName, type) {
  if (!window.showSaveFilePicker) return null;
  try {
    return await window.showSaveFilePicker({
      suggestedName,
      types: [
        {
          description: type.description,
          accept: { [type.mime]: [type.ext] },
        },
      ],
    });
  } catch (err) {
    if (err.name === "AbortError") throw err; // 使用者取消存檔
    console.warn("另存新檔失敗，改用瀏覽器預設下載", err);
    return null;
  }
}

function uniqueFileName(usedNames, name) {
  let candidate = name;
  let n = 1;
  while (usedNames.has(candidate)) {
    const dot = name.lastIndexOf(".");
    const base = dot === -1 ? name : name.slice(0, dot);
    const ext = dot === -1 ? "" : name.slice(dot);
    candidate = `${base}_${n}${ext}`;
    n++;
  }
  usedNames.add(candidate);
  return candidate;
}

// 整理要匯出的照片清單：一律用「加入當下」的原始檔案（不壓縮、不改變），
// 只有頁面重新整理過、記憶體中已遺失原圖的項目，才退回用 Word 裡那份壓縮版本頂替
// 統一讀成 Uint8Array（而不是直接傳 Blob/File），避免不同環境對 Blob 支援不一致
async function collectExportFiles(reservedNames) {
  const usedNames = new Set(reservedNames);
  const fallbackNames = [];
  const files = [];
  for (let i = 0; i < images.length; i++) {
    const img = images[i];
    let data;
    let filename;
    if (img.file) {
      data = new Uint8Array(await img.file.arrayBuffer());
      filename = sanitizeFilename(
        img.file.name || img.name || `photo_${i + 1}`,
      );
    } else {
      data = await dataUrlToBytes(img.dataUrl);
      const rawName = sanitizeFilename(img.name || `photo_${i + 1}`);
      const withoutExt = rawName.replace(/\.[^.]+$/, "") || `photo_${i + 1}`;
      filename = `${withoutExt}.jpg`;
      fallbackNames.push(img.name || filename);
    }
    const name = uniqueFileName(usedNames, filename || `photo_${i + 1}`);
    files.push({ name, data });
  }
  return { files, fallbackNames };
}

function warnFallbackPhotos(fallbackNames) {
  if (!fallbackNames.length) return;
  alert(
    `以下照片因頁面曾重新整理、記憶體中已無原圖，改用 Word 裡的壓縮版本：\n${fallbackNames.join("\n")}`,
  );
}

// 打包成 zip：folderName/ 資料夾裡放 docx 與所有照片
// 檔案直接放在 zip 根目錄（不在裡面再包一層資料夾）：
// zip 本身檔名就是 folderName，解壓縮工具（Windows／macOS）預設就會
// 建立一個同名資料夾把這些檔案放進去，避免解壓縮後要多點兩層資料夾
async function buildZip({ docxBlob, docxFilename }) {
  const zip = new JSZip();
  zip.file(docxFilename, docxBlob);

  const { files, fallbackNames } = await collectExportFiles([docxFilename]);
  files.forEach((f) => zip.file(f.name, f.data));
  warnFallbackPhotos(fallbackNames);

  return zip.generateAsync({ type: "blob" });
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

async function runExport() {
  if (typeof docx === "undefined") {
    console.error(
      "docx 全域變數不存在——本機 docx.js 不在同資料夾，且 CDN 腳本可能被封鎖或 404。請確認 docx.js 與本 HTML 檔放在一起，或開啟開發者工具的 Network 分頁檢查請求狀態。",
    );
    alert(
      "Word 匯出元件尚未載入。請確認 docx.js 檔案與本頁面放在同一個資料夾（或連上網路後重新整理頁面）再試一次。",
    );
    return;
  }

  const format = exportFormat.value; // "docx" | "zip"
  if (format === "zip" && typeof JSZip === "undefined") {
    console.error(
      "JSZip 全域變數不存在——本機 jszip.js 不在同資料夾，且 CDN 腳本可能被封鎖或 404。請確認 jszip.js 與本 HTML 檔放在一起，或開啟開發者工具的 Network 分頁檢查請求狀態。",
    );
    alert(
      "ZIP 匯出元件尚未載入。請確認 jszip.js 檔案與本頁面放在同一個資料夾（或連上網路後重新整理頁面）再試一次。",
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
    const docxFilename = baseName + ".docx";

    if (format === "zip") {
      // zip 本身的檔名＝資料夾名稱（地點樓層＋項目，不含日期），解壓縮後
      // 才會剛好只長出一層同名資料夾，不會多包一層
      const folderName =
        sanitizeFilename([loc, item].filter(Boolean).join("_")) || "領料記錄表";
      const zipSuggestedName = folderName + ".zip";

      // 「直接下載」時不跳另存新檔視窗，交給 writeFile 退回瀏覽器預設下載
      let handle = null;
      if (saveMode.value === "prompt") {
        try {
          handle = await pickSaveHandle(zipSuggestedName, SAVE_TYPES.zip);
        } catch (err) {
          if (err.name === "AbortError") return; // 使用者取消存檔
          throw err;
        }
      }

      const docxBlob = await buildDocx({ item, dateStr, loc, grid, cols });
      const zipBlob = await buildZip({ docxBlob, docxFilename });
      await writeFile(handle, zipBlob, zipSuggestedName);
    } else {
      let handle = null;
      if (saveMode.value === "prompt") {
        try {
          handle = await pickSaveHandle(docxFilename, SAVE_TYPES.docx);
        } catch (err) {
          if (err.name === "AbortError") return; // 使用者取消存檔
          throw err;
        }
      }

      const blob = await buildDocx({ item, dateStr, loc, grid, cols });
      await writeFile(handle, blob, docxFilename);
    }
  } catch (err) {
    console.error("匯出失敗：", err);
    alert("匯出失敗，請再試一次。（詳細錯誤請見瀏覽器主控台）");
  } finally {
    exportBtn.disabled = false;
  }
}

exportBtn.addEventListener("click", runExport);
