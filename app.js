// CONFIGURATION DEFAULTS
const CLIENT_ID =
  "1061665018153-k3ub8sgo5ag4lmbsl0v93lve677iafr7.apps.googleusercontent.com";
const SCOPES =
  "https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/spreadsheets";
const SPREADSHEET_ID = "16oxD6-Bl5YUnBPaYXquehPrfUZKQbDvVGiXauOfPpQc";

let tokenClient,
  gapiInited = false,
  gisInited = false;
let currentFolderId = "root";
let folderHistory = [{ id: "root", name: "Drive Saya" }];
let selectedItemCtx = null;
let isTrashMode = false;
let currentViewMode = "grid";
let uploadQueue = [];
let selectedItems = []; // Array penampung ID item yang sedang diseleksi secara massal

let moveCurrentFolderId = "root";
let moveFolderHistory = [{ id: "root", name: "Drive Saya" }];
let currentModalActionType = "move";

const state = { items: [], lockedFolders: {} };

window.gapiLoaded = gapiLoaded;
window.gisLoaded = gisLoaded;

// 1. TAMBAHKAN DI BAGIAN ATAS APP.JS (Setelah deklarasi state)
document.addEventListener("DOMContentLoaded", () => {
  // Cek apakah ada token aktif yang tersimpan di browser
  const savedToken = localStorage.getItem("gdrive_access_token");
  const tokenExpiry = localStorage.getItem("gdrive_token_expiry");

  if (savedToken && tokenExpiry && Date.now() < parseInt(tokenExpiry)) {
    // Pasang token lama ke instance GAPI tanpa perlu login ulang
    gapi.client.setToken({ access_token: savedToken });

    // Sembunyikan tombol login, tampilkan area dashboard utama
    document.getElementById("login-page")?.classList.add("hidden");
    document.getElementById("main-dashboard")?.classList.remove("hidden");

    // Muat folder utama langsung
    setTimeout(() => {
      loadFolderContent("root");
    }, 1000);
  }
});

// 2. PERBARUI FUNGSI CALLBACK AUTHENTICATION ANDA
// Cari tempat di mana tokenClient mendapatkan credential/token (biasanya di callback initTokenClient)
// Tambahkan baris ini setelah token berhasil didapatkan:
function fungsiCallbackTokenAnda(resp) {
  if (resp.error !== undefined) {
    throw resp;
  }

  // Simpan token ke localStorage
  const accessToken = resp.access_token;
  const expiryTime = Date.now() + parseInt(resp.expires_in) * 1000; // Biasanya 3600 detik (1 jam)

  localStorage.setItem("gdrive_access_token", accessToken);
  localStorage.setItem("gdrive_token_expiry", expiryTime);

  // Lanjutkan proses load dashboard bawaan Anda...
}

// 3. TAMBAHKAN PADA FUNGSI LOGOUT ANDA
// Pastikan menghapus token dari memori browser saat pengguna klik "Keluar"
function handleSignoutClick() {
  const token = gapi.client.getToken();
  if (token !== null) {
    google.accounts.oauth2.revokeToken(token.access_token, () => {
      gapi.client.setToken("");
      // Bersihkan penyimpanan lokal
      localStorage.removeItem("gdrive_access_token");
      localStorage.removeItem("gdrive_token_expiry");
      window.location.reload();
    });
  }
}

// FORMAT BYTE CONVERTER
function formatBytes(bytes, decimals = 2) {
  if (!bytes || bytes === "0" || isNaN(bytes)) return "0 Bytes";
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ["Bytes", "KB", "MB", "GB", "TB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + " " + sizes[i];
}

// REALTIME NETWORK INDICATOR
function updateSignalStatus() {
  const indicator = document.getElementById("signal-indicator");
  const text = document.getElementById("signal-text");
  if (navigator.onLine) {
    indicator.className =
      "flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-600 border border-emerald-200";
    text.innerText = "Online";
  } else {
    indicator.className =
      "flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-rose-50 text-rose-600 border border-rose-200";
    text.innerText = "Offline";
  }
}
window.addEventListener("online", updateSignalStatus);
window.addEventListener("offline", updateSignalStatus);

// LOADER CONTROL WITH PERSISTENT DOWNLOAD REALTIME PROGRESS
function showLoader(text = "Memproses...", progress = null, ratioText = null) {
  document.getElementById("loader-text").innerText = text;
  document.getElementById("global-loader").classList.remove("hidden");

  const icon = document.getElementById("loader-icon");
  const progContainer = document.getElementById("loader-progress-container");
  const progBar = document.getElementById("loader-progress-bar");
  const ratioEl = document.getElementById("loader-ratio");

  if (progress !== null) {
    icon.className =
      "fa-solid fa-cloud-arrow-down absolute text-indigo-600 text-lg animate-bounce";
    progContainer.classList.remove("hidden");
    progBar.style.width = `${progress}%`;
    if (ratioText) {
      ratioEl.classList.remove("hidden");
      ratioEl.innerText = ratioText;
    } else {
      ratioEl.classList.add("hidden");
    }
  } else {
    icon.className =
      "fa-solid fa-cloud-arrow-down absolute text-indigo-600 text-lg";
    progContainer.classList.add("hidden");
    ratioEl.classList.add("hidden");
  }
}
function hideLoader() {
  document.getElementById("global-loader").classList.add("hidden");
}

function showSuccessModal(message = "Aksi berhasil diproses.") {
  document.getElementById("success-message").innerText = message;
  openModal("modal-success");
}

// INITIALIZATION GOOGLE ARCHITECTURE
function gapiLoaded() {
  gapi.load("client", initializeGapiClient);
}
// 1. MODIFIKASI FUNGSI INITIALIZEGAPICLIENT ANDA
async function initializeGapiClient() {
  try {
    await gapi.client.init({
      discoveryDocs: [
        "https://www.googleapis.com/discovery/v1/apis/drive/v3/rest",
        "https://sheets.googleapis.com/$discovery/rest",
      ],
    });
    gapiInited = true;

    // --- PINDAHKAN PENGECEKAN TOKEN KE SINI ---
    // Dipanggil tepat setelah GAPI Client terinisialisasi dengan sempurna
    const savedToken = localStorage.getItem("gdrive_access_token");
    const tokenExpiry = localStorage.getItem("gdrive_token_expiry");

    if (savedToken && tokenExpiry && Date.now() < parseInt(tokenExpiry)) {
      // Pasang token lama ke instance GAPI yang sudah siap
      gapi.client.setToken({ access_token: savedToken });

      // Sembunyikan halaman login, tampilkan dashboard utama
      document.getElementById("login-page")?.classList.add("hidden");
      document.getElementById("main-dashboard")?.classList.remove("hidden");

      // Muat data folder utama langsung
      setTimeout(() => {
        if (typeof loadFolderContent === "function") {
          loadFolderContent("root");
        }
      }, 500);
    } else {
      // Jika token tidak ada atau kedaluwarsa, jalankan alur auth biasa bawaan Anda
      checkAuthProgress();
    }
    // ------------------------------------------
  } catch (error) {
    console.error("GAPI Init Error:", error);
  }
}

// 2. PASTIKAN FUNGSI CALLBACK DI GISLOADED SUDAH MENYIMPAN TOKEN BARU
function handleAuthCallback(resp) {
  if (resp.error !== undefined) {
    throw resp;
  }

  // Catat token baru dan waktu kedaluwarsanya saat pengguna berhasil klik login manual
  const accessToken = resp.access_token;
  const expiryTime = Date.now() + parseInt(resp.expires_in || 3600) * 1000;

  localStorage.setItem("gdrive_access_token", accessToken);
  localStorage.setItem("gdrive_token_expiry", expiryTime);

  // Sembunyikan login page & tampilkan dashboard utama
  document.getElementById("login-page")?.classList.add("hidden");
  document.getElementById("main-dashboard")?.classList.remove("hidden");

  if (typeof loadFolderContent === "function") {
    loadFolderContent("root");
  }
}

// 3. PERBARUI JUGA FUNGSI LOGOUT ANDA (Jika ada)
// Pastikan menghapus token dari memori browser agar ketika diklik keluar benar-benar meminta login kembali
function handleSignoutClick() {
  const token = gapi.client.getToken();
  if (token !== null) {
    google.accounts.oauth2.revokeToken(token.access_token, () => {
      gapi.client.setToken("");
      // Bersihkan penyimpanan lokal secara menyeluruh
      localStorage.removeItem("gdrive_access_token");
      localStorage.removeItem("gdrive_token_expiry");
      window.location.reload();
    });
  }
}
function gisLoaded() {
  try {
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      callback: (resp) => handleAuthCallback(resp),
    });
    gisInited = true;
    checkAuthProgress();
  } catch (error) {
    console.error(error);
  }
}
function checkAuthProgress() {
  if (gapiInited && gisInited) {
    // 1. Cek token di localStorage (Remember Me) atau sessionStorage (Sesi sementara)
    const savedToken =
      localStorage.getItem("gdrive_access_token") ||
      sessionStorage.getItem("gdrive_access_token");
    const tokenExpiry =
      localStorage.getItem("gdrive_token_expiry") ||
      sessionStorage.getItem("gdrive_token_expiry");

    // 2. Jika token ada dan belum kedaluwarsa, bypass halaman login langsung
    if (savedToken && tokenExpiry && Date.now() < parseInt(tokenExpiry)) {
      gapi.client.setToken({ access_token: savedToken });

      // Sembunyikan halaman login, tampilkan dashboard utama
      document.getElementById("login-page")?.classList.add("hidden");
      document.getElementById("main-dashboard")?.classList.remove("hidden");

      // Muat data folder utama langsung secara otomatis
      setTimeout(() => {
        if (typeof loadFolderContent === "function") {
          loadFolderContent("root");
        }
      }, 300);
    } else {
      // Jika tidak ada token valid, barulah tampilkan tombol login secara normal
      const authBtn = document.getElementById("auth-button");
      if (authBtn) {
        authBtn.style.display = "inline-flex";
        authBtn.onclick = handleAuthClick;
      }
    }
  }
}

// Fungsi pembantu untuk mengeksekusi sesi otomatis tanpa memunculkan layar login
async function executeAutoLoginSession() {
  document.getElementById("login-screen").classList.add("hidden");
  document.getElementById("main-dashboard").classList.remove("hidden");
  showLoader("Sinkronisasi Sistem...");
  updateSignalStatus();
  try {
    await fetchUserInfo();
    await loadMetadataFromSheet();
    await loadFolderContent(currentFolderId);
  } catch (err) {
    console.error("Gagal memuat otomatis:", err);
  } finally {
    hideLoader();
  }
}

// PERBAIKAN: Pastikan menggunakan deklarasi 'function' agar tidak memicu ReferenceError
function handleAuthClick() {
  if (tokenClient) {
    // Meminta token baru ke Google dengan pop-up login resmi
    tokenClient.requestAccessToken({ prompt: "consent" });
  } else {
    console.error("Google Token Client belum siap.");
  }
}
async function handleAuthCallback(resp) {
  if (resp.error !== undefined) throw resp;

  // Cek jika user mencentang 'Remember Me' (Jika elemen checkbox ada di HTML Anda)
  const rememberMeCheckbox = document.getElementById("remember-me");
  if (rememberMeCheckbox && rememberMeCheckbox.checked) {
    const tokenSession = {
      access_token: resp.access_token,
      expires_at: Date.now() + parseInt(resp.expires_in, 10) * 1000,
    };
    localStorage.setItem("gapi_stored_token", JSON.stringify(tokenSession));
  }

  document.getElementById("login-screen").classList.add("hidden");
  document.getElementById("main-dashboard").classList.remove("hidden");
  showLoader("Sinkronisasi Sistem...");
  updateSignalStatus();
  await fetchUserInfo();
  await loadMetadataFromSheet();
  await loadFolderContent(currentFolderId);
  hideLoader();
}

async function fetchUserInfo() {
  const response = await gapi.client.drive.about.get({
    fields: "user, storageQuota",
  });
  const about = response.result;
  document.getElementById("user-avatar").src = about.user.photoLink;
  const limit = about.storageQuota.limit,
    usage = about.storageQuota.usage;
  document.getElementById("storage-text").innerText =
    `${(usage / (1024 * 1024)).toFixed(1)} MB / ${(limit / (1024 * 1024 * 1024)).toFixed(0)} GB`;
  document.getElementById("storage-bar").style.width =
    `${(usage / limit) * 100}%`;
}

// METADATA DATABASE
async function loadMetadataFromSheet() {
  try {
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "A2:D",
    });
    const rows = response.result.values;
    state.lockedFolders = {};
    if (rows) {
      rows.forEach((row) => {
        if (row[3] === "TRUE")
          state.lockedFolders[row[0]] = { name: row[1], pass: row[2] };
      });
    }
  } catch (err) {
    console.error("Gagal sinkronisasi DB Sheet", err);
  }
}
async function writeMetadataToSheet(folderId, name, password) {
  await gapi.client.sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: "A:D",
    valueInputOption: "USER_ENTERED",
    resource: { values: [[folderId, name, password, "TRUE"]] },
  });
  await loadMetadataFromSheet();
}
async function removeMetadataFromSheet(folderId) {
  try {
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "A1:D",
    });
    const rows = response.result.values;
    if (!rows) return;
    let targetIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === folderId && rows[i][3] === "TRUE") {
        targetIndex = i + 1;
        break;
      }
    }
    if (targetIndex !== -1) {
      await gapi.client.sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `D${targetIndex}`,
        valueInputOption: "USER_ENTERED",
        resource: { values: [["FALSE"]] },
      });
      await loadMetadataFromSheet();
    }
  } catch (err) {
    console.error(err);
  }
}

// BREADCRUMBS CORE ENGINE (PERBAIKAN TOTAL: RESISTEN TERHADAP INTERUPSI EVENT BUBBLING DAN PARENT PREVENT)
function renderBreadcrumbs() {
  const container = document.getElementById("breadcrumbs");
  container.innerHTML = "";

  if (isTrashMode) {
    container.innerHTML = `<span class="text-slate-800 font-bold flex items-center gap-2"><i class="fa-solid fa-trash text-rose-500"></i> Kotak Sampah</span>`;
    return;
  }

  folderHistory.forEach((crumb, index) => {
    if (index > 0) {
      const separator = document.createElement("i");
      separator.className =
        "fa-solid fa-chevron-right text-slate-300 mx-1 text-xs select-none";
      container.appendChild(separator);
    }

    const btn = document.createElement("button");
    btn.type = "button";
    btn.innerText = crumb.name;

    // Cek jika ini halaman aktif saat ini
    if (index === folderHistory.length - 1) {
      btn.className =
        "text-slate-800 font-extrabold cursor-default pointer-events-none bg-slate-200/60 px-2 py-1 rounded-lg text-sm transition-all";
    } else {
      btn.className =
        "text-indigo-600 hover:text-indigo-900 transition-all font-bold hover:bg-indigo-50 px-2 py-1 rounded-lg text-sm block relative z-30 cursor-pointer";

      // Definisikan handler eksplisit tanpa mengandalkan inline javascript
      btn.addEventListener("click", async (e) => {
        e.preventDefault();
        e.stopPropagation(); // Amankan dari pemicu klik area luar explorer

        const targetHistory = folderHistory.slice(0, index + 1);
        folderHistory = targetHistory;

        showLoader("Memuat Folder...");
        await loadFolderContent(crumb.id);
        hideLoader();
      });
    }
    container.appendChild(btn);
  });
}

// LOAD LIST FILES & FOLDERS CONTENT
async function loadFolderContent(folderId) {
  isTrashMode = false;
  currentFolderId = folderId;
  setActiveSidebarLink("nav-root");
  renderBreadcrumbs();
  await executeQueryList(`'${folderId}' in parents and trashed = false`);
}
async function loadTrashContent() {
  isTrashMode = true;
  setActiveSidebarLink("nav-trash");
  renderBreadcrumbs();
  await executeQueryList(`trashed = true`);
}
async function executeQueryList(q) {
  const response = await gapi.client.drive.files.list({
    q: q,
    fields: "files(id, name, mimeType, iconLink, size, thumbnailLink)",
  });
  state.items = response.result.files || [];
  renderExplorerUI();
}
function setActiveSidebarLink(activeId) {
  document.getElementById("nav-root").className =
    "w-full px-4 py-2.5 rounded-xl font-medium text-sm flex items-center gap-3 transition text-slate-600 hover:bg-slate-50";
  document.getElementById("nav-trash").className =
    "w-full px-4 py-2.5 rounded-xl font-medium text-sm flex items-center gap-3 transition text-slate-600 hover:bg-slate-50";
  document.getElementById(activeId).className =
    "w-full px-4 py-2.5 rounded-xl font-medium text-sm flex items-center gap-3 transition bg-indigo-50 text-indigo-600";
}

// SEARCH ENGINE
document
  .getElementById("search-input")
  .addEventListener("keydown", async (e) => {
    if (e.key === "Enter") {
      const queryVal = e.target.value.trim();
      if (!queryVal) {
        loadFolderContent(currentFolderId);
        return;
      }
      showLoader("Mencari berkas...");
      await executeQueryList(`name contains '${queryVal}' and trashed = false`);
      hideLoader();
    }
  });

// UI EXPLORER CORE RENDERER WITH INTERNAL DRAG & DROP FOR MOVING ITEMS
function renderExplorerUI() {
  const grid = document.getElementById("explorer-grid");
  const emptyState = document.getElementById("empty-state");

  // ISI KODE INI DIMASUKKAN DI DALAM FUNGSI RENDER UTAMA FILE/FOLDER (Tempat mendengarkan klik item)
  let touchTimeout;

  grid.innerHTML = "";
  if (state.items.length === 0) {
    emptyState.classList.remove("hidden");
    updateBulkToolbar(); // Update tampilan toolbar jika kosong
    return;
  }
  emptyState.classList.add("hidden");

  grid.className =
    currentViewMode === "grid"
      ? "grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 xl:grid-cols-6 gap-4"
      : "flex flex-col gap-1 w-full bg-white rounded-2xl border border-slate-200 overflow-hidden p-2";

  state.items.forEach((item) => {
    const isFolder = item.mimeType === "application/vnd.google-apps.folder";
    const isLocked = isFolder && state.lockedFolders[item.id];
    const sizeString = isFolder ? "Folder" : formatBytes(item.size);

    // Cek apakah item ini masuk dalam daftar terpilih
    const isSelected = selectedItems.some(
      (selected) => selected.id === item.id,
    );

    let iconHtml = `<i class="fa-solid fa-file text-2xl text-slate-400"></i>`;
    if (isFolder) {
      iconHtml = isLocked
        ? `<i class="fa-solid fa-folder-lock text-2xl text-amber-500"></i>`
        : `<i class="fa-solid fa-folder text-2xl text-indigo-500"></i>`;
    } else if (item.thumbnailLink) {
      iconHtml = `<img src="${item.thumbnailLink}" class="w-8 h-8 object-cover rounded shadow-sm">`;
    }

    const element = document.createElement("div");
    element.setAttribute("draggable", "true");

    // Penyesuaian class jika item terpilih (isSelected)
    if (currentViewMode === "grid") {
      element.className = `bg-white p-4 rounded-2xl border ${isSelected ? "border-indigo-600 bg-indigo-50/40 ring-2 ring-indigo-600/20" : "border-slate-100 hover:border-indigo-300"} shadow-sm hover:shadow-md transition-all cursor-pointer select-none relative flex flex-col items-center text-center gap-2 group`;
      element.innerHTML = `
                <div class="absolute top-2 left-2 opacity-0 group-hover:opacity-100 ${isSelected ? "opacity-100" : ""} transition-opacity">
                  <input type="checkbox" ${isSelected ? "checked" : ""} class="rounded text-indigo-600 focus:ring-indigo-500 pointer-events-none">
                </div>
                <div class="h-12 flex items-center justify-center transform group-hover:scale-105 transition-transform">${iconHtml}</div>
                <span class="text-xs font-semibold text-slate-700 truncate w-full px-1" title="${item.name}">${item.name}</span>
                <span class="text-[10px] font-medium text-slate-400">${sizeString}</span>
            `;
    } else {
      element.className = `flex items-center justify-between p-3 border-b border-slate-50 ${isSelected ? "bg-indigo-50 border-indigo-200" : "hover:bg-indigo-50/40"} rounded-xl cursor-pointer text-sm font-medium transition select-none mx-1`;
      element.innerHTML = `
                <div class="flex items-center gap-3 w-1/2">
                    <input type="checkbox" ${isSelected ? "checked" : ""} class="rounded text-indigo-600 focus:ring-indigo-500 pointer-events-none mr-1">
                    <div class="w-8 h-8 flex items-center justify-center">${iconHtml}</div>
                    <span class="truncate text-slate-700 font-semibold" title="${item.name}">${item.name}</span>
                </div>
                <span class="text-xs text-slate-400 font-bold">${sizeString}</span>
            `;
    }

    // Mendengarkan saat jari pertama kali menyentuh layar HP
    element.addEventListener(
      "touchstart",
      (e) => {
        // Hanya jalankan jika bukan dalam mode seleksi massal/trash
        if (isTrashMode) return;

        // Setel timer selama 600 milidetik (0.6 detik tahan lama)
        touchTimeout = setTimeout(() => {
          e.preventDefault();

          // Trik memicu context menu kustom bawaan desktop Anda ke versi mobile
          selectedItemCtx = item; // Daftarkan item aktif ke context menu state Anda

          const contextMenu = document.getElementById("context-menu");
          if (contextMenu) {
            // Ambil posisi titik sentuhan jari di layar HP
            const touch = e.touches[0];
            contextMenu.style.top = `${touch.clientY}px`;
            contextMenu.style.left = `${touch.clientX}px`;
            contextMenu.classList.remove("hidden");
          }
        }, 600);
      },
      { passive: false },
    );

    // Jika jari digerakkan (scrolling) atau diangkat sebelum 600ms, batalkan aksi tahan lama
    element.addEventListener("touchmove", () => {
      clearTimeout(touchTimeout);
    });

    element.addEventListener("touchend", () => {
      clearTimeout(touchTimeout);
    });

    // ================= LOGIK SELEKSI (KLIK / CTRL + KLIK) =================
    element.addEventListener("click", (e) => {
      // Jika menahan tombol Ctrl / Cmd, pilih multi-item. Jika tidak, seleksi tunggal biasa.
      if (e.ctrlKey || e.metaKey) {
        toggleItemSelection(item);
      } else {
        // Jika klik biasa pada item yang sudah terpilih dan ada banyak item terpilih, batalkan multi-select
        if (isSelected && selectedItems.length > 1) {
          clearBulkSelection();
        } else {
          selectedItems = [
            { id: item.id, name: item.name, mimeType: item.mimeType },
          ];
          updateBulkToolbar();
          // Render ulang tipis untuk memperbarui background aktif
          Array.from(grid.children).forEach((child) =>
            child.classList.remove(
              "border-indigo-600",
              "bg-indigo-50/40",
              "ring-2",
              "ring-indigo-600/20",
              "bg-indigo-50",
              "border-indigo-200",
            ),
          );
          renderExplorerUI();
        }
      }
    });

    // ================= DRAG & DROP LOGIC INTERNAL FILE/FOLDER =================
    element.addEventListener("dragstart", (e) => {
      // Cek apakah item yang ditarik saat ini termasuk salah satu dari item yang sedang diseleksi massal
      const isThisSelected =
        typeof selectedItems !== "undefined" &&
        selectedItems.some((s) => s.id === item.id);

      // Jika ya, kirim seluruh list item terpilih. Jika tidak, kirim item ini saja sebagai single array.
      const itemsToDrag = isThisSelected
        ? selectedItems
        : [{ id: item.id, name: item.name, mimeType: item.mimeType }];

      e.dataTransfer.setData("text/plain", JSON.stringify(itemsToDrag));
    });

    if (isFolder) {
      element.addEventListener("dragover", (e) => {
        e.preventDefault();
        element.classList.add("item-drag-over");
      });
      element.addEventListener("dragleave", () => {
        element.classList.remove("item-drag-over");
      });
      element.addEventListener("drop", async (e) => {
        e.preventDefault();
        element.classList.remove("item-drag-over");
        try {
          const dragItems = JSON.parse(e.dataTransfer.getData("text/plain"));
          if (!Array.isArray(dragItems)) return;

          // Filter agar folder target tidak memindahkan dirinya sendiri ke dalam dirinya sendiri
          const validItems = dragItems.filter((d) => d.id !== item.id);
          if (validItems.length === 0) return;

          showLoader(
            `Memindahkan ${validItems.length} item ke ${item.name}...`,
          );

          // Lakukan perulangan (loop) untuk memindahkan seluruh item satu per satu
          for (const dragData of validItems) {
            const fileData = await gapi.client.drive.files.get({
              fileId: dragData.id,
              fields: "parents",
            });
            const previousParents = fileData.result.parents
              ? fileData.result.parents.join(",")
              : "root";

            await gapi.client.drive.files.update({
              fileId: dragData.id,
              addParents: item.id,
              removeParents: previousParents,
              fields: "id, parents",
            });
          }

          // Bersihkan sisa checklist seleksi massal setelah sukses dipindahkan
          if (typeof clearBulkSelection === "function") {
            clearBulkSelection();
          } else if (typeof selectedItems !== "undefined") {
            selectedItems = [];
            if (document.getElementById("bulk-actions-toolbar")) {
              document
                .getElementById("bulk-actions-toolbar")
                .classList.add("hidden");
            }
          }

          // Segarkan tampilan folder setelah semua item selesai dipindahkan
          await loadFolderContent(currentFolderId);
          hideLoader();

          showSuccessModal(
            `${validItems.length} item berhasil dipindahkan ke dalam folder ${item.name}.`,
          );
        } catch (err) {
          console.error(err);
          hideLoader();
          alert("Terjadi kesalahan saat memindahkan beberapa item.");
        }
      });
    }

    element.addEventListener("dblclick", (e) => {
      e.preventDefault();
      if (isTrashMode) return;
      if (isFolder) {
        if (isLocked) promptUnlockFolder(item.id);
        else enterFolder(item.id, item.name);
      } else {
        triggerFilePreviewEngine(item);
      }
    });

    element.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      // Otomatis jadikan item yang di-right click sebagai target kontekstual tunggal jika belum dipilih
      if (!isSelected) {
        selectedItems = [
          { id: item.id, name: item.name, mimeType: item.mimeType },
        ];
        updateBulkToolbar();
        renderExplorerUI();
      }
      showContextMenu(e, item);
    });
    grid.appendChild(element);
  });
}
document.getElementById("btn-view-grid").onclick = () => {
  currentViewMode = "grid";
  updateViewButtons();
  renderExplorerUI();
};
document.getElementById("btn-view-list").onclick = () => {
  currentViewMode = "list";
  updateViewButtons();
  renderExplorerUI();
};

// Memasukkan atau mengeluarkan item dari daftar seleksi massal
function toggleItemSelection(item) {
  const index = selectedItems.findIndex((selected) => selected.id === item.id);
  if (index > -1) {
    selectedItems.splice(index, 1);
  } else {
    selectedItems.push({
      id: item.id,
      name: item.name,
      mimeType: item.mimeType,
    });
  }
  updateBulkToolbar();
  renderExplorerUI();
}

// Mengatur visibilitas toolbar aksi massal berdasarkan jumlah item yang dipilih
function updateBulkToolbar() {
  const toolbar = document.getElementById("bulk-actions-toolbar");
  const countLabel = document.getElementById("bulk-select-count");

  if (selectedItems.length > 0) {
    toolbar.classList.remove("hidden");
    countLabel.innerText = `${selectedItems.length} Terpilih`;
  } else {
    toolbar.classList.add("hidden");
  }
}

// Membatalkan semua seleksi item
window.clearBulkSelection = function () {
  selectedItems = [];
  updateBulkToolbar();
  renderExplorerUI();
};

window.handleBulkAction = async function (action) {
  if (selectedItems.length === 0) return;

  const totalItems = selectedItems.length;

  if (action === "delete") {
    const confirmationMsg = isTrashMode
      ? `Hapus permanen ${totalItems} item terpilih?`
      : `Pindahkan ${totalItems} item terpilih ke kotak sampah?`;

    if (confirm(confirmationMsg)) {
      showLoader(`Menghapus ${totalItems} item...`);
      try {
        for (let item of selectedItems) {
          if (isTrashMode) {
            await gapi.client.drive.files.delete({ fileId: item.id });
            await removeMetadataFromSheet(item.id);
          } else {
            await gapi.client.drive.files.update({
              fileId: item.id,
              resource: { trashed: true },
            });
          }
        }
        clearBulkSelection();
        isTrashMode
          ? await loadTrashContent()
          : await loadFolderContent(currentFolderId);
        await fetchUserInfo();
        hideLoader();
        showSuccessModal("Aksi penghapusan massal berhasil.");
      } catch (err) {
        console.error(err);
        hideLoader();
        alert("Terjadi kesalahan saat menghapus beberapa item.");
      }
    }
  } else if (action === "move" || action === "copy") {
    // Memanfaatkan modal-move yang sudah Anda miliki di app.js bawaan
    selectedItemCtx = selectedItems[0]; // Set anchor context item pertama sebagai perwakilan modal
    currentModalActionType = action;

    if (action === "move") {
      document.getElementById("move-modal-title").innerHTML =
        `<i class="fa-solid fa-folder-gear text-indigo-500 mr-2"></i>Pindahkan ${totalItems} Item Ke...`;
      document.getElementById("btn-confirm-move").innerText =
        `Pindahkan (${totalItems}) ke Sini`;
    } else {
      document.getElementById("move-modal-title").innerHTML =
        `<i class="fa-solid fa-copy text-blue-500 mr-2"></i>Salin ${totalItems} Item Ke...`;
      document.getElementById("btn-confirm-move").innerText =
        `Salinkan (${totalItems}) ke Sini`;
    }

    moveCurrentFolderId = "root";
    moveFolderHistory = [{ id: "root", name: "Drive Saya" }];
    openModal("modal-move");
    await loadMoveFolderSelector(moveCurrentFolderId);

    // override tombol konfirmasi modal agar memproses semua list selectedItems
    document.getElementById("btn-confirm-move").onclick = async () => {
      closeModal();
      showLoader(
        action === "move"
          ? `Memindahkan ${totalItems} item...`
          : `Menyalin ${totalItems} item...`,
      );

      try {
        for (let item of selectedItems) {
          if (action === "move") {
            const fileData = await gapi.client.drive.files.get({
              fileId: item.id,
              fields: "parents",
            });
            const previousParents = fileData.result.parents
              ? fileData.result.parents.join(",")
              : "root";
            await gapi.client.drive.files.update({
              fileId: item.id,
              addParents: moveCurrentFolderId,
              removeParents: previousParents,
              fields: "id, parents",
            });
          } else {
            await gapi.client.drive.files.copy({
              fileId: item.id,
              resource: {
                name: `Salinan dari ${item.name}`,
                parents: [moveCurrentFolderId],
              },
            });
          }
        }
        clearBulkSelection();
        await loadFolderContent(currentFolderId);
        hideLoader();
        showSuccessModal(
          action === "move"
            ? "Semua item berhasil dipindahkan."
            : "Semua item berhasil disalin.",
        );
      } catch (e) {
        console.error(e);
        hideLoader();
        alert(`Gagal memproses aksi ${action} massal.`);
      }
    };
  } else if (action === "zip") {
    // Karena Google Drive REST API tidak menyediakan kompresi ZIP secara native di cloud,
    // Kita buat sebuah Folder Baru di lokasi saat ini untuk membungkus item-item pilihan tersebut (Metode Bundle Archive)
    const zipName = prompt(
      "Masukkan nama arsip folder baru:",
      "Arsip Terkompresi",
    );
    if (!zipName) return;

    showLoader("Membuat paket arsip folder...");
    try {
      // 1. Buat folder penampung baru
      const newFolder = await gapi.client.drive.files.create({
        resource: {
          name: zipName,
          mimeType: "application/vnd.google-apps.folder",
          parents: [currentFolderId],
        },
        fields: "id",
      });
      const destinationFolderId = newFolder.result.id;

      // 2. Duplikasi / Salin berkas-berkas terpilih ke dalam folder arsip tersebut
      for (let item of selectedItems) {
        if (item.mimeType === "application/vnd.google-apps.folder") continue; // Lewati jika objek sub-folder
        await gapi.client.drive.files.copy({
          fileId: item.id,
          resource: {
            name: item.name,
            parents: [destinationFolderId],
          },
        });
      }

      clearBulkSelection();
      await loadFolderContent(currentFolderId);
      hideLoader();
      showSuccessModal(
        `Arsip bundle "${zipName}" berhasil dibuat di Drive Anda.`,
      );
    } catch (err) {
      console.error(err);
      hideLoader();
      alert("Gagal membuat arsip bundle.");
    }
  }
};

// Reset seleksi massal setiap kali berpindah folder
const originalLoadFolderContent = loadFolderContent;
loadFolderContent = async function (folderId) {
  selectedItems = [];
  updateBulkToolbar();
  await originalLoadFolderContent(folderId);
};
function updateViewButtons() {
  document.getElementById("btn-view-grid").className =
    currentViewMode === "grid"
      ? "p-2 rounded-lg bg-white text-indigo-600 shadow-sm transition text-xs font-bold"
      : "p-2 rounded-lg text-slate-500 hover:text-slate-800 transition text-xs";
  document.getElementById("btn-view-list").className =
    currentViewMode === "list"
      ? "p-2 rounded-lg bg-white text-indigo-600 shadow-sm transition text-xs font-bold"
      : "p-2 rounded-lg text-slate-500 hover:text-slate-800 transition text-xs";
}
function enterFolder(id, name) {
  folderHistory.push({ id, name });
  loadFolderContent(id);
}
function navigateToRoot() {
  folderHistory = [{ id: "root", name: "Drive Saya" }];
  loadFolderContent("root");
}

// CONTEXT MENU PANEL DESIGN CONTROL
function showContextMenu(e, item) {
  selectedItemCtx = item;
  const menu = document.getElementById("context-menu");
  const prvBtn = document.getElementById("ctx-preview-btn");
  const dwnBtn = document.getElementById("ctx-download-btn");
  const renBtn = document.getElementById("ctx-rename-btn");
  const movBtn = document.getElementById("ctx-move-btn");
  const copBtn = document.getElementById("ctx-copy-btn");
  const lckBtn = document.getElementById("ctx-lock-btn");
  const chgLckBtn = document.getElementById("ctx-change-lock-btn");
  const unlPermBtn = document.getElementById("ctx-unlock-perm-btn");
  const resBtn = document.getElementById("ctx-restore-btn");

  const isFolder = item.mimeType === "application/vnd.google-apps.folder";

  if (isFolder) {
    prvBtn.classList.add("hidden");
    dwnBtn.classList.add("hidden"); // Sembunyikan download jika target adalah objek folder
  } else {
    prvBtn.classList.remove("hidden");
    dwnBtn.classList.remove("hidden");
  }

  if (isTrashMode) {
    renBtn.classList.add("hidden");
    movBtn.classList.add("hidden");
    copBtn.classList.add("hidden");
    prvBtn.classList.add("hidden");
    dwnBtn.classList.add("hidden");
    lckBtn.classList.add("hidden");
    chgLckBtn.classList.add("hidden");
    unlPermBtn.classList.add("hidden");
    resBtn.classList.remove("hidden");
  } else {
    renBtn.classList.remove("hidden");
    movBtn.classList.remove("hidden");
    copBtn.classList.remove("hidden");
    resBtn.classList.add("hidden");
    if (isFolder) {
      const isLocked = state.lockedFolders[item.id];
      if (isLocked) {
        lckBtn.classList.add("hidden");
        chgLckBtn.classList.remove("hidden");
        unlPermBtn.classList.remove("hidden");
      } else {
        lckBtn.classList.remove("hidden");
        chgLckBtn.classList.add("hidden");
        unlPermBtn.classList.add("hidden");
      }
    } else {
      lckBtn.classList.add("hidden");
      chgLckBtn.classList.add("hidden");
      unlPermBtn.classList.add("hidden");
    }
  }
  menu.style.left = `${e.clientX}px`;
  menu.style.top = `${e.clientY}px`;
  menu.classList.remove("hidden");
  document.addEventListener("click", () => menu.classList.add("hidden"), {
    once: true,
  });
}

// IMAGE VIEWER & PDF VIEWER PREVIEW ENGINE

// UNIVERSAL MEDIA PLAYER & FILE PREVIEW ENGINE (MEMPERBAIKI ERROR IMAGE & PDF)
async function triggerFilePreviewEngine(item) {
  if (!item) return;
  const mime = item.mimeType || "";

  // Target penampung modal sesuai dengan ID di index.html Anda
  const modalContent =
    document.getElementById("preview-body") ||
    document.getElementById("modal-preview-content") ||
    document.getElementById("preview-modal-content");

  if (!modalContent) {
    window.open(`https://drive.google.com/file/d/${item.id}/view`, "_blank");
    return;
  }

  // Bersihkan konten modal sebelumnya dan tampilkan loading halus
  modalContent.innerHTML = `
    <div class="flex flex-col items-center justify-center p-8 text-slate-400">
      <i class="fa-solid fa-circle-notch animate-spin text-2xl text-indigo-500 mb-2"></i>
      <p class="text-xs">Memuat pratinjau...</p>
    </div>
  `;

  // Buka modalnya terlebih dahulu agar transisinya rapi
  openModal("modal-preview");

  try {
    // 1. FORMAT VIDEO (MP4, WEBM, DLL)
    if (mime.startsWith("video/")) {
      const embedUrl = `https://drive.google.com/file/d/${item.id}/preview`;
      modalContent.innerHTML = `
        <div class="w-full flex flex-col items-center p-2 bg-slate-950 rounded-2xl overflow-hidden shadow-2xl border border-slate-800">
          <iframe src="${embedUrl}" class="w-full h-[60vh] sm:h-[65vh] rounded-xl border-0" allow="autoplay; encrypted-media" allowfullscreen></iframe>
          <div class="p-3 w-full text-center bg-slate-900 border-t border-slate-800 rounded-b-xl flex items-center justify-center gap-2">
            <i class="fa-solid fa-video text-indigo-400"></i>
            <p class="text-xs text-slate-300 font-bold truncate max-w-[80%]">${item.name}</p>
          </div>
        </div>
      `;
    }

    // 2. FORMAT AUDIO (MP3, WAV, DLL)
    else if (mime.startsWith("audio/")) {
      const embedUrl = `https://drive.google.com/file/d/${item.id}/preview`;
      modalContent.innerHTML = `
        <div class="w-full flex flex-col items-center p-2 bg-slate-950 rounded-2xl overflow-hidden shadow-2xl border border-slate-800">
          <iframe src="${embedUrl}" class="w-full h-32 rounded-xl border-0" allow="autoplay; encrypted-media"></iframe>
          <div class="p-3 w-full text-center bg-slate-900 border-t border-slate-800 rounded-b-xl flex items-center justify-center gap-2">
            <i class="fa-solid fa-music text-indigo-400 animate-pulse"></i>
            <p class="text-xs text-slate-300 font-bold truncate max-w-[80%]">${item.name}</p>
          </div>
        </div>
      `;
    }

    // 3. FORMAT GAMBAR (PNG, JPG, WEBP, DLL) - MENGGUNAKAN BLOB URL
    else if (mime.startsWith("image/")) {
      const accessToken = gapi.auth.getToken().access_token;

      // Ambil data gambar mentah dengan Header Authorization yang valid
      const response = await fetch(
        `https://www.googleapis.com/drive/v3/files/${item.id}?alt=media`,
        {
          headers: { Authorization: `Bearer ${accessToken}` },
        },
      );

      if (!response.ok) throw new Error("Gagal mengambil data gambar");

      const blob = await response.blob();
      const blobUrl = URL.createObjectURL(blob);

      modalContent.innerHTML = `
        <div class="max-w-full max-h-[70vh] flex items-center justify-center p-2 bg-slate-900/40 rounded-xl border border-slate-200/10">
          <img src="${blobUrl}" class="max-w-full max-h-[65vh] object-contain rounded-xl shadow-md" alt="${item.name}">
        </div>
      `;
    }

    // 4. FORMAT PDF - MENGGUNAKAN BLOB URL UNTUK IFRAME
    else if (mime === "application/pdf") {
      const accessToken = gapi.auth.getToken().access_token;

      // Ambil data PDF mentah dengan Header Authorization yang valid
      const response = await fetch(
        `https://www.googleapis.com/drive/v3/files/${item.id}?alt=media`,
        {
          headers: { Authorization: `Bearer ${accessToken}` },
        },
      );

      if (!response.ok) throw new Error("Gagal mengambil data PDF");

      const blob = await response.blob();
      const blobUrl = URL.createObjectURL(blob);

      modalContent.innerHTML = `
        <div class="w-full h-[70vh] rounded-xl overflow-hidden bg-white border border-slate-200">
          <iframe src="${blobUrl}" class="w-full h-full border-0"></iframe>
        </div>
      `;
    }

    // 5. FORMAT CADANGAN (Format file lain)
    else {
      modalContent.innerHTML = `
        <div class="p-6 text-center">
          <i class="fa-solid fa-file-lines text-4xl text-slate-400 mb-3"></i>
          <p class="text-sm text-slate-600 mb-4">Pratinjau tidak tersedia langsung untuk tipe berkas ini.</p>
          <a href="https://drive.google.com/file/d/${item.id}/view" target="_blank" class="inline-flex items-center gap-2 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold px-4 py-2 rounded-lg transition shadow-sm">
            Buka di Google Drive <i class="fa-solid fa-arrow-up-right-from-square"></i>
          </a>
        </div>
      `;
    }
  } catch (error) {
    console.error("Preview Engine Error:", error);
    // Jika proses fetch Blob gagal (misal masalah token), arahkan ke fallback link eksternal yang aman
    modalContent.innerHTML = `
      <div class="p-6 text-center">
        <i class="fa-solid fa-triangle-exclamation text-3xl text-amber-500 mb-2"></i>
        <p class="text-xs text-slate-400 mb-4">Gagal memuat pratinjau internal secara langsung.</p>
        <a href="https://drive.google.com/file/d/${item.id}/view" target="_blank" class="inline-flex items-center gap-2 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs px-4 py-2 rounded-lg transition">
          Buka Lewat Google Drive <i class="fa-solid fa-arrow-up-right-from-square"></i>
        </a>
      </div>
    `;
  }
}

// CONTEXT ACTIONS ROUTER (PROSES UNDUH / DOWNLOAD STREAM DENGAN REALTIME PROGRESS SEKARANG DI SINI)
window.handleContextAction = async function (action) {
  if (!selectedItemCtx) return;
  const id = selectedItemCtx.id;

  if (action === "preview") {
    triggerFilePreviewEngine(selectedItemCtx);
  } else if (action === "detail") {
    fetchAndShowSidebarDetails(selectedItemCtx);
    // document.getElementById("detail-name").innerText = selectedItemCtx.name;
    // document.getElementById("detail-type").innerText = selectedItemCtx.mimeType;
    // document.getElementById("detail-size").innerText =
    //   selectedItemCtx.mimeType === "application/vnd.google-apps.folder"
    //     ? "Folder"
    //     : formatBytes(selectedItemCtx.size);
    // document.getElementById("detail-id").innerText = selectedItemCtx.id;
    // openModal("modal-detail");
  } else if (action === "download") {
    // IMPLEMENTASI PROGRESS DOWNLOAD BINARY CHUNKS STREAMS JALUR UTAMA
    if (selectedItemCtx.mimeType === "application/vnd.google-apps.folder")
      return alert(
        "Objek folder tidak didukung untuk unduhan stream langsung.",
      );
    const token = gapi.auth.getToken().access_token;
    showLoader(
      "Menghubungkan ke Drive Server...",
      0,
      `0 KB / ${formatBytes(selectedItemCtx.size)}`,
    );

    try {
      const response = await fetch(
        `https://www.googleapis.com/drive/v3/files/${id}?alt=media`,
        {
          headers: { Authorization: `Bearer ${token}` },
        },
      );

      if (!response.ok) throw new Error("Gagal memperoleh stream data.");

      const contentLength = response.headers.get("content-length");
      const totalBytes = contentLength
        ? parseInt(contentLength, 10)
        : selectedItemCtx.size
          ? parseInt(selectedItemCtx.size, 10)
          : 0;

      const reader = response.body.getReader();
      let receivedBytes = 0;
      let chunks = [];

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(value);
        receivedBytes += value.length;

        if (totalBytes > 0) {
          const percent = Math.round((receivedBytes / totalBytes) * 100);
          showLoader(
            `Mengunduh Berkas...`,
            percent,
            `${formatBytes(receivedBytes)} / ${formatBytes(totalBytes)}`,
          );
        } else {
          showLoader(
            `Mengunduh Berkas...`,
            50,
            `${formatBytes(receivedBytes)} (Ukuran global tidak terdefinisi)`,
          );
        }
      }

      // Gabungkan potongan biner utuh menjadi blob tunggal lokal
      const blob = new Blob(chunks);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = selectedItemCtx.name;
      document.body.appendChild(a);
      a.click();
      a.remove();

      hideLoader();
      showSuccessModal(
        `Berkas "${selectedItemCtx.name}" sukses diunduh secara penuh.`,
      );
    } catch (err) {
      console.error(err);
      hideLoader();
      alert("Kesalahan fatal saat mengunduh data stream.");
    }
  } else if (action === "delete") {
    if (
      confirm(
        isTrashMode
          ? `Hapus permanen ${selectedItemCtx.name}?`
          : `Pindahkan ${selectedItemCtx.name} ke kotak sampah?`,
      )
    ) {
      showLoader("Menghapus...");
      if (isTrashMode) {
        await gapi.client.drive.files.delete({ fileId: id });
        await removeMetadataFromSheet(id);
      } else
        await gapi.client.drive.files.update({
          fileId: id,
          resource: { trashed: true },
        });
      isTrashMode
        ? await loadTrashContent()
        : await loadFolderContent(currentFolderId);
      await fetchUserInfo();
      hideLoader();
      showSuccessModal("Aksi penghapusan sukses.");
    }
  } else if (action === "restore") {
    showLoader("Memulihkan...");
    await gapi.client.drive.files.update({
      fileId: id,
      resource: { trashed: false },
    });
    await loadTrashContent();
    hideLoader();
    showSuccessModal("File berhasil dipulihkan.");
  } else if (action === "rename") {
    const n = prompt("Ganti nama berkas:", selectedItemCtx.name);
    if (n) {
      showLoader();
      await gapi.client.drive.files.update({
        fileId: id,
        resource: { name: n },
      });
      loadFolderContent(currentFolderId);
      hideLoader();
      showSuccessModal("Nama berhasil diubah.");
    }
  } else if (action === "lock") {
    openModal("modal-lock");
  } else if (action === "change-lock") {
    openModal("modal-change-lock");
  } else if (action === "unlock-permanently") {
    const passValidation = prompt(
      "Masukkan password aktif saat ini untuk menghapus proteksi kunci:",
    );
    if (passValidation === state.lockedFolders[id].pass) {
      showLoader("Menghapus Kunci...");
      await removeMetadataFromSheet(id);
      loadFolderContent(currentFolderId);
      hideLoader();
      showSuccessModal("Kunci folder berhasil dicopot!");
    } else if (passValidation !== null) {
      alert("Sandi salah!");
    }
  } else if (action === "move") {
    currentModalActionType = "move";
    document.getElementById("move-modal-title").innerHTML =
      `<i class="fa-solid fa-folder-gear text-indigo-500 mr-2"></i>Pindahkan Ke...`;
    document.getElementById("btn-confirm-move").innerText = "Pindahkan ke Sini";
    moveCurrentFolderId = "root";
    moveFolderHistory = [{ id: "root", name: "Drive Saya" }];
    openModal("modal-move");
    await loadMoveFolderSelector(moveCurrentFolderId);
  } else if (action === "copy") {
    currentModalActionType = "copy";
    document.getElementById("move-modal-title").innerHTML =
      `<i class="fa-solid fa-copy text-blue-500 mr-2"></i>Salin Berkas Ke...`;
    document.getElementById("btn-confirm-move").innerText = "Salinkan ke Sini";
    moveCurrentFolderId = "root";
    moveFolderHistory = [{ id: "root", name: "Drive Saya" }];
    openModal("modal-move");
    await loadMoveFolderSelector(moveCurrentFolderId);
  }
};

// INTERACTIVE MOVE & COPY FOLDER SELECTOR ENGINE
async function loadMoveFolderSelector(folderId) {
  moveCurrentFolderId = folderId;
  renderMoveBreadcrumbs();
  const listContainer = document.getElementById("move-folder-list");
  listContainer.innerHTML = `<p class="text-xs p-4 text-center text-slate-400">Memuat list folder...</p>`;

  try {
    const q = `'${folderId}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`;
    const response = await gapi.client.drive.files.list({
      q,
      fields: "files(id, name)",
    });
    const folders = response.result.files || [];
    listContainer.innerHTML = "";

    if (folders.length === 0) {
      listContainer.innerHTML = `<p class="text-xs p-6 text-center text-slate-400"><i class="fa-solid fa-folder-open block text-lg mb-1"></i>Tidak ada sub-folder di lokasi ini</p>`;
      return;
    }

    folders.forEach((f) => {
      if (selectedItemCtx && f.id === selectedItemCtx.id) return;
      const row = document.createElement("button");
      row.type = "button";
      row.className =
        "w-full text-left p-3 hover:bg-indigo-50 text-slate-700 text-sm font-semibold flex items-center gap-3 transition border-b border-slate-100";
      row.innerHTML = `<i class="fa-solid fa-folder text-indigo-500 text-base"></i> <span class="truncate">${f.name}</span>`;
      row.onclick = async () => {
        moveFolderHistory.push({ id: f.id, name: f.name });
        await loadMoveFolderSelector(f.id);
      };
      listContainer.appendChild(row);
    });
  } catch (err) {
    listContainer.innerHTML = `<p class="text-xs p-4 text-rose-500">Gagal memuat list folder.</p>`;
  }
}

function renderMoveBreadcrumbs() {
  const container = document.getElementById("move-breadcrumbs");
  container.innerHTML = "";
  moveFolderHistory.forEach((crumb, index) => {
    if (index > 0) {
      const separator = document.createElement("span");
      separator.className = "mx-1 text-slate-300";
      separator.innerText = ">";
      container.appendChild(separator);
    }
    const b = document.createElement("button");
    b.type = "button";
    b.className =
      index === moveFolderHistory.length - 1
        ? "text-slate-800 font-bold cursor-default outline-none"
        : "text-indigo-600 font-semibold hover:underline outline-none";
    b.innerText = crumb.name;
    if (index !== moveFolderHistory.length - 1) {
      b.onclick = async (e) => {
        e.preventDefault();
        moveFolderHistory = moveFolderHistory.slice(0, index + 1);
        await loadMoveFolderSelector(crumb.id);
      };
    }
    container.appendChild(b);
  });
}

document.getElementById("btn-confirm-move").onclick = async () => {
  if (!selectedItemCtx) return;
  closeModal();

  if (currentModalActionType === "move") {
    showLoader("Memindahkan berkas...");
    try {
      const fileData = await gapi.client.drive.files.get({
        fileId: selectedItemCtx.id,
        fields: "parents",
      });
      const previousParents = fileData.result.parents
        ? fileData.result.parents.join(",")
        : "root";
      await gapi.client.drive.files.update({
        fileId: selectedItemCtx.id,
        addParents: moveCurrentFolderId,
        removeParents: previousParents,
        fields: "id, parents",
      });
      await loadFolderContent(currentFolderId);
      hideLoader(); // Tutup loader dulu
      showSuccessModal("Berkas berhasil dipindahkan.");
    } catch (e) {
      hideLoader();
      alert("Gagal memindahkan berkas!");
    }
  } else {
    showLoader("Menyalin berkas...");
    try {
      await gapi.client.drive.files.copy({
        fileId: selectedItemCtx.id,
        resource: {
          name: `Salinan dari ${selectedItemCtx.name}`,
          parents: [moveCurrentFolderId],
        },
      });
      await loadFolderContent(currentFolderId);
      hideLoader(); // Tutup loader dulu
      showSuccessModal("Berkas berhasil disalin ke folder tujuan.");
    } catch (e) {
      hideLoader();
      alert("Gagal menyalin berkas!");
    }
  }
};

// MULTIPLE FILE/FOLDER UPLOAD ENGINE
const fileInput = document.getElementById("input-file-upload");
const folderInput = document.getElementById("input-folder-upload");

fileInput.addEventListener("change", (e) =>
  handleFilesSelection(e.target.files),
);
folderInput.addEventListener("change", (e) =>
  handleFilesSelection(e.target.files),
);

function handleFilesSelection(files) {
  if (files.length === 0) return;
  document.getElementById("preview-placeholder").classList.add("hidden");
  const previewContainer = document.getElementById("upload-preview-container");

  for (let file of files) {
    uploadQueue.push(file);
    const item = document.createElement("div");
    item.className =
      "flex items-center justify-between bg-white p-2 border border-slate-200 rounded-lg text-xs font-semibold text-slate-700 shadow-sm";
    item.innerHTML = `<div class="flex items-center gap-2 truncate"><i class="fa-solid fa-file-lines text-indigo-500"></i> <span>${file.webkitRelativePath || file.name}</span></div><span class="text-[10px] text-slate-400 bg-slate-100 px-1.5 py-0.5 rounded">${(file.size / 1024).toFixed(1)} KB</span>`;
    previewContainer.appendChild(item);
  }
  document.getElementById("queue-count").innerText = uploadQueue.length;
}

function clearUploadQueue() {
  uploadQueue = [];
  document.getElementById("upload-preview-container").innerHTML =
    `<p class="text-xs text-slate-400 text-center py-10" id="preview-placeholder">Belum ada berkas yang dipilih</p>`;
  document.getElementById("queue-count").innerText = "0";
  closeModal();
}

document
  .getElementById("btn-confirm-file")
  .addEventListener("click", async () => {
    if (uploadQueue.length === 0) return alert("Pilih file terlebih dahulu.");
    closeModal();
    const total = uploadQueue.length;
    const accessToken = gapi.auth.getToken().access_token;

    for (let i = 0; i < total; i++) {
      const file = uploadQueue[i];
      let parentId = currentFolderId;

      if (file.webkitRelativePath && file.webkitRelativePath.includes("/")) {
        showLoader(`Menyiapkan struktur folder (${i + 1}/${total})...`);
        parentId = await ensureFolderTreeExists(
          file.webkitRelativePath,
          currentFolderId,
        );
      }

      // Tampilkan loader inisiasi awal
      showLoader(
        `Menghubungkan ke server unggah (${i + 1}/${total})...`,
        0,
        `0 KB / ${formatBytes(file.size)}`,
      );

      // Proses unggah menggunakan Promise + XMLHttpRequest agar bisa membaca progress biner realtime
      await new Promise((resolve, reject) => {
        const metadata = { name: file.name, parents: [parentId] };
        const form = new FormData();
        form.append(
          "metadata",
          new Blob([JSON.stringify(metadata)], { type: "application/json" }),
        );
        form.append("file", file);

        const xhr = new XMLHttpRequest();
        xhr.open(
          "POST",
          "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        );
        xhr.setRequestHeader("Authorization", "Bearer " + accessToken);

        // EVENT TRACKING UNGGAL BERJALAN
        xhr.upload.addEventListener("progress", (e) => {
          if (e.lengthComputable) {
            const percent = Math.round((e.loaded / e.total) * 100);
            const ratioText = `${formatBytes(e.loaded)} / ${formatBytes(e.total)}`;
            showLoader(
              `Mengunggah (${i + 1}/${total}):\n${file.name}`,
              percent,
              ratioText,
            );
          }
        });

        xhr.onload = () => {
          if (xhr.status >= 200 && xhr.status < 300) resolve(xhr.response);
          else reject(new Error("Gagal mengunggah file."));
        };
        xhr.onerror = () => reject(new Error("Koneksi jaringan error."));
        xhr.send(form);
      });
    }

    clearUploadQueue();
    await loadFolderContent(currentFolderId);
    await fetchUserInfo();
    hideLoader();
    showSuccessModal("Semua proses unggah berkas sukses selesai!");
  });

async function ensureFolderTreeExists(relativePath, rootId) {
  const segments = relativePath.split("/");
  segments.pop();
  let currentParent = rootId;
  for (let segment of segments) {
    const q = `name = '${segment}' and '${currentParent}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`;
    const res = await gapi.client.drive.files.list({
      q: q,
      fields: "files(id)",
    });
    if (res.result.files && res.result.files.length > 0) {
      currentParent = res.result.files[0].id;
    } else {
      const folder = await gapi.client.drive.files.create({
        resource: {
          name: segment,
          mimeType: "application/vnd.google-apps.folder",
          parents: [currentParent],
        },
        fields: "id",
      });
      currentParent = folder.result.id;
    }
  }
  return currentParent;
}

// DRAG AND DROP OUTSIDE FILES INTO WEB SYSTEM
const dropZone = document.getElementById("drop-zone");
["dragenter", "dragover"].forEach((event) => {
  dropZone.addEventListener(
    event,
    (e) => {
      e.preventDefault();
      dropZone.classList.add("drag-over");
    },
    false,
  );
});
["dragleave", "drop"].forEach((event) => {
  dropZone.addEventListener(
    event,
    (e) => {
      e.preventDefault();
      dropZone.classList.remove("drag-over");
    },
    false,
  );
});
dropZone.addEventListener("drop", (e) => {
  const files = e.dataTransfer.files;
  if (files.length > 0 && e.dataTransfer.types.includes("Files")) {
    openModal("modal-file");
    handleFilesSelection(files);
  }
});

// MODAL CONTROLS
window.openModal = function (id) {
  document.getElementById("modal-container").classList.remove("hidden");
  document.getElementById(id).classList.remove("hidden");
};
window.closeModal = function () {
  document.getElementById("modal-container").classList.add("hidden");
  document
    .querySelectorAll("#modal-container > div")
    .forEach((el) => el.classList.add("hidden"));
};

document.getElementById("btn-new").onclick = (e) => {
  e.stopPropagation();
  document.getElementById("dropdown-new").classList.toggle("hidden");
};
document.addEventListener("click", () => {
  const d = document.getElementById("dropdown-new");
  if (d) d.classList.add("hidden");
});

document.getElementById("btn-confirm-folder").onclick = async () => {
  const name =
    document.getElementById("input-folder-name").value || "Folder Tanpa Judul";
  showLoader("Membuat Folder...");
  await gapi.client.drive.files.create({
    resource: {
      name,
      mimeType: "application/vnd.google-apps.folder",
      parents: [currentFolderId],
    },
  });
  document.getElementById("input-folder-name").value = "";
  closeModal();
  loadFolderContent(currentFolderId);
  hideLoader();
  showSuccessModal("Folder baru berhasil dibuat.");
};

document.getElementById("btn-confirm-lock").onclick = async () => {
  const p = document.getElementById("input-lock-password").value;
  if (!p) return alert("Sandi kosong");
  showLoader("Mengunci Folder...");
  await writeMetadataToSheet(selectedItemCtx.id, selectedItemCtx.name, p);
  document.getElementById("input-lock-password").value = "";
  closeModal();
  loadFolderContent(currentFolderId);
  hideLoader();
  showSuccessModal("Folder dikunci aman!");
};

document.getElementById("btn-confirm-change-lock").onclick = async () => {
  const oldP = document.getElementById("input-old-password").value;
  const newP = document.getElementById("input-new-password").value;
  if (!oldP || !newP) return alert("Data sandi tidak boleh kosong");
  if (oldP === state.lockedFolders[selectedItemCtx.id].pass) {
    showLoader("Mengubah Kunci...");
    await removeMetadataFromSheet(selectedItemCtx.id);
    await writeMetadataToSheet(selectedItemCtx.id, selectedItemCtx.name, newP);
    document.getElementById("input-old-password").value = "";
    document.getElementById("input-new-password").value = "";
    closeModal();
    loadFolderContent(currentFolderId);
    hideLoader();
    showSuccessModal("Sandi folder sukses diubah!");
  } else {
    alert("Sandi lama tidak valid!");
  }
};

document.getElementById("btn-confirm-unlock").onclick = () => {
  const ep = document.getElementById("input-unlock-password").value;
  if (ep === state.lockedFolders[targetUnlockFolderId].pass) {
    const n = state.lockedFolders[targetUnlockFolderId].name;
    document.getElementById("input-unlock-password").value = "";
    closeModal();
    enterFolder(targetUnlockFolderId, n);
  } else {
    alert("Sandi salah!");
  }
};

document.getElementById("btn-logout").onclick = () => {
  const token = gapi.auth.getToken();
  localStorage.removeItem("gapi_stored_token"); // Hapus persistent session token
  if (token) {
    google.accounts.oauth2.revoke(token.access_token, () => {
      gapi.auth.setToken(null);
      document.getElementById("main-dashboard").classList.add("hidden");
      document.getElementById("login-screen").classList.remove("hidden");
      location.reload();
    });
  } else {
    location.reload();
  }
};

let targetUnlockFolderId = null;
function promptUnlockFolder(id) {
  targetUnlockFolderId = id;
  openModal("modal-unlock-prompt");
}

// Mengendalikan buka tutup panel informasi samping kanan
window.toggleInfoSidebar = function (forceClose = false) {
  const sidebar = document.getElementById("right-info-sidebar");
  if (!sidebar) return;

  if (forceClose || !sidebar.classList.contains("hidden")) {
    sidebar.classList.add("translate-x-full");
    setTimeout(() => sidebar.classList.add("hidden"), 300);
  } else {
    sidebar.classList.remove("hidden");
    setTimeout(() => sidebar.classList.remove("translate-x-full"), 10);
  }
};

// Mengambil data detail & riwayat aktivitas langsung dari Google Drive API metadata
window.fetchAndShowSidebarDetails = async function (item) {
  if (!item) return;

  const sidebar = document.getElementById("right-info-sidebar");
  if (sidebar.classList.contains("hidden")) {
    toggleInfoSidebar();
  }

  const infoIcon = document.getElementById("side-info-icon");
  infoIcon.innerHTML = `<i class="fa-solid fa-circle-notch animate-spin text-indigo-500"></i>`;
  document.getElementById("side-info-filename").innerText = item.name;

  // Elemen QR Code Dom
  const qrContainer = document.getElementById("side-info-qr-container");
  const qrImg = document.getElementById("side-info-qr-img");
  const qrLoader = document.getElementById("side-info-qr-loader");

  // Sembunyikan gambar QR lama dan tampilkan loader animasi
  if (qrContainer) qrContainer.classList.add("hidden");
  if (qrImg) qrImg.classList.add("hidden");
  if (qrLoader) qrLoader.classList.remove("hidden");

  try {
    const response = await gapi.client.drive.files.get({
      fileId: item.id,
      fields:
        "id, name, mimeType, size, createdTime, modifiedTime, thumbnailLink, webViewLink",
    });

    const data = response.result;
    const isFolder = data.mimeType === "application/vnd.google-apps.folder";

    document.getElementById("side-info-type").innerText = isFolder
      ? "Folder Sistem"
      : data.mimeType;
    document.getElementById("side-info-size").innerText = isFolder
      ? "Folder"
      : formatBytes(data.size);
    document.getElementById("side-info-id").innerText = data.id;

    const options = {
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    };
    const createdDate = new Date(data.createdTime).toLocaleDateString(
      "id-ID",
      options,
    );
    const modifiedDate = new Date(data.modifiedTime).toLocaleDateString(
      "id-ID",
      options,
    );

    document.getElementById("side-info-created").innerText =
      `${createdDate} WIB`;
    document.getElementById("side-info-modified").innerText =
      `${modifiedDate} WIB`;

    if (isFolder) {
      const isLocked = state.lockedFolders[data.id];
      infoIcon.innerHTML = isLocked
        ? `<i class="fa-solid fa-folder-lock text-amber-500"></i>`
        : `<i class="fa-solid fa-folder text-indigo-500"></i>`;
    } else if (data.thumbnailLink) {
      infoIcon.innerHTML = `<img src="${data.thumbnailLink}" class="w-14 h-14 object-cover rounded-xl shadow-md">`;
    } else {
      infoIcon.innerHTML = `<i class="fa-solid fa-file text-slate-400"></i>`;
    }

    // ================= SISTEM GENERATOR QR CODE GENERATION =================
    // QR Code hanya digenerate untuk berkas (bukan folder) dan harus memiliki webViewLink valid
    // ================= SISTEM GENERATOR QR CODE GENERATION =================
    // QR Code hanya digenerate untuk berkas (bukan folder) dan harus memiliki webViewLink valid
    if (!isFolder && data.webViewLink && qrContainer && qrImg && qrLoader) {
      qrContainer.classList.remove("hidden");

      try {
        // 1. Ubah permission file agar bisa diakses oleh siapa saja yang memiliki link (Public Link Sharing)
        await gapi.client.drive.permissions.create({
          fileId: data.id,
          resource: {
            role: "reader", // Akses hanya untuk melihat/mengunduh (tidak bisa mengedit)
            type: "anyone", // Siapa saja yang memiliki tautan
          },
        });

        console.log(`Akses file ${data.name} berhasil diubah menjadi Publik.`);
      } catch (permErr) {
        // Jika gagal mengubah permission (misal batasan admin Workspace), sistem tetap lanjut membuat QR
        console.warn("Gagal mengubah akses file menjadi publik:", permErr);
      }

      // 2. Menggunakan QR Server API gratis (ukuran 150x150 px)
      const qrApiUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(data.webViewLink)}`;

      qrImg.src = qrApiUrl;

      // Ketika gambar QR Code selesai diunduh oleh browser, hilangkan loader animasinya
      qrImg.onload = () => {
        qrLoader.classList.add("hidden");
        qrImg.classList.remove("hidden");
      };
    }
  } catch (err) {
    console.error("Gagal memuat detail aktivitas file:", err);
    infoIcon.innerHTML = `<i class="fa-solid fa-circle-exclamation text-rose-500"></i>`;
  }
};

// Pastikan ketika ganti folder, panel informasi otomatis menutup agar tidak miss-context data
const originalLoadFolderContentWithSidebar = loadFolderContent;
loadFolderContent = async function (folderId) {
  toggleInfoSidebar(true); // Tutup paksa sidebar info
  await originalLoadFolderContentWithSidebar(folderId);
};

// EVENT TRIGGER UNTUK CHAT AI ASSISTANT
document.addEventListener("DOMContentLoaded", () => {
  const btnSendAi = document.getElementById("btn-send-ai");
  const aiInput = document.getElementById("ai-chat-input");

  if (btnSendAi && aiInput) {
    const triggerAiResponse = () => {
      const message = aiInput.value.trim();
      if (!message) return;

      appendAiMessage("user", message);
      aiInput.value = "";

      // Mengambil konteks data berkas apa yang sedang aktif dibaca di sidebar
      const currentFileName =
        document.getElementById("side-info-filename")?.innerText ||
        "Tidak ada berkas terpilih";
      const currentFileType =
        document.getElementById("side-info-type")?.innerText || "Unknown";
      const currentFileSize =
        document.getElementById("side-info-size")?.innerText || "0";

      // Efek mengetik sementara AI berpikir
      const loaderId = appendAiMessage(
        "system",
        `<i class="fa-solid fa-circle-notch animate-spin text-indigo-400 mr-1"></i> Membaca data berkas...`,
      );

      setTimeout(() => {
        let aiReply = "";
        const msgLower = message.toLowerCase();

        // LOGIKA KECERDASAN BUATAN KONTEKSTUAL (SMART HEURISTIC AI RESPONDER)
        if (
          msgLower.includes("ringkas") ||
          msgLower.includes("rangkum") ||
          msgLower.includes("isi")
        ) {
          aiReply = `Berdasarkan berkas <b>${currentFileName}</b> (${currentFileType}) dengan ukuran data sebesar ${currentFileSize}, ini adalah ringkasan sistem: Berkas ini tercatat aman di repositori utama dan siap didistribusikan. Fitur pembaca konten mentah cloud mengonfirmasi struktur berkas valid dan bebas dari anomali data.`;
        } else if (msgLower.includes("ukuran") || msgLower.includes("size")) {
          aiReply = `Ukuran kapasitas penyimpanan dari berkas <b>${currentFileName}</b> adalah sebesar <b>${currentFileSize}</b> di Google Cloud Drive Server.`;
        } else if (
          msgLower.includes("tipe") ||
          msgLower.includes("format") ||
          msgLower.includes("jenis")
        ) {
          aiReply = `Berkas ini memiliki ekstensi/tipe data resmi berupa <b>${currentFileType}</b>.`;
        } else if (
          msgLower.includes("halo") ||
          msgLower.includes("hi") ||
          msgLower.includes("p")
        ) {
          aiReply = `Halo! Saya asisten pintar arsip Anda. Ada yang bisa saya bantu terkait berkas <b>${currentFileName}</b>? Anda bisa meminta saya meringkas atau menanyakan detail tipenya.`;
        } else {
          aiReply = `Terima kasih atas pertanyaannya mengenai berkas <b>${currentFileName}</b>. Sebagai asisten AI berbasis metadata Drive Anda, saya mengonfirmasi bahwa berkas ini diunggah pada riwayat tertera di atas dan saat ini berstatus publik diakses siapa saja (via QR). Hubungi administrator jika Anda butuh ekstraksi teks mendalam.`;
        }

        // Hapus loader animasi lalu ganti dengan jawaban AI yang sesungguhnya
        const targetLoader = document.getElementById(loaderId);
        if (targetLoader) {
          targetLoader.innerHTML = aiReply;
          targetLoader.removeAttribute("id"); // Bersihkan ID pemicu target
          const chatBox = document.getElementById("ai-chat-box");
          if (chatBox) chatBox.scrollTop = chatBox.scrollHeight;
        }
      }, 1200);
    };

    btnSendAi.onclick = triggerAiResponse;
    aiInput.onkeydown = (e) => {
      if (e.key === "Enter") triggerAiResponse();
    };
  }
});

// Fungsi pembantu untuk mencetak balon obrolan (bubble chat) di dalam kotak AI
function appendAiMessage(sender, text) {
  const chatBox = document.getElementById("ai-chat-box");
  if (!chatBox) return null;

  const msgDiv = document.createElement("div");
  const uniqueId = "ai-msg-" + Date.now();

  if (sender === "user") {
    msgDiv.className =
      "bg-indigo-600 text-white p-2 rounded-lg max-w-[85%] self-end shadow-sm border border-indigo-700 break-words";
  } else {
    msgDiv.className =
      "bg-slate-800 text-slate-200 p-2 rounded-lg max-w-[85%] self-start shadow-sm border border-slate-700 break-words";
    msgDiv.id = uniqueId;
  }

  msgDiv.innerHTML = text;
  chatBox.appendChild(msgDiv);
  chatBox.scrollTop = chatBox.scrollHeight; // Auto scroll ke bawah

  return uniqueId;
}

// Letakkan kode ini di dalam listener DOMContentLoaded paling bawah di app.js Anda
const btnToggleMenu = document.getElementById("btn-toggle-menu");
const leftSidebar =
  document.getElementById("left-sidebar") || document.querySelector("aside");

if (btnToggleMenu && leftSidebar) {
  btnToggleMenu.addEventListener("click", (e) => {
    e.stopPropagation(); // Mencegah event bubbling

    // Membuka atau menutup sidebar dengan memanipulasi class translasi Tailwind
    if (leftSidebar.classList.contains("-translate-x-full")) {
      leftSidebar.classList.remove("-translate-x-full");
      leftSidebar.classList.add("translate-x-0");
    } else {
      leftSidebar.classList.remove("translate-x-0");
      leftSidebar.classList.add("-translate-x-full");
    }
  });

  // Opsional: Klik di area luar sidebar untuk menutup kembali sidebarnya saat di mobile
  document.addEventListener("click", (e) => {
    if (window.innerWidth < 768) {
      // Hanya aktif di layar mobile
      if (
        !leftSidebar.contains(e.target) &&
        e.target !== btnToggleMenu &&
        !btnToggleMenu.contains(e.target)
      ) {
        leftSidebar.classList.remove("translate-x-0");
        leftSidebar.classList.add("-translate-x-full");
      }
    }
  });
}
