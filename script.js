// State variables for guests list
let guestsList = [];
let activePreviewGuest = "Tamu Undangan";

// Default sample guests matching Excel structure requirement
const sampleGuests = [
    { name: "Budi Santoso" },
    { name: "Siti Aminah" },
    { name: "Ahmad Fauzi" },
    { name: "Pak RT" },
    { name: "Bu RW" }
];

// Wishes / Guestbook dummy data
let wishesList = [
    {
        name: "Budi Santoso",
        status: "Hadir",
        message: "Selamat menempuh hidup baru Andi dan Siti! Semoga menjadi keluarga yang sakinah, mawaddah, warahmah. Amin!",
        time: "Baru saja"
    },
    {
        name: "Ahmad Fauzi",
        status: "Hadir",
        message: "Barakallahu lakuma wa baraka 'alaikuma wa jama'a bainakuma fii khair. Selamat bro Andi!",
        time: "2 jam yang lalu"
    },
    {
        name: "Siti Aminah",
        status: "Tidak Hadir",
        message: "Selamat ya Siti dan Andi! Maaf sekali belum bisa hadir langsung karena masih di luar kota. Doa terbaik untuk kalian berdua.",
        time: "4 jam yang lalu"
    }
];

// =================================================================
// 1. INITIALIZATION & ROUTING
// =================================================================
document.addEventListener("DOMContentLoaded", () => {
    // Check URL parameters for invitation view
    const urlParams = new URLSearchParams(window.location.search);
    const guestParam = urlParams.get("to");

    if (guestParam !== null) {
        // If 'to' param is present (even if empty), load FULLSCREEN INVITATION VIEW
        showInvitationView(guestParam);
    } else {
        // Otherwise, show the DASHBOARD VIEW
        showDashboardView();
    }

    // Set up all event listeners
    initDashboardListeners();
    initInvitationListeners();
});

// Show Fullscreen Invitation View
function showInvitationView(guestName) {
    document.getElementById("dashboard-view").classList.add("hidden");
    document.getElementById("wedding-invitation-view").classList.remove("hidden");
    
    // Set recipient name (decode first in case it's URL-encoded)
    const displayName = guestName.trim() === "" ? "Tamu Undangan" : decodeURIComponent(guestName);
    document.getElementById("invitation-recipient-name").textContent = displayName;
    
    // Autofill RSVP and Wish names
    const rsvpNameField = document.getElementById("rsvp-name");
    const wishNameField = document.getElementById("wish-name");
    if (rsvpNameField) rsvpNameField.value = displayName;
    if (wishNameField) wishNameField.value = displayName;

    // Render wishes in the invitation
    renderWishes("wishes-list-container");
    
    // Start countdown timer
    startCountdown();
}

// Show Dashboard View
function showDashboardView() {
    document.getElementById("wedding-invitation-view").classList.add("hidden");
    document.getElementById("dashboard-view").classList.remove("hidden");
    
    // Load default empty state
    updateDashboardStats();
    
    // Load a default demo inside the phone frame simulator
    loadDefaultSimulatorPreview();
}

// =================================================================
// 2. DASHBOARD FUNCTIONALITY & EXCEL UPLOAD
// =================================================================
function initDashboardListeners() {
    const dropzone = document.getElementById("dropzone");
    const fileInput = document.getElementById("excel-file-input");
    const btnBrowse = document.getElementById("btn-browse-file");
    const btnProcess = document.getElementById("btn-process-excel");
    const btnSample = document.getElementById("btn-quick-sample");
    const downloadTemplateBtn = document.getElementById("download-template");
    const searchInput = document.getElementById("search-input");
    const btnFullscreenPreview = document.getElementById("btn-open-fullscreen-preview");

    // Click dropzone to browse
    dropzone.addEventListener("click", () => fileInput.click());
    btnBrowse.addEventListener("click", (e) => {
        e.stopPropagation();
        fileInput.click();
    });

    // Drag and Drop styles
    dropzone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropzone.classList.add("dragover");
    });

    dropzone.addEventListener("dragleave", () => {
        dropzone.classList.remove("dragover");
    });

    dropzone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropzone.classList.remove("dragover");
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            fileInput.files = files;
            handleFileSelection(files[0]);
        }
    });

    // File input change
    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleFileSelection(e.target.files[0]);
        }
    });

    // Process excel data
    btnProcess.addEventListener("click", () => {
        if (fileInput.files.length > 0) {
            parseExcel(fileInput.files[0]);
        }
    });

    // Quick Sample Data Button
    btnSample.addEventListener("click", () => {
        guestsList = sampleGuests.map(g => ({
            name: g.name,
            demoUrl: `https://demo-undangan.com/andi-siti?to=${encodeURIComponent(g.name)}`,
            localUrl: `${window.location.origin}${window.location.pathname}?to=${encodeURIComponent(g.name)}`
        }));
        renderGuestsTable();
        updateDashboardStats();
        
        // Auto preview first guest
        if (guestsList.length > 0) {
            previewGuestInSimulator(guestsList[0].name);
        }
    });

    // Download Sample Excel Template
    downloadTemplateBtn.addEventListener("click", (e) => {
        e.preventDefault();
        downloadExcelTemplate();
    });

    // Search bar functionality
    searchInput.addEventListener("input", (e) => {
        const query = e.target.value.toLowerCase();
        renderGuestsTable(query);
    });

    // Fullscreen simulation preview button
    btnFullscreenPreview.addEventListener("click", () => {
        const localUrl = `${window.location.origin}${window.location.pathname}?to=${encodeURIComponent(activePreviewGuest)}`;
        window.open(localUrl, "_blank");
    });
}

// File selected visual feedback
function handleFileSelection(file) {
    const dropzone = document.getElementById("dropzone");
    const btnProcess = document.getElementById("btn-process-excel");
    
    dropzone.innerHTML = `
        <i class="fa-solid fa-file-circle-check dropzone-icon" style="color: #4F46E5;"></i>
        <h3>File Terpilih: ${file.name}</h3>
        <p>Ukuran: ${(file.size / 1024).toFixed(2)} KB</p>
        <span class="file-info-badge" style="background-color: #EEF2F6; color: #4F46E5; border: 1px solid #E2E8F0;">Siap diproses</span>
    `;
    btnProcess.removeAttribute("disabled");
}

// Parse Excel file using SheetJS
function parseExcel(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Get first worksheet
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // Convert worksheet to JSON array
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            if (jsonData.length === 0) {
                alert("File Excel kosong!");
                return;
            }

            // Find key that contains "Nama" or matches "Nama Tamu"
            let guestNameKey = null;
            const firstRow = jsonData[0];
            
            // Log keys to find match
            const keys = Object.keys(firstRow);
            guestNameKey = keys.find(k => k.toLowerCase().includes("nama")) || keys[0];

            if (!guestNameKey) {
                alert("Kolom 'Nama Tamu' tidak ditemukan. Pastikan baris pertama Excel Anda memiliki judul kolom seperti 'Nama Tamu'.");
                return;
            }

            // Populate guests list
            guestsList = jsonData
                .map(row => {
                    const name = row[guestNameKey];
                    if (name && name.toString().trim() !== "") {
                        const trimmedName = name.toString().trim();
                        return {
                            name: trimmedName,
                            demoUrl: `https://demo-undangan.com/andi-siti?to=${encodeURIComponent(trimmedName)}`,
                            localUrl: `${window.location.origin}${window.location.pathname}?to=${encodeURIComponent(trimmedName)}`
                        };
                    }
                    return null;
                })
                .filter(g => g !== null);

            renderGuestsTable();
            updateDashboardStats();
            
            // Auto preview first guest
            if (guestsList.length > 0) {
                previewGuestInSimulator(guestsList[0].name);
            }
            
            // Toast Success
            showToast("Berhasil memuat " + guestsList.length + " tamu dari Excel!");
            
        } catch (error) {
            console.error(error);
            alert("Gagal membaca file Excel. Pastikan format file Anda benar (.xlsx atau .xls).");
        }
    };
    reader.readAsArrayBuffer(file);
}

// Render the Guest Table
function renderGuestsTable(filterQuery = "") {
    const tableBody = document.getElementById("guest-table-body");
    
    // Filter list if query exists
    const filteredList = guestsList.filter(g => g.name.toLowerCase().includes(filterQuery));

    if (filteredList.length === 0) {
        tableBody.innerHTML = `
            <tr class="empty-state-row">
                <td colspan="4">
                    <div class="empty-state">
                        <i class="fa-solid fa-magnifying-glass"></i>
                        <p>${guestsList.length === 0 ? "Belum ada data tamu. Upload Excel atau pilih data contoh." : "Tamu dengan nama '" + filterQuery + "' tidak ditemukan."}</p>
                    </div>
                </td>
            </tr>
        `;
        return;
    }

    tableBody.innerHTML = "";
    filteredList.forEach((guest, index) => {
        const row = document.createElement("tr");
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${escapeHTML(guest.name)}</strong></td>
            <td>
                <div class="link-display-wrapper">
                    <span class="link-text" id="demo-link-${index}">${escapeHTML(guest.demoUrl)}</span>
                    <button class="copy-icon-btn" onclick="copyLinkText('${escapeHTML(guest.demoUrl)}')" title="Salin Link Resmi">
                        <i class="fa-regular fa-copy"></i>
                    </button>
                </div>
            </td>
            <td class="text-center">
                <div class="action-buttons-cell">
                    <button class="btn btn-outline btn-sm" onclick="previewGuestInSimulator('${escapeJS(guest.name)}')" title="Preview di Ponsel Simulator">
                        <i class="fa-solid fa-mobile-screen"></i> Preview
                    </button>
                    <button class="btn btn-primary btn-sm" onclick="openFullscreenPreview('${escapeJS(guest.name)}')" title="Buka Fullscreen di Tab Baru">
                        <i class="fa-solid fa-arrow-up-right-from-square"></i> Buka
                    </button>
                </div>
            </td>
        `;
        tableBody.appendChild(row);
    });
}

// Update stats count widgets
function updateDashboardStats() {
    document.getElementById("stat-total-guests").textContent = guestsList.length;
    document.getElementById("stat-generated-links").textContent = guestsList.length;
    document.getElementById("stat-excel-status").textContent = guestsList.length > 0 ? "Aktif" : "Belum Ada";
}

// Generate & Download template Excel file for user convenience
function downloadExcelTemplate() {
    // Column header
    const data = [
        { "Nama Tamu": "Budi Santoso" },
        { "Nama Tamu": "Siti Aminah" },
        { "Nama Tamu": "Ahmad Fauzi" },
        { "Nama Tamu": "Pak RT" },
        { "Nama Tamu": "Bu RW" }
    ];
    
    // Create new sheet
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Daftar Tamu");
    
    // Save to file
    XLSX.writeFile(workbook, "template_tamu_undangan.xlsx");
}

// Copy URL to Clipboard
function copyLinkText(text) {
    navigator.clipboard.writeText(text).then(() => {
        showToast("Link berhasil disalin!");
    }).catch(err => {
        console.error('Failed to copy text: ', err);
    });
}

// Open Fullscreen local preview
function openFullscreenPreview(guestName) {
    const localUrl = `${window.location.origin}${window.location.pathname}?to=${encodeURIComponent(guestName)}`;
    window.open(localUrl, "_blank");
}

// Show Toast Alert
function showToast(message) {
    const toast = document.getElementById("toast");
    toast.textContent = message;
    toast.classList.add("show");
    setTimeout(() => {
        toast.classList.remove("show");
    }, 2500);
}

// =================================================================
// 3. INTERACTIVE SIMULATOR RENDERING
// =================================================================
function loadDefaultSimulatorPreview() {
    const simContainer = document.getElementById("mobile-preview-container");
    simContainer.innerHTML = `
        <div class="preview-placeholder">
            <i class="fa-solid fa-envelope-open-text"></i>
            <h3>Simulasi Undangan</h3>
            <p>Klik tombol <strong>Preview</strong> pada baris tabel tamu untuk melihat simulasi personalisasi undangan di sini.</p>
        </div>
    `;
    document.getElementById("previewing-indicator").textContent = "Menampilkan: Tamu Undangan";
}

// Render dynamic invitation view directly inside the smartphone mock frame
function previewGuestInSimulator(guestName) {
    activePreviewGuest = guestName;
    document.getElementById("previewing-indicator").textContent = `Menampilkan: ${guestName}`;
    
    const simContainer = document.getElementById("mobile-preview-container");
    
    // Generate inside simulator frame a full simulation of the invitation flow
    // We will render a micro cover page first. Clicking open will render the whole invitation page!
    renderSimulatorCover(guestName);
}

function renderSimulatorCover(guestName) {
    const simContainer = document.getElementById("mobile-preview-container");
    
    // HTML structure mimic for the cover inside smartphone
    simContainer.innerHTML = `
        <div class="sim-cover-wrap" style="
            position: relative; 
            height: 100%; 
            display: flex; 
            flex-direction: column; 
            justify-content: center; 
            align-items: center; 
            text-align: center; 
            padding: 20px; 
            color: #ffffff; 
            background: linear-gradient(to bottom, rgba(44, 38, 31, 0.4), rgba(44, 38, 31, 0.8)), url('assets/images/hero-wedding.png');
            background-size: cover;
            background-position: center;
        ">
            <span style="font-family: var(--font-display); font-size: 9px; letter-spacing: 2px; margin-bottom: 12px; display: block;">THE WEDDING OF</span>
            <h1 style="font-family: var(--font-script); font-size: 36px; margin-bottom: 20px;">Andi & Siti</h1>
            
            <div style="
                background: rgba(255, 255, 255, 0.15); 
                backdrop-filter: blur(8px); 
                -webkit-backdrop-filter: blur(8px); 
                border: 1px solid rgba(255, 255, 255, 0.25); 
                border-radius: var(--radius-md); 
                padding: 16px; 
                margin-bottom: 24px; 
                width: 100%;
            ">
                <p style="font-size: 9px; letter-spacing: 1px; text-transform: uppercase; color: var(--color-primary-light); margin-bottom: 8px;">Kepada Yth.</p>
                <h2 style="font-family: var(--font-serif); font-size: 18px; margin-bottom: 8px; color: #ffffff;">${escapeHTML(guestName)}</h2>
                <p style="font-size: 9px; opacity: 0.8; line-height: 1.4;">Di Tempat</p>
            </div>

            <button id="sim-btn-open" class="btn btn-gold btn-sm" style="font-size: 11px; padding: 10px 20px; box-shadow: 0 4px 10px rgba(0,0,0,0.3);">
                <i class="fa-solid fa-envelope-open"></i> Buka Undangan
            </button>
        </div>
    `;
    
    // Add listener for open button in simulator
    document.getElementById("sim-btn-open").addEventListener("click", () => {
        renderSimulatorMainContent(guestName);
    });
}

function renderSimulatorMainContent(guestName) {
    const simContainer = document.getElementById("mobile-preview-container");
    
    // Gather wishes list HTML
    let wishesHtml = wishesList.map(w => `
        <div class="wish-item" style="padding: 10px; margin-bottom: 10px; font-size: 11px;">
            <div class="wish-header" style="margin-bottom: 4px;">
                <span class="wish-author">${escapeHTML(w.name)}</span>
                <span class="wish-badge presence-${w.status}" style="font-size: 8px; padding: 1px 6px;">${w.status}</span>
            </div>
            <p class="wish-text" style="font-size: 11px;">${escapeHTML(w.message)}</p>
            <span class="wish-time" style="font-size: 8px; margin-top: 4px;">${w.time}</span>
        </div>
    `).join("");

    simContainer.innerHTML = `
        <div class="sim-main-wrap" style="background-color: var(--color-bg-cream); color: var(--color-text-dark); text-align: center;">
            
            <!-- Mobile Header Indicator -->
            <div style="background-color: var(--color-bg-darkcream); padding: 8px; font-size: 10px; border-bottom: 1px solid var(--color-primary-light); color: var(--color-text-muted);">
                <i class="fa-solid fa-circle-info"></i> Mode Simulasi (Tamu: <strong>${escapeHTML(guestName)}</strong>)
            </div>

            <!-- Hero Section -->
            <div style="position: relative; height: 250px; display: flex; align-items: center; justify-content: center; background-image: url('assets/images/hero-wedding.png'); background-size: cover; background-position: center; color: white;">
                <div style="position: absolute; top:0; left:0; width:100%; height:100%; background: rgba(44, 38, 31, 0.5);"></div>
                <div style="position: relative; z-index: 2; padding: 12px;">
                    <div class="flower-top" style="width: 80px; height: 50px; margin-bottom: 10px;"></div>
                    <h2 style="font-family: var(--font-script); font-size: 36px; margin-bottom: 4px;">Andi & Siti</h2>
                    <p style="font-family: var(--font-display); font-size: 10px; letter-spacing: 2px;">18 OKTOBER 2026</p>
                    <div class="flower-bottom" style="width: 80px; height: 50px; margin-top: 10px;"></div>
                </div>
            </div>

            <!-- Bride Groom -->
            <div style="padding: 30px 16px;">
                <h3 style="font-family: var(--font-serif); font-size: 18px; margin-bottom: 20px;">Mempelai</h3>
                
                <div style="margin-bottom: 24px;">
                    <div style="width: 100px; height: 100px; border-radius: 50%; border: 2px solid var(--color-primary); margin: 0 auto 12px; overflow: hidden;">
                        <img src="assets/images/gallery-1.png" style="width:100%; height:100%; object-fit:cover;">
                    </div>
                    <h4 style="font-family: var(--font-serif); font-size: 14px;">Andi Pratama, S.T.</h4>
                    <p style="font-size: 10px; color: var(--color-text-muted);">Putra dari Bapak H. Ahmad Pratama & Ibu Hj. Aminah</p>
                </div>
                
                <div style="font-family: var(--font-script); font-size: 32px; color: var(--color-primary); margin: 10px 0;">&</div>

                <div>
                    <div style="width: 100px; height: 100px; border-radius: 50%; border: 2px solid var(--color-primary); margin: 0 auto 12px; overflow: hidden;">
                        <img src="assets/images/gallery-4.png" style="width:100%; height:100%; object-fit:cover;">
                    </div>
                    <h4 style="font-family: var(--font-serif); font-size: 14px;">Siti Rahma, S.E.</h4>
                    <p style="font-size: 10px; color: var(--color-text-muted);">Putri dari Bapak Dr. H. M. Yusuf & Ibu Hj. Ratna</p>
                </div>
            </div>

            <!-- Events -->
            <div style="padding: 30px 16px; background-color: var(--color-bg-darkcream);">
                <h3 style="font-family: var(--font-serif); font-size: 18px; margin-bottom: 20px;">Acara</h3>
                
                <div class="event-card" style="padding: 16px; margin-bottom: 16px;">
                    <h4 style="font-family: var(--font-serif); font-size: 14px; margin-bottom: 8px;">Akad Nikah</h4>
                    <p style="font-size: 11px; margin-bottom: 4px;"><i class="fa-regular fa-clock"></i> 08.00 - 10.00 WIB</p>
                    <p style="font-size: 11px; color: var(--color-text-muted);">Masjid Raya Baiturrahman</p>
                </div>

                <div class="event-card premium" style="padding: 16px;">
                    <h4 style="font-family: var(--font-serif); font-size: 14px; margin-bottom: 8px;">Resepsi</h4>
                    <p style="font-size: 11px; margin-bottom: 4px;"><i class="fa-regular fa-clock"></i> 11.00 - 16.00 WIB</p>
                    <p style="font-size: 11px; color: var(--color-text-muted);">Hotel Mulia Senayan, Jakarta</p>
                </div>
            </div>

            <!-- RSVP Form Simulated inside phone -->
            <div style="padding: 30px 16px;">
                <h3 style="font-family: var(--font-serif); font-size: 18px; margin-bottom: 8px;">Konfirmasi RSVP</h3>
                <p style="font-size: 11px; color: var(--color-text-muted); margin-bottom: 16px;">Konfirmasi langsung di demo</p>
                
                <div class="rsvp-card" style="padding: 16px; text-align: left;">
                    <div class="form-group" style="margin-bottom: 12px;">
                        <label style="font-size: 9px;">Nama Tamu</label>
                        <input type="text" class="form-control" style="padding: 8px; font-size: 11px;" value="${escapeHTML(guestName)}" readonly>
                    </div>
                    <div class="form-group" style="margin-bottom: 12px;">
                        <label style="font-size: 9px;">Konfirmasi</label>
                        <select id="sim-rsvp-status" class="form-control" style="padding: 8px; font-size: 11px;">
                            <option value="Hadir">Saya Akan Hadir</option>
                            <option value="Tidak Hadir">Maaf, Tidak Bisa Hadir</option>
                        </select>
                    </div>
                    <button id="sim-rsvp-submit" class="btn btn-gold btn-block btn-sm"><i class="fa-regular fa-paper-plane"></i> Kirim RSVP</button>
                </div>
            </div>

            <!-- Wishes List inside phone -->
            <div style="padding: 30px 16px; background-color: var(--color-bg-darkcream);">
                <h3 style="font-family: var(--font-serif); font-size: 18px; margin-bottom: 16px;">Ucapan & Doa</h3>
                
                <div class="wishes-card" style="padding: 16px; text-align: left; margin-bottom: 16px;">
                    <div class="form-group" style="margin-bottom: 10px;">
                        <label style="font-size: 9px;">Ucapan Anda</label>
                        <textarea id="sim-wish-msg" class="form-control" style="padding: 8px; font-size: 11px;" rows="3" placeholder="Tulis doa restu..."></textarea>
                    </div>
                    <button id="sim-wish-submit" class="btn btn-gold btn-block btn-sm"><i class="fa-regular fa-comment-dots"></i> Kirim Ucapan</button>
                </div>

                <div id="sim-wishes-list-container">
                    ${wishesHtml}
                </div>
            </div>

            <!-- Digital Gift -->
            <div style="padding: 30px 16px;">
                <h3 style="font-family: var(--font-serif); font-size: 18px; margin-bottom: 12px;">Kado Digital</h3>
                <div class="gift-card" style="padding: 16px; margin-bottom: 12px;">
                    <p style="font-size: 12px; font-weight: bold;">BCA - 8720493294</p>
                    <p style="font-size: 10px; color: var(--color-text-muted);">A.n. Andi Pratama</p>
                </div>
                <div class="gift-card" style="padding: 16px;">
                    <p style="font-size: 12px; font-weight: bold;">MANDIRI - 132002938492</p>
                    <p style="font-size: 10px; color: var(--color-text-muted);">A.n. Siti Rahma</p>
                </div>
            </div>

            <!-- Footer -->
            <div style="background-color: #2C261F; color: white; padding: 24px 12px; font-size: 10px;">
                <h4 style="font-family: var(--font-script); font-size: 24px; color: var(--color-primary); margin-bottom: 4px;">Andi & Siti</h4>
                <p style="opacity: 0.6; font-size: 9px;">Terima kasih atas doa restunya.</p>
            </div>
        </div>
    `;

    // Add submit actions in the simulator to make it interactive!
    document.getElementById("sim-rsvp-submit").addEventListener("click", () => {
        const status = document.getElementById("sim-rsvp-status").value;
        alert(`[SIMULASI] RSVP Terkirim!\nStatus: ${status} untuk tamu ${guestName}`);
    });

    document.getElementById("sim-wish-submit").addEventListener("click", () => {
        const msg = document.getElementById("sim-wish-msg").value;
        if (!msg.trim()) {
            alert("Harap masukkan pesan ucapan.");
            return;
        }

        // Add to wishes list state
        wishesList.unshift({
            name: guestName,
            status: "Hadir", // default to hadir in simulator submit
            message: msg,
            time: "Baru saja"
        });

        // Re-render simulator main content to reflect new wish
        renderSimulatorMainContent(guestName);
        showToast("Ucapan berhasil ditambahkan ke simulasi!");
    });
}

// =================================================================
// 4. INVITATION VIEW LOGIC & INTERACTION
// =================================================================
function initInvitationListeners() {
    const btnOpen = document.getElementById("btn-open-invitation");
    const musicToggleBtn = document.getElementById("music-toggle-btn");
    const bgMusic = document.getElementById("bg-music");
    const rsvpForm = document.getElementById("rsvp-form");
    const wishForm = document.getElementById("wish-form");
    
    // Open Invitation Button Click
    if (btnOpen) {
        btnOpen.addEventListener("click", () => {
            const cover = document.getElementById("cover-section");
            const mainContent = document.getElementById("invitation-main-content");
            
            // Hide cover with sliding animation effect
            cover.style.transition = "transform 1s cubic-bezier(0.77, 0, 0.175, 1), opacity 0.8s ease";
            cover.style.transform = "translateY(-100%)";
            cover.style.opacity = "0";
            
            setTimeout(() => {
                cover.classList.add("hidden");
                mainContent.classList.remove("hidden");
                
                // Play Background Music
                playMusic();
            }, 1000);
        });
    }

    // Music control toggle
    if (musicToggleBtn && bgMusic) {
        musicToggleBtn.addEventListener("click", () => {
            if (bgMusic.paused) {
                playMusic();
            } else {
                pauseMusic();
            }
        });
    }

    // Copy Account Numbers
    document.querySelectorAll(".btn-copy-bank").forEach(btn => {
        btn.addEventListener("click", () => {
            const targetId = btn.getAttribute("data-target");
            const accountNum = document.getElementById(targetId).textContent;
            navigator.clipboard.writeText(accountNum).then(() => {
                showToast("Nomor rekening berhasil disalin!");
            });
        });
    });

    // Gallery Lightbox Modal
    const modal = document.getElementById("lightbox-modal");
    const modalImg = document.getElementById("lightbox-img");
    const closeBtn = document.querySelector(".close-lightbox");

    document.querySelectorAll(".gallery-item").forEach(item => {
        item.addEventListener("click", () => {
            const img = item.querySelector("img");
            modal.style.display = "block";
            modalImg.src = img.src;
        });
    });

    if (closeBtn) {
        closeBtn.addEventListener("click", () => {
            modal.style.display = "none";
        });
    }

    if (modal) {
        modal.addEventListener("click", (e) => {
            if (e.target === modal) {
                modal.style.display = "none";
            }
        });
    }

    // Real RSVP Form Submission
    if (rsvpForm) {
        rsvpForm.addEventListener("submit", (e) => {
            e.preventDefault();
            const guestName = document.getElementById("rsvp-name").value;
            const status = document.getElementById("rsvp-status").value;
            const guestsNum = document.getElementById("rsvp-guests").value;

            alert(`Terima kasih atas konfirmasi Anda!\n\nNama: ${guestName}\nStatus: ${status}\nJumlah: ${guestsNum} Orang`);
            
            // Add a temporary wish if they say they will attend
            if (status === "Hadir") {
                const autoWishMessage = "Terima kasih atas undangannya. Kami akan hadir membawa kebahagiaan bersama mempelai!";
                addCustomWish(guestName, status, autoWishMessage);
            }
        });
    }

    // Real Wish Form Submission
    if (wishForm) {
        wishForm.addEventListener("submit", (e) => {
            e.preventDefault();
            const name = document.getElementById("wish-name").value;
            const message = document.getElementById("wish-message").value;

            if (!message.trim()) return;

            addCustomWish(name, "Hadir", message);
            document.getElementById("wish-message").value = "";
            showToast("Ucapan doa restu berhasil dikirim!");
        });
    }
}

// Add Wish & refresh feeds
function addCustomWish(name, status, message) {
    wishesList.unshift({
        name: name,
        status: status,
        message: message,
        time: "Baru saja"
    });
    renderWishes("wishes-list-container");
}

// Render Wishes Feed
function renderWishes(containerId) {
    const wishesContainer = document.getElementById(containerId);
    if (!wishesContainer) return;

    wishesContainer.innerHTML = "";
    wishesList.forEach(wish => {
        const item = document.createElement("div");
        item.className = "wish-item";
        item.innerHTML = `
            <div class="wish-header">
                <span class="wish-author">${escapeHTML(wish.name)}</span>
                <span class="wish-badge presence-${wish.status}"><i class="fa-solid fa-circle-check"></i> ${wish.status}</span>
            </div>
            <p class="wish-text">${escapeHTML(wish.message)}</p>
            <span class="wish-time">${wish.time}</span>
        `;
        wishesContainer.appendChild(item);
    });
}

// Play Music logic
function playMusic() {
    const bgMusic = document.getElementById("bg-music");
    const btn = document.getElementById("music-toggle-btn");
    if (bgMusic && btn) {
        bgMusic.play().then(() => {
            btn.classList.add("playing");
        }).catch(err => {
            console.log("Audio autoplay was blocked. Waiting for user interaction.", err);
        });
    }
}

// Pause Music logic
function pauseMusic() {
    const bgMusic = document.getElementById("bg-music");
    const btn = document.getElementById("music-toggle-btn");
    if (bgMusic && btn) {
        bgMusic.pause();
        btn.classList.remove("playing");
    }
}

// =================================================================
// 5. HELPER FUNCTIONS & COUNTDOWN TIMER
// =================================================================

// Countdown Timer calculation (Target: October 18, 2026 at 08:00 WIB)
function startCountdown() {
    const targetDate = new Date("October 18, 2026 08:00:00").getTime();

    // Update countdown every second
    const timerInterval = setInterval(() => {
        const now = new Date().getTime();
        const distance = targetDate - now;

        if (distance < 0) {
            clearInterval(timerInterval);
            document.getElementById("cd-days").textContent = "00";
            document.getElementById("cd-hours").textContent = "00";
            document.getElementById("cd-minutes").textContent = "00";
            document.getElementById("cd-seconds").textContent = "00";
            return;
        }

        const days = Math.floor(distance / (1000 * 60 * 60 * 24));
        const hours = Math.floor((distance % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
        const minutes = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
        const seconds = Math.floor((distance % (1000 * 60)) / 1000);

        // Format to double digit
        document.getElementById("cd-days").textContent = String(days).padStart(2, '0');
        document.getElementById("cd-hours").textContent = String(hours).padStart(2, '0');
        document.getElementById("cd-minutes").textContent = String(minutes).padStart(2, '0');
        document.getElementById("cd-seconds").textContent = String(seconds).padStart(2, '0');
    }, 1000);
}

// Helper to escape HTML tags to prevent XSS
function escapeHTML(str) {
    return str.replace(/[&<>'"]/g, 
        tag => ({
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            "'": '&#39;',
            '"': '&quot;'
        }[tag] || tag)
    );
}

// Helper to escape single quotes for Inline Javascript
function escapeJS(str) {
    return str.replace(/'/g, "\\'");
}
