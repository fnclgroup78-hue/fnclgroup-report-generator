// Frontend Logic for Manulife Excel Automation Dashboard & Access Control System

document.addEventListener("DOMContentLoaded", () => {
    // Determine API Origin
    const apiOrigin = window.location.protocol === "file:" || window.location.origin === "null"
        ? "http://127.0.0.1:8080"
        : "";

    // Show protocol warning if loaded directly from file explorer
    const warningBanner = document.getElementById("protocol-warning");
    if (window.location.protocol === "file:" && warningBanner) {
        warningBanner.classList.remove("hidden");
    }

    checkBackendHealth();

    async function checkBackendHealth() {
        try {
            const res = await fetch(`${apiOrigin}/api/health`);
            if (res.ok) {
                if (window.location.protocol === "file:" && warningBanner) {
                    warningBanner.innerHTML = `✅ Connected to backend server on port 8080. (For full browser security and features, we still recommend opening <a href="http://127.0.0.1:8080" class="underline hover:text-white">http://127.0.0.1:8080</a>).`;
                    warningBanner.className = "bg-teal-950/80 border-b border-teal-500/50 text-teal-200 py-3.5 px-4 text-center text-sm font-semibold relative z-50";
                }
            } else {
                showOfflineWarning();
            }
        } catch (err) {
            showOfflineWarning();
        }
    }

    function showOfflineWarning() {
        if (warningBanner) {
            warningBanner.innerHTML = `⚠️ Backend Server Offline: Could not connect to the automation server on port 8080. Please ensure you double-click <strong>run.bat</strong> in the project folder and keep the terminal window open.`;
            warningBanner.className = "bg-amber-950/80 border-b border-amber-500/50 text-amber-200 py-3.5 px-4 text-center text-sm font-semibold relative z-50";
            warningBanner.classList.remove("hidden");
        }
    }

    // AUTHENTICATION STATE & DOM ELEMENTS
    let authToken = localStorage.getItem("fncl_auth_token") || "";
    let currentUser = null;

    const openLoginBtn = document.getElementById("open-login-btn");
    const userInfoBar = document.getElementById("user-info-bar");
    const userDisplayName = document.getElementById("user-display-name");
    const userRoleBadge = document.getElementById("user-role-badge");
    const adminPanelBtn = document.getElementById("admin-panel-btn");
    const logoutBtn = document.getElementById("logout-btn");

    const authModal = document.getElementById("auth-modal");
    const tabLogin = document.getElementById("tab-login");
    const tabRegister = document.getElementById("tab-register");
    const loginForm = document.getElementById("login-form");
    const registerForm = document.getElementById("register-form");
    const authStatusMsg = document.getElementById("auth-status-msg");

    const adminModal = document.getElementById("admin-modal");
    const closeAdminBtn = document.getElementById("close-admin-btn");
    const adminUserTableBody = document.getElementById("admin-user-table-body");
    const adminCreateForm = document.getElementById("admin-create-form");

    // Initialize Auth Status
    checkAuthStatus();

    async function checkAuthStatus() {
        if (!authToken) {
            showLoggedOutUI();
            showAuthModal(true);
            return;
        }

        try {
            const res = await fetch(`${apiOrigin}/api/auth/me`, {
                headers: { "Authorization": `Bearer ${authToken}` }
            });

            if (res.ok) {
                const data = await res.json();
                currentUser = data.user;
                showLoggedInUI();
                showAuthModal(false);
            } else {
                // Token invalid or revoked
                const errData = await res.json().catch(() => ({}));
                clearAuthSession();
                showLoggedOutUI();
                showAuthStatusMsg(errData.detail || "Session expired or access revoked. Please sign in again.", true);
                showAuthModal(true);
            }
        } catch (err) {
            console.error("Auth status check error:", err);
        }
    }

    function showLoggedInUI() {
        if (!currentUser) return;
        if (openLoginBtn) openLoginBtn.classList.add("hidden");
        if (userInfoBar) userInfoBar.classList.remove("hidden");

        userDisplayName.textContent = currentUser.full_name || currentUser.username;
        userRoleBadge.textContent = currentUser.role.toUpperCase();

        if (currentUser.role === "admin") {
            userRoleBadge.className = "ml-2 px-1.5 py-0.5 bg-purple-950 text-purple-300 border border-purple-500/30 text-[10px] rounded uppercase font-bold";
            adminPanelBtn.classList.remove("hidden");
        } else {
            userRoleBadge.className = "ml-2 px-1.5 py-0.5 bg-teal-950 text-teal-300 border border-teal-500/30 text-[10px] rounded uppercase font-bold";
            adminPanelBtn.classList.add("hidden");
        }
    }

    function showLoggedOutUI() {
        currentUser = null;
        if (openLoginBtn) openLoginBtn.classList.remove("hidden");
        if (userInfoBar) userInfoBar.classList.add("hidden");
        if (adminPanelBtn) adminPanelBtn.classList.add("hidden");
    }

    function clearAuthSession() {
        authToken = "";
        currentUser = null;
        localStorage.removeItem("fncl_auth_token");
    }

    function showAuthModal(show = true) {
        if (show) {
            authModal.classList.remove("hidden");
        } else {
            authModal.classList.add("hidden");
        }
    }

    function showAuthStatusMsg(msg, isError = false) {
        authStatusMsg.textContent = msg;
        authStatusMsg.classList.remove("hidden", "bg-red-950/60", "text-red-300", "border-red-500/30", "bg-teal-950/60", "text-teal-300", "border-teal-500/30", "bg-amber-950/60", "text-amber-300", "border-amber-500/30");

        if (isError) {
            authStatusMsg.classList.add("bg-red-950/60", "text-red-300", "border", "border-red-500/30");
        } else {
            authStatusMsg.classList.add("bg-teal-950/60", "text-teal-300", "border", "border-teal-500/30");
        }
    }

    // AUTH MODAL TAB SWITCHING
    tabLogin.addEventListener("click", () => {
        tabLogin.className = "flex-1 py-2 text-sm font-semibold text-teal-400 border-b-2 border-teal-400 transition-all";
        tabRegister.className = "flex-1 py-2 text-sm font-semibold text-slate-400 border-b-2 border-transparent hover:text-slate-200 transition-all";
        loginForm.classList.remove("hidden");
        registerForm.classList.add("hidden");
        authStatusMsg.classList.add("hidden");
    });

    tabRegister.addEventListener("click", () => {
        tabRegister.className = "flex-1 py-2 text-sm font-semibold text-teal-400 border-b-2 border-teal-400 transition-all";
        tabLogin.className = "flex-1 py-2 text-sm font-semibold text-slate-400 border-b-2 border-transparent hover:text-slate-200 transition-all";
        registerForm.classList.remove("hidden");
        loginForm.classList.add("hidden");
        authStatusMsg.classList.add("hidden");
    });

    if (openLoginBtn) {
        openLoginBtn.addEventListener("click", () => showAuthModal(true));
    }

    // LOGIN SUBMIT
    loginForm.addEventListener("submit", async (e) => {
        e.preventDefault();
        const username = document.getElementById("login-username").value.trim();
        const password = document.getElementById("login-password").value;

        showAuthStatusMsg("Authenticating...", false);

        try {
            const res = await fetch(`${apiOrigin}/api/auth/login`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ username_or_email: username, password: password })
            });

            const data = await res.json();
            if (res.ok) {
                authToken = data.token;
                currentUser = data.user;
                localStorage.setItem("fncl_auth_token", authToken);
                showLoggedInUI();
                showAuthModal(false);
                loginForm.reset();
                authStatusMsg.classList.add("hidden");
            } else {
                showAuthStatusMsg(data.detail || "Login failed.", true);
            }
        } catch (err) {
            showAuthStatusMsg("Failed to connect to server.", true);
        }
    });

    // REGISTER SUBMIT
    registerForm.addEventListener("submit", async (e) => {
        e.preventDefault();
        const fullName = document.getElementById("reg-name").value.trim();
        const username = document.getElementById("reg-username").value.trim();
        const email = document.getElementById("reg-email").value.trim();
        const password = document.getElementById("reg-password").value;

        showAuthStatusMsg("Creating account...", false);

        try {
            const res = await fetch(`${apiOrigin}/api/auth/register`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ full_name: fullName, username: username, email: email, password: password })
            });

            const data = await res.json();
            if (res.ok) {
                showAuthStatusMsg("⏳ " + (data.message || "Registration successful! Awaiting admin approval."), false);
                registerForm.reset();
            } else {
                showAuthStatusMsg(data.detail || "Registration failed.", true);
            }
        } catch (err) {
            showAuthStatusMsg("Failed to connect to server.", true);
        }
    });

    // LOGOUT HANDLER
    logoutBtn.addEventListener("click", async () => {
        if (authToken) {
            try {
                await fetch(`${apiOrigin}/api/auth/logout`, {
                    method: "POST",
                    headers: { "Authorization": `Bearer ${authToken}` }
                });
            } catch (err) {}
        }
        clearAuthSession();
        showLoggedOutUI();
        showAuthModal(true);
    });

    // ADMIN PANEL MODAL HANDLERS
    adminPanelBtn.addEventListener("click", () => {
        adminModal.classList.remove("hidden");
        loadAdminUserList();
    });

    closeAdminBtn.addEventListener("click", () => {
        adminModal.classList.add("hidden");
    });

    async function loadAdminUserList() {
        if (!authToken) return;
        adminUserTableBody.innerHTML = `<tr><td colspan="5" class="p-4 text-center text-slate-500">Loading user database...</td></tr>`;

        try {
            const res = await fetch(`${apiOrigin}/api/admin/users`, {
                headers: { "Authorization": `Bearer ${authToken}` }
            });

            if (!res.ok) {
                adminUserTableBody.innerHTML = `<tr><td colspan="5" class="p-4 text-center text-red-400">Failed to load user list. Admin access required.</td></tr>`;
                return;
            }

            const data = await res.json();
            renderUserTable(data.users || []);
        } catch (err) {
            adminUserTableBody.innerHTML = `<tr><td colspan="5" class="p-4 text-center text-red-400">Error connecting to server.</td></tr>`;
        }
    }

    function renderUserTable(users) {
        if (users.length === 0) {
            adminUserTableBody.innerHTML = `<tr><td colspan="5" class="p-4 text-center text-slate-500">No users found.</td></tr>`;
            return;
        }

        adminUserTableBody.innerHTML = users.map(u => {
            let statusBadge = "";
            let actionBtn = "";

            if (u.status === "active") {
                statusBadge = `<span class="px-2 py-0.5 bg-green-950 text-green-300 border border-green-500/30 rounded font-semibold text-[10px]">ACTIVE 🟢</span>`;
                actionBtn = `<button text-xs class="toggle-status-btn px-2.5 py-1 bg-red-950/60 hover:bg-red-900/80 border border-red-500/40 text-red-300 rounded font-semibold" data-id="${u.id}" data-target-status="revoked">🔴 Revoke Access</button>`;
            } else if (u.status === "pending") {
                statusBadge = `<span class="px-2 py-0.5 bg-amber-950 text-amber-300 border border-amber-500/30 rounded font-semibold text-[10px]">PENDING APPROVAL 🟡</span>`;
                actionBtn = `<button text-xs class="toggle-status-btn px-2.5 py-1 bg-teal-950/80 hover:bg-teal-900 border border-teal-500/40 text-teal-300 rounded font-semibold" data-id="${u.id}" data-target-status="active">🟢 Approve User</button>`;
            } else {
                statusBadge = `<span class="px-2 py-0.5 bg-red-950 text-red-300 border border-red-500/30 rounded font-semibold text-[10px]">REVOKED 🔴</span>`;
                actionBtn = `<button text-xs class="toggle-status-btn px-2.5 py-1 bg-teal-950/80 hover:bg-teal-900 border border-teal-500/40 text-teal-300 rounded font-semibold" data-id="${u.id}" data-target-status="active">🟢 Restore Access</button>`;
            }

            const deleteBtn = u.id === currentUser?.id
                ? `<span class="text-slate-600 text-[10px] ml-2">(You)</span>`
                : `<button class="delete-user-btn text-xs px-2 py-1 bg-slate-800 hover:bg-red-950 text-slate-400 hover:text-red-300 rounded ml-2 border border-slate-700 hover:border-red-500/40" data-id="${u.id}">🗑️ Delete</button>`;

            return `
                <tr class="hover:bg-slate-800/40">
                    <td class="p-3 font-semibold text-slate-200">${u.full_name}</td>
                    <td class="p-3 font-mono text-slate-400">${u.username} <br/><span class="text-[10px] text-slate-500">${u.email}</span></td>
                    <td class="p-3"><span class="uppercase text-[10px] font-bold ${u.role === 'admin' ? 'text-purple-400' : 'text-slate-400'}">${u.role}</span></td>
                    <td class="p-3">${statusBadge}</td>
                    <td class="p-3 text-right">${actionBtn}${deleteBtn}</td>
                </tr>
            `;
        }).join("");

        // Attach action listeners
        document.querySelectorAll(".toggle-status-btn").forEach(btn => {
            btn.addEventListener("click", async (e) => {
                const userId = e.currentTarget.getAttribute("data-id");
                const targetStatus = e.currentTarget.getAttribute("data-target-status");
                await updateUserStatus(userId, targetStatus);
            });
        });

        document.querySelectorAll(".delete-user-btn").forEach(btn => {
            btn.addEventListener("click", async (e) => {
                const userId = e.currentTarget.getAttribute("data-id");
                if (confirm("Are you sure you want to permanently delete this user?")) {
                    await deleteUser(userId);
                }
            });
        });
    }

    async function updateUserStatus(userId, newStatus) {
        try {
            const res = await fetch(`${apiOrigin}/api/admin/users/status`, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${authToken}`
                },
                body: JSON.stringify({ user_id: parseInt(userId), status: newStatus })
            });

            if (res.ok) {
                loadAdminUserList();
            } else {
                alert("Failed to update user status.");
            }
        } catch (err) {
            alert("Error communicating with server.");
        }
    }

    async function deleteUser(userId) {
        try {
            const res = await fetch(`${apiOrigin}/api/admin/users/${userId}`, {
                method: "DELETE",
                headers: { "Authorization": `Bearer ${authToken}` }
            });

            if (res.ok) {
                loadAdminUserList();
            } else {
                alert("Failed to delete user.");
            }
        } catch (err) {
            alert("Error communicating with server.");
        }
    }

    // ADMIN CREATE USER FORM SUBMIT
    adminCreateForm.addEventListener("submit", async (e) => {
        e.preventDefault();
        const fullName = document.getElementById("admin-add-name").value.trim();
        const username = document.getElementById("admin-add-username").value.trim();
        const email = document.getElementById("admin-add-email").value.trim();
        const password = document.getElementById("admin-add-password").value;

        try {
            const res = await fetch(`${apiOrigin}/api/admin/users/create`, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${authToken}`
                },
                body: JSON.stringify({
                    full_name: fullName,
                    username: username,
                    email: email,
                    password: password,
                    role: "user",
                    status: "active"
                })
            });

            const data = await res.json();
            if (res.ok) {
                adminCreateForm.reset();
                loadAdminUserList();
            } else {
                alert(data.detail || "Failed to create user.");
            }
        } catch (err) {
            alert("Error communicating with server.");
        }
    });

    // CHANGE CREDENTIALS FORM SUBMIT
    const changeCredForm = document.getElementById("change-credentials-form");
    const changeCredStatus = document.getElementById("change-cred-status");

    if (changeCredForm) {
        changeCredForm.addEventListener("submit", async (e) => {
            e.preventDefault();
            const newUsername = document.getElementById("change-new-username").value.trim();
            const newPassword = document.getElementById("change-new-password").value;
            const currentPassword = document.getElementById("change-curr-password").value;

            changeCredStatus.classList.remove("hidden", "text-red-400", "text-teal-400");
            changeCredStatus.classList.add("text-slate-400");
            changeCredStatus.textContent = "Updating credentials...";

            try {
                const res = await fetch(`${apiOrigin}/api/auth/change-credentials`, {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                        "Authorization": `Bearer ${authToken}`
                    },
                    body: JSON.stringify({
                        current_password: currentPassword,
                        new_username: newUsername || null,
                        new_password: newPassword || null
                    })
                });

                const data = await res.json();
                if (res.ok) {
                    changeCredStatus.classList.remove("text-slate-400");
                    changeCredStatus.classList.add("text-teal-400");
                    changeCredStatus.textContent = "✅ " + data.message;
                    changeCredForm.reset();
                    checkAuthStatus();
                } else {
                    changeCredStatus.classList.remove("text-slate-400");
                    changeCredStatus.classList.add("text-red-400");
                    changeCredStatus.textContent = "❌ " + (data.detail || "Failed to update credentials.");
                }
            } catch (err) {
                changeCredStatus.classList.remove("text-slate-400");
                changeCredStatus.classList.add("text-red-400");
                changeCredStatus.textContent = "Error communicating with server.";
            }
        });
    }

    // DOWNLOAD BLANK EXCEL TEMPLATE HANDLER
    const downloadTemplateBtn = document.getElementById("download-template-btn");
    if (downloadTemplateBtn) {
        downloadTemplateBtn.addEventListener("click", async (e) => {
            e.stopPropagation();
            if (!authToken) {
                showAuthStatusMsg("Please sign in to download the blank template.", true);
                showAuthModal(true);
                return;
            }

            try {
                const res = await fetch(`${apiOrigin}/api/template/download?type=main`, {
                    headers: { "Authorization": `Bearer ${authToken}` }
                });

                if (!res.ok) {
                    const errData = await res.json().catch(() => ({ detail: "Failed to download template." }));
                    alert(errData.detail || "Error downloading template.");
                    return;
                }

                const blob = await res.blob();
                const downloadUrl = window.URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = downloadUrl;
                a.download = "FNCL ML MAIN REPORT.xlsx";
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                window.URL.revokeObjectURL(downloadUrl);
            } catch (err) {
                alert("Error connecting to server to download template.");
            }
        });
    }


    // FILE DROP ZONE LOGIC & REPORT PROCESSING
    const pdfDropzone = document.getElementById("pdf-dropzone");
    const excelDropzone = document.getElementById("excel-dropzone");
    const pdfInput = document.getElementById("pdf-input");
    const excelInput = document.getElementById("excel-input");
    const pdfDefaultView = document.getElementById("pdf-default-view");
    const pdfSuccessView = document.getElementById("pdf-success-view");
    const excelDefaultView = document.getElementById("excel-default-view");
    const excelSuccessView = document.getElementById("excel-success-view");
    const pdfFileNameEl = document.getElementById("pdf-file-name");
    const pdfFileSizeEl = document.getElementById("pdf-file-size");
    const excelFileNameEl = document.getElementById("excel-file-name");
    const excelFileSizeEl = document.getElementById("excel-file-size");
    const clearPdf = document.getElementById("clear-pdf");
    const clearExcel = document.getElementById("clear-excel");
    const proceedBtn = document.getElementById("proceed-btn");
    const downloadBtn = document.getElementById("download-btn");
    const progressContainer = document.getElementById("progress-container");
    const progressText = document.getElementById("progress-text");
    const progressPercent = document.getElementById("progress-percent");
    const progressBar = document.getElementById("progress-bar");
    const errorMessage = document.getElementById("error-message");

    let selectedPdfFile = null;
    let selectedExcelFile = null;
    let generatedFileBlob = null;
    let generatedFileName = "FNCL Group - Generated_Report.xlsx";

    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    setupDragAndDrop(pdfDropzone, pdfInput, ".pdf", (file) => {
        selectedPdfFile = file;
        pdfFileNameEl.textContent = file.name;
        pdfFileSizeEl.textContent = formatFileSize(file.size);
        pdfDefaultView.classList.add("hidden");
        pdfSuccessView.classList.remove("hidden");
        pdfDropzone.classList.add("pdf-loaded");
        validateStates();
    });

    setupDragAndDrop(excelDropzone, excelInput, ".xlsx", (file) => {
        selectedExcelFile = file;
        excelFileNameEl.textContent = file.name;
        excelFileSizeEl.textContent = formatFileSize(file.size);
        excelDefaultView.classList.add("hidden");
        excelSuccessView.classList.remove("hidden");
        excelDropzone.classList.add("excel-loaded");
        validateStates();
    });

    clearPdf.addEventListener("click", (e) => {
        e.stopPropagation();
        selectedPdfFile = null;
        pdfInput.value = "";
        pdfDefaultView.classList.remove("hidden");
        pdfSuccessView.classList.add("hidden");
        pdfDropzone.classList.remove("pdf-loaded");
        validateStates();
        resetProgress();
    });

    clearExcel.addEventListener("click", (e) => {
        e.stopPropagation();
        selectedExcelFile = null;
        excelInput.value = "";
        excelDefaultView.classList.remove("hidden");
        excelSuccessView.classList.add("hidden");
        excelDropzone.classList.remove("excel-loaded");
        validateStates();
        resetProgress();
    });

    // PROCEED / GENERATE REPORT
    proceedBtn.addEventListener("click", async () => {
        if (!selectedPdfFile || !selectedExcelFile) return;

        if (!authToken) {
            showAuthStatusMsg("Please sign in to generate reports.", true);
            showAuthModal(true);
            return;
        }

        proceedBtn.disabled = true;
        proceedBtn.classList.remove("btn-active-teal");
        proceedBtn.classList.add("bg-slate-700", "text-slate-500", "cursor-not-allowed");

        errorMessage.classList.add("hidden");
        errorMessage.textContent = "";
        progressContainer.classList.remove("hidden");

        let progressVal = 0;
        let isBackendDone = false;

        const updateProgressUI = (val, text) => {
            progressPercent.textContent = `${val}%`;
            progressText.textContent = text;
            progressBar.style.width = `${val}%`;
        };

        const getStepText = (pct) => {
            if (pct >= 10 && pct < 40) return "Extracting and parsing multi-page PDF context.";
            if (pct >= 40 && pct < 70) return "Data structural cleaning, regex alignment, and edge-case correction.";
            if (pct >= 70 && pct < 90) return "Reading Excel template byte-by-byte and mapping values to rows.";
            if (pct >= 90 && pct <= 100) return "Finalizing file memory payload and activating download hook.";
            return "Initializing pipeline...";
        };

        const progressInterval = setInterval(() => {
            if (isBackendDone) return;
            if (progressVal < 95) {
                progressVal += 1;
                updateProgressUI(progressVal, getStepText(progressVal));
            }
        }, 100);

        const formData = new FormData();
        formData.append("pdf_file", selectedPdfFile);
        formData.append("excel_file", selectedExcelFile);

        try {
            const apiEndpoint = `${apiOrigin}/api/process`;

            const response = await fetch(apiEndpoint, {
                method: "POST",
                headers: {
                    "Authorization": `Bearer ${authToken}`
                },
                body: formData
            });

            isBackendDone = true;
            clearInterval(progressInterval);

            if (!response.ok) {
                const errJson = await response.json().catch(() => ({ detail: "An error occurred during sheet generation." }));

                // Handle Auth Expiry / Revocation during process call
                if (response.status === 401 || response.status === 403) {
                    clearAuthSession();
                    showLoggedOutUI();
                    showAuthStatusMsg(errJson.detail || "Access Revoked. Please sign in again.", true);
                    showAuthModal(true);
                }

                throw new Error(errJson.detail || "An error occurred during sheet generation.");
            }

            generatedFileBlob = await response.blob();

            const contentDisposition = response.headers.get("Content-Disposition");
            if (contentDisposition) {
                const match = contentDisposition.match(/filename="?([^"]+)"?/);
                if (match && match[1]) {
                    generatedFileName = match[1];
                }
            } else {
                const baseName = selectedExcelFile.name.substring(0, selectedExcelFile.name.lastIndexOf('.'));
                generatedFileName = `FNCL Group - ${baseName}.xlsx`;
            }

            const fastForwardInterval = setInterval(() => {
                if (progressVal < 100) {
                    progressVal += 1;
                    updateProgressUI(progressVal, getStepText(progressVal));
                } else {
                    clearInterval(fastForwardInterval);

                    downloadBtn.disabled = false;
                    downloadBtn.classList.remove("bg-slate-800", "text-slate-500", "cursor-not-allowed");
                    downloadBtn.classList.add("btn-active-green", "animate-scale-in");

                    const downloadSpan = downloadBtn.querySelector("span");
                    if (downloadSpan) downloadSpan.textContent = "Download Completed Report";

                    triggerFileDownload();
                }
            }, 15);

        } catch (err) {
            isBackendDone = true;
            clearInterval(progressInterval);
            progressContainer.classList.add("hidden");

            let displayMsg = err.message;
            if (err.name === "TypeError" || err.message === "Failed to fetch" || err.message.toLowerCase().includes("fetch")) {
                displayMsg = "Connection Error: Failed to communicate with the backend server. Please make sure that you launched the server by double-clicking 'run.bat' and that the terminal window is still open (running on Port 8080).";
            }
            errorMessage.textContent = displayMsg;
            errorMessage.classList.remove("hidden");

            validateStates();
        }
    });

    downloadBtn.addEventListener("click", () => {
        if (generatedFileBlob) {
            triggerFileDownload();
        }
    });

    function triggerFileDownload() {
        const downloadUrl = window.URL.createObjectURL(generatedFileBlob);
        const a = document.createElement("a");
        a.href = downloadUrl;
        a.download = generatedFileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(downloadUrl);
    }

    function validateStates() {
        if (selectedPdfFile && selectedExcelFile && authToken) {
            proceedBtn.disabled = false;
            proceedBtn.classList.remove("bg-slate-700", "text-slate-400", "cursor-not-allowed");
            proceedBtn.classList.add("btn-active-teal");
        } else {
            proceedBtn.disabled = true;
            proceedBtn.classList.add("bg-slate-700", "text-slate-400", "cursor-not-allowed");
            proceedBtn.classList.remove("btn-active-teal");
        }
    }

    function resetProgress() {
        progressContainer.classList.add("hidden");
        progressPercent.textContent = "0%";
        progressText.textContent = "Waiting...";
        progressBar.style.width = "0%";

        downloadBtn.disabled = true;
        downloadBtn.classList.add("bg-slate-800", "text-slate-500", "cursor-not-allowed");
        downloadBtn.classList.remove("btn-active-green", "animate-scale-in");
        generatedFileBlob = null;
    }

    function setupDragAndDrop(dropzone, input, acceptTypes, onFileSelect) {
        dropzone.addEventListener("click", () => {
            input.click();
        });

        input.addEventListener("change", (e) => {
            if (e.target.files.length > 0) {
                handleSelectedFile(e.target.files[0], acceptTypes, onFileSelect);
            }
        });

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
            if (e.dataTransfer.files.length > 0) {
                handleSelectedFile(e.dataTransfer.files[0], acceptTypes, onFileSelect);
            }
        });
    }

    function handleSelectedFile(file, acceptTypes, callback) {
        const allowedExtensions = acceptTypes.split(",").map(t => t.trim().toLowerCase());
        const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

        if (allowedExtensions.includes(fileExtension)) {
            callback(file);
        } else {
            alert(`Invalid file type. Please select a file matching: ${acceptTypes}`);
        }
    }
});
