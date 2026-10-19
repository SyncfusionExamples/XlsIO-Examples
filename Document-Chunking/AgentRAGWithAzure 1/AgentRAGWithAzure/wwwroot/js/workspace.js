const SUPPORTED = new Set([".pdf", ".doc", ".docx", ".xls", ".xlsx"]);
let hasIndexedFiles = false;
let busy = false;
let pdfViewer = null;
let wordContainer = null;
let spreadsheet = null;
let viewersReady = false;
let preparedFiles = [];
let lastOpenButton = null;
const SAMPLES = [
    "Summarize this document",
    "What are the main findings?",
    "List the important dates",
    "What does the document say about…?"
];

function extensionOf(name) {
    const i = (name || "").lastIndexOf(".");
    return i >= 0 ? name.slice(i).toLowerCase() : "";
}
function escapeHtml(text) {
    return String(text || "")
        .replace(/&/g, "&")
        .replace(/</g, "<")
        .replace(/>/g, ">")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}
function formatBytes(size) {
    const n = Number(size) || 0;
    if (n < 1024) return n + " B";
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
    return (n / (1024 * 1024)).toFixed(1) + " MB";
}
function typeLabel(name) {
    const ext = extensionOf(name);
    if (ext === ".pdf") return "PDF";
    if (ext === ".doc" || ext === ".docx") return "Word";
    if (ext === ".xls" || ext === ".xlsx") return "Excel";
    return "File";
}
function pluralDocs(count) {
    return count + " document" + (count === 1 ? " is" : "s are") + " available.";
}
function friendlyError(error, fallback) {
    const status = error && error.status;
    if (status === 429) return "Too many requests. Wait a moment and try again.";
    if (status === 408 || (error && error.name === "TimeoutError")) return "The service did not respond. Check your connection and try again.";
    if (error && error.name === "TypeError") return "The service did not respond. Check your connection and try again.";
    return fallback;
}
function setLive(el, message, kind) {
    if (!el) return;
    el.textContent = message || "";
    el.classList.remove("error", "ok", "info");
    if (kind) el.classList.add(kind);
}
function setStatus(el, message, kind) {
    if (!el) return;
    el.textContent = message || "";
    const keep = [];
    if (el.id === "viewerStatus" || el.id === "chatStatus") keep.push("sub");
    if (kind) keep.push("status-" + kind);
    el.className = keep.join(" ");
}

function setChatEnabled(isReady) {
    hasIndexedFiles = !!isReady;
    const send = document.getElementById("sendButton");
    const input = document.getElementById("questionInput");
    const hint = document.getElementById("composerHint");
    if (send) send.disabled = !hasIndexedFiles || busy;
    if (input) input.disabled = !hasIndexedFiles || busy;
    if (hint) {
        hint.textContent = hasIndexedFiles
            ? "Press Enter to send. Press Shift+Enter for a new line."
            : "Add a document before asking a question.";
    }
}
function setReadyUi(isReady, names, failedCheck) {
    const count = names.length;
    const pill = document.getElementById("readyPill");
    const pillText = document.getElementById("readyPillText");
    const chatStatus = document.getElementById("chatStatus");
    const title = document.getElementById("welcomeTitle");
    const body = document.getElementById("welcomeBody");
    const steps = document.getElementById("welcomeSteps");
    const samples = document.getElementById("sampleQuestions");
    if (pill) pill.classList.toggle("ready", isReady);
    if (pillText) pillText.textContent = failedCheck ? "Could not check documents" : (isReady ? "Ready. " + pluralDocs(count) : "Add a document to start");
    if (chatStatus) {
        chatStatus.textContent = failedCheck
            ? "Could not check your documents. Refresh the page and try again."
            : (isReady ? "Ready. " + pluralDocs(count) : "Add a PDF, Word, or Excel file, then ask a question.");
    }
    if (title && body && steps && samples) {
        if (isReady) {
            title.textContent = "Documents are ready";
            body.textContent = pluralDocs(count) + " Choose a sample question or write your own. Open the source to verify the answer.";
            steps.hidden = true;
            samples.hidden = false;
            if (!samples.childElementCount) {
                SAMPLES.forEach(function (q) {
                    const btn = document.createElement("button");
                    btn.type = "button";
                    btn.className = "sample-chip";
                    btn.textContent = q;
                    btn.addEventListener("click", function () {
                        const input = document.getElementById("questionInput");
                        input.value = q;
                        input.focus();
                    });
                    samples.appendChild(btn);
                });
            }
        } else {
            title.textContent = failedCheck ? "Documents could not be checked" : "Start with a document";
            body.textContent = failedCheck
                ? "Refresh the page and try again."
                : "Add a document, ask in plain language, then open the source to verify the answer.";
            steps.hidden = false;
            samples.hidden = true;
        }
    }
    renderReadyDocs(names);
}
function renderReadyDocs(names) {
    const host = document.getElementById("readyDocs");
    const list = document.getElementById("readyDocsList");
    if (!host || !list) return;
    list.replaceChildren();
    if (!names.length) {
        host.hidden = true;
        return;
    }
    host.hidden = false;
    names.forEach(function (name) {
        const li = document.createElement("li");
        li.className = "file-row";
        li.innerHTML = '<div><div class="file-name"></div><div class="file-meta"></div></div><div class="file-actions"><a class="text-btn doc-link">Download</a><button type="button" class="text-btn danger-btn">Delete</button></div>';
        li.querySelector(".file-name").textContent = name;
        li.querySelector(".file-meta").textContent = typeLabel(name);
        const link = li.querySelector(".doc-link");
        link.href = "/api/documents/" + encodeURIComponent(name) + "/stream";
        link.setAttribute("download", name);
        li.querySelector(".danger-btn").addEventListener("click", function () { deleteDocument(name); });
        list.appendChild(li);
    });
}
function renderSelected(files) {
    const list = document.getElementById("selectedList");
    const upload = document.getElementById("uploadButton");
    list.replaceChildren();
    const accepted = [];
    Array.from(files || []).forEach(function (file, index) {
        const ok = SUPPORTED.has(extensionOf(file.name));
        if (ok) accepted.push(file);
        const li = document.createElement("li");
        li.className = "file-row";
        const info = document.createElement("div");
        info.innerHTML = '<div class="file-name"></div><div class="file-meta"></div>';
        info.querySelector(".file-name").textContent = file.name;
        info.querySelector(".file-meta").textContent = typeLabel(file.name) + " · " + formatBytes(file.size);
        const status = document.createElement("span");
        status.className = "file-status " + (ok ? "ok" : "bad");
        status.textContent = ok ? "Ready to add" : "Not supported. Use PDF, Word, or Excel.";
        const remove = document.createElement("button");
        remove.type = "button";
        remove.className = "icon-btn";
        remove.textContent = "Remove";
        remove.addEventListener("click", function () {
            const input = document.getElementById("fileInput");
            const next = new DataTransfer();
            Array.from(input.files).forEach(function (f, i) { if (i !== index) next.items.add(f); });
            input.files = next.files;
            renderSelected(input.files);
        });
        li.append(info, status, remove);
        list.appendChild(li);
    });
    upload.disabled = accepted.length === 0;
}

async function refreshReadyState() {
    setChatEnabled(false);
    setReadyUi(false, [], false);
    try {
        const response = await fetch("/api/files", { cache: "no-store" });
        if (!response.ok) throw new Error("check");
        const data = await response.json();
        preparedFiles = Array.isArray(data.files) ? data.files : [];
        setChatEnabled(preparedFiles.length > 0);
        setReadyUi(preparedFiles.length > 0, preparedFiles, false);
    } catch (e) {
        setChatEnabled(false);
        setReadyUi(false, [], true);
    }
}

function assignFiles(fileList) {
    const input = document.getElementById("fileInput");
    const next = new DataTransfer();
    Array.from(fileList || []).forEach(function (file) { next.items.add(file); });
    input.files = next.files;
    renderSelected(input.files);
    const rejected = Array.from(input.files).filter(function (f) { return !SUPPORTED.has(extensionOf(f.name)); });
    if (rejected.length) {
        setLive(document.getElementById("uploadStatus"), rejected.map(function (f) { return f.name + " is not supported. Use PDF, Word, or Excel."; }).join(" "), "error");
    } else {
        setLive(document.getElementById("uploadStatus"), "", "");
    }
}

async function uploadFiles() {
    const input = document.getElementById("fileInput");
    const uploadButton = document.getElementById("uploadButton");
    const status = document.getElementById("uploadStatus");
    const progress = document.getElementById("uploadProgress");
    const files = Array.from(input.files || []);
    const accepted = files.filter(function (f) { return SUPPORTED.has(extensionOf(f.name)); });
    const rejected = files.filter(function (f) { return !SUPPORTED.has(extensionOf(f.name)); });
    if (!files.length) {
        setLive(status, "Choose a PDF, Word, or Excel file first.", "error");
        return;
    }
    if (!accepted.length) {
        setLive(status, rejected.map(function (f) { return f.name + " is not supported. Use PDF, Word, or Excel."; }).join(" "), "error");
        return;
    }
    const formData = new FormData();
    accepted.forEach(function (file) { formData.append("files", file); });
    uploadButton.disabled = true;
    progress.hidden = false;
    setLive(status, "", "");
    try {
        const response = await fetch("/api/upload", { method: "POST", body: formData });
        const payload = await response.json().catch(function () { return {}; });
        if (!response.ok) {
            const err = new Error("upload");
            err.status = response.status;
            throw err;
        }
        const added = payload.files || accepted.map(function (f) { return f.name; });
        let message = "Added " + added.join(", ") + ". You can ask a question.";
        if (rejected.length) {
            message += " " + rejected.map(function (f) { return f.name + " is not supported. Use PDF, Word, or Excel."; }).join(" ");
        }
        setLive(status, message, rejected.length ? "error" : "ok");
        input.value = "";
        renderSelected([]);
        await refreshReadyState();
    } catch (error) {
        setLive(status, friendlyError(error, "Could not prepare these documents. Try adding them again."), "error");
        uploadButton.disabled = accepted.length === 0;
    } finally {
        progress.hidden = true;
    }
}

function dismissWelcome() {
    const welcome = document.getElementById("welcomeCard");
    if (welcome) welcome.hidden = true;
}
function showWelcomeIfEmpty() {
    const welcome = document.getElementById("welcomeCard");
    const chat = document.getElementById("chatWindow");
    if (welcome && chat && !chat.querySelector(".msg")) welcome.hidden = false;
}
function scrollChatToBottom() {
    const chatWindow = document.getElementById("chatWindow");
    if (!chatWindow) return;
    requestAnimationFrame(function () { chatWindow.scrollTop = chatWindow.scrollHeight; });
}
function inlineFormat(text) {
    return escapeHtml(text)
        .replace(/`([^`]+)`/g, "<code>$1</code>")
        .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
        .replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2" rel="noopener noreferrer" target="_blank">$1</a>');
}
function renderMarkdown(src) {
    const lines = String(src || "").replace(/\r\n/g, "\n").split("\n");
    let html = "";
    let i = 0;
    function flushParagraph(buf) {
        if (!buf.length) return;
        html += "<p>" + buf.map(inlineFormat).join("<br>") + "</p>";
    }
    while (i < lines.length) {
        if (lines[i].startsWith("```")) {
            const buf = [];
            i += 1;
            while (i < lines.length && !lines[i].startsWith("```")) { buf.push(lines[i]); i += 1; }
            html += "<pre><code>" + escapeHtml(buf.join("\n")) + "</code></pre>";
            i += 1;
            continue;
        }
        if (lines[i].includes("|") && i + 1 < lines.length && /^\s*\|?\s*:?-{3,}/.test(lines[i + 1])) {
            const header = lines[i].split("|").map(function (c) { return c.trim(); }).filter(Boolean);
            i += 2;
            const rows = [];
            while (i < lines.length && lines[i].includes("|") && lines[i].trim()) {
                rows.push(lines[i].split("|").map(function (c) { return c.trim(); }).filter(function (c, idx, arr) { return !(idx === 0 && c === "") && !(idx === arr.length - 1 && c === ""); }));
                i += 1;
            }
            html += "<table><thead><tr>" + header.map(function (c) { return "<th>" + inlineFormat(c) + "</th>"; }).join("") + "</tr></thead><tbody>";
            rows.forEach(function (row) {
                html += "<tr>" + row.map(function (c) { return "<td>" + inlineFormat(c) + "</td>"; }).join("") + "</tr>";
            });
            html += "</tbody></table>";
            continue;
        }
        const heading = /^(#{1,3})\s+(.*)$/.exec(lines[i]);
        if (heading) {
            html += "<h" + heading[1].length + ">" + inlineFormat(heading[2]) + "</h" + heading[1].length + ">";
            i += 1;
            continue;
        }
        if (/^\s*[-*]\s+/.test(lines[i])) {
            html += "<ul>";
            while (i < lines.length && /^\s*[-*]\s+/.test(lines[i])) {
                html += "<li>" + inlineFormat(lines[i].replace(/^\s*[-*]\s+/, "")) + "</li>";
                i += 1;
            }
            html += "</ul>";
            continue;
        }
        if (/^\s*\d+\.\s+/.test(lines[i])) {
            html += "<ol>";
            while (i < lines.length && /^\s*\d+\.\s+/.test(lines[i])) {
                html += "<li>" + inlineFormat(lines[i].replace(/^\s*\d+\.\s+/, "")) + "</li>";
                i += 1;
            }
            html += "</ol>";
            continue;
        }
        if (!lines[i].trim()) { i += 1; continue; }
        const buf = [];
        while (i < lines.length && lines[i].trim() && !lines[i].startsWith("```") && !/^(#{1,3})\s+/.test(lines[i]) && !/^\s*[-*]\s+/.test(lines[i]) && !/^\s*\d+\.\s+/.test(lines[i])) {
            buf.push(lines[i]);
            i += 1;
        }
        flushParagraph(buf);
    }
    return '<div class="answer">' + (html || "<p></p>") + "</div>";
}

function formatCitationPlace(citation) {
    if (!citation) return "";
    if (citation.pageNumber) return "Page " + citation.pageNumber;
    if (citation.sheetName || citation.cellAddress) {
        if (citation.sheetName && citation.cellAddress) return citation.sheetName + "!" + citation.cellAddress;
        return citation.sheetName || citation.cellAddress || "";
    }
    if (citation.headingPath) return citation.headingPath;
    if (citation.citationDisplay && citation.sourceName && citation.citationDisplay !== citation.sourceName) {
        return citation.citationDisplay;
    }
    return "";
}

function addMessage(text, sender, citations, noSource) {
    dismissWelcome();
    const chatWindow = document.getElementById("chatWindow");
    const row = document.createElement("div");
    row.className = "msg " + sender;
    const avatar = document.createElement("div");
    avatar.className = "avatar";
    avatar.textContent = sender === "user" ? "You" : "AI";
    const wrap = document.createElement("div");
    wrap.className = "bubble-wrap";
    const bubble = document.createElement("div");
    bubble.className = "bubble";
    if (sender === "assistant") bubble.innerHTML = renderMarkdown(text);
    else bubble.textContent = text;
    wrap.appendChild(bubble);
    const time = document.createElement("div");
    time.className = "time";
    time.textContent = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    wrap.appendChild(time);
    if (sender === "assistant") {
        const copy = document.createElement("button");
        copy.type = "button";
        copy.className = "icon-btn copy-btn";
        copy.textContent = "Copy answer";
        const note = document.createElement("div");
        note.className = "copy-note";
        note.setAttribute("role", "status");
        copy.addEventListener("click", async function () {
            try {
                await navigator.clipboard.writeText(text || "");
                note.textContent = "Answer copied.";
            } catch (e) {
                note.textContent = "Could not copy the answer.";
            }
        });
        wrap.appendChild(copy);
        wrap.appendChild(note);
        const sources = Array.isArray(citations) ? citations.map(normalizeCitation).filter(function (c) { return c.sourceName || c.citationDisplay; }) : [];
        if (sources.length) {
            const list = document.createElement("div");
            list.className = "citations";
            const label = document.createElement("div");
            label.className = "citations-label";
            label.textContent = sources.length > 1 ? "Sources used for this answer" : "Source used for this answer";
            list.appendChild(label);
            sources.forEach(function (citation) {
                const card = document.createElement("div");
                card.className = "citation-card";
                const title = document.createElement("strong");
                title.textContent = citation.sourceName || "Source document";
                card.appendChild(title);
                const placeText = formatCitationPlace(citation);
                if (placeText && placeText !== "Exact location") {
                    const place = document.createElement("div");
                    place.className = "citation-place";
                    place.textContent = placeText;
                    card.appendChild(place);
                }
                if (citation.contentSnippet) {
                    const snip = document.createElement("div");
                    snip.className = "citation-snippet";
                    snip.textContent = citation.contentSnippet;
                    card.appendChild(snip);
                }
                const openBtn = document.createElement("button");
                openBtn.type = "button";
                openBtn.className = "open-btn";
                openBtn.textContent = "Open source";
                openBtn.addEventListener("click", function () {
                    lastOpenButton = openBtn;
                    openCitation(citation);
                });
                card.appendChild(openBtn);
                list.appendChild(card);
            });
            wrap.appendChild(list);
        } else if (noSource) {
            const empty = document.createElement("div");
            empty.className = "no-source";
            empty.textContent = "No matching passage was found in the added documents.";
            wrap.appendChild(empty);
        }
    }
    if (sender === "user") { row.append(wrap, avatar); }
    else { row.append(avatar, wrap); }
    chatWindow.appendChild(row);
    scrollChatToBottom();
}
function addLoading() {
    dismissWelcome();
    const chatWindow = document.getElementById("chatWindow");
    const row = document.createElement("div");
    row.className = "msg assistant loading-row";
    row.innerHTML = '<div class="avatar">AI</div><div class="bubble-wrap"><div class="bubble"><div class="typing" aria-hidden="true"><span></span><span></span><span></span></div></div><div class="time">Looking through your documents…</div></div>';
    chatWindow.appendChild(row);
    scrollChatToBottom();
}
function removeLoading() {
    const loading = document.querySelector(".loading-row");
    if (loading) loading.remove();
}
async function sendQuestion() {
    const input = document.getElementById("questionInput");
    const question = (input.value || "").trim();
    if (!hasIndexedFiles || busy) return;
    if (!question) return;
    busy = true;
    setChatEnabled(true);
    addMessage(question, "user");
    input.value = "";
    addLoading();
    try {
        const response = await fetch("/api/chat", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ question: question })
        });
        const payload = await response.json().catch(function () { return {}; });
        removeLoading();
        if (!response.ok) {
            const err = new Error(payload.error || payload.detail || payload.title || "Could not answer that question. Try again.");
            err.status = response.status;
            throw err;
        }
        const answer = (payload.answer || "").trim();
        const citations = Array.isArray(payload.citations) ? payload.citations : [];
        if (!answer) addMessage("No answer was returned. Try asking in a different way.", "assistant", [], true);
        else addMessage(answer, "assistant", citations, citations.length === 0);
    } catch (error) {
        removeLoading();
        addMessage(friendlyError(error, error && error.message ? error.message : "Could not answer that question. Try again."), "assistant");
    } finally {
        busy = false;
        setChatEnabled(hasIndexedFiles);
        if (hasIndexedFiles) input.focus();
    }
}
function clearConversation() {
    document.querySelectorAll("#chatWindow .msg").forEach(function (node) { node.remove(); });
    showWelcomeIfEmpty();
    document.getElementById("questionInput").focus();
}

        function initViewers() {
            if (viewersReady) {
                return;
            }

            if (typeof ej === "undefined") {
                console.warn("Syncfusion EJ2 is not loaded yet.");
                return;
            }

            try {
                // Server-backed PDF Viewer (Syncfusion.EJ2.PdfViewer.AspNet.Core via /api/PdfViewer).
                if (ej.pdfviewer && ej.pdfviewer.PdfViewer) {
                    const pdfServiceUrl = (window.location && window.location.origin
                        ? window.location.origin
                        : "") + "/api/PdfViewer";
                    pdfViewer = new ej.pdfviewer.PdfViewer({
                        enableToolbar: true,
                        enableNavigationToolbar: true,
                        serviceUrl: pdfServiceUrl,
                        ajaxRequestSettings: {
                            withCredentials: false
                        }
                    });
                    pdfViewer.ajaxRequestFailure = function (args) {
                        console.error("PDF Viewer ajaxRequestFailure", args);
                    };
                    // Surface load failures without leaving a stuck gray canvas + invalid-file dialog alone.
                    pdfViewer.documentLoadFailed = function (args) {
                        console.error("PDF Viewer documentLoadFailed", args);
                    };
                    pdfViewer.appendTo("#pdfViewer");
                }

                // Word DocumentEditor — selection + search modules are required for citation jump.
                if (ej.documenteditor && ej.documenteditor.DocumentEditorContainer) {
                    wordContainer = new ej.documenteditor.DocumentEditorContainer({
                        enableToolbar: true,
                        height: "100%",
                        restrictedEditing: false
                    });
                    wordContainer.appendTo("#wordViewer");
                    if (wordContainer.documentEditor) {
                        const de = wordContainer.documentEditor;
                        de.isReadOnly = false;
                        // Official find/replace sample requires these flags.
                        de.enableSelection = true;
                        de.enableSearch = true;
                        de.enableEditor = true;
                        de.enableOptionsPane = true;
                    }
                }

                // Spreadsheet viewer — openUrl is required so open({ file }) can convert XLSX server-side.
                if (ej.spreadsheet && ej.spreadsheet.Spreadsheet) {
                    spreadsheet = new ej.spreadsheet.Spreadsheet({
                        openUrl: "/api/spreadsheet/open",
                        allowOpen: true,
                        showRibbon: true,
                        showFormulaBar: true,
                        height: "100%"
                    });
                    spreadsheet.appendTo("#excelViewer");
                }

                viewersReady = true;
            } catch (e) {
                console.error("Viewer initialization failed", e);
            }
        }


function showViewer(type) {
    const empty = document.getElementById("viewerEmpty");
    if (empty) empty.classList.add("hidden");
    ["pdfPanel", "wordPanel", "excelPanel"].forEach(function (id) {
        const panel = document.getElementById(id);
        if (panel) panel.classList.remove("active");
    });
    const panelId = type === "pdf" ? "pdfPanel" : type === "word" ? "wordPanel" : type === "excel" ? "excelPanel" : "";
    const panel = panelId && document.getElementById(panelId);
    if (panel) panel.classList.add("active");
    setTimeout(resizeActiveViewer, 50);
}

function normalizeCitation(citation) {
    citation = citation || {};
    return {
        sourceName: citation.sourceName || citation.SourceName || "",
        fileExtension: citation.fileExtension || citation.FileExtension || "",
        viewerType: (citation.viewerType || citation.ViewerType || "").toLowerCase(),
        chunkIndex: citation.chunkIndex ?? citation.ChunkIndex ?? 0,
        citationDisplay: citation.citationDisplay || citation.CitationDisplay || "",
        contentSnippet: citation.contentSnippet || citation.ContentSnippet || "",
        score: citation.score ?? citation.Score ?? 0,
        pageNumber: citation.pageNumber ?? citation.PageNumber ?? null,
        sheetName: citation.sheetName || citation.SheetName || null,
        cellAddress: citation.cellAddress || citation.CellAddress || null,
        headingPath: citation.headingPath || citation.HeadingPath || null,
        searchText: citation.searchText || citation.SearchText || null
    };
}

function revealViewerPane() {
    const workspace = document.getElementById("workspace");
    const pane = document.getElementById("viewerPane");
    if (pane) pane.hidden = false;
    if (workspace) workspace.classList.remove("chat-only");
    document.body.classList.add("source-open");
    window.setTimeout(resizeActiveViewer, 60);
    window.setTimeout(resizeActiveViewer, 280);
    const close = document.getElementById("closeViewerButton");
    if (close) close.focus();
}
function hideViewerPane() {
    const workspace = document.getElementById("workspace");
    const pane = document.getElementById("viewerPane");
    if (workspace) workspace.classList.add("chat-only");
    if (pane) pane.hidden = true;
    document.body.classList.remove("source-open");
    if (lastOpenButton && document.contains(lastOpenButton)) lastOpenButton.focus();
}
function resizeActiveViewer() {
    const active = document.querySelector(".viewer-panel.active");
    if (!active) return;
    if (active.id === "wordPanel" && wordContainer && wordContainer.resize) wordContainer.resize();
    if (active.id === "excelPanel" && spreadsheet && spreadsheet.resize) spreadsheet.resize();
    if (active.id === "pdfPanel" && pdfViewer && pdfViewer.updateViewerContainer) pdfViewer.updateViewerContainer();
}


        async function openCitation(rawCitation) {
            revealViewerPane();
            initViewers();

            const citation = normalizeCitation(rawCitation);
            const status = document.getElementById("viewerStatus");
            const title = document.getElementById("viewerTitle");
            const fileName = citation.sourceName;

            if (!fileName) {
                setStatus(status, "Citation is missing a source file name.", "error");
                return;
            }

            let viewerType = citation.viewerType;
            if (!viewerType || viewerType === "unknown") {
                const ext = (citation.fileExtension || extensionOf(fileName)).toLowerCase();
                if (ext === ".pdf") viewerType = "pdf";
                else if (ext === ".doc" || ext === ".docx") viewerType = "word";
                else if (ext === ".xls" || ext === ".xlsx") viewerType = "excel";
            }

            title.textContent = fileName + (citation.citationDisplay ? " · " + citation.citationDisplay : "");
            setStatus(status, "Opening " + fileName + "…", "info");

            try {
                if (viewerType === "pdf") {
                    await openPdf(fileName, citation);
                } else if (viewerType === "word") {
                    await openWord(fileName, citation);
                } else if (viewerType === "excel") {
                    await openExcel(fileName, citation);
                } else {
                    throw new Error("Unsupported viewer type for this citation.");
                }

                setStatus(status, "Opened " + fileName, "ok");
            } catch (error) {
                console.error(error);
                setStatus(status, "Could not open this source. Try again.", "error");
            }
        }

        function cleanCitationText(value) {
            if (!value) return "";
            return String(value)
                // RAG chunks often use | as a field separator — not present in PDF text.
                .replace(/\|+/g, " ")
                .replace(/[\/\\]+/g, " ")
                .replace(/\s*[-–—]+\s*/g, " ")
                .replace(/\s+/g, " ")
                .trim();
        }

        function buildSearchTerm(citation, maxWords) {
            const raw = cleanCitationText(
                (citation && (citation.searchText || citation.headingPath || citation.contentSnippet)) ||
                ""
            );
            if (!raw) return "";

            // Prefer distinctive tokens (skip very short/stop-ish words) for PDF text search.
            const words = raw
                .split(" ")
                .map(function (w) { return w.replace(/^[^A-Za-z0-9]+|[^A-Za-z0-9]+$/g, ""); })
                .filter(function (w) {
                    if (!w || w.length < 3) return false;
                    if (/^(the|and|for|with|from|this|that|was|are|has|had|you|your)$/i.test(w)) {
                        return false;
                    }
                    return true;
                });

            if (!words.length) {
                const fallback = raw.split(" ").filter(Boolean);
                const takeFb = Math.max(2, Math.min(maxWords || 8, fallback.length));
                return fallback.slice(0, takeFb).join(" ");
            }

            const take = Math.max(2, Math.min(maxWords || 8, words.length));
            return words.slice(0, take).join(" ");
        }

        function navigatePdfPlace(citation) {
            if (!pdfViewer || !citation) return;

            const page = citation.pageNumber ? Number(citation.pageNumber) : null;
            // Short phrases used only to resolve a page when pageNumber is missing — never paint search highlights.
            const candidates = [
                buildSearchTerm(citation, 6),
                buildSearchTerm(citation, 4),
                buildSearchTerm(citation, 2)
            ].filter(function (t, i, arr) {
                return t && arr.indexOf(t) === i;
            });

            let cancelled = false;
            let done = false;

            function goPage(targetPage) {
                const p = targetPage || page;
                if (p && p > 0 && pdfViewer.navigation) {
                    try {
                        pdfViewer.navigation.goToPage(p);
                    } catch (e) {
                        console.warn("PDF goToPage failed", e);
                    }
                }
            }

            // Resolve page via silent find only when needed. Do NOT call searchText (yellow hits).
            async function tryResolvePage(term) {
                if (!term || done || cancelled) return false;
                if (!(pdfViewer.textSearchModule &&
                    typeof pdfViewer.textSearchModule.findTextAsync === "function")) {
                    return false;
                }

                try {
                    const results = await pdfViewer.textSearchModule.findTextAsync(term, false);
                    if (Array.isArray(results) && results.length > 0) {
                        let match = results[0];
                        if (page) {
                            const pageMatch = results.find(function (r) {
                                return (r.pageIndex + 1) === page;
                            });
                            if (pageMatch) match = pageMatch;
                        }
                        if (match && pdfViewer.navigation) {
                            goPage(match.pageIndex + 1);
                        }
                        // Clear any search UI state so matches are not left painted.
                        try {
                            if (pdfViewer.textSearch &&
                                typeof pdfViewer.textSearch.cancelTextSearch === "function") {
                                pdfViewer.textSearch.cancelTextSearch();
                            }
                        } catch (_) { /* ignore */ }
                        done = true;
                        return true;
                    }
                } catch (e) {
                    console.warn("PDF findTextAsync failed", e);
                }
                return false;
            }

            goPage();

            // If citation already has a page, navigation alone is enough.
            if (page && page > 0) {
                const delays = [300, 900];
                delays.forEach(function (delay) {
                    setTimeout(function () {
                        if (cancelled) return;
                        goPage();
                    }, delay);
                });
            } else if (candidates.length) {
                const delays = [400, 1000, 2000];
                delays.forEach(function (delay, attempt) {
                    setTimeout(async function () {
                        if (cancelled || done) return;
                        const term = candidates[Math.min(attempt, candidates.length - 1)];
                        await tryResolvePage(term);
                    }, delay);
                });
            }

            navigatePdfPlace._cancel = function () {
                cancelled = true;
            };
        }

        function softCleanWordText(value) {
            if (!value) return "";
            return String(value)
                .replace(/\u00a0/g, " ")
                .replace(/[\u2018\u2019]/g, "'")
                .replace(/[\u201C\u201D]/g, "\"")
                .replace(/[\u2013\u2014]/g, "-")
                .replace(/[•·▪◦]/g, " ")
                .replace(/\|+/g, " ")
                .replace(/\s+/g, " ")
                .trim();
        }

        function buildWordSearchCandidates(citation) {
            // Prefer body/snippet over searchText — keep contiguous DOCX phrases.
            const rawSources = [
                citation && citation.contentSnippet,
                citation && citation.searchText,
                citation && citation.headingPath
            ].filter(Boolean);

            const candidates = [];
            const seen = new Set();

            function pushTerm(term) {
                if (!term) return;
                // Prefer light cleaning so DOCX body text still matches.
                const variants = [
                    softCleanWordText(term),
                    cleanCitationText(term)
                ];
                variants.forEach(function (t) {
                    if (!t || t.length < 2) return;
                    // Drop trailing ellipsis from snippets.
                    t = t.replace(/[….]{1,3}$/g, "").trim();
                    if (!t || t.length < 2) return;
                    const key = t.toLowerCase();
                    if (seen.has(key)) return;
                    seen.add(key);
                    candidates.push(t);
                });
            }

            function pushWordWindows(src) {
                const cleaned = softCleanWordText(src);
                if (!cleaned) return;
                const words = cleaned
                    .split(" ")
                    .map(function (w) {
                        return w.replace(/^[^A-Za-z0-9'’-]+|[^A-Za-z0-9'’-]+$/g, "");
                    })
                    .filter(Boolean);

                if (!words.length) return;

                // Contiguous sliding windows — stopword-stripped phrases do NOT exist in DOCX.
                const sizes = [12, 10, 8, 6, 5, 4, 3, 2];
                sizes.forEach(function (size) {
                    if (words.length < size && size > 1) return;
                    const take = Math.min(size, words.length);
                    // Prefer later words too (chunk text often starts with labels).
                    pushTerm(words.slice(0, take).join(" "));
                    if (words.length > take + 2) {
                        pushTerm(words.slice(2, 2 + take).join(" "));
                    }
                    if (words.length > take + 5) {
                        pushTerm(words.slice(5, 5 + take).join(" "));
                    }
                    // Mid-chunk window for longer citations.
                    if (words.length > take + 10) {
                        const mid = Math.floor((words.length - take) / 2);
                        pushTerm(words.slice(mid, mid + take).join(" "));
                    }
                });

                // Long distinctive tokens (survive split runs better) — last resort.
                words
                    .filter(function (w) { return w.length >= 6; })
                    .sort(function (a, b) { return b.length - a.length; })
                    .slice(0, 6)
                    .forEach(pushTerm);
            }

            rawSources.forEach(function (src) {
                pushTerm(src);
                pushWordWindows(src);
            });

            return candidates;
        }

        function isWordDocumentStillLoading(editor) {
            try {
                const helper = editor && editor.documentHelper;
                if (helper && helper.isDocumentLoadAsynchronously === true) {
                    return true;
                }
            } catch (_) { /* ignore */ }
            return false;
        }

        function getWordSearchApi(editor) {
            if (!editor) return null;
            try {
                // Accessing .search while enableSearch is true injects the Search module.
                if (editor.enableSearch !== true) {
                    editor.enableSearch = true;
                }
            } catch (_) { /* ignore */ }

            // Prefer public search facade; fall back to searchModule used in Syncfusion samples.
            try {
                if (editor.search && typeof editor.search.findAll === "function") {
                    return editor.search;
                }
            } catch (e) {
                console.warn("Word editor.search access failed", e);
            }
            try {
                if (editor.searchModule && typeof editor.searchModule.findAll === "function") {
                    return editor.searchModule;
                }
            } catch (_) { /* ignore */ }
            return null;
        }

        function tryWordFind(editor, term) {
            if (!editor || !term) {
                return false;
            }

            // EJ2 Search.findAll is a no-op while async SFDT layout is still running.
            if (isWordDocumentStillLoading(editor)) {
                return false;
            }

            const searchApi = getWordSearchApi(editor);
            if (!searchApi) {
                return false;
            }

            // FindOption: None | WholeWord | CaseSensitive | CaseSensitiveWholeWord
            const findOption = "None";

            try {
                // Clear previous highlights so length reflects this findAll only.
                try {
                    const prior = searchApi.searchResults;
                    if (prior && typeof prior.clear === "function") {
                        prior.clear();
                    }
                } catch (_) { /* ignore */ }

                // findAll: EJ2 internally navigate()s + highlight()s the first hit when length > 0.
                // find() returns void — never treat undefined as a hit.
                if (typeof searchApi.findAll !== "function") {
                    return false;
                }

                searchApi.findAll(term, findOption);

                // Re-read after findAll; some builds expose a fresh SearchResults wrapper.
                const results = searchApi.searchResults;
                const count = results && typeof results.length === "number"
                    ? results.length
                    : 0;

                if (count <= 0) {
                    return false;
                }

                // Public API: setting index jumps to that match (docs). findAll already did index 0.
                try {
                    if (results && "index" in results) {
                        results.index = 0;
                    }
                } catch (idxErr) {
                    console.warn("Word searchResults.index failed", idxErr);
                }

                // Scroll into view, then drop yellow search highlights (navigate only).
                try {
                    if (editor.selection &&
                    itor.focusIn === "function") {
                        editor.focusIn();
                    }
                } catch (scrollErr) {
                    console.warn("Word scroll-to-match best-effort failed", scrollErr);
                }

                // Navigate only — clear search-result paint so content is not left highlighted.
                try {
                    const after = searchApi.searchResults;
                    if (after && typeof after.clear === "function") {
                        after.clear();
                    }
                } catch (_) { /* ignore */ }

                return true;
            } catch (e) {
                console.warn("Word find failed for term:", term, e);
                return false;
            }
        }

        function ensureWordSearchEnabled(editor) {
            if (!editor) return;
            try {
                editor.isReadOnly = false;
                editor.enableSelection = true;
                editor.enableSearch = true;
                editor.enableEditor = true;
                // Touch .search so EJ2 injects the Search module before findAll.
                try { void editor.search; } catch (_) { /* ignore */ }
            } catch (_) { /* ignore */ }
        }

        function waitForWordSearchReady(editor, timeoutMs) {
            return new Promise(function (resolve) {
                const started = Date.now();
                const ms = typeof timeoutMs === "number" ? timeoutMs : 8000;

                function check() {
                    if (!editor) {
                        resolve(false);
                        return;
                    }
                    ensureWordSearchEnabled(editor);
                    const stillLoading = isWordDocumentStillLoading(editor);
                    const searchApi = getWordSearchApi(editor);
                    if (!stillLoading && searchApi) {
                        resolve(true);
                        return;
                    }
                    if (Date.now() - started >= ms) {
                        resolve(!stillLoading && !!searchApi);
                        return;
                    }
                    setTimeout(check, 100);
                }

                check();
            });
        }

        function navigateWordPlace(editor, citation) {
            if (!editor || !citation) return;

            // Cancel any prior Word navigate sequence when opening a new citation.
            if (typeof navigateWordPlace._cancel === "function") {
                try { navigateWordPlace._cancel(); } catch (_) { /* ignore */ }
            }

            const candidates = buildWordSearchCandidates(citation);
            if (!candidates.length) {
                console.warn("Word citation has no usable search text.", citation);
                return;
            }

            console.debug("Word citation navigate candidates:", candidates.slice(0, 12));

            let cancelled = false;
            let done = false;
            const timers = [];

            navigateWordPlace._cancel = function () {
                cancelled = true;
                timers.forEach(function (id) { clearTimeout(id); });
            };

            function attemptRound(round) {
                if (cancelled || done || !editor) return;

                ensureWordSearchEnabled(editor);

                if (isWordDocumentStillLoading(editor)) {
                    console.debug("Word navigate waiting: document still loading asynchronously (round " + round + ")");
                    return;
                }

                const searchApi = getWordSearchApi(editor);
                if (!searchApi) {
                    console.warn("Word search API unavailable (enableSearch/selection may be off).");
                    return;
                }

                // Re-scan full candidate list each round. find is used only to jump; highlights are cleared.
                for (let i = 0; i < candidates.length; i++) {
                    const term = candidates[i];
                    if (tryWordFind(editor, term)) {
                        done = true;
                        console.info("Word citation navigated to term:", term);
                        return;
                    }
                }

                if (round >= 5) {
                    console.warn(
                        "Word citation: no navigate match for any candidate.",
                        { tried: candidates.length, citation: citation });
                }
            }

            // Document opens while the panel was hidden — layout/search need retries.
            const delays = [150, 400, 900, 1600, 2800, 4500];
            delays.forEach(function (delay, round) {
                timers.push(setTimeout(function () {
                    if (cancelled || done) return;
                    try {
                        if (wordContainer && typeof wordContainer.resize === "function") {
                            wordContainer.resize();
                        }
                    } catch (_) { /* ignore */ }
                    attemptRound(round);
                }, delay));
            });
        }

        function navigateExcelPlace(citation) {
            if (!spreadsheet || !citation) return;

            const sheet = (citation.sheetName || "").trim();
            const cell = (citation.cellAddress || "").trim();
            let address = "";
            if (sheet && cell) address = sheet + "!" + cell;
            else if (cell) address = cell;
            else if (sheet) address = sheet + "!A1";
            if (!address) return;

            const delays = [500, 1200, 2000];
            delays.forEach(function (delay) {
                setTimeout(function () {
                    try {
                        // Jump only — no yellow cell formatting.
                        if (typeof spreadsheet.goTo === "function") {
                            spreadsheet.goTo(address);
                        }
                        if (cell && typeof spreadsheet.selectRange === "function") {
                            spreadsheet.selectRange(cell);
                        }
                    } catch (e) {
                        console.warn("Excel goTo best-effort failed", e);
                    }
                }, delay);
            });
        }

        function waitForPdfDocumentLoad(timeoutMs) {
            return new Promise(function (resolve, reject) {
                if (!pdfViewer) {
                    reject(new Error("PDF viewer is not initialized."));
                    return;
                }

                let settled = false;
                const ms = typeof timeoutMs === "number" ? timeoutMs : 20000;

                function cleanup() {
                    try { pdfViewer.documentLoad = previousLoad; } catch (_) { /* ignore */ }
                    try { pdfViewer.documentLoadFailed = previousFail; } catch (_) { /* ignore */ }
                    if (timer) clearTimeout(timer);
                }

                const previousLoad = pdfViewer.documentLoad;
                const previousFail = pdfViewer.documentLoadFailed;

                const timer = setTimeout(function () {
                    if (settled) return;
                    settled = true;
                    cleanup();
                    // If toolbar already shows pageCount, treat as success even without the event.
                    if (pdfViewer.pageCount && pdfViewer.pageCount > 0) {
                        resolve({ timedOutWithPages: true, pageCount: pdfViewer.pageCount });
                    } else {
                        reject(new Error("PDF document load timed out."));
                    }
                }, ms);

                pdfViewer.documentLoad = function (args) {
                    if (typeof previousLoad === "function") {
                        try { previousLoad.call(pdfViewer, args); } catch (_) { /* ignore */ }
                    }
                    if (settled) return;
                    settled = true;
                    cleanup();
                    resolve(args || {});
                };

                pdfViewer.documentLoadFailed = function (args) {
                    if (typeof previousFail === "function") {
                        try { previousFail.call(pdfViewer, args); } catch (_) { /* ignore */ }
                    }
                    if (settled) return;
                    settled = true;
                    cleanup();
                    const msg = (args && (args.errorMessage || args.message)) || "PDF document load failed.";
                    reject(new Error(msg));
                };
            });
        }

        async function loadPdfAsBase64(fileName) {
            const response = await fetch(
                "/api/documents/" + encodeURIComponent(fileName) + "/base64",
                { cache: "no-store" });
            if (!response.ok) {
                const payload = await response.json().catch(function () { return {}; });
                throw new Error(payload.error || ("Failed to fetch PDF base64 for " + fileName));
            }

            const payload = await response.json();
            if (!payload || !payload.base64) {
                throw new Error("PDF base64 payload was empty.");
            }

            // EJ2 accepts raw base64 (isFileName=false) for client-side document bytes.
            pdfViewer.load(payload.base64, null);
        }

        async function openPdf(fileName, citation) {
            showViewer("pdf");

            if (!pdfViewer) {
                throw new Error("PDF viewer is not initialized.");
            }

            // Cancel any pending navigate timers from a previous citation.
            if (typeof navigatePdfPlace._cancel === "function") {
                try { navigatePdfPlace._cancel(); } catch (_) { /* ignore */ }
            }

            // Server-backed load: document name is resolved by PdfViewerController + RagService.
            const pdfServiceUrl = (window.location && window.location.origin
                ? window.location.origin
                : "") + "/api/PdfViewer";
            pdfViewer.serviceUrl = pdfServiceUrl;

            try {
                if (typeof pdfViewer.updateViewerContainer === "function") {
                    pdfViewer.updateViewerContainer();
                }
            } catch (e) {
                console.warn("updateViewerContainer failed", e);
            }

            // Prefer a fresh document session id per open to avoid stale server cache collisions
            // ("Document stream does not exist in the cache") when reopening the same file.
            try {
                const sessionId = "pdf_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
                if ("documentId" in pdfViewer) {
                    pdfViewer.documentId = sessionId;
                }
            } catch (e) {
                console.warn("Could not set PDF documentId", e);
            }

            let opened = false;
            let lastError = null;

            // Path 1: filename via serviceUrl Load (isFileName=true on server).
            try {
                const loadWait = waitForPdfDocumentLoad(18000);
                pdfViewer.load(fileName, null);
                await loadWait;
                opened = true;
            } catch (e) {
                lastError = e;
                console.warn("PDF filename load failed, trying base64 fallback", e);
            }

            // Path 2: fetch bytes and load as base64 (same approach as NetCoreLatest demo).
            if (!opened) {
                try {
                    const loadWait = waitForPdfDocumentLoad(20000);
                    await loadPdfAsBase64(fileName);
                    await loadWait;
                    opened = true;
                } catch (e) {
                    lastError = e;
                    console.error("PDF base64 load failed", e);
                }
            }

            if (!opened) {
                const message = (lastError && lastError.message) || ("Could not open PDF: " + fileName);
                throw new Error(message);
            }

            // Navigate to the cited page only (no text highlight).
            navigatePdfPlace(citation);
        }

        async function openWord(fileName, citation) {
            showViewer("word");

            if (!wordContainer || !wordContainer.documentEditor) {
                throw new Error("Word viewer is not initialized.");
            }

            // Cancel in-flight navigate from a previous Word citation.
            if (typeof navigateWordPlace._cancel === "function") {
                try { navigateWordPlace._cancel(); } catch (_) { /* ignore */ }
            }

            const response = await fetch(
                "/api/documents/" + encodeURIComponent(fileName) + "/sfdt",
                { cache: "no-store" });
            if (!response.ok) {
                const payload = await response.json().catch(function () { return {}; });
                throw new Error(payload.error || payload.title || "Failed to convert Word document.");
            }

            const sfdt = await response.text();
            if (!sfdt || !String(sfdt).trim()) {
                throw new Error("Word SFDT payload was empty.");
            }

            const editor = wordContainer.documentEditor;
            ensureWordSearchEnabled(editor);

            try {
                if (typeof wordContainer.resize === "function") {
                    wordContainer.resize();
                }
            } catch (e) {
                console.warn("Word pre-open setup failed", e);
            }

            // Wait for DocumentEditor to finish applying SFDT before searching.
            // EJ2 findAll no-ops while documentHelper.isDocumentLoadAsynchronously is true.
            await new Promise(function (resolve) {
                let settled = false;
                const previousChange = editor.documentChange;

                function finish() {
                    if (settled) return;
                    settled = true;
                    try { editor.documentChange = previousChange; } catch (_) { /* ignore */ }
                    resolve();
                }

                try {
                    editor.documentChange = function () {
                        if (typeof previousChange === "function") {
                            try { previousChange.apply(editor, arguments); } catch (_) { /* ignore */ }
                        }
                        // Layout + search index need a beat after SFDT apply.
                        setTimeout(finish, 250);
                    };
                } catch (_) {
                    /* fall through to timeout open */
                }

                try {
                    // open() accepts optimized {"sfdt":"<zip base64>"} or expanded SFDT JSON.
                    editor.open(sfdt);
                } catch (openErr) {
                    finish();
                    throw openErr;
                }

                // Fallback if documentChange never fires for this build.
                setTimeout(finish, 2500);
            });

            // Extra gate: findAll is a silent no-op until async layout completes.
            await waitForWordSearchReady(editor, 8000);

            try {
                if (typeof wordContainer.resize === "function") {
                    wordContainer.resize();
                }
                if (typeof editor.focusIn === "function") {
                    editor.focusIn();
                }
            } catch (_) { /* ignore */ }

            ensureWordSearchEnabled(editor);
            navigateWordPlace(editor, citation);
        }

        async function openExcel(fileName, citation) {
            showViewer("excel");

            // Server converts XLS/XLSX → EJ2 workbook JSON (same path openUrl uses).
            const response = await fetch(
                "/api/documents/" + encodeURIComponent(fileName) + "/workbook");
            if (!response.ok) {
                const payload = await response.json().catch(function () { return {}; });
                throw new Error(
                    payload.error ||
                    payload.title ||
                    payload.detail ||
                    "Failed to convert spreadsheet.");
            }

            const workbookJson = await response.json();
            if (typeof spreadsheet.openFromJson !== "function") {
                throw new Error("Spreadsheet openFromJson is not available.");
            }

            spreadsheet.openFromJson({ file: workbookJson });
            // Jump to sheet/cell only — no yellow cell format.
            navigateExcelPlace(citation);
        }


document.addEventListener("click", function (event) {
    const close = event.target.closest("#closeViewerButton");
    if (close) {
        event.preventDefault();
        hideViewerPane();
        return;
    }
    const add = event.target.closest("#addDocsButton");
    if (add) {
        event.preventDefault();
        showUpload();
        return;
    }
    const reset = event.target.closest("#clearChatButton");
    if (reset) {
        event.preventDefault();
        clearConversation();
    }
});
document.addEventListener("DOMContentLoaded", function () {
    refreshReadyState();
    const input = document.getElementById("fileInput");
    const drop = document.getElementById("dropZone");
    if (input) input.addEventListener("change", function () { assignFiles(input.files); });
    if (drop) {
        ["dragenter", "dragover"].forEach(function (name) {
            drop.addEventListener(name, function (event) { event.preventDefault(); drop.classList.add("dragover"); });
        });
        ["dragleave", "drop"].forEach(function (name) {
            drop.addEventListener(name, function (event) { event.preventDefault(); drop.classList.remove("dragover"); });
        });
        drop.addEventListener("drop", function (event) {
            assignFiles(event.dataTransfer.files);
        });
    }
    const uploadButton = document.getElementById("uploadButton");
    const questionForm = document.getElementById("questionForm");
    const questionInput = document.getElementById("questionInput");
    if (uploadButton) uploadButton.addEventListener("click", uploadFiles);
    if (questionForm) questionForm.addEventListener("submit", function (event) {
        event.preventDefault();
        sendQuestion();
    });
    if (questionInput) questionInput.addEventListener("keydown", function (event) {
        if (event.key === "Enter" && !event.shiftKey) {
            event.preventDefault();
            sendQuestion();
        }
    });
});
