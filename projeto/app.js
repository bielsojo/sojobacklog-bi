// Refined application logic for Low Cost Report Generator

// Global State
let equipments = [];
let uploadedFiles = []; // { name, size, buffer, extractedText, type: 'nf' | 'rat' | 'evidence' }
let nextEquipId = 1;

// Fixed Low Cost corporate data
const LOW_COST_INFO = {
    name: "LOWCOST GERENCIAMENTO DE SERVIÇOS LTDA",
    cnpj: "01.482.939/0001-00",
    address: "Av. Tamboré, 1400, Sala 39 – Tamboré – Barueri – SP – Cep: 06.460-000",
    footerAddress: "Rua Manuel de Nobrega, 986 – 4º andar – Paraiso SP – CEP: 04001-003"
};

// Initialize Application
document.addEventListener("DOMContentLoaded", () => {
    // Icons
    lucide.createIcons();
    
    // Set auto date (Today)
    const today = new Date();
    const dateFormatted = formatPortugueseDate(today.toISOString().split('T')[0]);
    document.getElementById("previewDocDate").innerText = dateFormatted;
    
    // Bind Event Listeners
    document.getElementById("operationType").addEventListener("change", handleOperationChange);
    document.getElementById("laudoNumber").addEventListener("input", updatePreview);
    document.getElementById("destName").addEventListener("input", updatePreview);
    document.getElementById("destCnpj").addEventListener("input", updatePreview);
    document.getElementById("destAddress").addEventListener("input", updatePreview);
    
    // Actions Buttons
    document.getElementById("addRowBtn").addEventListener("click", addNewRow);
    document.getElementById("clearBtn").addEventListener("click", clearAll);
    document.getElementById("generateBtn").addEventListener("click", generateReport);
    document.getElementById("themeToggleBtn").addEventListener("click", toggleTheme);
    
    // Excel Import Event Bind
    const excelBtn = document.getElementById("importExcelBtn");
    const excelInput = document.getElementById("fileExcel");
    excelBtn.addEventListener("click", () => excelInput.click());
    excelInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleExcelUpload(e.target.files);
        }
    });
    
    // Setup 3 Separate Upload Zones
    setupUploadZone("dropzoneNf", "fileNf", (files) => handleFilesUpload(files, 'nf'));
    setupUploadZone("dropzoneRat", "fileRat", (files) => handleFilesUpload(files, 'rat'));
    setupUploadZone("dropzoneEvidence", "fileEvidence", (files) => handleFilesUpload(files, 'evidence'));
    
    // Initial Render
    renderTable();
    updatePreview();
});

// Theme Toggle
function toggleTheme() {
    const html = document.documentElement;
    const currentTheme = html.getAttribute("data-theme");
    const newTheme = currentTheme === "dark" ? "light" : "dark";
    html.setAttribute("data-theme", newTheme);
    
    const themeIcon = document.getElementById("themeIcon");
    if (newTheme === "light") {
        themeIcon.setAttribute("data-lucide", "moon");
    } else {
        themeIcon.setAttribute("data-lucide", "sun");
    }
    lucide.createIcons();
}

// Operation Type change handler
// Changes observation cells in table between 'DOAÇÃO' and 'DESCARTE'
function handleOperationChange(e) {
    const op = e.target.value;
    const val = op === "doacao" ? "DOAÇÃO" : "DESCARTE";
    
    equipments.forEach(eq => {
        eq.observation = val;
    });
    
    renderTable();
    updatePreview();
}

// Drag & Drop Upload Zone Configuration
function setupUploadZone(zoneId, inputId, handler) {
    const zone = document.getElementById(zoneId);
    const input = document.getElementById(inputId);
    
    zone.addEventListener("click", () => input.click());
    
    zone.addEventListener("dragover", (e) => {
        e.preventDefault();
        zone.classList.add("dragover");
    });
    
    zone.addEventListener("dragleave", () => {
        zone.classList.remove("dragover");
    });
    
    zone.addEventListener("drop", (e) => {
        e.preventDefault();
        zone.classList.remove("dragover");
        if (e.dataTransfer.files.length > 0) {
            handler(e.dataTransfer.files);
        }
    });
    
    input.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handler(e.target.files);
        }
    });
}

// Show/Hide progress spinner
function toggleProgress(show, title = "", subtitle = "", percent = 0) {
    const overlay = document.getElementById("progressOverlay");
    if (show) {
        document.getElementById("progressTitle").innerText = title;
        document.getElementById("progressSubtitle").innerText = subtitle;
        document.getElementById("progressBarFill").style.width = `${percent}%`;
        overlay.classList.add("active");
    } else {
        overlay.classList.remove("active");
    }
}

// Process Excel Upload
function handleExcelUpload(files) {
    const file = files[0];
    if (!file) return;
    
    toggleProgress(true, "Lendo Planilha Excel", "Carregando dados dos equipamentos...", 20);
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Extract metadados (Laudo #)
            const j4Val = worksheet['J4'] ? worksheet['J4'].v : (worksheet['J3'] ? worksheet['J3'].v : null);
            if (j4Val) {
                document.getElementById("laudoNumber").value = j4Val;
            }
            
            // Parse table row headers
            const jsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            let startRowIdx = 18;
            let headerFound = false;
            
            for (let i = 0; i < Math.min(30, jsonRows.length); i++) {
                const row = jsonRows[i];
                if (row && row.some(cell => typeof cell === 'string' && cell.toUpperCase().includes('SERIAL'))) {
                    startRowIdx = i;
                    headerFound = true;
                    break;
                }
            }
            
            if (headerFound) {
                const headers = jsonRows[startRowIdx].map(h => h ? h.toString().toUpperCase().trim() : '');
                
                const idxSerial = headers.findIndex(h => h.includes('SERIAL'));
                const idxOrigem = headers.findIndex(h => h.includes('ORIGEM') || h.includes('LOCALIDADE'));
                const idxModelo = headers.findIndex(h => h.includes('MARCA') || h.includes('MODELO'));
                const idxContador = headers.findIndex(h => h.includes('CONTADOR') || h.includes('CÓD.'));
                const idxIdade = headers.findIndex(h => h.includes('IDADE'));
                const idxNfEnvio = headers.findIndex(h => h.includes('NF DE ENVIO') || h.includes('NF ENVIO') || h.includes('NF DE SAÍDA') || h.includes('ENVIO'));
                const idxNfDevolucao = headers.findIndex(h => h.includes('NF DE DEVOLUÇÃO') || h.includes('DEVOLUÇÃO') || h.includes('RETORNO'));
                
                equipments = [];
                nextEquipId = 1;
                
                const currentOp = document.getElementById("operationType").value;
                const opText = currentOp === "doacao" ? "DOAÇÃO" : "DESCARTE";
                
                for (let r = startRowIdx + 1; r < jsonRows.length; r++) {
                    const row = jsonRows[r];
                    if (!row || !row[idxSerial]) continue;
                    
                    const serial = row[idxSerial].toString().trim();
                    const origin = idxOrigem !== -1 && row[idxOrigem] ? row[idxOrigem].toString().trim() : '';
                    const model = idxModelo !== -1 && row[idxModelo] ? row[idxModelo].toString().trim() : '';
                    const counter = idxContador !== -1 && row[idxContador] ? row[idxContador].toString().trim() : '0';
                    const age = idxIdade !== -1 && row[idxIdade] ? row[idxIdade].toString().trim() : '';
                    const nfEnvio = idxNfEnvio !== -1 && row[idxNfEnvio] ? row[idxNfEnvio].toString().trim() : '';
                    const nfDevolucao = idxNfDevolucao !== -1 && row[idxNfDevolucao] ? row[idxNfDevolucao].toString().trim() : '';
                    
                    equipments.push({
                        id: nextEquipId++,
                        origin: origin,
                        model: model,
                        serial: serial,
                        counter: counter,
                        age: age,
                        nfEnvio: nfEnvio,
                        nfDevolucao: nfDevolucao,
                        observation: opText,
                        nfEnvioFile: null,
                        nfDevolucaoFile: null,
                        ratFile: null
                    });
                }
            } else {
                alert("Não foi possível encontrar a linha de cabeçalho na planilha.");
            }
            
            if (equipments.length === 0) addNewRow();
            
            autoMatchFiles();
            renderTable();
            updatePreview();
            toggleProgress(false);
            
        } catch (error) {
            console.error(error);
            alert("Erro ao processar o arquivo Excel.");
            toggleProgress(false);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Process separate uploads for NF, RAT, Evidence
async function handleFilesUpload(files, type) {
    if (files.length === 0) return;
    
    document.getElementById("filesTrackerCard").style.display = "block";
    toggleProgress(true, "Processando Documentos", "Lendo arquivos selecionados...", 0);
    
    const total = files.length;
    for (let i = 0; i < total; i++) {
        const file = files[i];
        const percent = Math.round((i / total) * 100);
        toggleProgress(true, "Processando Documentos", `Lendo ${file.name} (${i+1}/${total})...`, percent);
        
        try {
            const buffer = await file.arrayBuffer();
            let text = "";
            if (file.type === "application/pdf") {
                text = await extractTextFromPdf(buffer.slice(0));
            }
            
            // Keep tracker, separate by type
            if (!uploadedFiles.some(f => f.name === file.name && f.size === file.size)) {
                uploadedFiles.push({
                    name: file.name,
                    size: file.size,
                    buffer: buffer,
                    extractedText: text,
                    type: type // 'nf' | 'rat' | 'evidence'
                });
            }
        } catch (e) {
            console.error("Error reading file:", file.name, e);
        }
    }
    
    updateFilesTracker();
    
    toggleProgress(true, "Associação Automática", "Mapeando notas fiscais e relatórios técnicos aos equipamentos...", 85);
    autoMatchFiles();
    
    renderTable();
    updatePreview();
    toggleProgress(false);
}

// Text extraction using PDF.js
function extractTextFromPdf(arrayBuffer) {
    return new Promise((resolve, reject) => {
        pdfjsLib.getDocument({ data: arrayBuffer }).promise.then(async (pdf) => {
            let fullText = "";
            const maxPages = Math.min(5, pdf.numPages);
            for (let i = 1; i <= maxPages; i++) {
                try {
                    const page = await pdf.getPage(i);
                    const content = await page.getTextContent();
                    const pageText = content.items.map(item => item.str).join(" ");
                    fullText += pageText + "\n";
                } catch (e) {
                    console.error("Error text extraction:", e);
                }
            }
            resolve(fullText);
        }).catch(err => {
            reject(err);
        });
    });
}

// Refresh files tracker dashboard
function updateFilesTracker() {
    const list = document.getElementById("uploadedFilesList");
    const count = document.getElementById("uploadedFilesCount");
    
    count.innerText = uploadedFiles.length;
    list.innerHTML = "";
    
    uploadedFiles.forEach((file, index) => {
        const item = document.createElement("div");
        item.className = "uploaded-file-item";
        
        let typeBadge = "";
        if (file.type === 'nf') typeBadge = '<span style="font-size: 0.65rem; background-color: var(--color-primary-light); color: var(--text-primary); padding: 0.1rem 0.3rem; border-radius: 3px; margin-right: 4px;">NF</span>';
        else if (file.type === 'rat') typeBadge = '<span style="font-size: 0.65rem; background-color: var(--color-success-light); color: var(--color-success); padding: 0.1rem 0.3rem; border-radius: 3px; margin-right: 4px;">RAT</span>';
        else typeBadge = '<span style="font-size: 0.65rem; background-color: rgba(255,255,255,0.1); color: var(--text-secondary); padding: 0.1rem 0.3rem; border-radius: 3px; margin-right: 4px;">EVID</span>';

        item.innerHTML = `
            ${typeBadge}
            <span>${file.name}</span>
            <button onclick="removeUploadedFile(${index})">&times;</button>
        `;
        list.appendChild(item);
    });
    
    lucide.createIcons();
}

// Remove uploaded file
function removeUploadedFile(index) {
    const file = uploadedFiles[index];
    if (!file) return;
    
    equipments.forEach(eq => {
        if (eq.nfEnvioFile === file.name) eq.nfEnvioFile = null;
        if (eq.nfDevolucaoFile === file.name) eq.nfDevolucaoFile = null;
        if (eq.ratFile === file.name) eq.ratFile = null;
    });
    
    uploadedFiles.splice(index, 1);
    updateFilesTracker();
    renderTable();
    updatePreview();
    
    if (uploadedFiles.length === 0) {
        document.getElementById("filesTrackerCard").style.display = "none";
    }
}

// Auto match uploaded files based on type and content
function autoMatchFiles() {
    if (uploadedFiles.length === 0 || equipments.length === 0) return;
    
    equipments.forEach(eq => {
        const serialClean = eq.serial ? eq.serial.toUpperCase().replace(/[^A-Z0-9]/g, '') : '';
        const nfEnvioClean = eq.nfEnvio ? eq.nfEnvio.toUpperCase().replace(/[^0-9]/g, '') : '';
        const nfDevolucaoClean = eq.nfDevolucao ? eq.nfDevolucao.toUpperCase().replace(/[^0-9]/g, '') : '';
        
        if (!serialClean) return;
        
        uploadedFiles.forEach(file => {
            const fileNameUpper = file.name.toUpperCase();
            const textUpper = file.extractedText.toUpperCase();
            const textClean = textUpper.replace(/[^A-Z0-9]/g, '');
            
            // 1. Match RAT / OS by Serial Number in RAT-type files
            if (file.type === 'rat') {
                if (fileNameUpper.includes(serialClean) || textClean.includes(serialClean)) {
                    if (!eq.ratFile) eq.ratFile = file.name;
                }
            }
            
            // 2. Match NFs in NF-type files by Invoice Numbers
            if (file.type === 'nf') {
                if (nfEnvioClean && (fileNameUpper.includes(nfEnvioClean) || textClean.includes(nfEnvioClean))) {
                    if (!eq.nfEnvioFile) eq.nfEnvioFile = file.name;
                }
                if (nfDevolucaoClean && (fileNameUpper.includes(nfDevolucaoClean) || textClean.includes(nfDevolucaoClean))) {
                    if (!eq.nfDevolucaoFile) eq.nfDevolucaoFile = file.name;
                }
            }
        });
    });
}

// Add equipment row manually
function addNewRow() {
    const currentOp = document.getElementById("operationType").value;
    const opText = currentOp === "doacao" ? "DOAÇÃO" : "DESCARTE";
    
    equipments.push({
        id: nextEquipId++,
        origin: "",
        model: "",
        serial: "",
        counter: "",
        age: "",
        nfEnvio: "",
        nfDevolucao: "",
        observation: opText,
        nfEnvioFile: null,
        nfDevolucaoFile: null,
        ratFile: null
    });
    
    renderTable();
    updatePreview();
}

function removeRow(id) {
    equipments = equipments.filter(eq => eq.id !== id);
    if (equipments.length === 0) {
        addNewRow();
    } else {
        renderTable();
        updatePreview();
    }
}

function updateCell(id, field, value) {
    const eq = equipments.find(eq => eq.id === id);
    if (eq) {
        eq[field] = value;
        updatePreview();
        // If file relationship is updated, re-render the table to refresh status badges
        if (field === 'nfEnvioFile' || field === 'nfDevolucaoFile' || field === 'ratFile') {
            renderTable();
        }
    }
}

// Render dynamic elements inside interactive table
function renderTable() {
    const tbody = document.getElementById("equipTableBody");
    tbody.innerHTML = "";
    
    // Create dropdown selections for NF files and RAT files
    let nfOptions = `<option value="">-- Selecione a NF --</option>`;
    let ratOptions = `<option value="">-- Selecione o RAT --</option>`;
    
    uploadedFiles.forEach(f => {
        if (f.type === 'nf') {
            nfOptions += `<option value="${f.name}">${f.name}</option>`;
        } else if (f.type === 'rat') {
            ratOptions += `<option value="${f.name}">${f.name}</option>`;
        }
    });
    
    equipments.forEach((eq, index) => {
        const tr = document.createElement("tr");
        
        tr.innerHTML = `
            <td style="font-weight: 600; text-align: center;">${index + 1}</td>
            <td><input type="text" class="cell-input" value="${eq.origin}" oninput="updateCell(${eq.id}, 'origin', this.value)"></td>
            <td><input type="text" class="cell-input" value="${eq.model}" oninput="updateCell(${eq.id}, 'model', this.value)"></td>
            <td><input type="text" class="cell-input" value="${eq.serial}" oninput="updateCell(${eq.id}, 'serial', this.value)"></td>
            <td><input type="text" class="cell-input" value="${eq.counter}" oninput="updateCell(${eq.id}, 'counter', this.value)"></td>
            <td><input type="text" class="cell-input" value="${eq.age}" oninput="updateCell(${eq.id}, 'age', this.value)"></td>
            <td>
                <div style="display: flex; flex-direction: column; gap: 0.25rem;">
                    <input type="text" class="cell-input" style="font-weight: 500;" value="${eq.nfEnvio}" oninput="updateCell(${eq.id}, 'nfEnvio', this.value)">
                    <select class="cell-select" onchange="updateCell(${eq.id}, 'nfEnvioFile', this.value)">
                        ${nfOptions}
                    </select>
                </div>
            </td>
            <td>
                <div style="display: flex; flex-direction: column; gap: 0.25rem;">
                    <input type="text" class="cell-input" style="font-weight: 500;" value="${eq.nfDevolucao}" oninput="updateCell(${eq.id}, 'nfDevolucao', this.value)">
                    <select class="cell-select" onchange="updateCell(${eq.id}, 'nfDevolucaoFile', this.value)">
                        ${nfOptions}
                    </select>
                </div>
            </td>
            <td><span style="font-weight: 600; padding: 0.5rem; display: block; text-align: center; font-size: 0.75rem;">${eq.observation}</span></td>
            <td>
                <div style="display: flex; justify-content: center; align-items: center; gap: 0.5rem;">
                    <button class="btn btn-danger btn-icon-only tooltip" data-tooltip="Remover" onclick="removeRow(${eq.id})">
                        <i data-lucide="trash-2" style="width: 14px; height: 14px;"></i>
                    </button>
                </div>
            </td>
        `;
        
        tbody.appendChild(tr);
        
        // Match selects values
        const selects = tr.querySelectorAll("select");
        if (eq.nfEnvioFile) selects[0].value = eq.nfEnvioFile;
        if (eq.nfDevolucaoFile) selects[1].value = eq.nfDevolucaoFile;
        
        // Create Status row
        const trStatus = document.createElement("tr");
        trStatus.style.backgroundColor = "rgba(11, 19, 41, 0.1)";
        
        const badgeEnvio = eq.nfEnvioFile 
            ? `<span class="badge badge-success"><i data-lucide="check" style="width: 10px; height: 10px;"></i> NF Envio associada</span>` 
            : `<span class="badge badge-danger"><i data-lucide="x" style="width: 10px; height: 10px;"></i> Sem NF Envio (Obrigatorio)</span>`;
            
        const badgeRat = eq.ratFile 
            ? `<span class="badge badge-success"><i data-lucide="check" style="width: 10px; height: 10px;"></i> RAT associado</span>` 
            : `<span class="badge badge-danger"><i data-lucide="x" style="width: 10px; height: 10px;"></i> Sem RAT (Obrigatorio)</span>`;
            
        trStatus.innerHTML = `
            <td></td>
            <td colspan="5" style="padding: 0.35rem 1rem;">
                <div style="display: flex; gap: 0.75rem; align-items: center;">
                    ${badgeEnvio}
                    ${badgeRat}
                </div>
            </td>
            <td colspan="4" style="padding: 0.35rem 1rem; text-align: right;">
                <span style="font-size: 0.75rem; color: var(--text-secondary); margin-right: 0.5rem;">Vincular RAT do Equipamento:</span>
                <select class="cell-select" style="width: 220px; display: inline-block; vertical-align: middle;" onchange="updateCell(${eq.id}, 'ratFile', this.value)">
                    ${ratOptions}
                </select>
            </td>
        `;
        
        tbody.appendChild(trStatus);
        
        // Select RAT value
        const ratSelect = trStatus.querySelector("select");
        if (eq.ratFile) ratSelect.value = eq.ratFile;
    });
    
    lucide.createIcons();
}

function base64ToUint8Array(base64Str) {
    const rawBase64 = base64Str.includes(',') ? base64Str.split(',')[1] : base64Str;
    const binaryString = atob(rawBase64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes;
}

// Convert ISO date string to written Portuguese
function formatPortugueseDate(dateStr) {
    if (!dateStr) return "";
    const dateParts = dateStr.split('-');
    if (dateParts.length !== 3) return dateStr;
    
    const year = parseInt(dateParts[0], 10);
    const month = parseInt(dateParts[1], 10) - 1;
    const day = parseInt(dateParts[2], 10);
    
    const date = new Date(year, month, day);
    
    const weekDays = [
        "domingo", "segunda-feira", "terça-feira", "quarta-feira",
        "quinta-feira", "sexta-feira", "sábado"
    ];
    
    const months = [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ];
    
    const weekDay = weekDays[date.getDay()];
    const dayVal = date.getDate();
    const monthName = months[date.getMonth()];
    const yearVal = date.getFullYear();
    
    return `${weekDay}, ${dayVal} de ${monthName} de ${yearVal}`;
}

// Generate the text block containing destination company details
function generateTextDeclaration() {
    const destName = document.getElementById("destName").value.trim();
    const destCnpj = document.getElementById("destCnpj").value.trim();
    const destAddress = document.getElementById("destAddress").value.trim();
    
    let base = `Declaramos a quem possa interessar que a empresa Low Cost Gerenciamento de Serviços Ltda. – CNPJ 01.482.939/0001-00, situada na Av. Tamboré, 1400, Sala 39 – Tamboré – Barueri – SP – Cep: 06.460-000; vem através desta doar os equipamentos abaixo`;
    
    if (destName) {
        base += ` para ${destName}`;
        if (destCnpj) base += `.- CNPJ: ${destCnpj}`;
        if (destAddress) base += ` - Endereço: ${destAddress}`;
    } else {
        base += `.`;
    }
    
    return base;
}

// Update live preview layout
function updatePreview() {
    // Number
    const number = document.getElementById("laudoNumber").value || "---";
    document.getElementById("previewDocNumber").innerText = number;
    
    // Logo & Banner
    document.getElementById("previewLogo").src = ASSETS.logo;
    document.getElementById("previewFooterBanner").src = ASSETS.bannerClassic;
    
    // Title
    document.getElementById("previewDocTitle").innerText = "Declaração de Doação";
    
    // Text Declaration
    const textVal = generateTextDeclaration();
    document.getElementById("previewDocText").innerText = textVal;
    document.getElementById("declarationText").value = textVal;
    
    // Today's date automatically
    const today = new Date();
    const dateFormatted = formatPortugueseDate(today.toISOString().split('T')[0]);
    document.getElementById("previewDocDate").innerText = dateFormatted;
    
    // Address
    document.getElementById("previewFooterAddress").innerText = LOW_COST_INFO.footerAddress;
    
    // Populate preview table rows
    const tableBody = document.getElementById("previewTableBody");
    tableBody.innerHTML = "";
    
    equipments.forEach((eq, index) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td style="text-align: center;">${index + 1}</td>
            <td>${eq.origin || '-'}</td>
            <td>${eq.model || '-'}</td>
            <td style="font-weight: bold;">${eq.serial || '-'}</td>
            <td style="text-align: center;">${eq.counter || '0'}</td>
            <td style="text-align: center;">${eq.age || '-'}</td>
            <td style="text-align: center;">${eq.nfEnvio || '-'}</td>
            <td style="text-align: center;">${eq.nfDevolucao || '-'}</td>
            <td style="text-align: center; font-weight: bold;">${eq.observation || '-'}</td>
        `;
        tableBody.appendChild(tr);
    });
}

// Clear all inputs
function clearAll() {
    if (confirm("Tem certeza que deseja limpar os dados do laudo?")) {
        equipments = [];
        uploadedFiles = [];
        nextEquipId = 1;
        document.getElementById("laudoNumber").value = "";
        document.getElementById("destName").value = "";
        document.getElementById("destCnpj").value = "";
        document.getElementById("destAddress").value = "";
        
        // Reset file input elements values
        document.getElementById("fileExcel").value = "";
        document.getElementById("fileNf").value = "";
        document.getElementById("fileRat").value = "";
        document.getElementById("fileEvidence").value = "";
        
        document.getElementById("filesTrackerCard").style.display = "none";
        
        renderTable();
        updatePreview();
    }
}

// Consolidated PDF generation client-side using pdf-lib
async function generateReport() {
    // Filter out completely empty rows (where all main fields are empty)
    const activeEquipments = equipments.filter(eq => 
        (eq.serial || '').trim() || 
        (eq.model || '').trim() || 
        (eq.origin || '').trim()
    );
    
    if (activeEquipments.length === 0) {
        alert("Adicione pelo menos um equipamento válido ao laudo.");
        return;
    }
    
    // 1. Mandatory Files Validations
    let missingNf = [];
    let missingRat = [];
    
    activeEquipments.forEach((eq, idx) => {
        if (!eq.nfEnvioFile && eq.nfEnvio && eq.nfEnvio !== '0') {
            missingNf.push(idx + 1);
        }
        if (!eq.ratFile) {
            missingRat.push(idx + 1);
        }
    });
    
    if (missingNf.length > 0) {
        alert(`Geração bloqueada! A(s) Nota(s) Fiscal(is) correspondente(s) aos itens Nº: [ ${missingNf.join(', ')} ] não foram carregadas ou vinculadas.`);
        return;
    }
    
    if (missingRat.length > 0) {
        alert(`Geração bloqueada! Os Relatórios Técnicos (RAT) correspondentes aos itens Nº: [ ${missingRat.join(', ')} ] não foram carregados ou vinculados.`);
        return;
    }
    
    toggleProgress(true, "Gerando Laudo PDF", "Renderizando a capa oficial e carregando anexos...", 10);
    
    try {
        const { PDFDocument, rgb, StandardFonts } = PDFLib;
        
        // Create PDF
        const mergedPdf = await PDFDocument.create();
        
        // Cover A4 Landscape Page (Page 1)
        const coverPage = mergedPdf.addPage([841.89, 595.27]);
        
        // Embed assets
        const logoImg = await mergedPdf.embedJpg(base64ToUint8Array(ASSETS.logo));
        const bannerImg = await mergedPdf.embedJpg(base64ToUint8Array(ASSETS.bannerClassic));
        
        // Embed font
        const fontNormal = await mergedPdf.embedFont(StandardFonts.Helvetica);
        const fontBold = await mergedPdf.embedFont(StandardFonts.HelveticaBold);
        
        // Draw Header Logo
        coverPage.drawImage(logoImg, {
            x: 40,
            y: 508,
            width: 178,
            height: 48
        });
        
        // Draw Top Right Meta
        const metaText = "F.8.5A_ATIM002";
        const metaW = fontBold.widthOfTextAtSize(metaText, 6);
        coverPage.drawText(metaText, {
            x: 695 + (110 - metaW) / 2,
            y: 544,
            size: 6,
            font: fontBold,
            color: rgb(0.2, 0.2, 0.2)
        });
        
        const labelText = "Nº:";
        const labelW = fontBold.widthOfTextAtSize(labelText, 8);
        coverPage.drawText(labelText, {
            x: 695 + (110 - labelW) / 2,
            y: 532,
            size: 8,
            font: fontBold,
            color: rgb(0, 0, 0)
        });
        
        const numText = document.getElementById("laudoNumber").value || "000";
        // Draw a light gray box for the number
        coverPage.drawRectangle({
            x: 695,
            y: 505,
            width: 110,
            height: 22,
            color: rgb(0.95, 0.95, 0.95),
            borderColor: rgb(0, 0, 0),
            borderWidth: 1.5
        });
        
        const numW = fontBold.widthOfTextAtSize(numText, 14);
        coverPage.drawText(numText, {
            x: 695 + (110 - numW) / 2,
            y: 512,
            size: 14,
            font: fontBold,
            color: rgb(0, 0, 0)
        });
        
        // Draw Header bottom line
        coverPage.drawLine({
            start: { x: 40, y: 472 },
            end: { x: 801.89, y: 472 },
            thickness: 2,
            color: rgb(0, 0, 0)
        });
        
        // Title: always "Declaração de Doação"
        const titleText = "Declaração de Doação";
        const titleWidth = fontBold.widthOfTextAtSize(titleText, 20);
        coverPage.drawText(titleText, {
            x: (841.89 - titleWidth) / 2,
            y: 512,
            size: 20,
            font: fontBold,
            color: rgb(0, 0, 0)
        });
        
        // Date
        const today = new Date();
        const dateText = formatPortugueseDate(today.toISOString().split('T')[0]);
        const dateWidth = fontNormal.widthOfTextAtSize(dateText, 7.5);
        coverPage.drawText(dateText, {
            x: 801.89 - dateWidth,
            y: 483,
            size: 7.5,
            font: fontNormal,
            color: rgb(0, 0, 0)
        });
        
        // Paragraph text formatting
        const decText = generateTextDeclaration();
        const paragraphs = decText.split('\n');
        let currentY = 445;
        
        paragraphs.forEach(p => {
            const words = p.split(' ');
            let line = '';
            const maxWidth = 762;
            
            for (let n = 0; n < words.length; n++) {
                const testLine = line + words[n] + ' ';
                const testWidth = fontNormal.widthOfTextAtSize(testLine, 9.5);
                if (testWidth > maxWidth && n > 0) {
                    coverPage.drawText(line.trim(), {
                        x: 40,
                        y: currentY,
                        size: 9.5,
                        font: fontNormal,
                        color: rgb(0, 0, 0)
                    });
                    line = words[n] + ' ';
                    currentY -= 14;
                } else {
                    line = testLine;
                }
            }
            if (line.trim().length > 0) {
                coverPage.drawText(line.trim(), {
                    x: 40,
                    y: currentY,
                    size: 9.5,
                    font: fontNormal,
                    color: rgb(0, 0, 0)
                });
                currentY -= 18;
            }
        });
        
        // Table layout definitions
        const colWidths = [20, 185, 175, 90, 60, 65, 55, 55, 57]; // total 762
        const colHeaders = [
            ["Nº"],
            ["ORIGEM /", "LOCALIDADE"],
            ["MARCA/MODELO"],
            ["SERIAL"],
            ["ÚLTIMO", "CONTADOR"],
            ["IDADE DO", "EQUIPAMENTO"],
            ["NF DE", "ENVIO"],
            ["NF DE", "DEVOLUÇÃO"],
            ["OBSERVAÇÃO", "CONTADOR"]
        ];
        const tableStartX = 40;
        let tableY = currentY - 5;
        const headerHeight = 22;
        
        // Header rectangle
        coverPage.drawRectangle({
            x: tableStartX,
            y: tableY - headerHeight,
            width: 762,
            height: headerHeight,
            color: rgb(0.95, 0.96, 0.98),
            borderColor: rgb(0, 0, 0),
            borderWidth: 1
        });
        
        let curX = tableStartX;
        colHeaders.forEach((lines, idx) => {
            const colW = colWidths[idx];
            const size = 6.5;
            
            // Center text block vertically
            const totalTextHeight = lines.length * size + (lines.length - 1) * 2;
            const startY = tableY - (headerHeight - totalTextHeight) / 2 - size;
            
            lines.forEach((lineText, lineIdx) => {
                const textW = fontBold.widthOfTextAtSize(lineText, size);
                coverPage.drawText(lineText, {
                    x: curX + (colW - textW) / 2,
                    y: startY - lineIdx * (size + 2),
                    size: size,
                    font: fontBold,
                    color: rgb(0, 0, 0)
                });
            });
            
            // Draw column separator lines
            if (idx > 0) {
                coverPage.drawLine({
                    start: { x: curX, y: tableY },
                    end: { x: curX, y: tableY - headerHeight },
                    thickness: 1,
                    color: rgb(0, 0, 0)
                });
            }
            curX += colW;
        });
        
        tableY -= headerHeight;
        
        // Draw rows
        activeEquipments.forEach((eq, index) => {
            coverPage.drawRectangle({
                x: tableStartX,
                y: tableY - 16,
                width: 762,
                height: 16,
                color: rgb(1, 1, 1),
                borderColor: rgb(0, 0, 0),
                borderWidth: 1
            });
            
            const vals = [
                (index + 1).toString(),
                eq.origin || '',
                eq.model || '',
                eq.serial || '',
                eq.counter || '0',
                eq.age || '',
                eq.nfEnvio || '',
                eq.nfDevolucao || '',
                eq.observation || ''
            ];
            
            let curX = tableStartX;
            vals.forEach((val, idx) => {
                const colW = colWidths[idx];
                const isBold = idx === 3 || idx === 8;
                const font = isBold ? fontBold : fontNormal;
                const size = idx === 1 || idx === 2 ? 7 : 7.5;
                
                // Truncate text if it overflows columns
                let displayVal = val;
                const maxChars = idx === 1 ? 40 : (idx === 2 ? 35 : 20);
                if (displayVal.length > maxChars) {
                    displayVal = displayVal.substring(0, maxChars - 3) + "...";
                }
                
                const textW = font.widthOfTextAtSize(displayVal, size);
                
                // Align text
                const alignX = curX + (colW - textW) / 2;
                
                coverPage.drawText(displayVal, {
                    x: alignX,
                    y: tableY - 11,
                    size: size,
                    font: font,
                    color: rgb(0, 0, 0)
                });
                
                // Draw inner column borders
                if (idx > 0) {
                    coverPage.drawLine({
                        start: { x: curX, y: tableY },
                        end: { x: curX, y: tableY - 16 },
                        thickness: 1,
                        color: rgb(0, 0, 0)
                    });
                }
                curX += colW;
            });
            
            tableY -= 16;
        });
        
        // Draw Footer Centralized Address
        const fAddress = LOW_COST_INFO.footerAddress;
        const fAddressW = fontNormal.widthOfTextAtSize(fAddress, 8);
        coverPage.drawText(fAddress, {
            x: (841.89 - fAddressW) / 2,
            y: 76,
            size: 8,
            font: fontNormal,
            color: rgb(0.2, 0.2, 0.2)
        });
        
        // Draw Extended Footer Banner (spanning de margem a margem)
        // Spans from x=40 to x=801.89 (width = 762 pt)
        coverPage.drawImage(bannerImg, {
            x: 40,
            y: 10,
            width: 762,
            height: 57.15 // 762 * 60 / 800 (Banner clássico 800x60)
        });
        
        // 3. Document merging
        toggleProgress(true, "Mesclando Arquivos", "Copiando páginas das Notas Fiscais e RATs...", 40);
        
        // Unique NF Envio (Notas Espelhos)
        const uniqueNfEnvioFiles = new Set();
        activeEquipments.forEach(eq => {
            if (eq.nfEnvioFile) uniqueNfEnvioFiles.add(eq.nfEnvioFile);
        });

        // Unique NF Devolucao (Notas de Retorno)
        const uniqueNfDevolucaoFiles = new Set();
        activeEquipments.forEach(eq => {
            if (eq.nfDevolucaoFile) uniqueNfDevolucaoFiles.add(eq.nfDevolucaoFile);
        });
        
        let stepCount = 0;
        const totalSteps = uniqueNfEnvioFiles.size + uniqueNfDevolucaoFiles.size + activeEquipments.length;
        
        // Merge NF Envio (Notas Espelhos)
        for (const nfName of uniqueNfEnvioFiles) {
            stepCount++;
            const percent = 40 + Math.round((stepCount / totalSteps) * 40);
            toggleProgress(true, "Mesclando Arquivos", `Mesclando nota de envio: ${nfName}...`, percent);
            
            const fileObj = uploadedFiles.find(f => f.name === nfName);
            if (fileObj) {
                const extDoc = await PDFDocument.load(fileObj.buffer.slice(0));
                const pages = await mergedPdf.copyPages(extDoc, extDoc.getPageIndices());
                pages.forEach(p => mergedPdf.addPage(p));
            }
        }

        // Merge NF Devolucao (Notas de Retorno)
        for (const nfName of uniqueNfDevolucaoFiles) {
            stepCount++;
            const percent = 40 + Math.round((stepCount / totalSteps) * 40);
            toggleProgress(true, "Mesclando Arquivos", `Mesclando nota de retorno: ${nfName}...`, percent);
            
            const fileObj = uploadedFiles.find(f => f.name === nfName);
            if (fileObj) {
                const extDoc = await PDFDocument.load(fileObj.buffer.slice(0));
                const pages = await mergedPdf.copyPages(extDoc, extDoc.getPageIndices());
                pages.forEach(p => mergedPdf.addPage(p));
            }
        }
        
        // RATs
        for (let i = 0; i < activeEquipments.length; i++) {
            const eq = activeEquipments[i];
            stepCount++;
            const percent = 40 + Math.round((stepCount / totalSteps) * 40);
            
            if (eq.ratFile) {
                toggleProgress(true, "Mesclando Arquivos", `Mesclando RAT (${i+1}/${activeEquipments.length}): ${eq.ratFile}...`, percent);
                const fileObj = uploadedFiles.find(f => f.name === eq.ratFile);
                if (fileObj) {
                    const extDoc = await PDFDocument.load(fileObj.buffer.slice(0));
                    const pages = await mergedPdf.copyPages(extDoc, extDoc.getPageIndices());
                    pages.forEach(p => mergedPdf.addPage(p));
                }
            }
        }
        
        // Evidences (Optional, appended at the end)
        const evidenceFiles = uploadedFiles.filter(f => f.type === 'evidence');
        for (const fileObj of evidenceFiles) {
            toggleProgress(true, "Mesclando Arquivos", `Mesclando evidência: ${fileObj.name}...`, 90);
            const extDoc = await PDFDocument.load(fileObj.buffer.slice(0));
            const pages = await mergedPdf.copyPages(extDoc, extDoc.getPageIndices());
            pages.forEach(p => mergedPdf.addPage(p));
        }
        
        // Save and Download
        toggleProgress(true, "Finalizando PDF", "Gerando laudo consolidado para download...", 95);
        
        const pdfBytes = await mergedPdf.save();
        const blob = new Blob([pdfBytes], { type: "application/pdf" });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = `LAUDO - ${numText}.pdf`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        toggleProgress(false);
        
    } catch (e) {
        console.error(e);
        alert("Ocorreu um erro ao gerar o PDF. Verifique os logs do console.");
        toggleProgress(false);
    }
}
