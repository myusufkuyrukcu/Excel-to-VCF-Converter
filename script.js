document.getElementById('fileInput').addEventListener('change', handleFile);
document.getElementById('headerCheckbox').addEventListener('change', handleHeaderChange);
document.getElementById('showAllRowsCheckbox').addEventListener('change', handleShowAllRowsChange);
document.getElementById('addAreaCodeCheckbox').addEventListener('change', handleAddAreaCodeChange);
document.getElementById('areaCode').addEventListener('input', handleAreaCodeChange);
document.getElementById('onlyTenPlusCheckbox').addEventListener('change', handleOnlyTenPlusChange);
document.getElementById('minLength').addEventListener('input', handleMinLengthChange);

let globalData = [];
let types = [];
let showAllRows = false;
let areaCode = "+90";
let onlyTenPlus = false;
let minLength = 10;

function handleFile(event) {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        globalData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });

        document.getElementById('totalRows').textContent = globalData.length;
        document.getElementById('showAllRowsCheckbox').disabled = false;
        document.getElementById('addAreaCodeCheckbox').disabled = false;
        document.getElementById('onlyTenPlusCheckbox').disabled = !document.getElementById('addAreaCodeCheckbox').checked;
        document.getElementById('minLength').disabled = !document.getElementById('onlyTenPlusCheckbox').checked;
        
        displayPreview(globalData);
    };

    reader.readAsArrayBuffer(file);
}

function handleHeaderChange() {
    displayPreview(globalData);
}

function handleShowAllRowsChange() {
    showAllRows = document.getElementById('showAllRowsCheckbox').checked;
    displayPreview(globalData);
}

function handleAddAreaCodeChange() {
    const addAreaCode = document.getElementById('addAreaCodeCheckbox').checked;
    document.getElementById('areaCode').disabled = !addAreaCode;
    document.getElementById('onlyTenPlusCheckbox').disabled = !addAreaCode;
    document.getElementById('minLength').disabled = !addAreaCode || !document.getElementById('onlyTenPlusCheckbox').checked;
    displayPreview(globalData);
}

function handleOnlyTenPlusChange() {
    onlyTenPlus = document.getElementById('onlyTenPlusCheckbox').checked;
    document.getElementById('minLength').disabled = !onlyTenPlus;
    displayPreview(globalData);
}

function handleAreaCodeChange() {
    areaCode = document.getElementById('areaCode').value;
    displayPreview(globalData);
}

function handleMinLengthChange() {
    minLength = parseInt(document.getElementById('minLength').value, 10) || 0;
    displayPreview(globalData);
}

function displayPreview(data) {
    const previewDiv = document.getElementById('preview');
    previewDiv.innerHTML = '';

    if (data.length > 0) {
        const table = document.createElement('table');
        const headerRow = document.createElement('tr');
        const hasHeader = document.getElementById('headerCheckbox').checked;
        const addAreaCode = document.getElementById('addAreaCodeCheckbox').checked;

        const firstRow = hasHeader ? data[0] : Array.from({ length: data[0].length }, (_, i) => `Column ${i + 1}`);

        firstRow.forEach((cell, index) => {
            const th = document.createElement('th');
            const div = document.createElement('div');
            div.innerHTML = cell;
            th.appendChild(div);
            const select = document.createElement('select');
            select.innerHTML = `
                <option value="none">Select Type</option>
                <option value="firstName">First Name</option>
                <option value="middleName">Middle Name</option>
                <option value="lastName">Last Name</option>
                <option value="email">E-mail Address</option>
                <option value="mobile">Mobile Phone</option>
                <option value="home">Home Phone</option>
                <option value="home2">Home Phone 2</option>
                <option value="business">Business Phone</option>
                <option value="business2">Business Phone 2</option>
                <option value="other">Other Phone</option>
                <option value="notes">Notes</option>
            `;
            select.dataset.index = index;
            if (types[index]) {
                select.value = types[index];
            }
            select.addEventListener('change', () => {
                types[index] = select.value;
                displayPreview(globalData);
            });
            th.appendChild(select);
            headerRow.appendChild(th);
        });

        table.appendChild(headerRow);

        const startIndex = hasHeader ? 1 : 0;
        const endIndex = showAllRows ? data.length : Math.min(startIndex + 20, data.length);

        for (let i = startIndex; i < endIndex; i++) {
            const row = document.createElement('tr');
            data[i].forEach((cell, index) => {
                const td = document.createElement('td');
                if (addAreaCode && ['mobile', 'home', 'home2', 'business', 'business2', 'other'].includes(types[index]) && cell) {
                    if (!onlyTenPlus || cell.toString().length >= minLength) {
                        td.innerHTML = `${areaCode}${cell}`;
                    } else {
                        td.innerHTML = cell;
                    }
                } else {
                    td.innerHTML = cell;
                }
                row.appendChild(td);
            });
            table.appendChild(row);
        }

        previewDiv.appendChild(table);
        document.getElementById('exportButton').style.display = 'block';
    }
}

document.getElementById('exportButton').addEventListener('click', exportToVCF);

function exportToVCF() {
    const file = document.getElementById('fileInput').files[0];
    const reader = new FileReader();

    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const sheetData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });

        const contacts = [];
        const startIndex = document.getElementById('headerCheckbox').checked ? 1 : 0;

        for (let i = startIndex; i < sheetData.length; i++) {
            const contact = {};
            types.forEach((type, index) => {
                if (type !== 'none' && sheetData[i][index]) {
                    if (['mobile', 'home', 'home2', 'business', 'business2', 'other'].includes(type) && document.getElementById('addAreaCodeCheckbox').checked) {
                        if (!onlyTenPlus || sheetData[i][index].toString().length >= minLength) {
                            contact[type] = `${areaCode}${sheetData[i][index]}`;
                        } else {
                            contact[type] = sheetData[i][index];
                        }
                    } else {
                        contact[type] = sheetData[i][index];
                    }
                }
            });
            if (Object.keys(contact).length > 0) {
                contacts.push(contact);
            }
        }

        const vcfContent = generateVCF(contacts);
        downloadVCF(vcfContent);
    };

    reader.readAsArrayBuffer(file);
}

function generateVCF(contacts) {
    let vcf = '';
    contacts.forEach(contact => {
        const firstName = contact.firstName || '';
        const middleName = contact.middleName || '';
        const lastName = contact.lastName || '';
        const fullName = [firstName, middleName, lastName].filter(Boolean).join(' ');

        vcf += 'BEGIN:VCARD\n';
        vcf += 'VERSION:3.0\n';
        if (fullName) vcf += `FN:${fullName}\n`;
        if (lastName || firstName) vcf += `N:${lastName};${firstName};;;\n`;
        if (contact.email) vcf += `EMAIL;TYPE=INTERNET:${contact.email}\n`;
        if (contact.mobile) vcf += `TEL;TYPE=CELL:${contact.mobile}\n`;
        if (contact.home) vcf += `TEL;TYPE=HOME:${contact.home}\n`;
        if (contact.home2) vcf += `TEL;TYPE=HOME:${contact.home2}\n`;
        if (contact.business) vcf += `TEL;TYPE=WORK:${contact.business}\n`;
        if (contact.business2) vcf += `TEL;TYPE=WORK:${contact.business2}\n`;
        if (contact.other) vcf += `TEL;TYPE=OTHER:${contact.other}\n`;
        if (contact.notes) vcf += `NOTE:${contact.notes}\n`;
        vcf += 'END:VCARD\n';
    });
    return vcf;
}

function downloadVCF(content) {
    const blob = new Blob([content], { type: 'text/vcard' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'contacts.vcf';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
