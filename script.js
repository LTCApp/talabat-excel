class ExcelProcessor {
    constructor() {
        this.init();
        this.originalData = [];
        this.processedData = [];
    }

    init() {
        const fileInput = document.getElementById('fileInput');
        const uploadArea = document.getElementById('uploadArea');

        fileInput.addEventListener('change', (e) => this.handleFile(e.target.files[0]));

        // Drag and drop functionality
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.handleFile(files[0]);
            }
        });

        uploadArea.addEventListener('click', () => {
            fileInput.click();
        });
    }

    async handleFile(file) {
        if (!file) return;

        // Validate file type
        const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                           'application/vnd.ms-excel'];
        if (!validTypes.includes(file.type)) {
            this.showNotification('يرجى اختيار ملف Excel صحيح (.xlsx أو .xls)', 'error');
            return;
        }

        this.showLoading(true);

        try {
            const data = await this.readExcelFile(file);
            this.originalData = data;
            this.processData();
            this.displayResults();
        } catch (error) {
            console.error('Error processing file:', error);
            this.showNotification('حدث خطأ في معالجة الملف. يرجى التأكد من تنسيق الملف.', 'error');
        } finally {
            this.showLoading(false);
        }
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    resolve(jsonData);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = () => reject(new Error('فشل في قراءة الملف'));
            reader.readAsArrayBuffer(file);
        });
    }

    processData() {
        if (this.originalData.length < 2) {
            throw new Error('الملف لا يحتوي على بيانات كافية');
        }

        // Skip header row
        const dataRows = this.originalData.slice(1);
        this.processedData = [];

        let removedCount = 0;
        let priceIncreasedCount = 0;

        dataRows.forEach((row, index) => {
            try {
                // Extract columns (0-indexed)
                const itemCode = row[0]; // كود الصنف
                const itemName = row[1]; // اسم الصنف
                const unitPrice = parseFloat(row[2]) || 0; // سعر الوحدة
                const unit = row[3]; // الوحدة
                const section = row[8]; // القسم

                // Skip rows where unit is not 1
                if (unit != 1) {
                    removedCount++;
                    return;
                }

                // Calculate new price
                let newPrice = unitPrice;
                let priceIncreased = false;

                // Add 7.5% if section is not 52
                if (section != 52) {
                    newPrice = unitPrice * 1.075;
                    priceIncreased = true;
                    priceIncreasedCount++;
                }

                // Apply custom rounding
                newPrice = this.customRound(newPrice);

                this.processedData.push({
                    itemCode: itemCode || '',
                    itemName: itemName || '',
                    originalPrice: unitPrice,
                    newPrice: newPrice,
                    priceIncreased: priceIncreased,
                    section: section
                });
            } catch (error) {
                console.warn(`خطأ في السطر ${index + 2}:`, error);
            }
        });

        this.stats = {
            total: this.processedData.length,
            removed: removedCount,
            priceIncreased: priceIncreasedCount,
            originalTotal: dataRows.length
        };
    }

    // Custom rounding function
    customRound(price) {
        const wholePart = Math.floor(price);
        const decimalPart = price - wholePart;
        
        if (decimalPart === 0) {
            return price; // Already a whole number
        } else if (decimalPart <= 0.49) {
            return wholePart + 0.5; // Round to .5
        } else if (decimalPart === 0.5) {
            return price; // Keep as .5
        } else {
            return wholePart + 1; // Round up to next whole number
        }
    }

    displayResults() {
        // Show stats
        const statsDiv = document.getElementById('stats');
        statsDiv.innerHTML = `
            <div class="stat-item">
                <span class="stat-number">${this.stats.originalTotal}</span>
                <div class="stat-label">إجمالي السطور الأصلية</div>
            </div>
            <div class="stat-item">
                <span class="stat-number">${this.stats.removed}</span>
                <div class="stat-label">السطور المحذوفة</div>
            </div>
            <div class="stat-item">
                <span class="stat-number">${this.stats.total}</span>
                <div class="stat-label">السطور المعروضة</div>
            </div>
            <div class="stat-item">
                <span class="stat-number">${this.stats.priceIncreased}</span>
                <div class="stat-label">الأسعار المعدلة (+7.5%)</div>
            </div>
        `;

        // Display table
        const tableBody = document.getElementById('tableBody');
        tableBody.innerHTML = '';

        this.processedData.forEach((item, index) => {
            const row = document.createElement('tr');
            
            const priceClass = item.priceIncreased ? 'price-increased' : 'price-original';
            // Display price: if it's a whole number, show without decimals, otherwise show with 1 decimal
            const priceDisplay = item.newPrice % 1 === 0 ? item.newPrice.toString() : item.newPrice.toFixed(1);
            
            row.innerHTML = `
                <td>
                    <span class="clickable" onclick="copyToClipboard('${item.itemCode}')">
                        ${item.itemCode}
                    </span>
                </td>
                <td>${item.itemName}</td>
                <td>
                    <span class="clickable ${priceClass}" onclick="copyToClipboard('${priceDisplay}')">
                        ${priceDisplay}
                    </span>
                </td>
            `;
            
            tableBody.appendChild(row);
        });

        // Show results section
        document.getElementById('resultsSection').style.display = 'block';
        
        this.showNotification(`تم معالجة ${this.stats.total} صنف بنجاح`, 'success');
    }

    showLoading(show) {
        document.getElementById('loading').style.display = show ? 'block' : 'none';
    }

    showNotification(message, type = 'success') {
        const notification = document.getElementById('notification');
        notification.textContent = message;
        notification.className = `notification ${type} show`;
        
        setTimeout(() => {
            notification.classList.remove('show');
        }, 3000);
    }

    // Export processed data to Excel
    exportToExcel() {
        if (this.processedData.length === 0) {
            this.showNotification('لا توجد بيانات للتصدير', 'error');
            return;
        }

        // Prepare data for export
        const exportData = [
            ['كود الصنف', 'اسم الصنف', 'سعر الوحدة'] // Header
        ];

        this.processedData.forEach(item => {
            const priceDisplay = item.newPrice % 1 === 0 ? item.newPrice.toString() : item.newPrice.toFixed(1);
            exportData.push([
                item.itemCode,
                item.itemName,
                parseFloat(priceDisplay)
            ]);
        });

        // Create workbook
        const ws = XLSX.utils.aoa_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'الأصناف المعدلة');

        // Generate filename with timestamp
        const now = new Date();
        const timestamp = now.toISOString().slice(0, 19).replace(/[:.]/g, '-');
        const filename = `الأصناف_المعدلة_${timestamp}.xlsx`;

        // Download file
        XLSX.writeFile(wb, filename);
        
        this.showNotification('تم تحميل الملف بنجاح', 'success');
    }
}

// Global variable to access processor instance
let processorInstance = null;

// Copy to clipboard function
function copyToClipboard(text) {
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(text).then(() => {
            showCopyNotification('تم نسخ النص: ' + text);
        }).catch(err => {
            console.error('فشل في النسخ:', err);
            fallbackCopyTextToClipboard(text);
        });
    } else {
        fallbackCopyTextToClipboard(text);
    }
}

// Fallback copy function for older browsers
function fallbackCopyTextToClipboard(text) {
    const textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.top = '0';
    textArea.style.left = '0';
    textArea.style.position = 'fixed';
    
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    
    try {
        const successful = document.execCommand('copy');
        if (successful) {
            showCopyNotification('تم نسخ النص: ' + text);
        } else {
            showCopyNotification('فشل في النسخ', 'error');
        }
    } catch (err) {
        console.error('فشل في النسخ:', err);
        showCopyNotification('فشل في النسخ', 'error');
    }
    
    document.body.removeChild(textArea);
}

// Show copy notification
function showCopyNotification(message, type = 'success') {
    const notification = document.getElementById('notification');
    notification.textContent = message;
    notification.className = `notification ${type} show`;
    
    setTimeout(() => {
        notification.classList.remove('show');
    }, 2000);
}

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    processorInstance = new ExcelProcessor();
});

// Global download function
function downloadExcel() {
    if (processorInstance) {
        processorInstance.exportToExcel();
    }
}