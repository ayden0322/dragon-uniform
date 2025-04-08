// 公用函數模組

/**
 * 顯示警告提示
 * @param {string} message - 提示訊息
 * @param {string} type - 提示類型 (success, danger, warning, info)
 * @param {number} duration - 持續時間(毫秒)
 */
export function showAlert(message, type = 'info', duration = 3000) {
    const alertContainer = document.getElementById('alertContainer');
    const alertElement = document.createElement('div');
    
    // 設置提示樣式
    alertElement.className = `alert alert-${type} alert-dismissible fade show`;
    alertElement.role = 'alert';
    alertElement.innerHTML = `
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
    `;
    
    // 加入頁面
    alertContainer.appendChild(alertElement);
    
    // 設定自動消失
    setTimeout(() => {
        const bsAlert = bootstrap.Alert.getOrCreateInstance(alertElement);
        bsAlert.close();
    }, duration);
}

/**
 * 儲存資料到本地儲存
 * @param {string} key - 儲存鍵值
 * @param {any} data - 要儲存的資料
 */
export function saveToLocalStorage(key, data) {
    try {
        const serializedData = JSON.stringify(data);
        localStorage.setItem(key, serializedData);
        return true;
    } catch (error) {
        console.error(`儲存資料到本地儲存失敗: ${error.message}`);
        return false;
    }
}

/**
 * 從本地儲存讀取資料
 * @param {string} key - 儲存鍵值
 * @param {any} defaultValue - 預設值(如果沒有找到資料)
 * @returns {any} 讀取的資料或預設值
 */
export function loadFromLocalStorage(key, defaultValue = null) {
    try {
        const serializedData = localStorage.getItem(key);
        if (serializedData === null) {
            return defaultValue;
        }
        return JSON.parse(serializedData);
    } catch (error) {
        console.error(`從本地儲存讀取資料失敗: ${error.message}`);
        return defaultValue;
    }
}

/**
 * 格式化日期時間
 * @param {Date} date - 日期物件
 * @returns {string} 格式化的日期字串 (YYYY-MM-DD HH:MM:SS)
 */
export function formatDateTime(date = new Date()) {
    const pad = (num) => num.toString().padStart(2, '0');
    
    const year = date.getFullYear();
    const month = pad(date.getMonth() + 1);
    const day = pad(date.getDate());
    const hours = pad(date.getHours());
    const minutes = pad(date.getMinutes());
    const seconds = pad(date.getSeconds());
    
    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

/**
 * 建立CSV資料
 * @param {Array} headers - 表頭欄位
 * @param {Array} rows - 資料行
 * @returns {string} CSV格式化字串
 */
export function createCSV(headers, rows) {
    const headerRow = headers.join(',');
    const dataRows = rows.map(row => row.join(',')).join('\n');
    return `${headerRow}\n${dataRows}`;
}

/**
 * 下載檔案
 * @param {string} filename - 檔案名稱
 * @param {string} content - 檔案內容
 * @param {string} type - 檔案類型
 */
export function downloadFile(filename, content, type = 'text/csv') {
    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

/**
 * 建立和下載CSV檔案
 * @param {string} filename - 檔案名稱
 * @param {Array} headers - 表頭欄位
 * @param {Array} rows - 資料行
 */
export function downloadCSV(filename, headers, rows) {
    const csvContent = createCSV(headers, rows);
    downloadFile(filename, csvContent);
}

/**
 * 建立和下載Excel檔案
 * @param {string} filename - 檔案名稱
 * @param {Array} headers - 表頭欄位
 * @param {Array} rows - 資料行
 */
export function downloadExcel(filename, headers, rows) {
    // 確保檔名有 .xlsx 副檔名
    if (!filename.endsWith('.xlsx')) {
        filename += '.xlsx';
    }
    
    // 創建新的工作簿
    const wb = XLSX.utils.book_new();
    
    // 創建一個工作表
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    
    // 將工作表添加到工作簿
    XLSX.utils.book_append_sheet(wb, ws, "學生資料");
    
    // 將工作簿轉換為 ArrayBuffer
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    
    // 將 ArrayBuffer 轉換為 Blob
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    // 創建下載連結
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

/**
 * 隨機生成唯一ID
 * @returns {string} 唯一ID
 */
export function generateUniqueId() {
    return Date.now().toString(36) + Math.random().toString(36).substring(2);
} 