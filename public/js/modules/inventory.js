// 庫存相關功能模組
import { SIZES, tableIdToInventoryType, inventoryTypeToTableId, UNIFORM_TYPES, formatSize, RESERVE_RATIO } from './config.js';
import { saveToLocalStorage, loadFromLocalStorage, showAlert } from './utils.js';
import { updateAllocationRatios } from './ui.js';
import { demandData } from './students.js';

// 庫存資料結構
export let inventoryData = {
    shortSleeveShirt: {},
    shortSleevePants: {},
    longSleeveShirt: {},
    longSleevePants: {}
};

// 新增: 可分配數量與預留數量的手動調整資料
export let manualAdjustments = {
    shortSleeveShirt: {},
    shortSleevePants: {},
    longSleeveShirt: {},
    longSleevePants: {}
};

/**
 * 初始化庫存功能
 */
export function initInventoryFeatures() {
    // 初始化庫存表格
    initInventoryTables();
    
    // 綁定庫存匯入/匯出按鈕事件
    document.getElementById('importInventoryBtn')?.addEventListener('click', handleInventoryImport);
    document.getElementById('downloadInventoryTemplateBtn')?.addEventListener('click', downloadInventoryTemplate);
    document.getElementById('clearInventoryDataBtn')?.addEventListener('click', confirmClearInventoryData);
}

/**
 * 初始化庫存表格
 */
export function initInventoryTables() {
    const tableIds = [
        'shortSleeveShirtTable',
        'shortSleevePantsTable',
        'longSleeveShirtTable',
        'longSleevePantsTable'
    ];

    tableIds.forEach(tableId => {
        const table = document.getElementById(tableId);
        if (!table) return;

        // 清空表格
        const tbody = table.querySelector('tbody');
        tbody.innerHTML = '';

        // 為每個尺寸添加行
        SIZES.forEach(size => {
            const row = document.createElement('tr');
            // 使用尺寸格式化函數來顯示尺寸
            row.innerHTML = `
                <td data-original-size="${size}">${formatSize(size)}</td>
                <td><input type="number" class="form-control total-inventory" min="0" value="0"></td>
            `;
            tbody.appendChild(row);

            // 為輸入欄位添加事件監聽器
            const totalInput = row.querySelector('.total-inventory');

            // 更新庫存數據
            totalInput.addEventListener('change', () => {
                updateInventoryData(tableId, size, 'total', parseInt(totalInput.value) || 0);
            });
        });

        // 加載保存的數據
        loadInventoryData(tableId);
    });
}

/**
 * 更新庫存資料
 * @param {string} tableId - 表格ID
 * @param {string} size - 尺寸
 * @param {string} field - 欄位 (total)
 * @param {number} value - 數值
 */
export function updateInventoryData(tableId, size, field, value) {
    const type = tableIdToInventoryType(tableId);
    if (!type) return;

    // 確保該類型的庫存物件存在
    if (!inventoryData[type]) {
        inventoryData[type] = {};
    }

    // 確保該尺寸的庫存物件存在
    if (!inventoryData[type][size]) {
        inventoryData[type][size] = { total: 0 };
    }
    
    // 檢查是否更新總數，如果是則清除該尺寸的手動調整值
    if (field === 'total' && manualAdjustments[type] && manualAdjustments[type][size] !== undefined) {
        // 清除該尺寸的手動調整值
        delete manualAdjustments[type][size];
        // 保存到本地儲存
        saveToLocalStorage('manualAdjustments', manualAdjustments);
    }

    // 更新指定欄位
    inventoryData[type][size][field] = value;

    // 更新UI
    updateInventoryUI(tableId, size);

    // 保存到本地儲存
    saveToLocalStorage('inventoryData', inventoryData);
    
    // 更新分配比率
    updateAllocationRatios();
}

/**
 * 更新庫存UI
 * @param {string} tableId - 表格ID
 * @param {string} size - 尺寸
 */
export function updateInventoryUI(tableId, size) {
    const table = document.getElementById(tableId);
    if (!table) return;
    
    const type = tableIdToInventoryType(tableId);
    if (!type || !inventoryData[type] || !inventoryData[type][size]) return;
    
    // 尋找對應的行
    const rows = table.querySelectorAll('tbody tr');
    let targetRow = null;
    
    for (const row of rows) {
        const sizeCell = row.querySelector('td:first-child');
        if (sizeCell && sizeCell.dataset.originalSize === size) {
            // 更新尺寸顯示
            sizeCell.textContent = formatSize(size);
            targetRow = row;
            break;
        }
    }
    
    if (!targetRow) return;
    
    const data = inventoryData[type][size];
    
    // 更新UI顯示
    const totalInput = targetRow.querySelector('.total-inventory');
    
    if (totalInput) totalInput.value = data.total;
}

/**
 * 加載庫存資料
 * @param {string} tableId - 表格ID
 */
export function loadInventoryData(tableId) {
    const savedData = loadFromLocalStorage('inventoryData', null);
    if (!savedData) return;

    const type = tableIdToInventoryType(tableId);
    if (!type || !savedData[type]) return;

    // 更新內存中的庫存資料
    inventoryData = savedData;

    // 更新UI
    const table = document.getElementById(tableId);
    if (!table) return;

    SIZES.forEach(size => {
        // 查找對應行
        const rows = table.querySelectorAll('tbody tr');
        let row = null;

        for (const r of rows) {
            const sizeCell = r.querySelector('td:first-child');
            if (sizeCell && sizeCell.dataset.originalSize === size) {
                row = r;
                break;
            }
        }

        if (!row) return;

        // 更新輸入欄位
        const input = row.querySelector('.total-inventory');
        if (!input) return;

        const total = savedData[type][size]?.total || 0;
        input.value = total;
    });
}

/**
 * 計算預留數量 - 使用固定預留比例10%
 * @param {Object} inventoryData - 庫存資料
 * @param {Object} demandData - 需求資料
 * @param {Object} manualAdjustments - 手動調整資料
 */
export function calculateReservedQuantities(inventoryData, demandData, manualAdjustments) {
    // 處理每種制服類型
    for (const type in inventoryData) {
        if (!inventoryData.hasOwnProperty(type)) continue;
        
        // 計算每個尺寸
        for (const size in inventoryData[type]) {
            if (!inventoryData[type].hasOwnProperty(size)) continue;
            
            const totalInventory = inventoryData[type][size].total || 0;
            let allocatable = 0;
            
            // 檢查是否有手動調整
            if (manualAdjustments && 
                manualAdjustments[type] && 
                manualAdjustments[type][size] !== undefined) {
                // 使用手動調整的值
                allocatable = manualAdjustments[type][size];
            } else {
                // 計算預留數量（無條件進位）
                const reserved = Math.ceil(totalInventory * RESERVE_RATIO);
                // 計算可分配數量
                allocatable = totalInventory - reserved;
            }
            
            // 確保可分配數量不超過總庫存且不小於0
            allocatable = Math.min(Math.max(allocatable, 0), totalInventory);
            
            // 更新庫存資料
            inventoryData[type][size].allocatable = allocatable;
            inventoryData[type][size].reserved = totalInventory - allocatable;
        }
    }
    
    // 保存到本地儲存
    saveToLocalStorage('inventoryData', inventoryData);
    
    return inventoryData;
}

/**
 * 計算庫存類型的總數
 * @param {Object} typeInventory - 某類型的庫存資料
 * @returns {number} 總庫存數量
 */
export function calculateTotalInventory(typeInventory) {
    let total = 0;
    for (const size in typeInventory) {
        if (typeInventory.hasOwnProperty(size)) {
            total += typeInventory[size].total || 0;
        }
    }
    return total;
}

/**
 * 更新尺寸表格
 * @param {string} type - 制服類型
 */
export function updateSizeTable(type) {
    const tableId = `${type}AdjustTable`;
    const table = document.getElementById(tableId);
    if (!table) return;

    const tbody = table.querySelector('tbody');
    tbody.innerHTML = '';

    // 使用固定的預留比例10%
    const reserveRatio = RESERVE_RATIO; // 10%

    // 初始化總數計數器
    let totalInventory = 0;
    let totalCalculatedAllocatable = 0;
    let totalManualAdjustment = 0;
    let totalAdjustedAllocatable = 0;
    let totalReserved = 0;

    // 為每個尺寸添加行
    for (const size of SIZES) {
        if (!inventoryData[type][size]) continue;
        
        const data = inventoryData[type][size];
        const total = data.total || 0;
        
        // 使用固定預留比例計算預留數量（無條件進位）
        const reservedQuantity = Math.ceil(total * reserveRatio);
        // 計算可分配數量
        const calculatedAllocatable = total - reservedQuantity;
        
        // 檢查是否有手動調整，如果沒有則使用計算的可分配數
        let manualAdjustment = calculatedAllocatable;
        let adjustedAllocatable = calculatedAllocatable;
        
        if (manualAdjustments[type] && manualAdjustments[type][size] !== undefined) {
            manualAdjustment = manualAdjustments[type][size];
            adjustedAllocatable = manualAdjustment;
        }
        
        const reserved = total - adjustedAllocatable;
        
        // 更新總數
        totalInventory += total;
        totalCalculatedAllocatable += calculatedAllocatable;
        totalManualAdjustment += manualAdjustment;
        totalAdjustedAllocatable += adjustedAllocatable;
        totalReserved += reserved;

        // 創建行
        const row = document.createElement('tr');
        row.dataset.size = size;
        row.innerHTML = `
            <td data-original-size="${size}">${formatSize(size)}</td>
            <td>${total}</td>
            <td>${calculatedAllocatable}</td>
            <td class="adjustment-cell">
                <div class="input-group adjustment-group">
                    <div class="input-group-prepend">
                        <button type="button" class="btn btn-outline-primary decrement-btn">−</button>
                    </div>
                    <input type="number" class="form-control manual-adjustment" 
                        value="${manualAdjustment}" 
                        min="0" max="${total}" step="1" 
                        data-original="${calculatedAllocatable}">
                    <div class="input-group-append">
                        <button type="button" class="btn btn-outline-primary increment-btn">+</button>
                    </div>
                </div>
            </td>
            <td class="adjusted-allocatable">${adjustedAllocatable}</td>
            <td class="reserved">${reserved}</td>
        `;
        
        tbody.appendChild(row);

        // 添加事件監聽器
        const input = row.querySelector('.manual-adjustment');
        const incrementBtn = row.querySelector('.increment-btn');
        const decrementBtn = row.querySelector('.decrement-btn');

        // 數字輸入框事件監聽器
        input.addEventListener('input', function() {
            const value = parseInt(this.value) || 0;
            const max = parseInt(this.getAttribute('max')) || 0;
            
            // 確保值不超過最大值
            if (value > max) {
                this.value = max;
                return;
            }
            
            // 更新手動調整數據
            if (!manualAdjustments[type]) {
                manualAdjustments[type] = {};
            }
            manualAdjustments[type][size] = value;
            
            // 更新顯示
            const row = this.closest('tr');
            row.querySelector('.adjusted-allocatable').textContent = value;
            row.querySelector('.reserved').textContent = total - value;
            
            // 直接更新總計行（立即反映變更）
            manualUpdateTotalRow(type, table);
            
            // 保存手動調整到本地儲存
            saveToLocalStorage('manualAdjustments', manualAdjustments);
        });

        // 增加按鈕事件監聽器
        incrementBtn.addEventListener('click', function() {
            const currentValue = parseInt(input.value) || 0;
            const max = parseInt(input.getAttribute('max')) || 0;
            
            if (currentValue < max) {
                input.value = currentValue + 1;
                // 觸發 input 事件以更新數據
                input.dispatchEvent(new Event('input'));
            }
        });
        
        // 減少按鈕事件監聽器
        decrementBtn.addEventListener('click', function() {
            const currentValue = parseInt(input.value) || 0;
            if (currentValue > 0) {
                input.value = currentValue - 1;
                // 觸發 input 事件以更新數據
                input.dispatchEvent(new Event('input'));
            }
        });
    }
    
    // 添加總計行
    const totalRow = document.createElement('tr');
    totalRow.classList.add('table-info', 'total-row');
    totalRow.innerHTML = `
        <td><strong>總計</strong></td>
        <td><strong>${totalInventory}</strong></td>
        <td><strong>${totalCalculatedAllocatable}</strong></td>
        <td><strong>${totalManualAdjustment}</strong></td>
        <td><strong>${totalAdjustedAllocatable}</strong></td>
        <td><strong>${totalReserved}</strong></td>
    `;
    tbody.appendChild(totalRow);
}

/**
 * 手動更新總計行
 * @param {string} type - 制服類型
 * @param {HTMLElement} table - 表格元素
 */
function manualUpdateTotalRow(type, table) {
    if (!table) return;
    
    // 獲取總計行
    const totalRow = table.querySelector('.total-row');
    if (!totalRow) return;
    
    // 計算總數
    let totalInventory = 0;
    let totalCalculatedAllocatable = 0;
    let totalManualAdjustment = 0;
    let totalAdjustedAllocatable = 0;
    let totalReserved = 0;
    
    // 從表格中獲取數據
    const rows = table.querySelectorAll('tbody tr:not(.total-row)');
    for (const row of rows) {
        const total = parseInt(row.querySelector('td:nth-child(2)').textContent) || 0;
        const calculatedAllocatable = parseInt(row.querySelector('td:nth-child(3)').textContent) || 0;
        const manualAdjustment = parseInt(row.querySelector('.manual-adjustment').value) || 0;
        const adjustedAllocatable = parseInt(row.querySelector('.adjusted-allocatable').textContent) || 0;
        const reserved = parseInt(row.querySelector('.reserved').textContent) || 0;
        
        totalInventory += total;
        totalCalculatedAllocatable += calculatedAllocatable;
        totalManualAdjustment += manualAdjustment;
        totalAdjustedAllocatable += adjustedAllocatable;
        totalReserved += reserved;
    }
    
    // 更新總計行
    totalRow.innerHTML = `
        <td><strong>總計</strong></td>
        <td><strong>${totalInventory}</strong></td>
        <td><strong>${totalCalculatedAllocatable}</strong></td>
        <td><strong>${totalManualAdjustment}</strong></td>
        <td><strong>${totalAdjustedAllocatable}</strong></td>
        <td><strong>${totalReserved}</strong></td>
    `;
}

/**
 * 更新總計行 - 使用固定預留比例10%計算總需求和庫存
 * @param {string} type - 制服類型
 */
function updateTotalRow(type) {
    const table = document.getElementById(`${type}AdjustTable`);
    if (!table) return;
    
    // 計算總數
    let totalInventory = 0;
    let totalCalculatedAllocatable = 0;
    let totalManualAdjustment = 0;
    let totalAdjustedAllocatable = 0;
    let totalReserved = 0;
    
    // 從庫存數據中獲取數據
    for (const size in inventoryData[type]) {
        if (!inventoryData[type].hasOwnProperty(size)) continue;
        
        const total = inventoryData[type][size].total || 0;
        totalInventory += total;
        
        let calculatedAllocatable = 0;
        let manualAdjustment = 0;
        let adjustedAllocatable = 0;
        let reserved = 0;
        
        // 使用固定的預留比例10%
        // 計算預留數量（無條件進位）
        const reservedQuantity = Math.ceil(total * RESERVE_RATIO);
        // 計算可分配數量
        calculatedAllocatable = total - reservedQuantity;
        
        // 檢查是否有手動調整
        if (manualAdjustments[type] && manualAdjustments[type][size] !== undefined) {
            manualAdjustment = manualAdjustments[type][size];
            adjustedAllocatable = manualAdjustment;
        } else {
            manualAdjustment = calculatedAllocatable;
            adjustedAllocatable = calculatedAllocatable;
        }
        
        reserved = total - adjustedAllocatable;
        
        totalCalculatedAllocatable += calculatedAllocatable;
        totalManualAdjustment += manualAdjustment;
        totalAdjustedAllocatable += adjustedAllocatable;
        totalReserved += reserved;
    }
    
    // 更新總計行
    const totalRow = table.querySelector('.total-row');
    if (totalRow) {
        totalRow.innerHTML = `
            <td><strong>總計</strong></td>
            <td><strong>${totalInventory}</strong></td>
            <td><strong>${totalCalculatedAllocatable}</strong></td>
            <td><strong>${totalManualAdjustment}</strong></td>
            <td><strong>${totalAdjustedAllocatable}</strong></td>
            <td><strong>${totalReserved}</strong></td>
        `;
    }
}

/**
 * 保存手動調整
 */
export function saveManualAdjustments() {
    let hasWarning = false;
    let warningMessages = [];
    
    // 遍歷所有制服類型並更新庫存資料
    for (const type in manualAdjustments) {
        if (!manualAdjustments.hasOwnProperty(type)) continue;
        
        let typeTotal = 0; // 追蹤每種制服類型的總可分配數
        
        // 遍歷所有尺寸，包括沒有手動調整的尺寸
        for (const size in inventoryData[type]) {
            if (!inventoryData[type].hasOwnProperty(size)) continue;
            
            const total = inventoryData[type][size].total || 0;
            let allocatable;
            
            // 檢查該尺寸是否有手動調整
            if (manualAdjustments[type] && manualAdjustments[type][size] !== undefined) {
                // 使用手動調整的值
                allocatable = Math.min(manualAdjustments[type][size], total);
            } else {
                // 使用原來計算的值（比例計算）
                const ratioElem = document.getElementById(`${type}Ratio`);
                let ratio = 0;
                if (ratioElem) {
                    const percentText = ratioElem.textContent || '0%';
                    ratio = parseFloat(percentText) / 100 || 0;
                }
                allocatable = Math.round(total * ratio);
            }
            
            // 更新庫存資料
            inventoryData[type][size].allocatable = allocatable;
            inventoryData[type][size].reserved = total - allocatable;
            
            // 累加可分配數
            typeTotal += allocatable;
        }
        
        // 更新對應表格的UI
        const tableId = inventoryTypeToTableId(type);
        if (tableId) {
            for (const size in inventoryData[type]) {
                updateInventoryUI(tableId, size);
            }
        }

        // 檢查當前制服類型的可分配數是否小於需求量
        const demand = demandData[type]?.totalDemand || 0;
        if (typeTotal < demand) {
            const uniformTypeName = UNIFORM_TYPES[type];
            const shortfall = demand - typeTotal;
            const message = `${uniformTypeName}：可分配數總和(${typeTotal})小於需求量(${demand})，差額${shortfall}件`;
            console.warn(message);
            warningMessages.push(message);
            hasWarning = true;
        }

        // 更新表格
        updateSizeTable(type);
    }
    
    // 保存到本地儲存
    saveToLocalStorage('inventoryData', inventoryData);
    saveToLocalStorage('manualAdjustments', manualAdjustments);
    
    // 更新分配比率
    updateAllocationRatios();
    
    // 顯示訊息
    if (hasWarning) {
        const warningHTML = `
            <div>手動調整已保存，但以下制服類型的可分配數小於需求量：</div>
            <ul style="text-align:left; margin-top:8px; padding-left:20px;">
                ${warningMessages.map(msg => `<li>${msg}</li>`).join('')}
            </ul>
            <div style="margin-top:5px;">請注意檢查並調整。</div>
        `;
        showAlert(warningHTML, 'warning', 8000);
    } else {
        showAlert('手動調整已保存', 'success');
    }
    
    return { hasWarning, warningMessages };
}

/**
 * 保存手動調整（無通知版本）
 * 與原始saveManualAdjustments功能相同，但不顯示任何通知
 */
export function saveManualAdjustmentsSilent() {
    let hasWarning = false;
    let warningMessages = [];
    
    // 遍歷所有制服類型並更新庫存資料
    for (const type in manualAdjustments) {
        if (!manualAdjustments.hasOwnProperty(type)) continue;
        
        let typeTotal = 0; // 追蹤每種制服類型的總可分配數
        
        // 遍歷所有尺寸，包括沒有手動調整的尺寸
        for (const size in inventoryData[type]) {
            if (!inventoryData[type].hasOwnProperty(size)) continue;
            
            const total = inventoryData[type][size].total || 0;
            let allocatable;
            
            // 檢查該尺寸是否有手動調整
            if (manualAdjustments[type] && manualAdjustments[type][size] !== undefined) {
                // 使用手動調整的值
                allocatable = Math.min(manualAdjustments[type][size], total);
            } else {
                // 使用原來計算的值（比例計算）
                const ratioElem = document.getElementById(`${type}Ratio`);
                let ratio = 0;
                if (ratioElem) {
                    const percentText = ratioElem.textContent || '0%';
                    ratio = parseFloat(percentText) / 100 || 0;
                }
                allocatable = Math.round(total * ratio);
            }
            
            // 更新庫存資料
            inventoryData[type][size].allocatable = allocatable;
            inventoryData[type][size].reserved = total - allocatable;
            
            // 累加可分配數
            typeTotal += allocatable;
        }
        
        // 更新對應表格的UI
        const tableId = inventoryTypeToTableId(type);
        if (tableId) {
            for (const size in inventoryData[type]) {
                updateInventoryUI(tableId, size);
            }
        }

        // 檢查當前制服類型的可分配數是否小於需求量
        const demand = demandData[type]?.totalDemand || 0;
        if (typeTotal < demand) {
            const uniformTypeName = UNIFORM_TYPES[type];
            const shortfall = demand - typeTotal;
            const message = `${uniformTypeName}：可分配數總和(${typeTotal})小於需求量(${demand})，差額${shortfall}件`;
            console.warn(message);
            warningMessages.push(message);
            hasWarning = true;
        }

        // 更新表格
        updateSizeTable(type);
    }
    
    // 保存到本地儲存
    saveToLocalStorage('inventoryData', inventoryData);
    saveToLocalStorage('manualAdjustments', manualAdjustments);
    
    // 更新分配比率
    updateAllocationRatios();
    
    // 不顯示任何通知，但仍返回警告信息以供調用方使用
    return { hasWarning, warningMessages };
}

/**
 * 重置手動調整
 */
export function resetManualAdjustments() {
    // 清空手動調整資料
    manualAdjustments = {
        shortSleeveShirt: {},
        shortSleevePants: {},
        longSleeveShirt: {},
        longSleevePants: {}
    };
    
    // 保存到本地儲存
    saveToLocalStorage('manualAdjustments', manualAdjustments);
    
    // 重新載入所有調整表格
    updateSizeTable('shortSleeveShirt');
    updateSizeTable('shortSleevePants');
    updateSizeTable('longSleeveShirt');
    updateSizeTable('longSleevePants');
    
    // 更新分配比率
    updateAllocationRatios();
    
    // 顯示成功訊息
    showAlert('手動調整已重置', 'info');
}

/**
 * 載入庫存和手動調整資料
 */
export function loadInventoryAndAdjustments() {
    // 載入庫存資料
    const savedInventory = loadFromLocalStorage('inventoryData', null);
    if (savedInventory) {
        inventoryData = savedInventory;
    }
    
    // 載入手動調整資料
    const savedAdjustments = loadFromLocalStorage('manualAdjustments', null);
    if (savedAdjustments) {
        manualAdjustments = savedAdjustments;
    }
}

/**
 * 處理庫存資料匯入
 */
function handleInventoryImport() {
    // 創建檔案輸入元素
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx,.xls,.csv';
    fileInput.style.display = 'none';
    document.body.appendChild(fileInput);
    
    // 監聽檔案選擇事件
    fileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) {
            return;
        }
        
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // 解析Excel數據
                importInventoryFromWorkbook(workbook);
            } catch (error) {
                showAlert('error', `匯入失敗: ${error.message}`);
                console.error('匯入錯誤:', error);
            }
        };
        
        reader.readAsArrayBuffer(file);
        
        // 移除檔案輸入元素
        document.body.removeChild(fileInput);
    });
    
    // 觸發檔案選擇對話框
    fileInput.click();
}

/**
 * 從工作簿導入庫存數據
 * @param {Object} workbook - XLSX 工作簿對象
 */
function importInventoryFromWorkbook(workbook) {
    // 確認工作簿有工作表
    if (workbook.SheetNames.length === 0) {
        showAlert('error', '匯入失敗: Excel 檔案不包含任何工作表');
        return;
    }
    
    // 初始化用於保存匯入數據的臨時變數
    const importedData = {
        shortSleeveShirt: {},
        shortSleevePants: {},
        longSleeveShirt: {},
        longSleevePants: {}
    };
    
    // 檢查工作表是否存在
    const sheetNames = ['短衣', '短褲', '長衣', '長褲'];
    const typeMapping = {
        '短衣': 'shortSleeveShirt',
        '短褲': 'shortSleevePants',
        '長衣': 'longSleeveShirt',
        '長褲': 'longSleevePants'
    };
    
    let importedSheets = 0;
    
    // 處理每個工作表
    for (const sheetName of sheetNames) {
        if (workbook.SheetNames.includes(sheetName)) {
            const worksheet = workbook.Sheets[sheetName];
            const sheetData = XLSX.utils.sheet_to_json(worksheet);
            
            // 處理工作表數據
            const type = typeMapping[sheetName];
            if (type && sheetData.length > 0) {
                sheetData.forEach(row => {
                    // 檢查數據是否符合預期格式
                    if (row['尺寸'] && row['總數'] !== undefined) {
                        const size = row['尺寸'];
                        const total = parseInt(row['總數']) || 0;
                        
                        // 檢查尺寸是否有效
                        if (SIZES.includes(size)) {
                            importedData[type][size] = { total };
                            importedSheets++;
                        }
                    }
                });
            }
        }
    }
    
    // 如果沒有匯入任何有效數據，顯示錯誤
    if (importedSheets === 0) {
        showAlert('error', '匯入失敗: 找不到有效的庫存數據');
        return;
    }
    
    // 更新庫存數據
    inventoryData = importedData;
    
    // 保存到本地儲存
    saveToLocalStorage('inventoryData', inventoryData);
    
    // 更新UI
    const tableIds = [
        'shortSleeveShirtTable',
        'shortSleevePantsTable', 
        'longSleeveShirtTable',
        'longSleevePantsTable'
    ];
    
    // 重新初始化表格，確保結構正確
    tableIds.forEach(tableId => {
        const type = tableIdToInventoryType(tableId);
        // 確保各尺碼的屬性結構完整
        SIZES.forEach(size => {
            if (inventoryData[type][size]) {
                // 確保有 total 屬性
                inventoryData[type][size].total = inventoryData[type][size].total || 0;
            }
        });
        
        initInventoryTable(tableId);
        loadInventoryData(tableId);
    });
    
    // 更新分配比率
    updateAllocationRatios();
    
    showAlert('success', '庫存數據匯入成功');
}

/**
 * 下載庫存範本
 */
function downloadInventoryTemplate() {
    // 創建工作簿
    const workbook = XLSX.utils.book_new();
    
    // 為每種制服類型創建一個工作表
    const types = [
        { name: '短衣', id: 'shortSleeveShirt' },
        { name: '短褲', id: 'shortSleevePants' },
        { name: '長衣', id: 'longSleeveShirt' },
        { name: '長褲', id: 'longSleevePants' }
    ];
    
    // 為每種類型創建範本數據
    types.forEach(type => {
        const data = [];
        
        // 添加標題行
        SIZES.forEach(size => {
            data.push({
                '尺寸': size,
                '總數': 0
            });
        });
        
        // 創建工作表
        const worksheet = XLSX.utils.json_to_sheet(data);
        
        // 添加到工作簿
        XLSX.utils.book_append_sheet(workbook, worksheet, type.name);
    });
    
    // 下載工作簿
    XLSX.writeFile(workbook, '庫存範本.xlsx');
}

/**
 * 確認清除庫存數據
 */
function confirmClearInventoryData() {
    if (confirm('確定要清除所有庫存數據嗎？此操作無法復原。')) {
        clearInventoryData();
    }
}

/**
 * 清空庫存資料
 */
export function clearInventoryData() {
    // 重置庫存資料
    inventoryData = {
        shortSleeveShirt: {},
        shortSleevePants: {},
        longSleeveShirt: {},
        longSleevePants: {}
    };
    
    // 初始化每種類型的每個尺寸
    const types = ['shortSleeveShirt', 'shortSleevePants', 'longSleeveShirt', 'longSleevePants'];
    types.forEach(type => {
        SIZES.forEach(size => {
            if (!inventoryData[type]) {
                inventoryData[type] = {};
            }
            inventoryData[type][size] = { total: 0 };
        });
    });

    // 重置手動調整
    manualAdjustments = {
        shortSleeveShirt: {},
        shortSleevePants: {},
        longSleeveShirt: {},
        longSleevePants: {}
    };
    
    // 保存到本地存儲
    saveToLocalStorage('inventoryData', inventoryData);
    saveToLocalStorage('manualAdjustments', manualAdjustments);
    
    // 更新各表格UI
    const tableIds = [
        'shortSleeveShirtTable',
        'shortSleevePantsTable',
        'longSleeveShirtTable',
        'longSleevePantsTable'
    ];
    
    tableIds.forEach(tableId => {
        initInventoryTable(tableId);
    });
    
    // 更新分配比率
    updateAllocationRatios();
    
    // 顯示成功提示
    showAlert('已清空所有庫存資料', 'success');
}

/**
 * 初始化單個庫存表格
 * @param {string} tableId - 表格ID
 */
function initInventoryTable(tableId) {
    const table = document.getElementById(tableId);
    if (!table) return;

    // 清空表格
    const tbody = table.querySelector('tbody');
    tbody.innerHTML = '';

    // 為每個尺寸添加行
    SIZES.forEach(size => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td data-original-size="${size}">${formatSize(size)}</td>
            <td><input type="number" class="form-control total-inventory" min="0" value="0"></td>
        `;
        tbody.appendChild(row);

        // 為輸入欄位添加事件監聽器
        const totalInput = row.querySelector('.total-inventory');

        // 更新庫存數據
        totalInput.addEventListener('change', () => {
            updateInventoryData(tableId, size, 'total', parseInt(totalInput.value) || 0);
        });
    });
} 