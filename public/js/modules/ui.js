// UI相關功能模組
import { saveToLocalStorage, loadFromLocalStorage, showAlert } from './utils.js';
import { inventoryData, updateSizeTable, saveManualAdjustments, resetManualAdjustments } from './inventory.js';
import { handleExcelImport, exportStudentData, downloadTemplate, clearStudentData, updateAdjustmentPage, demandData } from './students.js';
import { startAllocation, updateAllocationResults } from './allocation.js';
import { UNIFORM_TYPES } from './config.js';

/**
 * 設置匯入匯出按鈕功能
 */
export function setupImportExportButtons() {
    // 匯入按鈕
    const importExcelBtn = document.getElementById('importExcelBtn');
    if (importExcelBtn) {
        importExcelBtn.addEventListener('click', function() {
            // 創建隱藏的檔案輸入元素
            const fileInput = document.createElement('input');
            fileInput.type = 'file';
            fileInput.accept = '.xlsx,.xls,.csv';
            fileInput.style.display = 'none';
            document.body.appendChild(fileInput);
            
            // 監聽檔案選擇事件
            fileInput.addEventListener('change', async function(e) {
                if (e.target.files.length > 0) {
                    const file = e.target.files[0];
                    try {
                        await handleExcelImport(file);
                        // 更新調整頁面資料
                        updateAdjustmentPage();
                    } catch (error) {
                        console.error('匯入處理失敗:', error);
                        showAlert('匯入處理失敗: ' + error.message, 'danger');
                    }
                }
                // 移除檔案輸入元素
                document.body.removeChild(fileInput);
            });
            
            // 觸發檔案選擇對話框
            fileInput.click();
        });
    }
    
    // 下載範本按鈕
    const downloadTemplateBtn = document.getElementById('downloadTemplateBtn');
    if (downloadTemplateBtn) {
        // 更改按鈕顏色為藍色系，與背景有更明顯的區別
        downloadTemplateBtn.classList.remove('btn-outline-secondary');
        downloadTemplateBtn.classList.add('btn-info');
        
        downloadTemplateBtn.addEventListener('click', function() {
            downloadTemplate();
        });
    }
    
    // 清除資料按鈕
    const clearStudentDataBtn = document.getElementById('clearStudentDataBtn');
    if (clearStudentDataBtn) {
        clearStudentDataBtn.addEventListener('click', function() {
            clearStudentData();
            // 更新調整頁面資料
            updateAdjustmentPage();
        });
    }
    
    // 如果表格中沒有資料，顯示提示訊息
    const studentTable = document.getElementById('studentTable');
    if (studentTable) {
        const tbody = studentTable.querySelector('tbody');
        if (tbody && !tbody.firstChild) {
            const row = document.createElement('tr');
            row.innerHTML = '<td colspan="12" class="text-center">無資料，請匯入或手動新增</td>';
            tbody.appendChild(row);
        }
    }
}

/**
 * 設置頁籤切換事件
 */
export function setupTabEvents() {
    // 學生資料頁籤
    const studentTabButton = document.getElementById('student-tab');
    if (studentTabButton) {
        studentTabButton.addEventListener('click', function() {
            // 更新學生資料表格
            // 這裡不需要額外操作，因為學生資料表格在初始化時已經設置好
        });
    }
    
    // 庫存頁籤
    const inventoryTabButton = document.getElementById('inventory-tab');
    if (inventoryTabButton) {
        inventoryTabButton.addEventListener('click', function() {
            // 更新庫存表格
            // 這裡不需要額外操作，因為庫存表格在初始化時已經設置好
        });
    }
    
    // 調整頁籤
    const adjustmentTabButton = document.getElementById('adjustment-tab');
    if (adjustmentTabButton) {
        adjustmentTabButton.addEventListener('click', function() {
            try {
                // 更新調整頁面資料
                updateAdjustmentPage();
                
                // 更新各尺寸表格
                updateSizeTable('shortSleeveShirt');
                updateSizeTable('shortSleevePants');
                updateSizeTable('longSleeveShirt');
                updateSizeTable('longSleevePants');
            } catch (error) {
                console.error('更新調整頁面失敗:', error);
            }
        });
    }
    
    // 分配頁籤
    const allocationTabButton = document.getElementById('allocation-tab');
    if (allocationTabButton) {
        allocationTabButton.addEventListener('click', function() {
            // 目前不需要任何操作
        });
    }
    
    // 結果頁籤
    const resultTabButton = document.getElementById('result-tab');
    if (resultTabButton) {
        resultTabButton.addEventListener('click', function() {
            // 不再自動更新分配結果頁面，僅在點擊開始分配按鈕後才更新
            console.log('結果頁籤點擊: 不自動更新分配結果頁面');
        });
    }
}

/**
 * 設置分配按鈕功能
 */
export function setupAllocationButton() {
    const allocateButton = document.getElementById('allocateButton');
    const resultsTab = document.getElementById('result-tab');
    
    if (allocateButton) {
        allocateButton.addEventListener('click', async () => {
            try {
                // 禁用按鈕
                allocateButton.disabled = true;
                allocateButton.textContent = '分配中...';
                
                // 開始分配 (這個函數會先重置所有之前的分配結果)
                await startAllocation();
                
                // 更新分配結果頁面
                await updateAllocationResults();
                
                // 切換到結果頁籤
                if (resultsTab) {
                    resultsTab.click();
                }
                
                // 顯示成功訊息
                showAlert('制服分配完成', 'success');
            } catch (error) {
                console.error('分配過程發生錯誤:', error);
                showAlert('分配過程發生錯誤: ' + error.message, 'danger');
            } finally {
                // 恢復按鈕狀態
                allocateButton.disabled = false;
                allocateButton.textContent = '開始分配';
            }
        });
    }
}

/**
 * 設置調整頁面按鈕
 */
export function setupAdjustmentPageButtons() {
    // 保存調整按鈕
    const saveAdjustmentBtn = document.getElementById('saveAdjustmentBtn');
    if (saveAdjustmentBtn) {
        saveAdjustmentBtn.addEventListener('click', function() {
            saveManualAdjustments();
        });
    }
    
    // 重置調整按鈕
    const resetAdjustmentBtn = document.getElementById('resetAdjustmentBtn');
    if (resetAdjustmentBtn) {
        resetAdjustmentBtn.addEventListener('click', function() {
            resetManualAdjustments();
        });
    }
    
    // 尺寸標籤切換
    setupSizeTabsEvents();
}

/**
 * 設置尺寸標籤切換事件
 */
function setupSizeTabsEvents() {
    const sizeTabsContainer = document.getElementById('sizeTabs');
    if (!sizeTabsContainer) return;
    
    // 所有尺寸標籤
    const sizeTabs = sizeTabsContainer.querySelectorAll('.nav-link');
    
    // 為每個標籤添加點擊事件
    sizeTabs.forEach(tab => {
        tab.addEventListener('click', function() {
            // 取得目標表格的類型
            const targetType = this.getAttribute('data-target-type');
            
            // 更新標籤激活狀態
            sizeTabs.forEach(t => t.classList.remove('active'));
            this.classList.add('active');
            
            // 更新尺寸表格內容
            if (targetType) {
                updateSizeTable(targetType);
            }
        });
    });
}

/**
 * 更新分配比率
 */
export function updateAllocationRatios() {
    try {
        // 獲取所有制服類型的需求和庫存總量
        const ratioElements = {
            shortSleeveShirt: document.getElementById('shortSleeveShirtRatio'),
            shortSleevePants: document.getElementById('shortSleevePantsRatio'),
            longSleeveShirt: document.getElementById('longSleeveShirtRatio'),
            longSleevePants: document.getElementById('longSleevePantsRatio')
        };
        
        // 記錄比例是否變更
        let ratioChanged = false;
        const oldRatios = {};
        
        for (const type in ratioElements) {
            if (!ratioElements[type]) continue;
            
            // 儲存舊的比例值用於比較
            if (ratioElements[type]) {
                const oldPercentText = ratioElements[type].textContent || '0%';
                oldRatios[type] = parseFloat(oldPercentText) / 100 || 0;
            }
            
            // 直接從 demandData 獲取需求數量，而不是從 DOM 元素獲取
            let demand = 0;
            if (demandData && demandData[type] && demandData[type].totalDemand !== undefined) {
                demand = demandData[type].totalDemand;
            }
            
            // 同時更新 DOM 元素中的需求數量顯示，確保一致性
            const demandElem = document.getElementById(`${type}Demand`);
            if (demandElem) {
                demandElem.textContent = demand;
            }
            
            // 計算庫存總量
            let totalInventory = 0;
            if (inventoryData[type]) {
                for (const size in inventoryData[type]) {
                    if (inventoryData[type][size] && inventoryData[type][size].total) {
                        totalInventory += inventoryData[type][size].total;
                    }
                }
            }
            
            // 計算比率
            let ratio = 0;
            if (totalInventory > 0) {
                ratio = demand / totalInventory;
            }
            
            // 檢查比例是否變更
            if (Math.abs(ratio - oldRatios[type]) > 0.0001) {
                ratioChanged = true;
            }
            
            // 格式化顯示
            const formattedRatio = (ratio * 100).toFixed(1) + '%';
            
            // 更新UI
            ratioElements[type].textContent = formattedRatio;
            
            // 設置顏色
            if (ratio > 1) {
                ratioElements[type].classList.add('text-danger');
                ratioElements[type].classList.remove('text-success');
            } else {
                ratioElements[type].classList.remove('text-danger');
                ratioElements[type].classList.add('text-success');
            }
        }
        
        // 在更新了所有比例後，始終更新所有尺寸表格
        updateSizeTable('shortSleeveShirt');
        updateSizeTable('shortSleevePants');
        updateSizeTable('longSleeveShirt');
        updateSizeTable('longSleevePants');
    } catch (error) {
        console.error('更新分配比率時發生錯誤:', error);
    }
}

/**
 * 格式化尺寸顯示，包含褲長調整標記
 * @param {Object} student - 學生資料
 * @param {string} size - 尺寸字串
 * @param {string} fieldType - 欄位類型，例如 'allocatedShirtSize'
 * @returns {string} - 格式化後的尺寸字串
 */
export function formatSizeWithAdjustment(student, size, fieldType) {
    if (!size) return '';
    
    // 根據欄位類型檢查是否有褲長調整
    let isAdjusted = false;
    if (fieldType === 'allocatedShirtSize' && student.isShirtSizeAdjustedForPantsLength) {
        isAdjusted = true;
    }
    
    // 如果有調整，添加上箭頭標記
    const formattedSize = formatSize(size);
    return isAdjusted ? `${formattedSize}↑` : formattedSize;
} 