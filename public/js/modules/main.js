// 主程式入口模組
import { inventoryData, initInventoryFeatures, loadInventoryAndAdjustments } from './inventory.js';
import { studentData, loadStudentData, updateAdjustmentPage } from './students.js';
import { loadAllocationResults } from './allocation.js';
import { setupImportExportButtons, setupTabEvents, setupAllocationButton, setupAdjustmentPageButtons, updateAllocationRatios } from './ui.js';
import { showAlert } from './utils.js';
import { currentSizeDisplayMode, SIZE_DISPLAY_MODES, setSizeDisplayMode, initSizeDisplayMode } from './config.js';

/**
 * 設置尺寸顯示模式UI
 */
function setupSizeDisplayModeUI() {
    // 先從localStorage載入保存的顯示模式
    initSizeDisplayMode();
    
    // 設置選擇器的初始值
    const sizeDisplaySelect = document.getElementById('sizeDisplayMode');
    if (sizeDisplaySelect) {
        sizeDisplaySelect.value = currentSizeDisplayMode;
        
        // 添加事件監聽器 - 僅影響庫存相關頁面
        sizeDisplaySelect.addEventListener('change', (event) => {
            const newMode = event.target.value;
            if (setSizeDisplayMode(newMode)) {
                // 僅重新渲染庫存相關表格
                refreshInventoryTables();
                showAlert('庫存頁面尺寸顯示模式已變更', 'success');
            }
        });
    }
}

/**
 * 重新渲染庫存相關的表格
 */
function refreshInventoryTables() {
    // 重新初始化庫存表格以更新尺寸顯示
    initInventoryFeatures();
    
    // 更新調整頁面
    updateAdjustmentPage();
}



/**
 * 應用程式初始化
 */
export function initApp() {
    try {
        console.log('應用程式初始化中...');
        
        // 設置尺寸顯示模式UI
        setupSizeDisplayModeUI();
        
        // 載入庫存資料和手動調整資料
        loadInventoryAndAdjustments();
        
        // 載入學生資料
        loadStudentData();
        
        // 不再載入分配結果，等用戶點擊「開始分配」按鈕時才載入
        // loadAllocationResults();
        
        // 初始化庫存表格和功能
        initInventoryFeatures();
        
        // 設置頁籤切換事件
        setupTabEvents();
        
        // 設置匯入匯出按鈕
        setupImportExportButtons();
        
        // 設置分配按鈕
        setupAllocationButton();
        
        // 設置調整頁面按鈕
        setupAdjustmentPageButtons();
        
        // 更新調整頁面
        try {
            updateAdjustmentPage();
            // 更新分配比率
            updateAllocationRatios();
        } catch (error) {
            console.error('更新調整頁面失敗:', error);
        }
        
        console.log('應用程式初始化完成');
    } catch (error) {
        console.error('應用程式初始化失敗:', error);
        showAlert('應用程式初始化失敗: ' + error.message, 'danger');
    }
}

// 當DOM完全載入後初始化應用程式
document.addEventListener('DOMContentLoaded', initApp); 