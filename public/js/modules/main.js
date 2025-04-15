// 主程式入口模組
import { inventoryData, initInventoryFeatures, loadInventoryAndAdjustments } from './inventory.js';
import { studentData, loadStudentData, updateAdjustmentPage } from './students.js';
import { loadAllocationResults } from './allocation.js';
import { setupImportExportButtons, setupTabEvents, setupAllocationButton, setupAdjustmentPageButtons, updateAllocationRatios } from './ui.js';
import { showAlert } from './utils.js';
import { currentSizeDisplayMode, SIZE_DISPLAY_MODES, setSizeDisplayMode } from './config.js';

/**
 * 初始化尺寸顯示模式
 */
function initSizeDisplayMode() {
    // 從localStorage獲取保存的顯示模式
    const savedMode = localStorage.getItem('sizeDisplayMode');
    if (savedMode && Object.values(SIZE_DISPLAY_MODES).includes(savedMode)) {
        setSizeDisplayMode(savedMode);
    }
    
    // 設置選擇器的初始值
    const sizeDisplaySelect = document.getElementById('sizeDisplayMode');
    if (sizeDisplaySelect) {
        sizeDisplaySelect.value = currentSizeDisplayMode;
        
        // 添加事件監聽器
        sizeDisplaySelect.addEventListener('change', (event) => {
            const newMode = event.target.value;
            if (setSizeDisplayMode(newMode)) {
                // 重新渲染所有表格
                refreshAllTables();
                showAlert('尺寸顯示模式已變更', 'success');
            }
        });
    }
}

/**
 * 重新渲染所有包含尺寸的表格
 */
function refreshAllTables() {
    // 重新初始化庫存表格
    initInventoryFeatures();
    
    // 更新調整頁面
    updateAdjustmentPage();
    
    // 更新學生表格（如果需要）
    // reloadStudentTable();
    
    // 如果已經有分配結果，則重新載入
    loadAllocationResults();
}

/**
 * 應用程式初始化
 */
export function initApp() {
    try {
        console.log('應用程式初始化中...');
        
        // 初始化尺寸顯示模式
        initSizeDisplayMode();
        
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