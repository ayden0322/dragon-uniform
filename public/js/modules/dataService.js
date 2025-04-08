// 數據服務模塊 - 兼容本地存儲和API
import * as apiClient from './api/index.js';
import { showAlert, saveToLocalStorage, loadFromLocalStorage } from './utils.js';

// 使用API模式或本地存儲模式 (true=API, false=本地存儲)
const useApiMode = true;

// 當前選中的學校ID
let currentSchoolId = 'dragon';

/**
 * 設置當前學校ID
 * @param {string} schoolId - 學校ID
 */
export function setCurrentSchool(schoolId) {
    currentSchoolId = schoolId;
}

/**
 * 從服務器或本地存儲加載配置數據
 * @param {string} configType - 配置類型 (students, inventory, adjustments)
 * @param {string} schoolId - 學校ID（可選）
 * @returns {Promise<Object>} 配置數據
 */
export async function loadConfigData(configType, schoolId = currentSchoolId) {
    try {
        if (useApiMode) {
            try {
                // 嘗試從API加載
                const data = await apiClient.loadConfig(schoolId, configType);
                return data;
            } catch (apiError) {
                console.warn(`API加載失敗，嘗試從本地存儲加載 ${configType}:`, apiError);
                // API加載失敗，回退到本地存儲
                return loadFromLocalStorage(`${schoolId}_${configType}`, null);
            }
        } else {
            // 直接從本地存儲加載
            return loadFromLocalStorage(`${schoolId}_${configType}`, null);
        }
    } catch (error) {
        console.error(`加載 ${configType} 數據失敗:`, error);
        showAlert(`無法加載 ${configType} 數據: ${error.message}`, 'warning');
        return null;
    }
}

/**
 * 保存配置數據到服務器或本地存儲
 * @param {string} configType - 配置類型 (students, inventory, adjustments)
 * @param {Object} data - 配置數據
 * @param {string} schoolId - 學校ID（可選）
 * @returns {Promise<boolean>} 是否成功
 */
export async function saveConfigData(configType, data, schoolId = currentSchoolId) {
    try {
        // 始終保存到本地存儲作為備份
        saveToLocalStorage(`${schoolId}_${configType}`, data);
        
        if (useApiMode) {
            try {
                // 嘗試保存到API
                await apiClient.saveConfig(schoolId, configType, data);
                return true;
            } catch (apiError) {
                console.warn(`API保存失敗 ${configType}:`, apiError);
                showAlert(`數據已保存到本地，但無法保存到服務器: ${apiError.message}`, 'warning');
                return false;
            }
        } else {
            return true; // 本地存儲模式已保存
        }
    } catch (error) {
        console.error(`保存 ${configType} 數據失敗:`, error);
        showAlert(`保存 ${configType} 數據失敗: ${error.message}`, 'danger');
        return false;
    }
}

/**
 * 保存分配結果
 * @param {Object} studentData - 學生數據
 * @param {Object} inventoryData - 庫存數據
 * @param {Object} adjustmentData - 調整數據
 * @param {Object} results - 分配結果
 * @param {string} schoolId - 學校ID（可選）
 * @returns {Promise<boolean>} 是否成功
 */
export async function saveAllocationData(studentData, inventoryData, adjustmentData, results, schoolId = currentSchoolId) {
    try {
        // 創建一個本地歷史記錄
        const timestamp = new Date().toISOString();
        const historyData = {
            allocationDate: timestamp,
            studentData,
            inventoryData,
            adjustmentData,
            results
        };
        
        // 保存到本地作為備份
        const localHistory = loadFromLocalStorage(`${schoolId}_allocationHistory`, []);
        localHistory.unshift(historyData);
        // 只保留最近10條記錄
        if (localHistory.length > 10) {
            localHistory.length = 10;
        }
        saveToLocalStorage(`${schoolId}_allocationHistory`, localHistory);
        
        // 保存最新的配置數據
        saveToLocalStorage(`${schoolId}_students`, studentData);
        saveToLocalStorage(`${schoolId}_inventory`, inventoryData);
        saveToLocalStorage(`${schoolId}_adjustments`, adjustmentData);
        
        if (useApiMode) {
            try {
                // 嘗試保存到API
                await apiClient.saveAllocation(schoolId, studentData, inventoryData, adjustmentData, results);
                return true;
            } catch (apiError) {
                console.warn('API保存分配結果失敗:', apiError);
                showAlert(`分配結果已保存到本地，但無法保存到服務器: ${apiError.message}`, 'warning');
                return false;
            }
        } else {
            return true; // 本地存儲模式已保存
        }
    } catch (error) {
        console.error('保存分配結果失敗:', error);
        showAlert(`保存分配結果失敗: ${error.message}`, 'danger');
        return false;
    }
}

/**
 * 獲取分配歷史
 * @param {string} schoolId - 學校ID（可選）
 * @returns {Promise<Array>} 歷史記錄數組
 */
export async function getAllocationHistory(schoolId = currentSchoolId) {
    try {
        if (useApiMode) {
            try {
                // 嘗試從API獲取
                return await apiClient.getAllocationHistory(schoolId);
            } catch (apiError) {
                console.warn('API獲取分配歷史失敗，回退到本地存儲:', apiError);
                // 回退到本地存儲
                const localHistory = loadFromLocalStorage(`${schoolId}_allocationHistory`, []);
                return localHistory.map((item, index) => ({
                    id: `local_${index}`,
                    allocationDate: item.allocationDate,
                    studentCount: item.studentData.students ? item.studentData.students.length : 0
                }));
            }
        } else {
            // 直接從本地存儲獲取
            const localHistory = loadFromLocalStorage(`${schoolId}_allocationHistory`, []);
            return localHistory.map((item, index) => ({
                id: `local_${index}`,
                allocationDate: item.allocationDate,
                studentCount: item.studentData.students ? item.studentData.students.length : 0
            }));
        }
    } catch (error) {
        console.error('獲取分配歷史失敗:', error);
        showAlert(`獲取分配歷史失敗: ${error.message}`, 'danger');
        return [];
    }
}

/**
 * 獲取特定分配的詳細信息
 * @param {string} id - 分配ID
 * @param {string} schoolId - 學校ID（可選）
 * @returns {Promise<Object>} 分配詳情
 */
export async function getAllocationDetail(id, schoolId = currentSchoolId) {
    try {
        if (useApiMode && !id.startsWith('local_')) {
            try {
                // 嘗試從API獲取
                return await apiClient.getAllocationDetail(schoolId, id);
            } catch (apiError) {
                console.warn('API獲取分配詳情失敗，回退到本地存儲:', apiError);
                // 如果ID不是本地ID格式，則無法從本地獲取
                showAlert('無法從服務器獲取分配詳情', 'warning');
                return null;
            }
        } else {
            // 從本地存儲獲取
            const localIndex = id.replace('local_', '');
            const localHistory = loadFromLocalStorage(`${schoolId}_allocationHistory`, []);
            if (localHistory[localIndex]) {
                return {
                    schoolId,
                    allocationDate: localHistory[localIndex].allocationDate,
                    studentCount: localHistory[localIndex].studentData.students ? localHistory[localIndex].studentData.students.length : 0,
                    studentSnapshot: localHistory[localIndex].studentData,
                    inventorySnapshot: localHistory[localIndex].inventoryData,
                    adjustmentSnapshot: localHistory[localIndex].adjustmentData,
                    results: localHistory[localIndex].results
                };
            }
            return null;
        }
    } catch (error) {
        console.error('獲取分配詳情失敗:', error);
        showAlert(`獲取分配詳情失敗: ${error.message}`, 'danger');
        return null;
    }
}
