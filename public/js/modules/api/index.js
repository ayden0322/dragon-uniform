// API 客戶端模塊

/**
 * 從服務器加載配置
 * @param {string} schoolId - 學校ID
 * @param {string} configType - 配置類型 (students, inventory, adjustments)
 * @returns {Promise<Object>} 配置數據
 */
export async function loadConfig(schoolId, configType) {
    try {
        const response = await fetch(`/api/config/${schoolId}/${configType}`);
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || '加載數據失敗');
        }
        
        const data = await response.json();
        return data.data;
    } catch (error) {
        console.error(`加載 ${configType} 配置失敗:`, error);
        throw error;
    }
}

/**
 * 保存配置到服務器
 * @param {string} schoolId - 學校ID
 * @param {string} configType - 配置類型 (students, inventory, adjustments)
 * @param {Object} data - 配置數據
 * @returns {Promise<boolean>} 是否成功
 */
export async function saveConfig(schoolId, configType, data) {
    try {
        const response = await fetch(`/api/config/${schoolId}/${configType}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ data }),
        });
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || '保存數據失敗');
        }
        
        const result = await response.json();
        return result.success;
    } catch (error) {
        console.error(`保存 ${configType} 配置失敗:`, error);
        throw error;
    }
}

/**
 * 保存分配結果到服務器
 * @param {string} schoolId - 學校ID
 * @param {Object} studentData - 學生數據
 * @param {Object} inventoryData - 庫存數據
 * @param {Object} adjustmentData - 調整數據
 * @param {Object} results - 分配結果
 * @returns {Promise<Object>} 結果對象
 */
export async function saveAllocation(schoolId, studentData, inventoryData, adjustmentData, results) {
    try {
        const response = await fetch(`/api/allocation/${schoolId}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                studentData,
                inventoryData,
                adjustmentData,
                results
            }),
        });
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || '保存分配結果失敗');
        }
        
        return await response.json();
    } catch (error) {
        console.error('保存分配結果失敗:', error);
        throw error;
    }
}

/**
 * 獲取分配歷史記錄
 * @param {string} schoolId - 學校ID
 * @param {number} limit - 限制數量
 * @returns {Promise<Array>} 歷史記錄數組
 */
export async function getAllocationHistory(schoolId, limit = 10) {
    try {
        const response = await fetch(`/api/allocation/${schoolId}/history?limit=${limit}`);
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || '獲取分配歷史失敗');
        }
        
        return await response.json();
    } catch (error) {
        console.error('獲取分配歷史失敗:', error);
        throw error;
    }
}

/**
 * 獲取特定分配的詳細信息
 * @param {string} schoolId - 學校ID
 * @param {string} id - 分配ID
 * @returns {Promise<Object>} 分配詳情
 */
export async function getAllocationDetail(schoolId, id) {
    try {
        const response = await fetch(`/api/allocation/${schoolId}/detail/${id}`);
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || '獲取分配詳情失敗');
        }
        
        return await response.json();
    } catch (error) {
        console.error('獲取分配詳情失敗:', error);
        throw error;
    }
}
