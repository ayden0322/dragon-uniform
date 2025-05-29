const fs = require('fs');
const path = require('path');
const { ensureSchoolDir } = require('./utils');

/**
 * @typedef {Object} Student
 * @property {string} id - 學生ID
 * @property {number} chest - 胸圍尺寸
 * @property {number} waist - 腰圍尺寸
 * @property {number} pantsLength - 褲長
 * @property {string} shirtSize - 上衣尺寸
 * @property {string} pantsSize - 短褲尺寸
 * @property {string} longPantsSize - 長褲尺寸
 * @property {number} [requiredQuantity] - 需求件數，可選
 */

/**
 * @typedef {Object} InventoryItem
 * @property {string} id - 庫存項目ID
 * @property {string} size - 尺寸編號
 * @property {number} quantity - 剩餘數量
 */

/**
 * @typedef {Object} Inventory
 * @property {InventoryItem[]} shirts - 上衣庫存
 * @property {InventoryItem[]} pants - 短褲庫存
 * @property {InventoryItem[]} longPants - 長褲庫存
 */

/**
 * @typedef {Object} AllocationResult
 * @property {string} studentId - 學生ID
 * @property {string} [shirtSize] - 分配的上衣尺寸
 * @property {string} [pantsSize] - 分配的短褲尺寸
 * @property {string} [longPantsSize] - 分配的長褲尺寸
 * @property {boolean} allocated - 是否成功分配
 * @property {string} [shirtId] - 分配的上衣ID
 * @property {string} [pantsId] - 分配的短褲ID
 * @property {string} [longPantsId] - 分配的長褲ID
 */

/**
 * 驗證相關功能
 */
const validationService = {
  /**
   * 驗證學生數據
   * @param {Student[]} students - 學生數據
   * @returns {boolean|string} - 成功返回true，失敗返回錯誤訊息
   */
  validateStudents: (students) => {
    if (!Array.isArray(students) || students.length === 0) {
      return '學生數據無效或為空';
    }
    
    // 檢查必要屬性
    for (const student of students) {
      if (!student.id || 
          typeof student.chest !== 'number' || 
          typeof student.waist !== 'number' ||
          typeof student.pantsLength !== 'number') {
        return `學生ID: ${student.id || '未知'} 的數據不完整`;
      }
    }
    
    return true;
  },
  
  /**
   * 驗證分配結果數據
   * @param {object} results - 分配結果數據
   * @returns {boolean|string} - 成功返回true，失敗返回錯誤訊息
   */
  validateAllocationResults: (results) => {
    if (!results || typeof results !== 'object') {
      return '分配結果數據無效';
    }
    
    // 檢查是否有必要的結果類型（前端可能傳入不同格式的結果）
    const hasResults = Object.keys(results).length > 0;
    
    if (!hasResults) {
      return '分配結果不能為空';
    }
    
    return true;
  },
  
  /**
   * 驗證分配請求
   * @param {object} data - 請求數據
   * @returns {object} - { isValid: boolean, error: string }
   */
  validateAllocationRequest: (data) => {
    const { studentData, inventoryData, results } = data;
    
    if (!studentData || !inventoryData || !results) {
      return { isValid: false, error: '缺少必要數據' };
    }
    
    const studentsValidation = validationService.validateStudents(
      Array.isArray(studentData.students) ? studentData.students : studentData
    );
    
    if (studentsValidation !== true) {
      return { isValid: false, error: studentsValidation };
    }
    
    const resultsValidation = validationService.validateAllocationResults(results);
    if (resultsValidation !== true) {
      return { isValid: false, error: resultsValidation };
    }
    
    return { isValid: true, error: null };
  }
};

/**
 * 數據存儲和檔案操作相關功能
 */
const storageService = {
  /**
   * 創建並保存分配歷史記錄
   * @param {string} schoolId - 學校ID
   * @param {string} historyDir - 歷史記錄目錄路徑
   * @param {object} data - 分配數據
   * @returns {string} - 時間戳記錄ID
   * @throws {Error} - 如果寫入文件失敗
   */
  createHistoryRecord: (schoolId, historyDir, data) => {
    const { studentData, inventoryData, adjustmentData, results } = data;
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        
    try {
        // 確保歷史記錄目錄存在
        if (!fs.existsSync(historyDir)) {
            fs.mkdirSync(historyDir, { recursive: true });
        }
        
        // 創建分配歷史記錄
        const historyFile = path.join(historyDir, `allocation-${timestamp}.json`);
      const studentCount = storageService._getStudentCount(studentData);
      
        const history = {
            schoolId,
            allocationDate: new Date().toISOString(),
        studentCount,
            inventorySnapshot: inventoryData,
            studentSnapshot: studentData,
            adjustmentSnapshot: adjustmentData || {},
            results,
        };
        
        fs.writeFileSync(historyFile, JSON.stringify(history, null, 2));
      return timestamp;
    } catch (error) {
      console.error('創建歷史記錄失敗:', error);
      throw new Error(`創建歷史記錄失敗: ${error.message}`);
    }
  },
  
  /**
   * 獲取學生數量
   * @param {object|Array} studentData - 學生數據
   * @returns {number} - 學生數量
   * @private
   */
  _getStudentCount: (studentData) => {
    if (Array.isArray(studentData)) {
      return studentData.length;
    } else if (studentData && Array.isArray(studentData.students)) {
      return studentData.students.length;
    }
    return 0;
  },
  
  /**
   * 更新學校配置文件
   * @param {string} schoolDir - 學校目錄路徑
   * @param {string} schoolId - 學校ID
   * @param {object} data - 配置數據
   * @throws {Error} - 如果寫入文件失敗
   */
  updateSchoolConfigurations: (schoolDir, schoolId, data) => {
    const { studentData, inventoryData, adjustmentData } = data;
    
    try {
        const configurations = [
            { type: 'students', data: studentData },
            { type: 'inventory', data: inventoryData },
            { type: 'adjustments', data: adjustmentData || {} }
        ];
        
        configurations.forEach(config => {
            const configFile = path.join(schoolDir, `${config.type}.json`);
            const configData = {
                schoolId,
                configType: config.type,
                data: config.data,
                updatedAt: new Date().toISOString()
            };
            fs.writeFileSync(configFile, JSON.stringify(configData, null, 2));
        });
    } catch (error) {
      console.error('更新配置文件失敗:', error);
      throw new Error(`更新配置文件失敗: ${error.message}`);
    }
  },
  
  /**
   * 獲取分配歷史記錄列表
   * @param {string} historyDir - 歷史記錄目錄路徑
   * @param {number} limit - 返回記錄的最大數量
   * @returns {Array} - 歷史記錄列表
   */
  getHistoryRecords: (historyDir, limit) => {
    try {
        if (!fs.existsSync(historyDir)) {
        return [];
        }
        
        const files = fs.readdirSync(historyDir)
            .filter(file => file.startsWith('allocation-') && file.endsWith('.json'))
            .sort((a, b) => b.localeCompare(a)) // 按文件名降序排序（較新的文件在前）
            .slice(0, parseInt(limit));
        
      return files.map(file => {
        try {
            const filePath = path.join(historyDir, file);
            const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
            
            // 返回簡要信息
            return {
                id: file.replace('allocation-', '').replace('.json', ''),
                allocationDate: data.allocationDate,
                studentCount: data.studentCount
            };
        } catch (err) {
          console.warn(`讀取歷史記錄 ${file} 失敗: ${err.message}`);
          return null;
        }
      }).filter(record => record !== null);
    } catch (error) {
      console.error('獲取歷史記錄列表失敗:', error);
      return [];
    }
  },
  
  /**
   * 獲取特定分配歷史的詳細信息
   * @param {string} historyFile - 歷史記錄文件路徑
   * @returns {object} - 分配歷史詳細數據
   * @throws {Error} - 如果讀取文件失敗
   */
  getHistoryDetail: (historyFile) => {
    try {
      return JSON.parse(fs.readFileSync(historyFile, 'utf8'));
    } catch (error) {
      console.error('讀取歷史記錄詳情失敗:', error);
      throw new Error(`讀取歷史記錄詳情失敗: ${error.message}`);
    }
  }
};

/**
 * 錯誤處理服務
 */
const errorService = {
  /**
   * 處理API錯誤
   * @param {Error} error - 錯誤對象
   * @param {string} operation - 操作名稱
   * @returns {object} - 適合API返回的錯誤對象
   */
  handleApiError: (error, operation) => {
    console.error(`${operation}失敗:`, error);
    return {
      error: `${operation}失敗`,
      message: error.message,
      code: error.code || 'UNKNOWN_ERROR'
    };
  }
};

// API 端點處理函數
/**
 * 保存分配結果（前端計算完成後調用）
 * @param {object} req - HTTP請求對象
 * @param {object} res - HTTP回應對象
 */
exports.saveAllocation = (req, res) => {
  try {
    const { schoolId } = req.params;
    const requestData = {
      studentData: req.body.studentData,
      inventoryData: req.body.inventoryData,
      adjustmentData: req.body.adjustmentData,
      results: req.body.results
    };
    
    // 驗證數據
    const validation = validationService.validateAllocationRequest(requestData);
    if (!validation.isValid) {
      return res.status(400).json({ error: validation.error });
    }
    
    const schoolDir = ensureSchoolDir(schoolId);
    const historyDir = path.join(schoolDir, 'history');
    
    // 保存歷史記錄
    const timestamp = storageService.createHistoryRecord(
      schoolId, 
      historyDir, 
      requestData
    );
    
    // 更新最新配置
    storageService.updateSchoolConfigurations(
      schoolDir, 
      schoolId, 
      requestData
    );
    
    res.json({ 
      success: true, 
      allocationId: timestamp,
      message: '分配結果已成功保存'
    });
  } catch (err) {
    const errorResponse = errorService.handleApiError(err, '保存分配結果');
    res.status(500).json(errorResponse);
  }
};

/**
 * 獲取分配歷史
 * @param {object} req - HTTP請求對象
 * @param {object} res - HTTP回應對象
 */
exports.getAllocationHistory = (req, res) => {
  try {
    const { schoolId } = req.params;
    const { limit = 10 } = req.query;
    
    const schoolDir = ensureSchoolDir(schoolId);
    const historyDir = path.join(schoolDir, 'history');
    
    const history = storageService.getHistoryRecords(historyDir, limit);
        res.json(history);
    } catch (err) {
    const errorResponse = errorService.handleApiError(err, '獲取分配歷史');
    res.status(500).json(errorResponse);
  }
};

/**
 * 獲取特定分配的詳細信息
 * @param {object} req - HTTP請求對象
 * @param {object} res - HTTP回應對象
 */
exports.getAllocationDetail = (req, res) => {
    try {
        const { schoolId, id } = req.params;
        
        const schoolDir = ensureSchoolDir(schoolId);
        const historyFile = path.join(schoolDir, 'history', `allocation-${id}.json`);
        
        if (!fs.existsSync(historyFile)) {
            return res.status(404).json({ error: '找不到分配記錄' });
        }
        
    const allocation = storageService.getHistoryDetail(historyFile);
        res.json(allocation);
    } catch (err) {
    const errorResponse = errorService.handleApiError(err, '獲取分配詳細信息');
    res.status(500).json(errorResponse);
  }
};

// 導出服務，供其他模組使用
exports.storageService = storageService;
exports.validationService = validationService;
exports.errorService = errorService;