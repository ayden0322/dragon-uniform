const fs = require('fs');
const path = require('path');
const { ensureSchoolDir } = require('./utils');

// 保存分配結果
exports.saveAllocation = (req, res) => {
    try {
        const { schoolId } = req.params;
        const { studentData, inventoryData, adjustmentData, results } = req.body;
        
        // 驗證數據
        if (!studentData || !inventoryData || !results) {
            return res.status(400).json({ error: '缺少必要數據' });
        }
        
        const schoolDir = ensureSchoolDir(schoolId);
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const historyDir = path.join(schoolDir, 'history');
        
        // 確保歷史記錄目錄存在
        if (!fs.existsSync(historyDir)) {
            fs.mkdirSync(historyDir, { recursive: true });
        }
        
        // 創建分配歷史記錄
        const historyFile = path.join(historyDir, `allocation-${timestamp}.json`);
        const history = {
            schoolId,
            allocationDate: new Date().toISOString(),
            studentCount: Array.isArray(studentData.students) ? studentData.students.length : (studentData.length || 0),
            inventorySnapshot: inventoryData,
            studentSnapshot: studentData,
            adjustmentSnapshot: adjustmentData || {},
            results,
        };
        
        fs.writeFileSync(historyFile, JSON.stringify(history, null, 2));
        
        // 更新最新配置
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
        
        res.json({ success: true, allocationId: timestamp });
    } catch (err) {
        console.error('保存分配結果失敗:', err);
        res.status(500).json({ error: '保存分配結果失敗' });
    }
};

// 獲取分配歷史
exports.getAllocationHistory = (req, res) => {
    try {
        const { schoolId } = req.params;
        const { limit = 10 } = req.query;
        
        const schoolDir = ensureSchoolDir(schoolId);
        const historyDir = path.join(schoolDir, 'history');
        
        // 如果沒有歷史記錄目錄或文件，返回空數組
        if (!fs.existsSync(historyDir)) {
            return res.json([]);
        }
        
        const files = fs.readdirSync(historyDir)
            .filter(file => file.startsWith('allocation-') && file.endsWith('.json'))
            .sort((a, b) => b.localeCompare(a)) // 按文件名降序排序（較新的文件在前）
            .slice(0, parseInt(limit));
        
        const history = files.map(file => {
            const filePath = path.join(historyDir, file);
            const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
            
            // 返回簡要信息
            return {
                id: file.replace('allocation-', '').replace('.json', ''),
                allocationDate: data.allocationDate,
                studentCount: data.studentCount
            };
        });
        
        res.json(history);
    } catch (err) {
        console.error('獲取分配歷史失敗:', err);
        res.status(500).json({ error: '獲取分配歷史失敗' });
    }
};

// 獲取特定分配的詳細信息
exports.getAllocationDetail = (req, res) => {
    try {
        const { schoolId, id } = req.params;
        
        const schoolDir = ensureSchoolDir(schoolId);
        const historyFile = path.join(schoolDir, 'history', `allocation-${id}.json`);
        
        if (!fs.existsSync(historyFile)) {
            return res.status(404).json({ error: '找不到分配記錄' });
        }
        
        const allocation = JSON.parse(fs.readFileSync(historyFile, 'utf8'));
        res.json(allocation);
    } catch (err) {
        console.error('獲取分配詳細信息失敗:', err);
        res.status(500).json({ error: '獲取分配詳細信息失敗' });
    }
};
