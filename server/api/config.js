const fs = require('fs');
const path = require('path');
const { ensureSchoolDir } = require('./utils');

// 獲取配置
exports.getConfig = (req, res) => {
    try {
        const { schoolId, configType } = req.params;
        
        // 驗證配置類型
        if (!['students', 'inventory', 'adjustments'].includes(configType)) {
            return res.status(400).json({ error: '無效的配置類型' });
        }
        
        const schoolDir = ensureSchoolDir(schoolId);
        const configFile = path.join(schoolDir, `${configType}.json`);
        
        if (!fs.existsSync(configFile)) {
            return res.json({ data: null });
        }
        
        const configData = JSON.parse(fs.readFileSync(configFile, 'utf8'));
        res.json(configData);
    } catch (err) {
        console.error('獲取配置失敗:', err);
        res.status(500).json({ error: '獲取配置失敗' });
    }
};

// 保存配置
exports.saveConfig = (req, res) => {
    try {
        const { schoolId, configType } = req.params;
        const { data } = req.body;
        
        // 驗證數據
        if (!data) {
            return res.status(400).json({ error: '缺少數據' });
        }
        
        // 驗證配置類型
        if (!['students', 'inventory', 'adjustments'].includes(configType)) {
            return res.status(400).json({ error: '無效的配置類型' });
        }
        
        const schoolDir = ensureSchoolDir(schoolId);
        const configFile = path.join(schoolDir, `${configType}.json`);
        
        // 添加版本號和更新時間
        const config = {
            schoolId,
            configType,
            data,
            updatedAt: new Date().toISOString()
        };
        
        fs.writeFileSync(configFile, JSON.stringify(config, null, 2));
        res.json({ success: true });
    } catch (err) {
        console.error('保存配置失敗:', err);
        res.status(500).json({ error: '保存配置失敗' });
    }
};
