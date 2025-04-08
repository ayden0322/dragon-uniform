const { getConfig, saveConfig } = require('./config');
const { saveAllocation, getAllocationHistory, getAllocationDetail } = require('./allocation');
const fs = require('fs');
const path = require('path');
const { dataDir } = require('./utils');

function setupRoutes(app) {
    // 配置相關路由
    app.get('/api/config/:schoolId/:configType', getConfig);
    app.post('/api/config/:schoolId/:configType', saveConfig);
    
    // 分配相關路由
    app.post('/api/allocation/:schoolId', saveAllocation);
    app.get('/api/allocation/:schoolId/history', getAllocationHistory);
    app.get('/api/allocation/:schoolId/detail/:id', getAllocationDetail);
    
    // 環境檢查路由
    app.get('/api/check-env', (req, res) => {
        try {
            // 檢查環境變量
            const envInfo = {
                NODE_ENV: process.env.NODE_ENV || 'not set',
                DATA_DIR: process.env.DATA_DIR || 'not set',
                PORT: process.env.PORT || 'not set',
                CWD: process.cwd(),
                dataDir: dataDir
            };
            
            // 檢查數據目錄
            const dataDirExists = fs.existsSync(dataDir);
            const dirContents = dataDirExists ? fs.readdirSync(dataDir) : [];
            
            // 檢查權限
            let writePermission = false;
            if (dataDirExists) {
                try {
                    const testFile = path.join(dataDir, 'test-write-permission.txt');
                    fs.writeFileSync(testFile, 'test');
                    fs.unlinkSync(testFile);
                    writePermission = true;
                } catch (err) {
                    writePermission = false;
                }
            }
            
            res.json({
                env: envInfo,
                dataDirectory: {
                    path: dataDir,
                    exists: dataDirExists,
                    writePermission,
                    contents: dirContents
                }
            });
        } catch (error) {
            res.status(500).json({ error: error.message });
        }
    });
}

module.exports = setupRoutes;
