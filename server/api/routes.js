const { getConfig, saveConfig } = require('./config');
const { saveAllocation, getAllocationHistory, getAllocationDetail } = require('./allocation');

function setupRoutes(app) {
    // 配置相關路由
    app.get('/api/config/:schoolId/:configType', getConfig);
    app.post('/api/config/:schoolId/:configType', saveConfig);
    
    // 分配相關路由
    app.post('/api/allocation/:schoolId', saveAllocation);
    app.get('/api/allocation/:schoolId/history', getAllocationHistory);
    app.get('/api/allocation/:schoolId/detail/:id', getAllocationDetail);
}

module.exports = setupRoutes;
