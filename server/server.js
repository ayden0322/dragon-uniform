const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
const fs = require('fs');
const app = express();
const port = process.env.PORT || 3000;

// 中間件設置
app.use(bodyParser.json({ limit: '10mb' }));

// 確保data目錄存在
const dataDir = path.join(__dirname, '../data');
if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true });
}

// 設置API路由
const setupRoutes = require('./api/routes');
setupRoutes(app);

// 靜態文件服務
app.use(express.static(path.join(__dirname, '../public')));

// 添加默認路由重定向到土庫國中
app.get('/', (req, res) => {
    res.redirect('/dragon');
});

// 校園特定路由
app.get('/:schoolId', (req, res) => {
    // 直接返回索引頁面，前端會處理數據加載
    res.sendFile(path.join(__dirname, '../public/index.html'));
});

// 啟動伺服器
app.listen(port, () => {
    console.log(`學生制服自動分配系統伺服器運行在 http://localhost:${port}`);
});
