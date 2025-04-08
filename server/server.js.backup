const express = require('express');
const path = require('path');
const app = express();
const port = process.env.PORT || 3000;

// 靜態文件服務
app.use(express.static(path.join(__dirname, '../public')));

// 啟動伺服器
app.listen(port, () => {
  console.log(`學生制服自動分配系統伺服器運行在 http://localhost:${port}`);
}); 