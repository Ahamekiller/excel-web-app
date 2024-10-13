const express = require('express');
const path = require('path');
const multer = require('multer');
const fs = require('fs');

const app = express();

// 设置ejs为模板引擎
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// 静态文件目录
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.urlencoded({ extended: true }));

// 设置文件上传目录
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadPath = path.join(__dirname, 'public/uploads');
    if (!fs.existsSync(uploadPath)) {
      fs.mkdirSync(uploadPath, { recursive: true });
    }
    cb(null, uploadPath);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, uniqueSuffix + '-' + file.originalname);
  }
});

const upload = multer({ storage: storage });

// 路由：显示上传页面
app.get('/', (req, res) => {
  res.render('upload');
});

// 路由：处理文件上传
app.post('/upload', upload.single('excelFile'), (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }
  res.redirect(`/view/${req.file.filename}`);
});

// 路由：展示上传的Excel文件内容
const viewRouter = require('./routes/view');  // 确保引用view.js
app.use('/view', viewRouter);

// 启动服务器
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
