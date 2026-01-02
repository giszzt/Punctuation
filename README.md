# 标点。Punctuation

一个极简的英文标点转中文标点工具，支持文本粘贴和 DOCX 文档上传。

![Preview](https://img.shields.io/badge/纯前端-无需后端-brightgreen)
![License](https://img.shields.io/badge/License-MIT-blue)

## ✨ 功能特性

- **智能识别**：仅在中文语境下转换标点，不影响纯英文内容
- **双模式支持**：
  - 📝 **文本模式**：实时粘贴转换，即时预览结果
  - 📄 **文档模式**：上传 DOCX 文件，保留原格式导出
- **可视化反馈**：
  - 进度环动画展示处理进度
  - 模糊背景预览文档内容
  - 统计数据显示转换成果
- **深色模式**：自动适配系统主题

## 🎯 支持的标点转换

| 英文 | 中文 |
|:---:|:---:|
| `,` | `，` |
| `.` | `。` |
| `:` | `：` |
| `;` | `；` |
| `?` | `？` |
| `!` | `！` |
| `()` | `（）` |
| `""` | `""` |
| `''` | `''` |

## 🚀 快速开始

### 在线使用

直接打开 `index.html` 即可使用，无需安装任何依赖。

### 本地开发

```bash
# 克隆仓库
git clone https://github.com/giszzt/Punctuation.git

# 进入目录
cd Punctuation

# 启动本地服务器（可选）
npx http-server -p 8080
```

然后访问 http://localhost:8080

## 🛠️ 技术栈

- **纯前端**：HTML + CSS + JavaScript
- **无框架依赖**：轻量、快速
- **第三方库**：
  - [JSZip](https://stuk.github.io/jszip/) - 处理 DOCX 文件
  - [FileSaver.js](https://github.com/eligrey/FileSaver.js/) - 文件下载

## 📖 使用说明

### 文本模式

1. 点击左侧文本区域
2. 粘贴或输入文本
3. 右侧实时显示转换结果
4. 点击"复制结果"按钮复制

### 文档模式

1. 点击右上角"文档"切换模式
2. 点击圆圈上传或拖拽 DOCX 文件
3. 等待处理完成
4. 点击"下载处理后的文档"

### Word 插件 (VBA 宏)

我们提供了一个 Word VBA 宏脚本，让您可以在 Word 中直接一键转换标点。

1. **获取代码**：在仓库中找到 `PunctuationConverter.bas` 文件。
2. **导入 Word**：
   - 在 Word 中按 `Alt + F11` 打开 VBA 编辑器。
   - 选择菜单 `文件` -> `导入文件`，选择 `PunctuationConverter.bas`。
   - 或者：选择 `插入` -> `模块`，然后将代码复制粘贴进去。
3. **运行**：
   - 按 `Alt + F8` 打开宏列表。
   - 选择 `ConvertPunctuation` 并点击运行。


## 📄 License

MIT License
