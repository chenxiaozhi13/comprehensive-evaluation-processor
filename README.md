# 综合测评处理系统
![演示动画](https://github.com/chenxiaozhi13/comprehensive-evaluation-processor/raw/main/cute.gif)
## 项目简介
这是一个基于 Flask 框架开发的综合测评处理系统后端。它能够自动解析学生的综合测评 Word 文档（.docx），提取学号、姓名以及各项测评成绩，然后将这些数据汇总并导出为结构化的 Excel 文件。系统支持个人自评模式和班级批量评估模式，旨在提高测评数据处理的效率和准确性。


-   **上传文件功能**
![演示1](https://github.com/chenxiaozhi13/comprehensive-evaluation-processor/raw/main/1.png)

-   **历史下载功能**
![演示2](https://github.com/chenxiaozhi13/comprehensive-evaluation-processor/raw/main/2.png)
## 🌟主要功能
-   **Word 文档解析**：自动从 Word 文档中识别并提取学生信息和各项测评得分。
-   **多模式评估**：支持单个学生自评和批量班级评估两种模式。
-   **Excel 报告生成**：将处理后的数据生成包含总评分和各模块（思想品德、专业科研、体艺、劳动实践）明细的 Excel 文件。
-   **历史记录管理**：记录每次处理的历史信息，方便用户查看和管理。
-   **文件下载**：提供处理后 Excel 文件的下载功能。
-   **数据统计**：提供全站文件处理总数和平均处理时间统计。
-   **安全机制**：班评模式下载和历史记录删除需要管理员密码验证。
-   **性能优化**：支持 Gzip 压缩，提高数据传输效率。

## 🛠️技术栈
-   **后端框架**：Flask
-   **文档处理**：python-docx
-   **数据处理**：pandas
-   **文件操作**：os, io
-   **其他**：re, time, datetime, json, uuid, functools, flask-compress

## 🚀 安装与运行

### 1. 克隆仓库
首先，将本项目从 GitHub 克隆到你的本地机器：
```bash
git clone <YOUR_GITHUB_REPO_URL>
cd <你的项目目录>
```

### 2. 安装依赖
项目所需的所有 Python 库都已列在 `requirements.txt` 文件中。使用 pip 安装它们：
```bash
pip install -r requirements.txt
```

### 3. 运行应用程序
确保你当前的工作目录是项目的根目录，然后运行 Flask 应用程序：
```bash
python app.py
```
应用程序将在 `http://0.0.0.0:80` 上运行，通常你在本地可以通过浏览器访问 `http://127.0.0.1` 来访问。

## 📖 使用说明
1.  打开浏览器访问应用程序地址。
2.  在页面上选择"自评模式"或"班评模式"。
3.  上传对应的学生综合测评 Word 文档。
4.  系统处理完成后，会显示汇总表格，并提供 Excel 文件的下载链接。
5.  在班评模式下，下载文件或删除历史记录需要管理员密码，请自行查看app.py文件修改

## 📁项目结构
```
.
├── app.py              # Flask 主应用程序文件
├── templates/          # HTML 模板文件目录
│    ── index.html
├── processed_files/    # 处理后的 Excel 文件存储目录
├── temp_uploads/       # 临时上传文件目录
├── requirements.txt    # Python 依赖列表
├── history.json        # 历史记录文件
└── README.md           # 项目说明文件
```

## 📧 联系方式
如果你有任何问题或建议，联系我。 
