
# 智能招聘管理系统

## 项目概述
本系统是一个自动化简历处理与候选人评估工具，支持以下核心功能：
- 多格式简历解析（PDF/DOCX）
- 职位说明书智能分析
- 候选人匹配度评估
- Excel报告自动生成
- 可视化操作界面

## 主要特性
✅ 智能信息提取：自动解析简历中的教育背景、工作经历、技能等关键信息  
✅ 深度学习评估：基于岗位要求进行候选人匹配度分析  
✅ 批量处理能力：支持多文件批量处理与自动归档  
✅ 格式自适应：智能处理扫描版PDF与复杂版式文档  
✅ 可视化报表：生成带格式的Excel评估报告

## 环境要求
- Python 3.8+
- Windows/Linux/macOS
- 4GB+ 可用内存

## 安装步骤

### 1. 克隆仓库
```bash
git clone https://github.com/yourusername/recruitment-system.git
cd recruitment-system
```

### 2. 安装Python依赖
```bash
pip install -r requirements.txt
```

### 3. 安装系统组件
- **Tesseract OCR** ([下载地址](https://github.com/UB-Mannheim/tesseract/wiki))
- **Poppler Tools** ([下载地址](https://poppler.freedesktop.org/))

## 配置说明
在`config.ini`中设置：
```ini
[API]
api_key = your_deepseek_api_key
base_url = https://api.deepseek.com

[PATHS]
tesseract_path = C:\Program Files\Tesseract-OCR\tesseract.exe
poppler_path = C:\path\to\poppler\bin
```

## 使用指南

1. **启动系统**
```bash
python main.py
```

2. **界面操作流程**：
   1. 设置工作目录与文件路径
   2. 添加职位说明书（支持多选）
   3. 选择简历目录
   4. 指定输出Excel路径
   5. 点击"处理简历"开始分析

3. **输出结果**：
   - 结构化数据表格
   - 自动生成的评估结论
   - 处理日志文件（recruitment_system.log）

## 模块说明
- **ResumeProcessor**: 简历解析引擎（支持PDF/DOCX）
- **DeepSeekEvaluator**: 候选人评估模型
- **JobDescriptionProcessor**: 岗位说明书分析器
- **ExcelGenerator**: 智能报表生成模块
- **RecruitmentSystemGUI**: 可视化界面

## 技术支持
- PDF解析：PyPDF2 + Tesseract OCR
- 文档处理：python-docx + pdf2image
- AI接口：DeepSeek Chat API
- 界面框架：Tkinter

## 注意事项
1. 请确保API密钥有效
2. 支持中英文简历混合处理
3. 输出Excel路径需具有写权限
4. 处理大型简历集时建议预留足够存储空间

