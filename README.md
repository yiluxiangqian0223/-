考场图像信息核对单生成工具 - 用于创建标准化考试管理文档

## 📋 目录

- [简介](#-简介)
- [功能特性](#-功能特性)
- [安装](#-安装)
- [使用方法](#-使用方法)
- [数据格式要求](#-数据格式要求)
- [配置](#-配置)
- [示例](#-示例)
- [API接口](#-api接口)
- [开发指南](#-开发指南)
- [贡献指南](#-贡献指南)
- [问题排查](#-问题排查)
- [许可证](#-许可证)

## 💡 简介

**Exam Image Generator** 是一个专业的Python工具，用于自动生成标准化的考场图像信息核对单。该工具专为教育考试场景设计，能够从Excel数据源读取考生信息，结合证件照片，批量生成符合考试管理规范的PDF文档。

## ✨ 功能特性

- 📊 **数据驱动**: 从Excel文件自动读取考生信息
- 📸 **照片集成**: 自动匹配考生证件照片
- 📄 **标准化输出**: 生成符合考试管理规范的PDF文档
- 🎨 **定制布局**: 支持竖排文字和多种信息展示方式
- 📈 **批量处理**: 一次性生成多个考场的核对单
- 🌐 **多语言支持**: 内置中文字符集支持
- 🛡️ **数据安全**: 本地处理，数据不上传云端
- 🔧 **高度可配置**: 灵活的布局和样式设置
- 📅 **时间戳标记**: 自动添加生成时间戳
- 📋 **日志记录**: 详细的处理过程日志

## 🛠️ 安装

### 系统要求
- Python 3.8 或更高版本
- 至少 100MB 可用磁盘空间
- 支持常见图片格式（JPG、PNG、BMP）

### 安装步骤

1. **克隆仓库**
bash
git clone https://github.com/yourusername/exam-image-generator.git
cd exam-image-generator
```

2. **创建虚拟环境（推荐）**
bash
# Linux/Mac
python -m venv venv
source venv/bin/activate

# Windows
python -m venv venv
venv\Scripts\activate
```

3. **安装依赖**
```bash
pip install -r requirements.txt
```

## 🚀 使用方法

### 基本使用

1. **准备数据文件**
   - 将考生信息保存为 `2026kctx.xlsx`，包含以下列：
     - `考场号` - 考场编号
     - `座位号` - 座位编号
     - `姓名` - 考生姓名
     - `身份证号` - 身份证号码

2. **准备照片文件**
   - 在 `photos` 目录下存放以身份证号命名的照片文件
   - 支持格式：JPG、PNG、BMP
   - 推荐尺寸：358x441像素（护照照片规格）

3. **运行生成器**
```bash
python exam_generator/generator.py
```

### 交互式配置
程序启动后会提示输入以下路径：
- 数据目录路径（默认：./data）
- 照片目录路径（默认：./photos） 
- 输出目录路径（默认：./output）

## 📁 数据格式要求

### Excel数据格式
确保Excel文件包含以下列（列名区分大小写）：

| 考场号 | 座位号 | 姓名 | 身份证号 |
|--------|--------|------|----------|
| 1      | 1      | 张三 | 123456789012345678 |
| 1      | 2      | 李四 | 123456789012345679 |

### 照片文件命名规则
- 文件名必须与身份证号完全一致（不含扩展名）
- 例如：`123456789012345678.jpg`
- 支持的格式：`.jpg`, `.jpeg`, `.png`, `.bmp`

### 目录结构示例
```
project-root/
├── exam_generator/
│   └── generator.py          # 主程序文件
├── data/
│   └── 2026kctx.xlsx         # 考生数据文件
├── photos/                   # 照片目录
│   ├── 123456789012345678.jpg
│   └── 123456789012345679.png
├── output/                   # 输出目录
├── requirements.txt          # 依赖文件
├── README.md                 # 说明文档
└── LICENSE                   # 许可证
```

## 🔧 配置

### PDF布局配置
- 每页显示30名考生信息（5行×6列）
- 照片尺寸：2.16cm × 2.74cm
- 左侧竖排显示座位号和姓名
- 底部显示考试注意事项

### 自定义配置
如需修改PDF布局，可以修改以下参数：
- `students_per_page`: 每页显示考生数量
- `rows`, `cols`: 表格行列数
- `photo_width`, `photo_height`: 照片尺寸
- `side_width`: 左侧竖排区域宽度

## 📸 示例

生成的PDF文档包含：
- 统一的标题和考试信息
- 30名考生的图像信息（5行×6列）
- 左侧竖排显示座位号和姓名
- 右侧显示照片和基本信息
- 底部考试注意事项

### 输出文件命名规则
```
考场图像单_XX考场_YYYYMMDD_HHMMSS.pdf
```
例如：`考场图像单_01考场_20240601_143022.pdf`

## 🔌 API接口

### 主要类

```python
from exam_generator.generator import ExamImageGenerator

# 初始化生成器
generator = ExamImageGenerator(
    data_dir="./data",
    photo_dir="./photos", 
    output_dir="./output"
)

# 生成单个考场
generator.generate_room_pdf(room_num=1, exam_info={
    'date': '2026-04-06上午09:00-11:30',
    'exam_site': '2幢教学楼',
    'exam_room': '01'
})

# 生成所有考场
files = generator.generate_all_rooms(
    exam_date='2026-04-06上午09:00-11:30',
    exam_site='2幢教学楼'
)
```

### 主要方法

| 方法 | 参数 | 返回值 | 说明 |
|------|------|--------|------|
| `__init__()` | data_dir, photo_dir, output_dir | - | 初始化生成器 |
| `generate_room_pdf()` | room_num, exam_info | str | 生成单个考场PDF |
| `generate_all_rooms()` | exam_date, exam_site | list[str] | 生成所有考场PDF |

## 🛠️ 开发指南

### 环境设置
```bash
# 克隆仓库
git clone https://github.com/yourusername/exam-image-generator.git
cd exam-image-generator

# 创建开发环境
python -m venv dev-env
source dev-env/bin/activate  # Linux/Mac
# 或
dev-env\Scripts\activate     # Windows

# 安装开发依赖
pip install -r requirements.txt
pip install pytest black flake8  # 开发工具
```
### 测试
```bash
# 运行测试
pytest tests/

# 代码格式化
black .

# 代码检查
flake8 .
```

## 🤝 贡献指南

我们欢迎各种形式的贡献！

### 贡献流程
1. Fork 仓库
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 创建 Pull Request

### 代码规范
- 使用 PEP 8 代码风格
- 添加类型注解
- 编写详细文档字符串
- 确保代码测试覆盖率

## 🛡️ 安全说明

- 所有数据处理均在本地进行
- 不收集或传输任何用户数据
- 请妥善保管包含个人信息的Excel文件和照片
- 生成的PDF文件包含敏感信息，请注意保密

## 🚨 问题排查

### 常见问题

**Q: 找不到中文字体怎么办？**
A: 程序会自动尝试多个常见字体路径，如果都不行，会使用默认字体。建议在系统中安装中文字体。

**Q: 照片无法显示？**
A: 检查照片文件名是否与身份证号完全一致，确保照片格式正确。

**Q: Excel文件读取失败？**
A: 确保Excel文件包含必需的列名，且文件格式为.xlsx。

**Q: PDF生成失败？**
A: 检查依赖是否正确安装，确保有足够的磁盘空间。

### 日志查看
程序会在控制台输出详细的处理信息，如需更详细日志，请修改日志级别。

## 📞 支持

遇到问题？请查看：
- [Issues](https://github.com/yourusername/exam-image-generator/issues) - 提交bug报告
- [Discussions](https://github.com/yourusername/exam-image-generator/discussions) - 讨论功能需求
- 邮箱：support@example.com

## 🙏 致谢

感谢以下开源项目的支持：
- [ReportLab](https://www.reportlab.com/) - PDF生成库
- [Pandas](https://pandas.pydata.org/) - 数据处理库
- [Python](https://www.python.org/) - 编程语言

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

MIT许可证允许：
- ✅ 商业使用
- ✅ 修改代码
- ✅ 分发副本
- ✅ 私人使用

---

⭐ 如果这个项目对你有帮助，请给我们一个星标！

🐛 发现问题？请提交 [Issue](https://github.com/yourusername/exam-image-generator/issues/new)

---

## 更新日志

### v1.0.0 (2024-06-01)
- ✅ 初始版本发布
- ✅ 支持Excel数据导入
- ✅ 支持照片集成
- ✅ 生成标准化PDF输出
- ✅ 支持批量处理
``````
