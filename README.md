# 🤖 汇享易报告自助生成智能体

基于Python + Gradio的Word报告自动生成系统，AI 自动解析Excel数据并自动按照格式生成Word报告。
<img width="2698" height="2753" alt="CleanShot 2025-11-07 at 19 51 57@2x" src="https://github.com/user-attachments/assets/bde5e49d-183d-4f7a-bab5-2fa63f672d5f" />


## ✨ 功能特性

- 📊 自动解析Excel数据（支持.xls和.xlsx格式）
- 🔐 支持密码保护的Excel文件
- 📅 智能识别本周/上周数据（周一到周日）
- 📝 基于模板自动生成Word报告
- 🎨 友好的Web界面操作
- ⚡ 快速生成，即时下载

## 📋 系统要求

- Python 3.8+
- 操作系统：macOS / Linux / Windows

## 🚀 快速开始

### 方法一：Docker部署（推荐）

```bash
# 使用快速启动脚本
./docker-start.sh

# 或使用docker-compose
docker-compose up -d
```

访问地址：http://localhost:7861

详细说明请查看 [Docker使用说明.md](Docker使用说明.md)

### 方法二：本地部署

#### 1. 安装依赖

```bash
# 创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
source venv/bin/activate  # macOS/Linux
# 或
venv\Scripts\activate  # Windows

# 安装依赖包
pip install -r requirements.txt
```

#### 2. 启动应用

**使用启动脚本（推荐）**：

```bash
chmod +x start.sh
./start.sh
```

**或直接运行**：

```bash
source venv/bin/activate
python app.py
```

#### 3. 访问应用

启动成功后，在浏览器中打开：

```
http://localhost:7861
```

## 📖 使用说明

### Excel文件要求

您的Excel文件需要满足以下格式：

1. **必须包含两个sheet**：
   - `阳光xf登记`
   - `gab上访`

2. **每个sheet的结构**：
   - A列：序号
   - B列：登记时间（格式：月.日，例如：1.2、2.5、12.25）
   - 其他列：可包含任意数据

3. **时间格式示例**：
   - 1.2 → 2025年1月2日
   - 2.17 → 2025年2月17日
   - 12.25 → 2025年12月25日

### 操作步骤

1. **上传Excel文件**
   - 点击"1. 选择Excel文件"按钮
   - 选择您维护的信访数据Excel文件
   - 文件会自动上传到`upload`目录
   - 上传成功后会显示文件信息

2. **输入密码**（如需要）
   - 如果文件有密码保护，请在"2. Excel密码"框中输入
   - 默认密码：110110

3. **设置输出文件名**（可选）
   - 在"3. 输出文件名"框中输入生成的Word报告文件名
   - 默认格式：`报告_YYYYMMDD.docx`

4. **生成报告**
   - 点击"🚀 4. 开始生成"按钮
   - 系统会自动处理并生成报告

5. **下载文档**
   - 生成成功后，从"生成结果"区域下载Word文档

### 数据统计逻辑

系统会自动根据当前日期计算：

- **本周范围**：本周一 00:00 到本周日 23:59
- **上周范围**：上周一 00:00 到上周日 23:59

统计指标：
- `阳光xf登记+gab登记`：本周两个sheet人数总和
- `阳光xf登记`：本周"阳光xf登记"sheet中的人数
- `上周阳光xf登记`：上周"阳光xf登记"sheet中的人数
- `gab上访`：本周"gab上访"sheet中的人数
- `上周gab上访`：上周"gab上访"sheet中的人数
- **环比趋势**：自动计算并显示"上升X人"、"下降X人"或"持平"

## 📁 项目结构

```
reportgene/
├── app.py                          # Gradio主应用
├── excel_parser.py                 # Excel解析模块
├── date_calculator.py              # 日期计算模块
├── word_generator.py               # Word生成模块
├── template.docx                   # Word模板文件
├── requirements.txt                # 依赖包列表
├── start.sh                        # 启动脚本
├── README.md                       # 使用说明
├── output/                         # 生成的报告输出目录
├── 2025年复盘人员明细9.22.xls    # 示例数据文件
└── venv/                           # Python虚拟环境
```

## 🔧 技术栈

- **Web框架**：Gradio 4.44.0
- **数据处理**：Pandas 2.2.2
- **Excel解析**：openpyxl 3.1.5、xlrd 2.0.1
- **Word生成**：python-docx 1.1.2
- **文件解密**：msoffcrypto-tool 5.4.2

## 📝 生成的报告内容

系统会在Word模板中的"人员基本情况"部分自动填充以下数据：

```
本周，我市在京信访登记[X]人，其中国家信访局登记[X]人，
环比（[X]人）[持平/上升X人/下降X人]；
公安部登记[X]人，环比（[X]人）[持平/上升X人/下降X人]。
```

## ⚠️ 注意事项

1. **Excel文件格式**：确保Excel文件包含正确的sheet名称和列结构
2. **日期格式**：登记时间必须使用"月.日"格式（如：1.2、12.25）
3. **周的定义**：系统按照周一到周日计算一周
4. **年份假设**：所有日期默认为2025年
5. **文件密码**：如果Excel有密码保护，请确保输入正确的密码

## 🐛 常见问题

### 1. 提示"解析Excel时遇到问题"

- 检查Excel文件是否包含"阳光xf登记"和"gab上访"两个sheet
- 确认B列是否为"登记时间"列
- 验证密码是否正确

### 2. 统计人数为0

- 检查Excel中的登记时间是否在本周/上周范围内
- 确认日期格式是否正确（月.日）

### 3. Word文档无法打开

- 确保系统已正确安装python-docx
- 检查template.docx模板文件是否存在

## 📞 技术支持

如遇到问题，请检查：
1. Python版本是否 >= 3.8
2. 所有依赖包是否正确安装
3. Excel文件格式是否符合要求
4. 终端输出的错误信息

## 📄 许可证

本项目仅供内部使用。

---

**最后更新**: 2025-10-14


