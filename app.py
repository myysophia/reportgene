"""
汇享易报告自助生成智能体
主应用程序 - Gradio界面
"""
import gradio as gr
import os
import shutil
from datetime import datetime
from excel_parser import ExcelParser
from word_generator import WordGenerator
from docx import Document
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 默认Excel密码
DEFAULT_PASSWORD = "110110"
UPLOAD_DIR = os.path.join(BASE_DIR, "upload")
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.docx")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
VERSION_FILE = os.path.join(BASE_DIR, "version.txt")


def get_version():
    """读取版本号"""
    try:
        with open(VERSION_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except:
        return "v1.0"


def upload_file(file):
    """
    上传Excel文件到upload目录
    
    Args:
        file: 上传的文件对象
    
    Returns:
        tuple: (上传后的文件路径, 状态消息)
    """
    try:
        if file is None:
            return None, "❌ 请选择要上传的Excel文件"
        
        # 检查文件类型
        if not (file.name.endswith('.xls') or file.name.endswith('.xlsx')):
            return None, "❌ 只支持.xls和.xlsx格式的Excel文件"
        
        # 确保upload目录存在
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        
        # 生成新的文件名（带时间戳，避免重复）
        filename = os.path.basename(file.name)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        name, ext = os.path.splitext(filename)
        new_filename = f"{name}_{timestamp}{ext}"
        upload_path = os.path.join(UPLOAD_DIR, new_filename)
        
        # 复制文件到upload目录
        shutil.copy2(file.name, upload_path)
        
        # 获取文件信息
        file_size = os.path.getsize(upload_path)
        file_size_mb = round(file_size / (1024 * 1024), 2)
        
        success_msg = f"""
✅ 文件上传成功！

📁 文件信息：
  • 文件名：{new_filename}
  • 大小：{file_size_mb} MB
  • 上传时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

现在可以输入输出文件名和密码，然后点击"开始生成"按钮。
"""
        
        return upload_path, success_msg
        
    except Exception as e:
        return None, f"❌ 文件上传失败: {str(e)}"


def preview_word_document(file_path):
    """
    预览Word文档内容（HTML格式，字符级动态内容标注）
    
    Args:
        file_path: Word文档路径
    
    Returns:
        str: HTML格式的文档预览
    """
    try:
        if not file_path or not os.path.exists(file_path):
            return """
            <div style="color: #e74c3c; padding: 20px; border: 1px solid #e74c3c; border-radius: 8px; background-color: #fdf2f2;">
                <h3>❌ 文档不存在，无法预览</h3>
            </div>
            """
        
        # 读取Word文档
        doc = Document(file_path)
        
        # 构建HTML内容
        html_content = """
        <div style="font-family: 'Microsoft YaHei', 'SimSun', serif; line-height: 1.6; color: #333;">
            <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px; border-radius: 8px 8px 0 0; margin-bottom: 0;">
                <h2 style="margin: 0; font-size: 18px;">📄 文档预览（标注）</h2>
            </div>
            <div style="border: 1px solid #ddd; border-top: none; border-radius: 0 0 8px 8px; padding: 20px; background-color: #fafafa; max-height: 600px; overflow-y: auto;">
        """
        
        # 提取文本内容
        content_lines = []
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                content_lines.append(paragraph.text.strip())
        
        # 显示全部内容，进行字符级标注
        for i, line in enumerate(content_lines):
            # 对每行进行字符级动态内容标注
            annotated_line = _annotate_dynamic_content(line)
            
            # 根据内容类型添加不同的样式
            if line.startswith('（一）') or line.startswith('（二）') or line.startswith('（三）'):
                html_content += f'<h4 style="color: #2c3e50; margin: 15px 0 8px 0; font-size: 16px;">{annotated_line}</h4>'
            elif line.startswith('1、') or line.startswith('2、') or line.startswith('3、'):
                html_content += f'<p style="margin: 8px 0; padding-left: 20px; color: #34495e;">{annotated_line}</p>'
            elif line.startswith('本周，我市在京信访登记') or line.startswith('从涉事地看') or line.startswith('从涉稳群体类型看') or line.startswith('从进京交通工具看'):
                html_content += f'<p style="margin: 10px 0; font-weight: 500; color: #2980b9;">{annotated_line}</p>'
            elif line.startswith('"情指行"机制复盘报告'):
                html_content += f'<h3 style="color: #8e44ad; text-align: center; margin: 10px 0;">{annotated_line}</h3>'
            elif line.startswith('第') and line.endswith('期'):
                html_content += f'<h4 style="color: #8e44ad; text-align: center; margin: 5px 0;">{annotated_line}</h4>'
            elif line.startswith('阳光信访登记复盘工作周报'):
                html_content += f'<h3 style="color: #8e44ad; text-align: center; margin: 10px 0;">{annotated_line}</h3>'
            else:
                html_content += f'<p style="margin: 8px 0; color: #2c3e50;">{annotated_line}</p>'
        
        # 添加动态内容说明
        html_content += '''
        <div style="margin-top: 20px; padding: 15px; background-color: #e8f5e8; border-radius: 5px; border-left: 4px solid #28a745;">
            <h4 style="margin: 0 0 10px 0; color: #155724;">📊 动态渲染内容说明</h4>
            <p style="margin: 5px 0; color: #155724;">
                <span style="background-color: #ffc107; color: #856404; padding: 2px 6px; border-radius: 3px; font-size: 12px;">黄色高亮</span> 
                表示从Excel中自动提取的动态数据
            </p>
            <p style="margin: 5px 0; color: #155724;">
                • 统计数据：<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px;">8人</span>、<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px;">6人</span>
            </p>
            <p style="margin: 5px 0; color: #155724;">
                • 人员信息：<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px;">贾汪XX</span>、<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px;">市直XX</span>
            </p>
            <p style="margin: 5px 0; color: #155724;">
                • 趋势变化：<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px;">上升2人</span>、<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px;">下降</span>
            </p>
            <p style="margin: 5px 0; color: #155724;">
                • 地区统计：<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px;">贾汪1人</span>、<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px;">铜山1人</span>
            </p>
        </div>
        '''
        
        html_content += """
            </div>
        </div>
        """
        
        return html_content
        
    except Exception as e:
        return f"""
        <div style="color: #e74c3c; padding: 20px; border: 1px solid #e74c3c; border-radius: 8px; background-color: #fdf2f2;">
            <h3>❌ 预览失败</h3>
            <p>错误信息: {str(e)}</p>
        </div>
        """

def _annotate_dynamic_content(text):
    """
    对文本进行字符级动态内容标注
    
    Args:
        text: 原始文本
    
    Returns:
        str: 标注后的HTML文本
    """
    import re
    
    # 定义动态内容模式
    patterns = [
        # 统计数据模式
        (r'(\d+人)', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1</span>'),
        
        # 人员信息模式（责任单位+姓名）
        (r'([贾汪市直铜山云龙经开区丰县沛县邳州泉山新沂睢宁鼓楼])([A-Za-z\u4e00-\u9fa5]{1,3}XX?)', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1\2</span>'),
        
        # 趋势变化模式
        (r'(上升\d+人|下降\d+人|持平)', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1</span>'),
        
        # 地区统计模式
        (r'([贾汪市直铜山云龙经开区丰县沛县邳州泉山新沂睢宁鼓楼])(\d+人)', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1\2</span>'),
        
        # 诉求类型模式
        (r'(征地拆迁|讨薪|拖欠工程款|失地保险|案件办理|截访)(\d+人)?', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1\2</span>'),
        
        # 进京方式模式
        (r'(公路|铁路|长期在京)(\d+人)?', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1\2</span>'),
        
        # 环比数据模式
        (r'(环比（\d+人）)', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1</span>'),
        
        # 百分比模式
        (r'(\d+\.\d+%)', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1</span>'),
        
        # 具体数字模式（在特定上下文中）
        (r'(本周登记\d+人中)', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1</span>'),
        (r'(在库\d+人中)', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1</span>'),
        (r'(\d+人触发平台)', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1</span>'),
        (r'(\d+人未触发)', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1</span>'),
        
        # 车辆数据模式
        (r'(\d+车\d+人)', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1</span>'),
        
        # 具体人员姓名模式（脱敏后）
        (r'([A-Za-z\u4e00-\u9fa5]{1,2}XX)', r'<span style="background-color: #ffc107; color: #856404; padding: 1px 3px; border-radius: 2px; font-weight: bold;">\1</span>'),
    ]
    
    # 应用所有模式进行标注
    annotated_text = text
    for pattern, replacement in patterns:
        annotated_text = re.sub(pattern, replacement, annotated_text)
    
    return annotated_text

def _is_dynamic_content(line):
    """
    判断是否为动态内容（从Excel中提取的数据）
    
    Args:
        line: 文本行
    
    Returns:
        bool: 是否为动态内容
    """
    # 检查是否包含动态数据的特征
    dynamic_indicators = [
        '本周，我市在京信访登记',  # 包含动态人数
        '从涉事地看',              # 包含动态地区统计
        '从涉稳群体类型看',        # 包含动态诉求统计
        '从进京交通工具看',        # 包含动态进京方式统计
        '环比（',                 # 包含环比数据
        '人（贾汪',               # 包含人员信息
        '人（市直',               # 包含人员信息
        '人（铜山',               # 包含人员信息
        '人（云龙',               # 包含人员信息
        '人（经开区',             # 包含人员信息
        '人（丰县',               # 包含人员信息
        '人（沛县',               # 包含人员信息
        '人（邳州',               # 包含人员信息
        '人（泉山',               # 包含人员信息
        '人（新沂',               # 包含人员信息
        '人（睢宁',               # 包含人员信息
        '人（鼓楼',               # 包含人员信息
        '人（贾汪',               # 包含人员信息
        '人（是',                 # 包含人员信息
        '征地拆迁',               # 包含诉求类型
        '讨薪',                   # 包含诉求类型
        '案件办理',               # 包含诉求类型
        '失地保险',               # 包含诉求类型
        '公路',                   # 包含进京方式
        '铁路',                   # 包含进京方式
        '长期在京',               # 包含进京方式
        '上升',                   # 包含趋势
        '下降',                   # 包含趋势
        '持平',                   # 包含趋势
    ]
    
    # 检查是否包含数字（表示统计数据）
    import re
    has_numbers = bool(re.search(r'\d+人', line))
    
    # 检查是否包含动态指标
    has_dynamic_indicators = any(indicator in line for indicator in dynamic_indicators)
    
    return has_numbers or has_dynamic_indicators


def generate_report(upload_path, output_filename, password):
    """
    生成报告的主函数
    
    Args:
        upload_path: 上传后的Excel文件路径
        output_filename: 输出文件名
        password: Excel密码
    
    Returns:
        tuple: (输出文件路径, 状态消息, 预览内容)
    """
    try:
        # 验证输入
        if not upload_path or not os.path.exists(upload_path):
            return None, "❌ 请先上传Excel文件", ""
        
        if not output_filename:
            # 如果未提供文件名，使用默认格式
            output_filename = f"报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        
        # 确保文件名以.docx结尾
        if not output_filename.endswith('.docx'):
            output_filename += '.docx'
        
        # 确保密码不为空
        if not password:
            password = DEFAULT_PASSWORD
        
        # 步骤1: 解析Excel文件
        status_msg = "📊 正在解析Excel数据..."
        print(status_msg)
        
        parser = ExcelParser(upload_path, password=password)
        data = parser.parse_all()
        
        # 检查是否有错误
        if data.get('errors'):
            error_msg = "⚠️ 解析Excel时遇到以下问题：\n" + "\n".join(data['errors'])
            return None, error_msg, ""
        
        # 显示统计结果
        stats_msg = f"""✅ Excel解析成功！

📅 统计时间范围：
  • 本周: {parser.current_week_start.strftime('%Y-%m-%d')} 至 {parser.current_week_end.strftime('%Y-%m-%d')}
  • 上周: {parser.last_week_start.strftime('%Y-%m-%d')} 至 {parser.last_week_end.strftime('%Y-%m-%d')}

📈 数据统计：
  • 阳光xf登记: 本周 {data['sunshine_current']} 人，上周 {data['sunshine_last']} 人，{data['sunshine_trend']}
  • gab上访: 本周 {data['gab_current']} 人，上周 {data['gab_last']} 人，{data['gab_trend']}
  • 本周总计: {data['total_current']} 人

📝 正在生成Word文档..."""
        print(stats_msg)
        
        # 步骤2: 生成Word文档
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        generator = WordGenerator(TEMPLATE_PATH)
        
        success = generator.generate(data, output_path)
        
        if success:
            # 生成预览内容
            preview_content = preview_word_document(output_path)
            
            final_msg = f"""✅ 报告生成成功！

📅 统计时间范围：
  • 本周: {parser.current_week_start.strftime('%Y-%m-%d')} 至 {parser.current_week_end.strftime('%Y-%m-%d')}
  • 上周: {parser.last_week_start.strftime('%Y-%m-%d')} 至 {parser.last_week_end.strftime('%Y-%m-%d')}

📊 统计数据：
  • 阳光xf登记: 本周 {data['sunshine_current']} 人，上周 {data['sunshine_last']} 人，{data['sunshine_trend']}
  • gab上访: 本周 {data['gab_current']} 人，上周 {data['gab_last']} 人，{data['gab_trend']}
  • 本周总计: {data['total_current']} 人

📄 文件已保存: {output_filename}
请查看下方预览，确认无误后点击下载。
"""
            return output_path, final_msg, preview_content
        else:
            return None, "❌ Word文档生成失败", ""
    
    except Exception as e:
        error_msg = f"❌ 生成报告时出错: {str(e)}"
        print(error_msg)
        return None, error_msg, ""


# 创建Gradio界面
def create_ui():
    """创建Gradio用户界面"""
    
    version = get_version()
    
    with gr.Blocks(
        title="汇享易报告自助生成智能体", 
        theme=gr.themes.Soft(),
        css="""
        footer {visibility: hidden}
        .version-badge {
            position: absolute;
            top: 20px;
            right: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 8px 16px;
            border-radius: 20px;
            font-size: 14px;
            font-weight: bold;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
            z-index: 1000;
        }
        """
    ) as app:
        gr.HTML(f'<div class="version-badge">版本 {version}</div>')
        
        gr.Markdown(
            """
            <div style="text-align: center;">
            <h1>🤖 汇享易报告自助生成智能体</h1>
            </div>
            
            上传您的Excel文件，系统将自动解析数据并生成Word报告
            """
        )
        
        with gr.Row():
            with gr.Column(scale=1):
                # 上传区域
                gr.Markdown("### 📤 上传Excel文件")
                
                excel_input = gr.File(
                    label="1. 选择Excel文件",
                    file_types=['.xls', '.xlsx'],
                    type="filepath"
                )
                
                upload_status = gr.Textbox(
                    label="上传状态",
                    lines=8,
                    interactive=False,
                    placeholder="请先选择Excel文件..."
                )
                
                # 隐藏的上传路径状态
                uploaded_path = gr.State(value=None)
                
                gr.Markdown("### ⚙️ 生成设置")
                
                password_input = gr.Textbox(
                    label="2. Excel密码（如有）",
                    value=DEFAULT_PASSWORD,
                    type="password",
                    placeholder="如果文件有密码保护，请输入密码"
                )
                
                output_name = gr.Textbox(
                    label="3. 输出文件名",
                    value=f"报告_{datetime.now().strftime('%Y%m%d')}.docx",
                    placeholder="例如: 报告_20251014.docx"
                )
                
                generate_btn = gr.Button("🚀 4. 开始生成", variant="primary", size="lg")
            
            with gr.Column(scale=1):
                # 输出组件
                gr.Markdown("### 📥 生成结果")
                
                status_output = gr.Textbox(
                    label="处理状态",
                    lines=8,
                    interactive=False
                )
                
                # 预览组件
                preview_output = gr.HTML(
                    label="📄 文档预览",
                    value="",
                    elem_id="preview"
                )
                
                file_output = gr.File(
                    label="📥 下载生成的报告"
                )
        
        # 使用说明
        with gr.Accordion("📖 使用说明", open=False):
            gr.Markdown(
                """
                ### 使用步骤：
                
                1. **上传Excel文件**：点击"选择Excel文件"按钮，选择您维护的信访数据Excel文件
                2. **输入密码**：如果文件有密码保护，请输入密码（默认：110110）
                3. **设置文件名**：输入生成的Word报告文件名（可选，默认自动命名）
                4. **生成报告**：点击"开始生成"按钮
                5. **预览文档**：生成成功后，在"文档预览"区域查看报告内容
                6. **下载文档**：确认预览无误后，点击"下载生成的报告"下载Word文档
                
                ### Excel文件要求：
                
                - 必须包含两个sheet：
                  - "阳光xf登记"
                  - "gab上访"
                - 每个sheet的B列为"登记时间"（格式：月.日，例如1.2、2.5）
                - 系统会自动统计本周和上周的数据
                
                ### 注意事项：
                
                - 系统根据当前日期自动计算本周和上周的范围（周一到周日）
                - 支持.xls和.xlsx格式的Excel文件
                - 如遇到问题，请检查Excel文件格式是否正确
                """
            )
        
        # 绑定事件
        # 文件上传时自动处理
        excel_input.change(
            fn=upload_file,
            inputs=[excel_input],
            outputs=[uploaded_path, upload_status]
        )
        
        # 生成报告
        generate_btn.click(
            fn=generate_report,
            inputs=[uploaded_path, output_name, password_input],
            outputs=[file_output, status_output, preview_output]
        )
    
    return app


if __name__ == '__main__':
    # 确保输出目录和上传目录存在
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    
    # 创建并启动应用
    app = create_ui()
    app.launch(
        server_name="0.0.0.0",
        server_port=7861,
        share=False,
        show_error=True
    )


