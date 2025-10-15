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


# 获取当前脚本所在目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 默认Excel密码
DEFAULT_PASSWORD = "110110"
UPLOAD_DIR = os.path.join(BASE_DIR, "upload")
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.docx")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")


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


def generate_report(upload_path, output_filename, password):
    """
    生成报告的主函数
    
    Args:
        upload_path: 上传后的Excel文件路径
        output_filename: 输出文件名
        password: Excel密码
    
    Returns:
        tuple: (输出文件路径, 状态消息)
    """
    try:
        # 验证输入
        if not upload_path or not os.path.exists(upload_path):
            return None, "❌ 请先上传Excel文件"
        
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
            return None, error_msg
        
        # 显示统计结果
        stats_msg = f"""
✅ Excel解析成功！

📈 数据统计：
  • 本周日期范围: {parser.current_week_start.strftime('%Y-%m-%d')} 到 {parser.current_week_end.strftime('%Y-%m-%d')}
  • 上周日期范围: {parser.last_week_start.strftime('%Y-%m-%d')} 到 {parser.last_week_end.strftime('%Y-%m-%d')}

  • 阳光xf登记: 本周 {data['sunshine_current']} 人，上周 {data['sunshine_last']} 人，{data['sunshine_trend']}
  • gab上访: 本周 {data['gab_current']} 人，上周 {data['gab_last']} 人，{data['gab_trend']}
  • 本周总计: {data['total_current']} 人

📝 正在生成Word文档...
"""
        print(stats_msg)
        
        # 步骤2: 生成Word文档
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        generator = WordGenerator(TEMPLATE_PATH)
        
        success = generator.generate(data, output_path)
        
        if success:
            final_msg = f"""
✅ 报告生成成功！

📊 统计数据：
  • 阳光xf登记: 本周 {data['sunshine_current']} 人，上周 {data['sunshine_last']} 人，{data['sunshine_trend']}
  • gab上访: 本周 {data['gab_current']} 人，上周 {data['gab_last']} 人，{data['gab_trend']}
  • 本周总计: {data['total_current']} 人

📄 文件已保存: {output_filename}
"""
            return output_path, final_msg
        else:
            return None, "❌ Word文档生成失败"
    
    except Exception as e:
        error_msg = f"❌ 生成报告时出错: {str(e)}"
        print(error_msg)
        return None, error_msg


# 创建Gradio界面
def create_ui():
    """创建Gradio用户界面"""
    
    with gr.Blocks(title="汇享易报告自助生成智能体", theme=gr.themes.Soft()) as app:
        gr.Markdown(
            """
            # 🤖 汇享易报告自助生成智能体
            
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
                    lines=15,
                    interactive=False
                )
                
                file_output = gr.File(
                    label="下载生成的报告"
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
                5. **下载文档**：生成成功后，从"生成结果"区域下载Word文档
                
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
            outputs=[file_output, status_output]
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


