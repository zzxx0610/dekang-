import streamlit as st
import pandas as pd
import os
import io
import zipfile
import time

# --- 页面基础设置 ---
st.set_page_config(
    page_title="Excel 表格拆分工具",
    page_icon="📊",
    layout="wide"
)

# --- 主函数，现在接收一个streamlit占位符来显示日志 ---
def process_and_zip(uploaded_file, column_name, log_container):
    """
    处理上传的Excel文件，将其拆分，并将结果打包成一个ZIP文件。
    同时，将处理日志实时更新到指定的Streamlit容器中。
    """
    logs = []  # 用来收集日志信息

    def log_message(message):
        """辅助函数，用于记录日志并更新界面"""
        logs.append(message)
        # 使用Markdown的代码块格式来显示日志
        log_container.markdown("```\n" + "\n".join(logs) + "\n```")

    try:
        # 获取源文件名（不含扩展名），用于日志和输出文件名
        source_filename = os.path.splitext(uploaded_file.name)[0]
        
        log_message(f"准备处理文件: {uploaded_file.name}")
        df = pd.read_excel(uploaded_file)
        
        total_rows = len(df)
        log_message(f"✅ 成功读取源文件，共包含 {total_rows} 条数据。")

        unique_values = df[column_name].dropna().unique()
        log_message(f"🔍 在“{column_name}”列中发现 {len(unique_values)} 个独立的收货单位，准备开始拆分...")
        log_message("-" * 40) # 分割线

        # 创建一个在内存中的ZIP文件
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            original_columns = df.columns.tolist()
            
            processed_rows_count = 0
            
            for i, value in enumerate(unique_values, 1):
                df_group = df[df[column_name] == value]
                num_rows_in_group = len(df_group)
                processed_rows_count += num_rows_in_group
                
                # 清理文件名
                safe_filename = "".join([c for c in str(value) if c.isalnum() or c in (' ', '_', '-')]).rstrip()
                if not safe_filename:
                    safe_filename = f"未命名项目_{i}"
                
                output_filename_in_zip = f"{safe_filename}.xlsx"
                
                # 将拆分出的Excel文件写入内存
                excel_buffer = io.BytesIO()
                df_group.reindex(columns=original_columns).to_excel(excel_buffer, index=False, engine='openpyxl')
                excel_buffer.seek(0)
                
                # 将内存中的Excel文件添加到ZIP包中
                zf.writestr(output_filename_in_zip, excel_buffer.read())
                
                # 记录这条处理日志
                log_message(f"({i}/{len(unique_values)}) 已生成文件: {output_filename_in_zip} (包含 {num_rows_in_group} 条数据)")
                time.sleep(0.01) # 短暂休眠，让前端有时间渲染，看起来更流畅

        log_message("-" * 40)
        log_message("✅ 所有表格拆分完成！")

        # 最终核对
        if total_rows == processed_rows_count:
            log_message(f"数据核对成功：原始 {total_rows} 条，已处理 {processed_rows_count} 条。")
        else:
            unprocessed_rows = total_rows - processed_rows_count
            log_message(f"⚠️ 警告：数据核对不匹配！有 {unprocessed_rows} 条数据未被处理。")
            log_message(f"   (原因通常是 '{column_name}' 列中存在空白单元格)")
        
        # 将ZIP文件的指针也重置到开头
        zip_buffer.seek(0)
        return zip_buffer, source_filename

    except Exception as e:
        log_message(f"❌ 处理过程中发生错误: {e}")
        log_message("   请检查上传的文件格式是否正确，以及指定的列名是否存在于文件中。")
        return None, None


# --- Streamlit 界面布局 ---

st.title("📊 Excel 表格按列拆分工具")
st.markdown("上传一个Excel总表，指定一个用于分类的列，工具会自动将表格拆分成多个独立的Excel文件，并打包成ZIP供您下载。")
st.markdown("---")

# 1. 文件上传控件
uploaded_file = st.file_uploader("上传您的 Excel 总表", type=['xlsx'])

if uploaded_file is not None:
    st.subheader("1. 设置拆分规则")
    
    try:
        temp_df = pd.read_excel(uploaded_file, nrows=0)
        column_options = temp_df.columns.tolist()
        default_index = column_options.index('收货单位名称') if '收货单位名称' in column_options else 0
        column_to_split = st.selectbox("请选择用于分类的列名:", options=column_options, index=default_index)
    except Exception:
        column_to_split = st.text_input("无法自动读取列名，请输入用于分类的列名:", value="收货单位名称")

    st.subheader("2. 开始处理并查看日志")
    
    # 创建一个用于显示日志的占位符
    log_container = st.empty()
    log_container.info("准备就绪，点击下方按钮开始处理。")

    if st.button("🚀 开始拆分", use_container_width=True):
        # 在点击按钮后，清空占位符，准备显示新日志
        log_container.empty()
        
        with st.spinner('正在处理中，请耐心等待...'):
            zip_buffer, source_filename = process_and_zip(uploaded_file, column_to_split, log_container)
        
        if zip_buffer and source_filename:
            st.success("🎉 处理完成！可以下载结果了。")
            
            st.subheader("3. 下载结果")
            st.download_button(
                label="📥 下载拆分结果 (ZIP)",
                data=zip_buffer,
                file_name=f'{source_filename}_拆分结果.zip',
                mime='application/zip',
                use_container_width=True
            )
else:
    st.info("请上传一个 .xlsx 文件以开始。")

st.markdown("---")
st.write("由 AI 与开发者共同构建的小工具。")
