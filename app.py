import streamlit as st
import pandas as pd
import os
import io
import zipfile

# --- 页面基础设置 ---
st.set_page_config(
    page_title="Excel 表格拆分工具",
    page_icon="📊",
    layout="wide"
)

# --- 主函数，包含之前脚本的核心逻辑 ---
def process_and_zip(uploaded_file, column_name):
    """
    处理上传的Excel文件，将其拆分，并将结果打包成一个ZIP文件。
    返回一个包含ZIP文件的内存中对象(BytesIO)。
    """
    try:
        df = pd.read_excel(uploaded_file)
        
        # 使用 st.info 在界面上显示反馈信息
        st.info(f"成功读取文件，共包含 {len(df)} 条数据。")

        unique_values = df[column_name].dropna().unique()
        st.info(f"在“{column_name}”列中发现 {len(unique_values)} 个独立的项目，准备开始拆分...")

        # 创建一个在内存中的ZIP文件
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            original_columns = df.columns.tolist()
            
            # 使用 st.progress 显示处理进度
            progress_bar = st.progress(0)
            
            for i, value in enumerate(unique_values, 1):
                df_group = df[df[column_name] == value]
                
                # 清理文件名
                safe_filename = "".join([c for c in str(value) if c.isalnum() or c in (' ', '_', '-')]).rstrip()
                if not safe_filename:
                    safe_filename = f"未命名项目_{i}"
                
                # 将拆分出的Excel文件写入内存
                excel_buffer = io.BytesIO()
                df_group.reindex(columns=original_columns).to_excel(excel_buffer, index=False, engine='openpyxl')
                excel_buffer.seek(0) # 重置指针到开头
                
                # 将内存中的Excel文件添加到ZIP包中
                zf.writestr(f"{safe_filename}.xlsx", excel_buffer.read())
                
                # 更新进度条
                progress_bar.progress(i / len(unique_values))

        # 将ZIP文件的指针也重置到开头
        zip_buffer.seek(0)
        return zip_buffer

    except Exception as e:
        st.error(f"处理过程中发生错误: {e}")
        st.error("请检查上传的文件格式是否正确，以及指定的列名是否存在于文件中。")
        return None


# --- Streamlit 界面布局 ---

st.title("📊 Excel 表格按列拆分工具")
st.markdown("上传一个Excel总表，指定一个用于分类的列，工具会自动将表格拆分成多个独立的Excel文件，并打包成ZIP供您下载。")
st.markdown("---")

# 1. 文件上传控件
uploaded_file = st.file_uploader("上传您的 Excel 总表", type=['xlsx'])

if uploaded_file is not None:
    # 让用户可以自定义列名
    st.subheader("设置拆分规则")
    
    # 尝试从文件中读取列名，提供给用户选择
    try:
        temp_df = pd.read_excel(uploaded_file, nrows=0) # 只读表头，速度快
        column_options = temp_df.columns.tolist()
        # 让用户选择列，默认推荐'收货单位名称'（如果存在的话）
        default_index = column_options.index('收货单位名称') if '收货单位名称' in column_options else 0
        column_to_split = st.selectbox("请选择用于分类的列名:", options=column_options, index=default_index)
    except Exception:
        # 如果读取失败，退回到手动输入
        column_to_split = st.text_input("无法自动读取列名，请输入用于分类的列名:", value="收货单位名称")

    # 2. “开始处理”按钮
    if st.button("🚀 开始拆分", use_container_width=True):
        with st.spinner('正在处理中，请稍候...'):
            zip_buffer = process_and_zip(uploaded_file, column_to_split)
        
        if zip_buffer:
            st.success("🎉 处理完成！可以下载结果了。")
            
            # 提取原始文件名用于命名ZIP包
            source_filename = os.path.splitext(uploaded_file.name)[0]
            
            # 3. 下载按钮
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
