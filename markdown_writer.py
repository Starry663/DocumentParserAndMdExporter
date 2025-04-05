# markdown_writer.py
import os

def save_as_markdown(output_md_path, text_runs, tables_data, image_files):
    # 打开输出Markdown文件
    with open(output_md_path, "w", encoding="utf-8") as md:
        # 写入文本段落
        for para in text_runs:
            # 确保段落不为空
            if para.strip():
                md.write(para.strip() + "\n\n")  # 写入段落后空一行

        # 写入表格数据（转换为Markdown表格格式）
        for table in tables_data:
            if not table:
                continue
            # 假定第一行是表头
            headers = table[0]
            md.write("| " + " | ".join(headers) + " |\n")
            # 写入表头与下一行的分隔线
            md.write("|" + " | ".join([" --- " for _ in headers]) + "|\n")
            # 写入表格剩余行
            for row in table[1:]:
                md.write("| " + " | ".join(row) + " |\n")
            md.write("\n")  # 每个表格后空一行

        # 写入图片引用
        if image_files:
            md.write("<!-- 提取的图片 -->\n")
            for img_path in image_files:
                md.write(f"![提取图片]({img_path})\n\n")
    print(f"Markdown 文件已保存: {output_md_path}")
