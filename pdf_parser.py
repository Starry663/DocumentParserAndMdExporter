# pdf_parser.py
import fitz  # PyMuPDF 的引用名通常为 fitz
import pdfplumber
import os
from PIL import Image

def parse_pdf(pdf_path, output_dir):
    # 确保输出目录和图片子目录存在
    img_dir = os.path.join(output_dir, "images")
    os.makedirs(img_dir, exist_ok=True)

    text_runs = []    # 存储PDF中所有页面的文本
    tables_data = []  # 存储PDF中所有表格的数据
    image_files = []  # 存储提取的图片文件路径

    # 打开 PDF 文件
    pdf_document = fitz.open(pdf_path)
    # 使用 pdfplumber 打开同一 PDF 以提取表格
    pdf_tables = pdfplumber.open(pdf_path)

    # 遍历 PDF 每一页
    for page_number in range(len(pdf_document)):
        page = pdf_document[page_number]
        # 提取文本
        text = page.get_text()  # 获取纯文本内容
        if text:
            # 为清晰起见，可以在每页文本后添加一个分页符或空行
            text_runs.append(text.strip())
        
        # 提取表格（可能有多个表格）
        page_tables = pdf_tables.pages[page_number].extract_tables()
        if page_tables:
            for table in page_tables:
                # 清洗表格数据的空白
                table_rows = []
                for row in table:
                    cells_text = [ (cell.strip() if cell else "") for cell in row ]
                    table_rows.append(cells_text)
                tables_data.append(table_rows)
        
        # 提取图片
        image_list = page.get_images(full=True)
        # get_images 返回一个列表，每个元素包含图片信息，例如xref等
        for img in image_list:
            xref = img[0]  # 第一个元素是 XREF id
            base_image = pdf_document.extract_image(xref)
            if base_image:
                image_bytes = base_image["image"]
                img_ext = base_image["ext"]  # 图片扩展名，如 png, jpg
                img_name = f"{os.path.basename(pdf_path)}_page{page_number+1}_img{xref}.{img_ext}"
                img_path = os.path.join(img_dir, img_name)
                # 将图片字节保存为文件
                with open(img_path, "wb") as f:
                    f.write(image_bytes)
                # 如果不是PNG格式，使用PIL转换为PNG（可选）
                if img_ext.lower() != "png":
                    try:
                        im = Image.open(img_path)
                        png_path = os.path.splitext(img_path)[0] + ".png"
                        im.save(png_path, format="PNG")
                        os.remove(img_path)
                        img_path = png_path
                    except Exception as e:
                        print(f"PDF图片转换PNG失败: {e}")
                # 添加图片的相对路径用于Markdown引用
                image_files.append(os.path.join("images", os.path.basename(img_path)))

    # 关闭 PDF 文件对象
    pdf_document.close()
    pdf_tables.close()
    return text_runs, tables_data, image_files
