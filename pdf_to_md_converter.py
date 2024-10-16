import fitz
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
from collections import Counter
import os
import io
from PIL import Image

def convert_pdf_to_md(pdf_path, md_path, progress_callback):
    doc = fitz.open(pdf_path)
    md_content = ""
    total_pages = len(doc)
    
    # 创建一个文件夹来保存图片，使用相对路径
    image_folder_name = os.path.splitext(os.path.basename(md_path))[0] + "_images"
    image_folder = os.path.join(os.path.dirname(md_path), image_folder_name)
    os.makedirs(image_folder, exist_ok=True)
    
    # 首先分析整个文档的字体大小和粗细
    font_sizes = []
    font_weights = []
    for page in doc:
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if block["type"] == 0:  # 文本块
                for line in block["lines"]:
                    for span in line["spans"]:
                        font_sizes.append(span["size"])
                        font_weights.append(span["flags"] & 2**4 != 0)  # 检查是否为粗体

    # 计算常见的字体大小和粗细
    common_size = Counter(font_sizes).most_common(1)[0][0]
    common_weight = Counter(font_weights).most_common(1)[0][0]
    
    print(f"Common font size: {common_size}, Common font weight: {common_weight}")
    
    # 现在处理文档内容
    for page_num, page in enumerate(doc):
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if block["type"] == 0:  # 文本块
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"].strip()
                        font_size = span["size"]
                        is_bold = span["flags"] & 2**4 != 0
                        
                        if (font_size > common_size * 1.1) or (is_bold and not common_weight):
                            heading_level = min(3, max(1, int((font_size - common_size) / 2) + 1))
                            md_content += f"{'#' * heading_level} {text}\n\n"
                        else:
                            md_content += f"{text} "
                md_content += "\n"
            elif block["type"] == 1:  # 图片块
                image_index = block["number"]
                image_rect = fitz.Rect(block["bbox"])
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=image_rect)
                img = Image.open(io.BytesIO(pix.tobytes()))
                img_filename = f"image_page{page_num+1}_{image_index}.png"
                img_path = os.path.join(image_folder, img_filename)
                img.save(img_path)
                # 使用相对路径在Markdown中引用图片
                md_content += f"![Image]({image_folder_name}/{img_filename})\n\n"
        
        # 更新进度
        progress = (page_num + 1) / total_pages * 100
        progress_callback(progress)
    
    # 清理多余的空白行和空格
    md_content = re.sub(r'\n{3,}', '\n\n', md_content)
    md_content = re.sub(r' +', ' ', md_content)
    
    with open(md_path, "w", encoding="utf-8") as md_file:
        md_file.write(md_content)

class PDFtoMDConverter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF to Markdown Converter")
        self.geometry("400x250")
        
        self.create_widgets()
    
    def create_widgets(self):
        self.select_button = tk.Button(self, text="选择PDF文件", command=self.select_pdf)
        self.select_button.pack(pady=20)
        
        self.convert_button = tk.Button(self, text="转换为Markdown", command=self.convert_to_md, state=tk.DISABLED)
        self.convert_button.pack(pady=20)
        
        self.progress_bar = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=20)
        
        self.status_label = tk.Label(self, text="")
        self.status_label.pack(pady=20)
    
    def select_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.pdf_path:
            self.convert_button["state"] = tk.NORMAL
            self.status_label["text"] = f"已选择: {self.pdf_path}"
    
    def update_progress(self, value):
        self.progress_bar["value"] = value
        self.update_idletasks()
    
    def convert_to_md(self):
        if hasattr(self, 'pdf_path'):
            md_path = filedialog.asksaveasfilename(defaultextension=".md", filetypes=[("Markdown Files", "*.md")])
            if md_path:
                try:
                    self.progress_bar["value"] = 0
                    convert_pdf_to_md(self.pdf_path, md_path, self.update_progress)
                    messagebox.showinfo("成功", "PDF已成功转换为Markdown!")
                    self.status_label["text"] = f"转换完成: {md_path}"
                except Exception as e:
                    messagebox.showerror("错误", f"转换过程中出现错误: {str(e)}")
        else:
            messagebox.showerror("错误", "请先选择一个PDF文件")

if __name__ == "__main__":
    app = PDFtoMDConverter()
    app.mainloop()
