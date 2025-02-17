import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import zipfile
import os
import shutil
from PIL import Image

class WordImageExtractorApp:
    def __init__(self, master):
        self.master = master
        master.title("Word图片提取工具 v1.0,by--drake")
        master.geometry("600x400")

        # 创建界面元素
        self.create_widgets()

    def create_widgets(self):
        # 文件选择区域
        tk.Label(self.master, text="Word文档路径:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.word_path_entry = tk.Entry(self.master, width=50)
        self.word_path_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.master, text="选择文件", command=self.select_word_file).grid(row=0, column=2, padx=5)

        # 保存路径区域
        tk.Label(self.master, text="保存路径:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.save_path_entry = tk.Entry(self.master, width=50)
        self.save_path_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(self.master, text="选择目录", command=self.select_save_dir).grid(row=1, column=2, padx=5)

        # 进度条
        self.progress = ttk.Progressbar(self.master, orient="horizontal", length=400, mode="determinate")
        self.progress.grid(row=2, column=0, columnspan=3, pady=10)

        # 日志区域
        self.log_text = tk.Text(self.master, height=10, width=70)
        self.log_text.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

        # 开始按钮
        tk.Button(self.master, text="开始提取", command=self.start_extraction,
                  bg="#4CAF50", fg="white").grid(row=4, column=1, pady=10)

    def select_word_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Word文档", "*.docx")])
        self.word_path_entry.delete(0, tk.END)
        self.word_path_entry.insert(0, file_path)

    def select_save_dir(self):
        dir_path = filedialog.askdirectory()
        self.save_path_entry.delete(0, tk.END)
        self.save_path_entry.insert(0, dir_path)

    def log_message(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.master.update()

    def start_extraction(self):
        try:
            word_path = self.word_path_entry.get()
            save_path = self.save_path_entry.get()

            if not word_path.endswith(".docx"):
                messagebox.showerror("错误", "请选择有效的Word文档(.docx)")
                return

            if not os.path.exists(save_path):
                os.makedirs(save_path)

            temp_dir = os.path.join(save_path, "temp_extract")
            self.progress["value"] = 20
            self.log_message("开始解压文档...")

            with zipfile.ZipFile(word_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            media_path = os.path.join(temp_dir, 'word/media')
            if os.path.exists(media_path):
                files = os.listdir(media_path)
                total = len(files)
                self.progress["value"] = 40
                self.log_message(f"发现{total}个图片文件")

                for idx, img_file in enumerate(files):
                    src = os.path.join(media_path, img_file)
                    dst = os.path.join(save_path, img_file)
                    shutil.copy(src, dst)

                    if img_file.endswith(".emf"):
                        self.log_message(f"转换EMF文件: {img_file}")
                        img = Image.open(dst)
                        png_path = dst.replace(".emf", ".png")
                        img.save(png_path, 'PNG')
                        os.remove(dst)  # 删除原始emf文件

                    self.progress["value"] = 40 + (idx + 1) / total * 50

            shutil.rmtree(temp_dir)
            self.progress["value"] = 100
            messagebox.showinfo("完成", "图片提取转换完成！")

        except Exception as e:
            messagebox.showerror("错误", f"处理过程中发生错误:\n{str(e)}")
            self.progress["value"] = 0
        finally:
            if 'temp_dir' in locals() and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
if __name__ == '__main__':
    root = tk.Tk()
    app = WordImageExtractorApp(root)
    root.mainloop()
