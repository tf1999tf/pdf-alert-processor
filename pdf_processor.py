import os
import re
import glob
import time
import threading
from datetime import datetime
from pdfplumber import open as pdf_open
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import sys

class PDFProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("贵阳龙洞堡机场PDF警报处理器")
        self.root.geometry("800x600")
        
        # 设置图标（如果有的话）
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        # 处理器实例
        self.processor = None
        self.monitoring = False
        self.monitor_thread = None
        
        # 创建界面
        self.setup_ui()
        
        # 自动初始化处理器
        self.initialize_processor()
        
    def setup_ui(self):
        """设置用户界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 设置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, 
                                text="贵阳龙洞堡机场PDF警报处理器", 
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 文件夹设置区域
        folder_frame = ttk.LabelFrame(main_frame, text="文件夹设置", padding="10")
        folder_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        folder_frame.columnconfigure(1, weight=1)
        
        # PDF文件夹
        ttk.Label(folder_frame, text="PDF文件夹:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.pdf_folder_var = tk.StringVar()
        self.pdf_folder_entry = ttk.Entry(folder_frame, textvariable=self.pdf_folder_var, width=50)
        self.pdf_folder_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)
        ttk.Button(folder_frame, text="浏览", command=self.browse_pdf_folder).grid(row=0, column=2, padx=(0, 0))
        
        # TXT文件夹
        ttk.Label(folder_frame, text="TXT文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.txt_folder_var = tk.StringVar()
        self.txt_folder_entry = ttk.Entry(folder_frame, textvariable=self.txt_folder_var, width=50)
        self.txt_folder_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)
        ttk.Button(folder_frame, text="浏览", command=self.browse_txt_folder).grid(row=1, column=2, padx=(0, 0))
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=(10, 10))
        
        # 处理按钮
        ttk.Button(button_frame, text="处理所有文件", 
                  command=self.process_all).pack(side=tk.LEFT, padx=5)
        
        # 开始监控按钮
        self.monitor_button = ttk.Button(button_frame, text="开始监控", 
                                        command=self.toggle_monitoring)
        self.monitor_button.pack(side=tk.LEFT, padx=5)
        
        # 停止按钮
        ttk.Button(button_frame, text="停止所有", 
                  command=self.stop_all).pack(side=tk.LEFT, padx=5)
        
        # 打开文件夹按钮
        ttk.Button(button_frame, text="打开PDF文件夹", 
                  command=self.open_pdf_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="打开TXT文件夹", 
                  command=self.open_txt_folder).pack(side=tk.LEFT, padx=5)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding="10")
        log_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # 日志文本框
        self.log_text = scrolledtext.ScrolledText(log_frame, width=80, height=15)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 添加右键菜单到日志框
        self.setup_context_menu()
        
    def setup_context_menu(self):
        """为日志文本框添加上下文菜单"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="复制", command=self.copy_log)
        self.context_menu.add_command(label="清空日志", command=self.clear_log)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="保存日志", command=self.save_log)
        
        # 绑定右键事件
        self.log_text.bind("<Button-3>", self.show_context_menu)
        
    def show_context_menu(self, event):
        """显示上下文菜单"""
        self.context_menu.post(event.x_root, event.y_root)
        
    def copy_log(self):
        """复制日志内容"""
        self.root.clipboard_clear()
        self.root.clipboard_append(self.log_text.get("1.0", tk.END))
        
    def clear_log(self):
        """清空日志"""
        self.log_text.delete("1.0", tk.END)
        
    def save_log(self):
        """保存日志到文件"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(self.log_text.get("1.0", tk.END))
            self.log_message(f"日志已保存到: {file_path}")
            
    def initialize_processor(self):
        """初始化处理器"""
        try:
            # 设置默认文件夹路径
            base_path = r"E:\代码\警报处理"
            
            # 如果程序被打包成exe，使用exe所在目录
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            
            pdf_folder = os.path.join(base_path, "PDF")
            txt_folder = os.path.join(base_path, "TXT")
            
            # 更新UI中的文件夹路径
            self.pdf_folder_var.set(pdf_folder)
            self.txt_folder_var.set(txt_folder)
            
            # 创建处理器实例
            self.processor = PDFProcessor(pdf_folder, txt_folder, self.log_message)
            
            self.log_message("程序初始化完成")
            self.update_status("就绪")
            
        except Exception as e:
            self.log_message(f"初始化失败: {str(e)}", "error")
            self.update_status("初始化失败")
            
    def browse_pdf_folder(self):
        """浏览选择PDF文件夹"""
        folder = filedialog.askdirectory(title="选择PDF文件夹")
        if folder:
            self.pdf_folder_var.set(folder)
            self.update_processor_folders()
            
    def browse_txt_folder(self):
        """浏览选择TXT文件夹"""
        folder = filedialog.askdirectory(title="选择TXT文件夹")
        if folder:
            self.txt_folder_var.set(folder)
            self.update_processor_folders()
            
    def update_processor_folders(self):
        """更新处理器的文件夹设置"""
        if self.processor:
            self.processor.pdf_folder = self.pdf_folder_var.get()
            self.processor.txt_folder = self.txt_folder_var.get()
            self.log_message(f"文件夹设置已更新")
            
    def process_all(self):
        """处理所有文件"""
        if not self.processor:
            messagebox.showerror("错误", "处理器未初始化")
            return
            
        def process():
            try:
                self.update_status("正在处理所有文件...")
                self.processor.process_all()
                self.update_status("处理完成")
            except Exception as e:
                self.log_message(f"处理失败: {str(e)}", "error")
                self.update_status("处理失败")
                
        # 在新线程中处理，避免界面卡顿
        thread = threading.Thread(target=process, daemon=True)
        thread.start()
        
    def toggle_monitoring(self):
        """切换监控状态"""
        if not self.processor:
            messagebox.showerror("错误", "处理器未初始化")
            return
            
        if not self.monitoring:
            # 开始监控
            self.monitoring = True
            self.monitor_button.config(text="停止监控")
            self.update_status("开始监控...")
            
            # 在新线程中启动监控
            self.monitor_thread = threading.Thread(target=self.start_monitoring, daemon=True)
            self.monitor_thread.start()
        else:
            # 停止监控
            self.stop_monitoring()
            
    def start_monitoring(self):
        """开始监控"""
        try:
            self.processor.start_monitoring(self.update_status_callback)
        except Exception as e:
            self.log_message(f"监控出错: {str(e)}", "error")
            self.stop_monitoring()
            
    def stop_monitoring(self):
        """停止监控"""
        if self.processor:
            self.processor.stop_monitoring()
        self.monitoring = False
        self.monitor_button.config(text="开始监控")
        self.update_status("监控已停止")
        
    def update_status_callback(self, message):
        """监控状态回调函数"""
        self.update_status(message)
        
    def stop_all(self):
        """停止所有操作"""
        self.stop_monitoring()
        self.update_status("已停止所有操作")
        
    def open_pdf_folder(self):
        """打开PDF文件夹"""
        folder = self.pdf_folder_var.get()
        if os.path.exists(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("警告", f"文件夹不存在: {folder}")
            
    def open_txt_folder(self):
        """打开TXT文件夹"""
        folder = self.txt_folder_var.get()
        if os.path.exists(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("警告", f"文件夹不存在: {folder}")
            
    def log_message(self, message, level="info"):
        """记录日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # 根据级别设置颜色
        if level == "error":
            tag = "error"
            prefix = "[错误]"
        elif level == "warning":
            tag = "warning"
            prefix = "[警告]"
        else:
            tag = "info"
            prefix = "[信息]"
            
        # 插入日志
        self.log_text.insert(tk.END, f"{timestamp} {prefix} {message}\n", tag)
        
        # 自动滚动到底部
        self.log_text.see(tk.END)
        
        # 更新状态栏
        self.update_status(message)
        
    def update_status(self, message):
        """更新状态栏"""
        self.status_var.set(message)
        self.root.update_idletasks()
        
    def on_closing(self):
        """窗口关闭时的处理"""
        self.stop_all()
        self.root.destroy()

class PDFProcessor:
    def __init__(self, pdf_folder, txt_folder, log_callback=None):
        """
        初始化PDF处理器
        
        Args:
            pdf_folder: PDF文件夹路径
            txt_folder: TXT输出文件夹路径
            log_callback: 日志回调函数
        """
        self.pdf_folder = pdf_folder
        self.txt_folder = txt_folder
        self.log_callback = log_callback
        self.stop_flag = False
        
        # 确保文件夹存在
        if not os.path.exists(self.pdf_folder):
            os.makedirs(self.pdf_folder)
            self.log(f"已创建PDF文件夹: {self.pdf_folder}")
            
        if not os.path.exists(self.txt_folder):
            os.makedirs(self.txt_folder)
            self.log(f"已创建TXT文件夹: {self.txt_folder}")
            
        # 记录已处理的文件
        self.processed_files = set()
        
    def log(self, message, level="info"):
        """记录日志"""
        if self.log_callback:
            self.log_callback(message, level)
            
    def extract_time_from_filename(self, filename):
        """从PDF文件名中提取时间信息"""
        pattern = r'9_FCST_C_ZUGY_(\d{14})'
        match = re.search(pattern, filename)
        return match.group(1) if match else None
    
    def format_timestamp(self, timestamp_str):
        """格式化时间戳"""
        if not timestamp_str or len(timestamp_str) != 14:
            return "unknown_time"
        
        try:
            dt = datetime.strptime(timestamp_str, "%Y%m%d%H%M%S")
            return dt.strftime("%Y%m%d_%H%M%S")
        except:
            return "unknown_time"
    
    def extract_content_from_pdf(self, pdf_path):
        """从PDF文件中提取内容"""
        try:
            with pdf_open(pdf_path) as pdf:
                all_text = ""
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        all_text += page_text + "\n"
            
            # 处理文本
            formatted_text, warning_number = self.process_full_text(all_text)
            return warning_number, formatted_text
            
        except Exception as e:
            self.log(f"提取PDF文件时出错: {str(e)}", "error")
            return None, ""
    
    def process_full_text(self, text):
        """处理整个文本"""
        result = []
        warning_number = "01"
        
        # 1. 提取标题
        title_match = re.search(r'贵阳龙洞堡机场天气警报', text)
        if title_match:
            result.append(title_match.group(0))
        
        # 2. 提取气象台信息
        weather_bureau_match = re.search(r'贵阳龙洞堡机场气象台', text)
        if weather_bureau_match:
            result.append(weather_bureau_match.group(0))
        
        # 3. 提取预警发布序号
        warning_no_match = re.search(r'预警发布序号[:：]\s*(\d+)', text)
        if warning_no_match:
            warning_number = warning_no_match.group(1)
            result.append(f"预警发布序号：{warning_number}")
        
        # 4. 提取发布时间
        release_time_match = re.search(r'发布时间[:：]\s*([^。\n\r]+?北京时[）\)])', text)
        if release_time_match:
            result.append(f"发布时间：{release_time_match.group(1)}")
        else:
            release_time_match2 = re.search(r'发布时间[:：]\s*([0-9]{4}-[0-9]{2}-[0-9]{2}\s+[0-9]{2}:[0-9]{2}\s*（北京时）)', text)
            if release_time_match2:
                result.append(f"发布时间：{release_time_match2.group(1)}")
        
        # 5. 提取正文内容
        content = self.extract_main_content(text)
        
        if content:
            # 删除所有空格、换行符和回车符
            content = content.replace(' ', '').replace('\n', '').replace('\r', '')
            
            # 确保以句号结尾
            if not content.endswith('。') and not content.endswith('.'):
                content += '。'
            elif content.endswith('.'):
                content = content[:-1] + '。'
                
            result.append(content)
        
        return '\n'.join(result), warning_number
    
    def extract_main_content(self, text):
        """提取正文主体部分"""
        pattern = r'发布时间[:：][^。\n\r]+?北京时[）\)](.*?)发布人'
        match = re.search(pattern, text, re.DOTALL)
        
        if match:
            return match.group(1).strip()
        
        release_time_pattern = r'发布时间[:：][^。\n\r]+?北京时[）\)]'
        release_match = re.search(release_time_pattern, text)
        
        if release_match:
            content = text[release_match.end():].strip()
            content = re.sub(r'电话[：:].*', '', content)
            content = re.sub(r'传真[：:].*', '', content)
            return content
        
        return ""
    
    def process_pdf(self, pdf_path):
        """处理单个PDF文件"""
        try:
            warning_number, formatted_text = self.extract_content_from_pdf(pdf_path)
            
            if not formatted_text:
                self.log(f"无法从 {os.path.basename(pdf_path)} 提取内容", "warning")
                return False
            
            # 生成输出文件名
            filename = os.path.basename(pdf_path)
            timestamp_str = self.extract_time_from_filename(filename)
            formatted_time = self.format_timestamp(timestamp_str)
            output_filename = f"ZUGY_{formatted_time}_{warning_number}.txt"
            
            output_path = os.path.join(self.txt_folder, output_filename)
            
            # 写入TXT文件
            with open(output_path, 'w', encoding='utf-8') as txt_file:
                txt_file.write(formatted_text)
            
            # 记录已处理的文件
            self.processed_files.add(os.path.basename(pdf_path))
            
            self.log(f"成功处理: {os.path.basename(pdf_path)} -> {output_filename}")
            return True
            
        except Exception as e:
            self.log(f"处理PDF文件 {os.path.basename(pdf_path)} 时出错: {str(e)}", "error")
            return False
    
    def get_all_pdfs(self):
        """获取所有PDF文件"""
        return glob.glob(os.path.join(self.pdf_folder, "*.pdf"))
    
    def process_all(self):
        """处理所有文件"""
        self.log("正在查找所有PDF文件...")
        all_pdfs = self.get_all_pdfs()
        
        if not all_pdfs:
            self.log("未找到PDF文件", "warning")
            return False
        
        self.log(f"找到 {len(all_pdfs)} 个PDF文件")
        
        success_count = 0
        for pdf_path in all_pdfs:
            if self.stop_flag:
                self.log("处理被用户中断")
                break
                
            pdf_name = os.path.basename(pdf_path)
            if pdf_name not in self.processed_files:
                if self.process_pdf(pdf_path):
                    success_count += 1
        
        self.log(f"处理完成: {success_count} 个文件处理成功")
        return success_count > 0
    
    def start_monitoring(self, status_callback=None):
        """开始监控文件夹"""
        self.stop_flag = False
        self.log(f"开始监控文件夹: {self.pdf_folder}")
        
        if status_callback:
            status_callback("开始监控...")
        
        # 初始检查并处理所有未处理文件
        self.process_all()
        
        try:
            while not self.stop_flag:
                # 获取当前所有PDF文件
                current_pdfs = self.get_all_pdfs()
                
                # 检查是否有新文件
                for pdf_path in current_pdfs:
                    if self.stop_flag:
                        break
                        
                    pdf_name = os.path.basename(pdf_path)
                    if pdf_name not in self.processed_files:
                        self.log(f"发现新文件: {pdf_name}")
                        self.process_pdf(pdf_path)
                
                # 等待下一次检查
                for _ in range(10):  # 每0.5秒检查一次停止标志
                    if self.stop_flag:
                        break
                    time.sleep(0.5)
                    
        except Exception as e:
            self.log(f"监控出错: {str(e)}", "error")
            
        self.log("监控已停止")
        
    def stop_monitoring(self):
        """停止监控"""
        self.stop_flag = True

def main():
    """主函数"""
    root = tk.Tk()
    app = PDFProcessorGUI(root)
    
    # 设置窗口关闭事件
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # 运行主循环
    root.mainloop()

if __name__ == "__main__":
    main()