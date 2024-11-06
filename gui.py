import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from sn_processor import SNProcessor

class SNInputGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("华为SN码批量查询系统")
        self.root.geometry("600x500")
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Instructions
        ttk.Label(self.main_frame, text="请输入SN码 (每行一个)", font=('Arial', 10, 'bold')).grid(row=0, column=0, pady=5, sticky=tk.W)
        
        # Text area for SN codes
        self.sn_text = scrolledtext.ScrolledText(self.main_frame, width=50, height=15)
        self.sn_text.grid(row=1, column=0, columnspan=2, pady=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(self.main_frame, length=400, mode='determinate')
        self.progress.grid(row=2, column=0, columnspan=2, pady=10)
        
        # Status label
        self.status_label = ttk.Label(self.main_frame, text="就绪")
        self.status_label.grid(row=3, column=0, columnspan=2, pady=5)
        
        # Buttons frame
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        # Submit button
        self.submit_btn = ttk.Button(button_frame, text="提交查询", command=self.process_input)
        self.submit_btn.pack(side=tk.LEFT, padx=5)
        
        # Clear button
        self.clear_btn = ttk.Button(button_frame, text="清除输入", command=self.clear_input)
        self.clear_btn.pack(side=tk.LEFT, padx=5)

    def process_input(self):
        """Process the input SN codes"""
        # Get input text and split into lines
        sn_text = self.sn_text.get("1.0", tk.END).strip()
        if not sn_text:
            messagebox.showwarning("警告", "请输入SN码")
            return
            
        sn_codes = [code.strip() for code in sn_text.split('\n') if code.strip()]
        
        if not sn_codes:
            messagebox.showwarning("警告", "没有有效的SN码")
            return
            
        # Disable buttons during processing
        self.submit_btn.state(['disabled'])
        self.clear_btn.state(['disabled'])
        
        try:
            # Initialize processor
            processor = SNProcessor(self.update_progress, self.update_status)
            
            # Process SN codes
            self.status_label.config(text="正在处理...")
            processor.process_sn_codes(sn_codes)
            
            messagebox.showinfo("完成", "处理完成！结果已导出到 query_results.xlsx")
            
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出现错误：{str(e)}")
            
        finally:
            # Re-enable buttons
            self.submit_btn.state(['!disabled'])
            self.clear_btn.state(['!disabled'])
            self.progress['value'] = 0
            self.status_label.config(text="就绪")

    def clear_input(self):
        """Clear the input text area"""
        self.sn_text.delete("1.0", tk.END)
        self.progress['value'] = 0
        self.status_label.config(text="就绪")

    def update_progress(self, value):
        """Update progress bar"""
        self.progress['value'] = value
        self.root.update_idletasks()

    def update_status(self, text):
        """Update status label"""
        self.status_label.config(text=text)
        self.root.update_idletasks()

def main():
    root = tk.Tk()
    app = SNInputGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()