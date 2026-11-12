import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client

# Biến toàn cục để lưu thư mục tạm thời trong phiên làm việc hiện tại
current_output_folder = ""

# --- CHỨC NĂNG CHÍNH ---
def select_output_folder():
    global current_output_folder
    folder = filedialog.askdirectory(title="Chọn thư mục lưu PDF mặc định")
    if folder:
        current_output_folder = folder
        lbl_output.config(text=f"Đang lưu tại: {folder}")

def start_convert():
    global current_output_folder
    
    # Bắt buộc phải có điểm lưu mới cho chạy
    if not current_output_folder:
        messagebox.showwarning("Cảnh báo", "Bạn chưa chọn điểm lưu! Vui lòng bấm Nút 1 để thiết lập trước.")
        return

    # Chọn file
    input_files = filedialog.askopenfilenames(
        title="Chọn các file Word / PowerPoint cần chuyển",
        filetypes=[
            ("Tất cả hỗ trợ", "*.doc *.docx *.ppt *.pptx"),
            ("Word Documents", "*.doc *.docx"),
            ("PowerPoint", "*.ppt *.pptx")
        ]
    )
    
    if not input_files:
        return

    # Chuẩn bị giao diện
    root.config(cursor="wait")
    btn_start.config(state=tk.DISABLED)
    btn_folder.config(state=tk.DISABLED)
    
    total_files = len(input_files)
    progress_bar['maximum'] = total_files
    progress_bar['value'] = 0
    
    word_app = None
    ppt_app = None
    
    try:
        for index, file in enumerate(input_files, 1):
            abs_input_path = os.path.abspath(file)
            filename = os.path.basename(abs_input_path)
            file_ext = os.path.splitext(filename)[1].lower()
            pdf_name = os.path.splitext(filename)[0] + ".pdf"
            abs_output_path = os.path.abspath(os.path.join(current_output_folder, pdf_name))
            
            # Cập nhật UI tiến trình
            lbl_progress.config(text=f"Đang xử lý ({index}/{total_files}): {filename}")
            root.update_idletasks()
            
            # --- XỬ LÝ ĐUÔI WORD ---
            if file_ext in ['.doc', '.docx']:
                if word_app is None:
                    word_app = win32com.client.Dispatch("Word.Application")
                    word_app.Visible = False
                    
                doc = word_app.Documents.Open(abs_input_path)
                doc.SaveAs(abs_output_path, FileFormat=17) 
                doc.Close()
                
            # --- XỬ LÝ ĐUÔI POWERPOINT ---
            elif file_ext in ['.ppt', '.pptx']:
                if ppt_app is None:
                    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                    
                presentation = ppt_app.Presentations.Open(abs_input_path, ReadOnly=1, Untitled=0, WithWindow=0)
                presentation.SaveAs(abs_output_path, 32)
                presentation.Close()
            
            # Cập nhật thanh tiến trình
            progress_bar['value'] = index
            root.update_idletasks()
            
        lbl_progress.config(text="Hoàn tất!")
        messagebox.showinfo("Thành công", f"Đã chuyển đổi thành công {total_files} file!\nLưu tại: {current_output_folder}")
        
    except Exception as e:
        lbl_progress.config(text="Xảy ra lỗi, tiến trình dừng lại.")
        error_log = traceback.format_exc()
        messagebox.showerror("Lỗi hệ thống", f"CHI TIẾT LỖI:\n\n{error_log}")
        
    finally:
        # Dọn dẹp tiến trình ngầm để giải phóng RAM
        if word_app is not None:
            word_app.Quit()
        if ppt_app is not None:
            ppt_app.Quit()
            
        # Khôi phục giao diện
        root.config(cursor="")
        btn_start.config(state=tk.NORMAL)
        btn_folder.config(state=tk.NORMAL)

# --- KHỞI TẠO GIAO DIỆN TKINTER ---
root = tk.Tk()
root.title("Tool Convert Office -> PDF")
root.geometry("480x270")
root.resizable(False, False)

tk.Label(root, text="CHUYỂN WORD / POWERPOINT SANG PDF", font=("Segoe UI", 12, "bold"), fg="#2b579a").pack(pady=10)

lbl_output = tk.Label(root, text="Chưa thiết lập điểm lưu (Yêu cầu chọn)", font=("Segoe UI", 9, "italic"), fg="#d9534f")
lbl_output.pack(pady=2)

btn_folder = tk.Button(root, text="1. Thiết lập thư mục lưu PDF (Bắt buộc)", command=select_output_folder, width=45, font=("Segoe UI", 9))
btn_folder.pack(pady=8)

btn_start = tk.Button(root, text="2. Chọn File Cần Chuyển & Bắt Đầu", command=start_convert, width=40, bg="#2b579a", fg="white", font=("Segoe UI", 10, "bold"))
btn_start.pack(pady=5)

lbl_progress = tk.Label(root, text="", font=("Segoe UI", 8, "italic"), fg="gray")
lbl_progress.pack(pady=2)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.pack(pady=5)

root.mainloop()