import os
import io
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import fitz  # PyMuPDF để đọc PDF

# ==========================================================
# 🧠 BIẾN TOÀN CỤC
# ==========================================================
rename_history = []      # Lưu lịch sử rename để hoàn tác
selection_order = []     # Lưu thứ tự click chọn file
select_all_mode = False  # Đánh dấu đang chọn tất cả

# ==========================================================
# 🔢 Natural Sort Helper
# ==========================================================
def natural_sort_key(s):
    """Tách chữ và số để sắp xếp tự nhiên: page_2 < page_10"""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('(\d+)', s)]

# ==========================================================
# 📁 CHỨC NĂNG CHỌN THƯ MỤC
# ==========================================================
def choose_folder():
    folder_selected = filedialog.askdirectory(title="※ PDFフォルダを選択してください")
    if folder_selected:
        folder_path_var.set(folder_selected)
        list_pdfs(folder_selected)

def list_pdfs(folder_path):
    listbox.delete(0, tk.END)
    if not os.path.isdir(folder_path):
        return

    pdf_files = [entry.name for entry in os.scandir(folder_path)
                 if entry.is_file() and entry.name.lower().endswith(".pdf")]

    # Natural sort
    pdf_files.sort(key=natural_sort_key)

    for f in pdf_files:
        listbox.insert(tk.END, f)

    update_selected_count()

# ==========================================================
# ✏️ ĐỔI TÊN FILE
# ==========================================================
def rename_pdfs():
    folder_path = folder_path_var.get()
    part1 = part1_var.get().strip()
    part3 = part3_var.get().strip()
    part4 = part4_var.get().strip()

    try:
        start_number = int(start_var.get())
    except ValueError:
        messagebox.showerror("エラー", "開始番号は数字で入力してください！")
        return

    digits = 6

    if not os.path.isdir(folder_path):
        messagebox.showerror("エラー", "有効なフォルダを選択してください！")
        return

    global rename_history, selection_order, select_all_mode
    selected_indices = listbox.curselection()
    if not selected_indices:
        messagebox.showinfo("通知", "少なくとも1つのPDFファイルを選択してください。")
        return

    # Dựa trên chế độ select_all hoặc chọn từng file
    if select_all_mode:
        ordered_indices = list(selected_indices)
    else:
        ordered_indices = [i for i in selection_order if i in selected_indices]

    rename_history.clear()

    for offset, idx in enumerate(ordered_indices):
        filename = listbox.get(idx)
        old_path = os.path.join(folder_path, filename)
        number_str = str(start_number + offset).zfill(digits)
        new_name = f"{part1}-{number_str}-{part3}-{part4}.pdf"
        new_path = os.path.join(folder_path, new_name)

        if os.path.exists(new_path):
            messagebox.showerror("エラー", f"{new_name} は既に存在します！")
            return

        os.rename(old_path, new_path)
        rename_history.append((new_path, old_path))
        listbox.delete(idx)
        listbox.insert(idx, new_name)

    start_var.set(str(start_number + len(ordered_indices)))
    listbox.selection_clear(0, tk.END)
    update_selected_count()
    output_label.config(text=f"※ {len(ordered_indices)} 個のPDFファイルの名前を変更しました！", fg="#1565c0")
    selection_order.clear()
    select_all_mode = False

# ==========================================================
# ↩️ HOÀN TÁC (UNDO)
# ==========================================================
def undo_rename():
    global rename_history
    if not rename_history:
        messagebox.showinfo("通知", "元に戻す操作がありません。")
        return

    failed = []
    for new_path, old_path in rename_history:
        if os.path.exists(new_path):
            os.rename(new_path, old_path)
        else:
            failed.append(os.path.basename(new_path))

    list_pdfs(folder_path_var.get())

    if failed:
        messagebox.showwarning("警告", f"一部のファイルを元に戻せませんでした: {failed}")
    else:
        messagebox.showinfo("完了", "前回の変更を元に戻しました。")

    rename_history.clear()
    output_label.config(text="※ 前回の名前変更を元に戻しました。", fg="#00796b")

# ==========================================================
# 🖱️ XỬ LÝ CHỌN FILE
# ==========================================================
def on_select(event=None):
    global selection_order, select_all_mode
    if select_all_mode:
        update_selected_count()
        return
    current_selection = listbox.curselection()
    new_selection = [i for i in current_selection if i not in selection_order]
    removed = [i for i in selection_order if i not in current_selection]
    for r in removed:
        selection_order.remove(r)
    selection_order.extend(new_selection)
    update_selected_count()

def select_all():
    global selection_order, select_all_mode
    select_all_mode = True
    listbox.select_set(0, tk.END)
    selection_order = list(range(listbox.size()))
    update_selected_count()

def deselect_all():
    global selection_order, select_all_mode
    select_all_mode = False
    selection_order.clear()
    listbox.selection_clear(0, tk.END)
    update_selected_count()

def update_selected_count(event=None):
    count = len(listbox.curselection())
    selected_count_var.set(str(count))

# ==========================================================
# 👁️ XEM TRƯỚC PDF KHI HOVER
# ==========================================================
def get_pdf_preview_image(pdf_path, width=900):
    """Trả về ảnh trang đầu tiên của PDF dưới dạng PhotoImage"""
    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        w_percent = width / float(img.size[0])
        h_size = int(float(img.size[1]) * w_percent)
        img = img.resize((width, h_size))
        return ImageTk.PhotoImage(img)
    except Exception:
        return None

class PDFPreviewer:
    def __init__(self, listbox, folder_var):
        self.listbox = listbox
        self.folder_var = folder_var
        self.popup = None
        self.preview_img = None
        listbox.bind("<Motion>", self.on_hover)
        listbox.bind("<Leave>", self.hide_preview)

    def on_hover(self, event):
        index = self.listbox.nearest(event.y)
        if index < 0:
            return
        filename = self.listbox.get(index)
        pdf_path = os.path.join(self.folder_var.get(), filename)
        if not os.path.isfile(pdf_path) or not pdf_path.lower().endswith(".pdf"):
            return

        img = get_pdf_preview_image(pdf_path)
        if not img:
            self.hide_preview()
            return

        if self.popup:
            self.popup.destroy()

        self.popup = tk.Toplevel()
        self.popup.overrideredirect(True)
        self.popup.geometry(f"+{event.x_root + 20}+{event.y_root + 10}")
        label = tk.Label(self.popup, image=img, bg="white", borderwidth=2, relief="solid")
        label.pack()
        self.preview_img = img

    def hide_preview(self, event=None):
        if self.popup:
            self.popup.destroy()
            self.popup = None

# ==========================================================
# 🎨 GIAO DIỆN CHÍNH
# ==========================================================
root = tk.Tk()
root.title("※ PDF 名前変更ツール + プレビュー機能")
root.geometry("700x780")
root.resizable(False, False)
root.configure(bg="#cce6ff")

folder_path_var = tk.StringVar()
part1_var = tk.StringVar(value="3301")
start_var = tk.StringVar(value="100011")
part3_var = tk.StringVar(value="A0")
part4_var = tk.StringVar(value="N0")
selected_count_var = tk.StringVar(value="0")

# --- Folder chọn ---
tk.Label(root, text="※ PDFフォルダを選択:", bg="#cce6ff", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 0))
folder_frame = tk.Frame(root, bg="#cce6ff")
folder_frame.pack(fill="x", padx=10)
tk.Entry(folder_frame, textvariable=folder_path_var, font=("Arial", 10)).pack(side="left", fill="x", expand=True, pady=5)
tk.Button(folder_frame, text="※ 選択...", bg="#4da6ff", fg="white", font=("Arial", 10, "bold"), command=choose_folder).pack(side="right", padx=5, pady=5)

# --- Nút chọn tất cả / bỏ chọn ---
button_frame = tk.Frame(root, bg="#cce6ff")
button_frame.pack(fill="x", padx=10)
tk.Button(button_frame, text="※ すべて選択", bg="#4da6ff", fg="white", font=("Arial", 10, "bold"), command=select_all).pack(side="left", fill="x", expand=True, padx=5)
tk.Button(button_frame, text="※ 選択解除", bg="#4da6ff", fg="white", font=("Arial", 10, "bold"), command=deselect_all).pack(side="left", fill="x", expand=True, padx=5)

# --- Danh sách PDF ---
tk.Label(root, text="※ PDFファイル一覧（複数選択可）:", bg="#cce6ff", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 0))
list_frame = tk.Frame(root)
list_frame.pack(fill="both", padx=10)
scrollbar = tk.Scrollbar(list_frame)
scrollbar.pack(side="right", fill="y")
listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, height=12, yscrollcommand=scrollbar.set, bg="#e6f2ff", font=("Arial", 10))
listbox.pack(side="left", fill="both", expand=True)
scrollbar.config(command=listbox.yview)
listbox.bind('<<ListboxSelect>>', on_select)

# --- Preview PDF khi hover ---
previewer = PDFPreviewer(listbox, folder_path_var)

# --- Số lượng file chọn ---
selected_count_frame = tk.Frame(root, bg="#cce6ff")
selected_count_frame.pack(anchor="w", fill="x", padx=10, pady=(5, 2))
tk.Label(selected_count_frame, text="※ 選択中のファイル数: ", bg="#cce6ff", font=("Arial", 10, "bold")).pack(side="left")
tk.Label(selected_count_frame, textvariable=selected_count_var, bg="#cce6ff", fg="#d32f2f", font=("Arial", 16, "bold")).pack(side="left")

# --- Hướng dẫn ---
tk.Label(root, text="💡 個別に選択した場合、選択順が名前変更の順番になります。", bg="#cce6ff", fg="#004d99", font=("Arial", 10, "italic")).pack(anchor="w", padx=20, pady=(0, 8))

# --- Cấu hình tên ---
input_frame = tk.LabelFrame(root, text="※ 名前フォーマット設定", bg="#cce6ff", font=("Arial", 10, "bold"), padx=10, pady=10)
input_frame.pack(fill="x", padx=10, pady=(5, 10))
for label_text, var in [("※ パート1（例: 3301）:", part1_var),
                        ("※ 開始番号（例: 100011）:", start_var),
                        ("※ パート3（例: A0）:", part3_var),
                        ("※ パート4（例: N0）:", part4_var)]:
    tk.Label(input_frame, text=label_text, bg="#cce6ff", font=("Arial", 10)).pack(anchor="w", pady=2)
    tk.Entry(input_frame, textvariable=var, font=("Arial", 10)).pack(fill="x", pady=2)

# --- Nút rename + Undo ---
action_frame = tk.Frame(root, bg="#cce6ff")
action_frame.pack(fill="x", padx=10, pady=10)
tk.Button(action_frame, text="※ 選択したPDFのみ名前を変更", bg="#3399ff", fg="white", font=("Arial", 12, "bold"), command=rename_pdfs).pack(side="left", fill="x", expand=True, padx=5)
tk.Button(action_frame, text="※ 元に戻す (Undo)", bg="#ff704d", fg="white", font=("Arial", 12, "bold"), command=undo_rename).pack(side="left", fill="x", expand=True, padx=5)

# --- Nhãn kết quả ---
output_label = tk.Label(root, text="", bg="#cce6ff", font=("Arial", 11, "bold"))
output_label.pack(pady=(0, 10))

root.mainloop()
