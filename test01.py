
import os
import pathlib
import pypdf
import shutil # フォルダ削除
import tkinter as tk
import subprocess # PDFファイルを開く
from functools import partial # ボタンクリック時固有のtextを取得
from pathlib import Path
from pdf2image import convert_from_path
from PIL import Image, ImageTk
from PyPDF2 import PdfMerger
from tkinter import messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
poppler_dir = "poppler/Library/bin"
os.environ["PATH"] += os.pathsep + str(poppler_dir)
acrobat_path = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe" # Acrobatアプリのパス
drop_PDF = None
btn = None
root = tk.Tk()
root.title("メインメニュー")
root.geometry("640x400+200+200")
selected_PDFFiles = [] # 選択されたPDFのファイル名を保持
# 複数選択したPDFを結合して表示
def btn_click():
    print(selected_PDFFiles)
    merged_PDF = PdfMerger()
    for file_name in selected_PDFFiles:
        merged_PDF.append(f"splited/{file_name}.pdf")
    merged_PDF.write("merged.pdf")
    merged_PDF.close()
    pdf_path = "merged.pdf"
    pdf_pro = subprocess.Popen([acrobat_path, pdf_path], shell=False)
def shift_click(_,event):
    print("Success")
def button_func(event,i):
    i.widget.config(bd=5)
    print(i.widget.config("text")[4])
    selected_PDFFiles.append(os.path.splitext(i.widget.config("text")[4])[0])
def test():
    global drop_PDF
    create_drop_PDF()
    drop_PDF.geometry("640x400")
    drop_PDF.title("Test-Window")
    label = tk.Label(drop_PDF,text="Test-Window")
    label.place(x=20, y=20)
    
def show_canvas():
    create_drop_PDF()
    global drop_PDF
    global canvas
    global btn
    scrollregion_y = 2000
    drop_PDF.geometry("850x500+300+100")
    # canvasの定義
    canvas = tk.Canvas(drop_PDF,  width=800, height=600
                       , scrollregion=(0,0,800,scrollregion_y)
                       )
    canvas.grid(row=10, column=0)
    # scrollbar設置
    scrollbar = tk.Scrollbar(drop_PDF, orient=tk.VERTICAL)
    scrollbar.grid(row=0, column=1, sticky=tk.N + tk.S)
    scrollbar.config(command=canvas.yview)
    canvas.config(yscrollcommand=scrollbar.set)
    frame = tk.Frame(canvas,width=800, height=2000)
    thumbs_path = "thumbs/"
    count = 0 # ボタンの座標を決めるための変数
    img_list = [] # ボタンに表示する画像を格納するリスト
    for i in Path(thumbs_path).glob("*/"):
        read_image = Image.open(f"{thumbs_path}{i.name}")
        img = ImageTk.PhotoImage(image=read_image,width=100, height=50,master=canvas)
        # partialを使用しボタンの固有のtextを取得
        btn = tk.Button(frame,image=img,relief="sunken", bd=2, text=i.name
                        # , command=partial(select_thumb, i.name)
                        )
        btn.place(x=20+150*int(count%5),y=20+210*int(count/5))
        btn.bind("<MouseWheel>", swipe_scroll)
        btn.bind("<Button-1>", partial(button_func, count)) # クリック
        btn.bind("<Shift-Button-1>", partial(shift_click, count)) # クリック
        btn.bind("<Double-Button-1>", partial(show_thumb, i.name)) # ダブルクリックでPDFを表示
        img_list.append(img)
        count += 1
    frame.place(x=20, y=20)
    canvas.create_window((0,0), window=frame, anchor="nw", width=800, height=600)
    
    # スワイプでスクロール
    frame.bind("<MouseWheel>", swipe_scroll)
    canvas.bind("<MouseWheel>", swipe_scroll)
    # mainloopがないと画像が表示されない。
    drop_PDF.mainloop()
def select_PDF(num, _):
    print(num)
# PDFを表示
def show_thumb(num, _):
    pdf_name = os.path.splitext(num)[0] # 拡張しなしのPDFファイル名
    
    pdf_path = "splited/" # サムネのパス
    pdf_pro = subprocess.Popen([acrobat_path, pdf_path+pdf_name+".pdf"], shell=False) # PDFファイルを表示
def swipe_scroll(event):
    if event.delta > 0:
        canvas.yview_scroll(-1, "units")
    elif event.delta < 0:
        canvas.yview_scroll(1, "units")
# ドロップウィンドウを生成
def create_drop_PDF():
    global drop_PDF
    drop_PDF = TkinterDnD.Tk()
    drop_PDF.geometry("320x200+300+300")
    drop_PDF.title("drop")
    drop_PDF.drop_target_register(DND_FILES)
    drop_PDF.dnd_bind("<<Drop>>", drop_PDFFiles)
    btn = tk.Button(drop_PDF,text="btn", command=btn_click)
    btn.place(x=20, y=20)
# ウィンドウにサムネイルを表示
def embed_thumbs():
    global drop_PDF
    thumbs_path = "thumbs/"
    count = 0
    thumbs_list = []
    for i in Path(thumbs_path).glob("*/"):
        print(f"{thumbs_path}{i.name}")
        img = ImageTk.PhotoImage(file=f"{thumbs_path}{i.name}",width=100, height=100, master=drop_PDF)
        btn = tk.Button(drop_PDF, image=img, relief="sunken")
        btn.place(x=20+int(count%5)*155, y=20+int(count/5)*215)
        thumbs_list.append(img)
        count += 1
    drop_PDF.mainloop()
def create_thumb():
    src_path = "splited/"
    thumb_path = "thumbs/"
    # フォルダの存在を確認し削除
    if os.path.exists(thumb_path):
        shutil.rmtree(thumb_path)
    
        os.mkdir(thumb_path)
    
    # PDFをサムネイル化し保存
    for i in Path(src_path).glob("*"):
        src_pdf = i.name
        pdf_pages = convert_from_path(src_path+src_pdf, 72, size=200)
        # print(f"{thumb_path}{os.path.splitext(i.name)[0]}.jpg")
        pdf_pages[0].save(f"{thumb_path}{os.path.splitext(i.name)[0]}.jpg", "JPEG")
    embed_thumbs()

# ドロップしてPDFを分割
def drop_PDFFiles(event):
    # PDFのパスを取得（バックスラッシュをスラッシュの置換し、波かっこを削除）
    reader = pypdf.PdfReader(event.data.replace("\\", "/").replace("{", "").replace("}", ""))
    total_pages = len(reader.pages)
    result_path = "./splited/"
    if os.path.exists(result_path):
        ret = messagebox.askyesno("確認", "フォルダを上書きしますか？")
        if ret:
            shutil.rmtree(result_path)
        else:
            exit()
    os.makedirs(result_path)
    # PDFを分割
    for i in range(total_pages):
        source_file = reader.pages[i]
        writer = pypdf.PdfWriter()
        writer.add_page(source_file)
        with open(f"{result_path}result{str(i).zfill(6)}.pdf", "wb") as fp:
            writer.write(fp)
        writer.close()
    create_thumb()




btn1 = tk.Button(root, text="btn1", command=create_drop_PDF)
btn1.place(x=20, y=20)
btn2 = tk.Button(root, text="btn2", command=show_canvas)
btn2.place(x=20, y=50)
btn3 = tk.Button(root, text="test", command=test)
btn3.place(x=20, y=80)
root.mainloop()
