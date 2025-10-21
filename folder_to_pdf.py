import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import inch
import threading
from docx import Document
from pptx import Presentation
import subprocess
import platform

class FolderToPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Klasör → PDF Dönüştürücü")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # Ana çerçeve
        main_frame = tk.Frame(root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Başlık
        title_label = tk.Label(main_frame, text="Klasör İçeriğini PDF'e Dönüştür", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Açıklama
        desc_label = tk.Label(main_frame, 
                             text="Bir klasör seçin. İçindeki tüm resimler ve metin dosyaları\ntek bir PDF dosyasına dönüştürülecek.",
                             font=("Arial", 10))
        desc_label.pack(pady=(0, 20))
        
        # Klasör seçim butonu
        select_btn = tk.Button(main_frame, text="📁 Klasör Seç", 
                              command=self.select_folder,
                              font=("Arial", 12, "bold"),
                              bg="#4CAF50", fg="white",
                              padx=30, pady=15,
                              cursor="hand2")
        select_btn.pack(pady=10)
        
        # Seçilen klasör gösterimi
        self.folder_label = tk.Label(main_frame, text="Henüz klasör seçilmedi", 
                                    font=("Arial", 9), fg="gray")
        self.folder_label.pack(pady=10)
        
        # İlerleme çubuğu
        self.progress = ttk.Progressbar(main_frame, length=400, mode='indeterminate')
        self.progress.pack(pady=20)
        
        # Durum etiketi
        self.status_label = tk.Label(main_frame, text="", font=("Arial", 9))
        self.status_label.pack()
        
        # Desteklenen formatlar
        self.image_formats = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'}
        self.text_formats = {'.txt', '.md', '.csv', '.json', '.xml', '.log', '.py', '.js', 
                           '.html', '.css', '.java', '.cpp', '.c', '.h'}
        self.office_formats = {'.docx', '.doc', '.pptx', '.ppt', '.xlsx', '.xls'}
    
    def select_folder(self):
        folder_path = filedialog.askdirectory(title="Dönüştürülecek Klasörü Seçin")
        
        if folder_path:
            self.folder_label.config(text=f"Seçilen: {folder_path}", fg="black")
            
            # PDF kayıt yeri sor
            pdf_path = filedialog.asksaveasfilename(
                title="PDF'i Kaydet",
                defaultextension=".pdf",
                filetypes=[("PDF Dosyası", "*.pdf")]
            )
            
            if pdf_path:
                # Ayrı thread'de dönüştürme işlemini başlat
                thread = threading.Thread(target=self.convert_to_pdf, 
                                        args=(folder_path, pdf_path))
                thread.start()
    
    def convert_to_pdf(self, folder_path, pdf_path):
        try:
            # İlerleme çubuğunu başlat
            self.progress.start()
            self.status_label.config(text="Dosyalar taranıyor...", fg="blue")
            
            # Tüm dosyaları topla ve sırala
            all_files = []
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    ext = os.path.splitext(file)[1].lower()
                    if ext in self.image_formats or ext in self.text_formats or ext in self.office_formats:
                        all_files.append(file_path)
            
            all_files.sort()
            
            if not all_files:
                self.progress.stop()
                self.status_label.config(text="Klasörde uygun dosya bulunamadı!", fg="red")
                messagebox.showwarning("Uyarı", "Klasörde dönüştürülebilir dosya bulunamadı!")
                return
            
            self.status_label.config(text=f"{len(all_files)} dosya bulundu. PDF oluşturuluyor...", fg="blue")
            
            # PDF oluştur
            c = canvas.Canvas(pdf_path, pagesize=A4)
            page_width, page_height = A4
            
            for idx, file_path in enumerate(all_files, 1):
                try:
                    ext = os.path.splitext(file_path)[1].lower()
                    filename = os.path.basename(file_path)
                    
                    if ext in self.image_formats:
                        self._add_image_page(c, file_path, filename, page_width, page_height)
                    elif ext in self.text_formats:
                        self._add_text_page(c, file_path, filename, page_width, page_height)
                    elif ext in self.office_formats:
                        self._add_office_page(c, file_path, filename, page_width, page_height)
                    
                    self.status_label.config(
                        text=f"İşleniyor: {idx}/{len(all_files)} - {filename}", 
                        fg="blue"
                    )
                
                except Exception as e:
                    print(f"Hata ({filename}): {str(e)}")
                    continue
            
            c.save()
            
            self.progress.stop()
            self.status_label.config(text="✓ PDF başarıyla oluşturuldu!", fg="green")
            messagebox.showinfo("Başarılı", f"PDF dosyası oluşturuldu:\n{pdf_path}")
            
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="❌ Hata oluştu!", fg="red")
            messagebox.showerror("Hata", f"PDF oluşturulurken hata:\n{str(e)}")
    
    def _add_image_page(self, c, image_path, filename, page_width, page_height):
        """Resim sayfası ekle"""
        c.setFont("Helvetica-Bold", 10)
        c.drawString(30, page_height - 30, f"Dosya: {filename}")
        
        img = Image.open(image_path)
        img_width, img_height = img.size
        
        # Sayfa içine sığdır
        max_width = page_width - 60
        max_height = page_height - 100
        
        ratio = min(max_width / img_width, max_height / img_height)
        new_width = img_width * ratio
        new_height = img_height * ratio
        
        x = (page_width - new_width) / 2
        y = page_height - 60 - new_height
        
        c.drawImage(ImageReader(img), x, y, width=new_width, height=new_height)
        c.showPage()
    
    def _add_text_page(self, c, text_path, filename, page_width, page_height):
        """Metin sayfası ekle"""
        try:
            with open(text_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except:
            try:
                with open(text_path, 'r', encoding='latin-1') as f:
                    content = f.read()
            except:
                content = "[Dosya okunamadı]"
        
        c.setFont("Helvetica-Bold", 10)
        c.drawString(30, page_height - 30, f"Dosya: {filename}")
        
        c.setFont("Courier", 8)
        
        lines = content.split('\n')
        y = page_height - 60
        
        for line in lines:
            if y < 50:
                c.showPage()
                c.setFont("Helvetica-Bold", 10)
                c.drawString(30, page_height - 30, f"Dosya: {filename} (devam)")
                c.setFont("Courier", 8)
                y = page_height - 60
            
            # Uzun satırları kes
            if len(line) > 95:
                line = line[:95] + "..."
            
            c.drawString(30, y, line)
            y -= 12
        
        c.showPage()

if __name__ == "__main__":
    root = tk.Tk()
    app = FolderToPDFConverter(root)
    root.mainloop()