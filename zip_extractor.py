import tkinter as tk
from tkinter import ttk, filedialog, Menu
import zipfile
import os
from datetime import datetime
import schedule
import time
import threading
import shutil
import logging

class ZipExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ZIP ve Excel Dosyası İşlemcisi")
        self.root.geometry("600x500")
        
        # Logging ayarları
        self.setup_logging()
        
        # Menü oluşturma
        self.create_menu()
        
        # Ana çerçeve
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Dosya seçim alanları
        ttk.Label(self.main_frame, text="ZIP Dosyası:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.zip_path = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.zip_path, width=50).grid(row=0, column=1, pady=5)
        ttk.Button(self.main_frame, text="Gözat", command=self.browse_zip).grid(row=0, column=2, padx=5, pady=5)
        
        # Hedef klasör seçimi
        ttk.Label(self.main_frame, text="Hedef Klasör:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.target_path = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.target_path, width=50).grid(row=1, column=1, pady=5)
        ttk.Button(self.main_frame, text="Gözat", command=self.browse_target).grid(row=1, column=2, padx=5, pady=5)
        
        # Excel'in yeni adı
        ttk.Label(self.main_frame, text="Excel'in Yeni Adı:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.excel_new_name = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.excel_new_name, width=50).grid(row=2, column=1, pady=5)
        
        # Zaman seçimi
        ttk.Label(self.main_frame, text="Çıkarma Zamanı (SS:DD):").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.extract_time = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.extract_time, width=10).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # Log görüntüleme alanı
        ttk.Label(self.main_frame, text="İşlem Logları:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.log_text = tk.Text(self.main_frame, height=8, width=60)
        self.log_text.grid(row=5, column=0, columnspan=3, pady=5)
        scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=5, column=3, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # Durum göstergesi
        self.status_var = tk.StringVar(value="Hazır")
        ttk.Label(self.main_frame, textvariable=self.status_var).grid(row=6, column=0, columnspan=3, pady=10)
        
        # Düğmeler
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)
        
        self.start_button = ttk.Button(button_frame, text="Başlat", command=self.start_scheduler)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="Durdur", command=self.stop_scheduler, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        self.scheduler_running = False
        self.scheduler_thread = None

        # Program başladığında otomatik olarak zamanlayıcıyı başlat
        self.root.after(1000, self.auto_start)

    def auto_start(self):
        """Program başladığında otomatik olarak zamanlayıcıyı başlatır"""
        if self.validate_inputs(show_errors=False):
            self.start_scheduler()

    def setup_logging(self):
        if not os.path.exists('logs'):
            os.makedirs('logs')
            
        log_file = os.path.join('logs', f'zip_extractor_{datetime.now().strftime("%Y%m%d")}.log')
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

    def create_menu(self):
        menubar = Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Dosya", menu=file_menu)
        file_menu.add_command(label="Logları Göster", command=self.show_logs)
        file_menu.add_separator()
        file_menu.add_command(label="Çıkış", command=self.root.quit)
        
        about_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Hakkında", menu=about_menu)
        about_menu.add_command(label="Hakkında", command=self.show_about)

    def show_logs(self):
        log_dir = 'logs'
        if os.path.exists(log_dir):
            os.startfile(log_dir)
        else:
            self.log_message("Log klasörü henüz oluşturulmamış.")

    def show_about(self):
        about_text = """
        ZIP ve Excel Dosyası İşlemcisi
        
        Yazar: Önder AKÖZ
        E-posta: onder7@gmail.com
        
        Bu program, ZIP dosyalarından Excel dosyalarını
        çıkarmak ve yeniden adlandırmak için
        tasarlanmıştır.
        
        © 2025 Tüm hakları saklıdır.
        """
        self.log_message("Hakkında bilgisi görüntülendi")

    def log_message(self, message, level="info"):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"{timestamp} - {message}\n"
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
        
        if level == "info":
            self.logger.info(message)
        elif level == "error":
            self.logger.error(message)
        elif level == "warning":
            self.logger.warning(message)

    def browse_zip(self):
        filename = filedialog.askopenfilename(filetypes=[("ZIP dosyaları", "*.zip")])
        if filename:
            self.zip_path.set(filename)
            self.log_message(f"ZIP dosyası seçildi: {filename}")

    def browse_target(self):
        dirname = filedialog.askdirectory()
        if dirname:
            self.target_path.set(dirname)
            self.log_message(f"Hedef klasör seçildi: {dirname}")

    def extract_zip(self):
        try:
            zip_path = self.zip_path.get()
            target_path = self.target_path.get()
            excel_new_name = self.excel_new_name.get()
            
            self.log_message(f"ZIP çıkarma işlemi başlatıldı: {zip_path}")
            
            temp_dir = os.path.join(target_path, "temp_extract")
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)
            
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            excel_found = False
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    if file.endswith(('.xlsx', '.xls')):
                        old_path = os.path.join(root, file)
                        new_name = f"{excel_new_name}.{file.split('.')[-1]}"
                        new_path = os.path.join(target_path, new_name)
                        shutil.move(old_path, new_path)
                        excel_found = True
                        self.log_message(f"Excel dosyası yeniden adlandırıldı: {new_name}")
                        break
                if excel_found:
                    break
            
            shutil.rmtree(temp_dir)
            
            if excel_found:
                self.log_message(f"Excel dosyası başarıyla çıkarıldı ve yeniden adlandırıldı: {new_name}")
            else:
                self.log_message("ZIP dosyasında Excel dosyası bulunamadı!", "warning")
            
        except Exception as e:
            self.log_message(f"İşlem sırasında hata oluştu: {str(e)}", "error")
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

    def validate_inputs(self, show_errors=True):
        if not self.zip_path.get():
            if show_errors:
                self.log_message("ZIP dosyası seçilmedi!", "error")
            return False
        if not self.target_path.get():
            if show_errors:
                self.log_message("Hedef klasör seçilmedi!", "error")
            return False
        if not self.excel_new_name.get():
            if show_errors:
                self.log_message("Excel için yeni ad belirlenmedi!", "error")
            return False
        if not self.extract_time.get():
            if show_errors:
                self.log_message("Çıkarma zamanı belirlenmedi!", "error")
            return False
        
        try:
            hour, minute = map(int, self.extract_time.get().split(':'))
            if not (0 <= hour <= 23 and 0 <= minute <= 59):
                raise ValueError
        except ValueError:
            if show_errors:
                self.log_message("Geçersiz zaman formatı!", "error")
            return False
            
        return True

    def scheduler_loop(self):
        while self.scheduler_running:
            schedule.run_pending()
            time.sleep(1)

    def start_scheduler(self):
        if not self.validate_inputs():
            return
        
        self.scheduler_running = True
        self.start_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)
        
        schedule.clear()
        extract_time = self.extract_time.get()
        schedule.every().day.at(extract_time).do(self.extract_zip)
        
        self.scheduler_thread = threading.Thread(target=self.scheduler_loop)
        self.scheduler_thread.start()
        
        self.log_message(f"Zamanlayıcı başlatıldı. Çıkarma zamanı: {extract_time}")
        self.status_var.set(f"Aktif - Her gün saat {extract_time}'de çalışacak")

    def stop_scheduler(self):
        self.scheduler_running = False
        if self.scheduler_thread:
            self.scheduler_thread.join()
        schedule.clear()
        
        self.start_button.configure(state=tk.NORMAL)
        self.stop_button.configure(state=tk.DISABLED)
        
        self.log_message("Zamanlayıcı durduruldu")
        self.status_var.set("Durduruldu")

if __name__ == "__main__":
    root = tk.Tk()
    app = ZipExtractorGUI(root)
    root.mainloop()
