import os
import win32com.client
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import time
import threading

class ScannerInfo:
    def __init__(self):
        pass
    
    def scan_devices(self):
        devices = []
        
        # WIA servislerine bağlan
        wia = win32com.client.Dispatch("WIA.DeviceManager")
        
        # Bağlı cihazları al
        all_devices = wia.DeviceInfos
        
        if all_devices.Count == 0:
            return None
        else:
            for i in range(1, all_devices.Count + 1):
                device_info = all_devices.Item(i)
                device = {
                    'DeviceID': device_info.DeviceID,
                    'Manufacturer': getattr(device_info, 'Manufacturer', 'Bilinmiyor'),
                    'Model': getattr(device_info, 'Model', 'Bilinmiyor'),
                    'DeviceType': getattr(device_info, 'Type', 'Bilinmiyor'),
                }
                devices.append(device)
        
        # Cihazları döndür
        return devices

class ScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Scanner Selection")
        
        # Dark mode için stil ayarları
        self.set_dark_mode()

        # Cihazları al
        scanner_info = ScannerInfo()
        self.devices = scanner_info.scan_devices()

        # Eğer cihaz yoksa uyarı göster
        if not self.devices:
            self.label = ttk.Label(self.root, text="Hiçbir cihaz bulunamadı.", background="#2E2E2E", foreground="white")
            self.label.grid(row=0, column=0, padx=10, pady=10)
            return

        # Cihazları listele
        self.device_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE, height=10, bg="#173B45", fg="#F8EDED", selectbackground="#F8EDED", selectforeground="#FFB200")
        self.device_listbox.grid(row=1, column=0, padx=10, pady=10)

        # Cihazları Listbox'a ekle
        for device in self.devices:
            display_text = f"{device['DeviceID']} - {device['Manufacturer']} - {device['Model']}"
            self.device_listbox.insert(tk.END, display_text)

        # Dosya adı girişi
        label_file_name = ttk.Label(self.root, text="Dosya Adı:", background="#173B45", foreground="#FFB200")
        label_file_name.grid(row=2, column=0, padx=10, pady=10)
        self.entry_file_name = ttk.Entry(self.root)
        self.entry_file_name.grid(row=2, column=1, padx=10, pady=10)

        # Dosya sahibi girişi
        label_file_owner = ttk.Label(self.root, text="Dosya Sahibi:", background="#173B45", foreground="#FFB200")
        label_file_owner.grid(row=3, column=0, padx=10, pady=10)
        self.entry_file_owner = ttk.Entry(self.root)
        self.entry_file_owner.grid(row=3, column=1, padx=10, pady=10)

        # Dosya konumu butonu
        label_file_location = ttk.Label(self.root, text="Dosya Konumu:", background="#173B45", foreground="#FFB200")
        label_file_location.grid(row=4, column=0, padx=10, pady=10)
        self.button_file_location = ttk.Button(self.root, text="Konum Seç", command=self.select_file_location, style="TButton")
        self.button_file_location.grid(row=4, column=1, padx=10, pady=10)

        # Dosya konumu görüntüleme (butonun altına alındı)
        self.entry_file_location = ttk.Entry(self.root, state="readonly", width=40)
        self.entry_file_location.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

        # Taramayı Başlat butonu
        self.scan_button = ttk.Button(self.root, text="Taramayı Başlat", command=self.start_scan, style="TButton")
        self.scan_button.grid(row=6, column=0, padx=10, pady=10)

        # Taramayı Bitir butonu
        self.stop_button = ttk.Button(self.root, text="Taramayı Bitir", command=self.stop_scan, state="disabled", style="TButton")
        self.stop_button.grid(row=6, column=1, padx=10, pady=10)

        # Durum mesajı
        self.status_label = ttk.Label(self.root, text="", foreground="red", background="#173B45")
        self.status_label.grid(row=7, column=0, columnspan=2)

        # Tarama durumu
        self.is_scanning = False

    def set_dark_mode(self):
        # Koyu tema için stil ayarları
        self.root.configure(bg="#173B45")
        
        style = ttk.Style()
        style.configure("TButton", background="#173B45", foreground="#FFB200", padding=6)  # Altın rengi buton, siyah yazı
        style.configure("TLabel", background="#F8EDED", foreground="#FFB200")  # Siyah yazı rengi
        style.configure("TEntry", fieldbackground="#F8EDED", foreground="#FFB200", padding=6)

    def select_file_location(self):
        # Dosya seçme penceresini aç
        folder_selected = filedialog.askdirectory(title="Dosya Konumu Seç")
        
        # Eğer bir klasör seçildiyse, yolunu Entry'ye yaz
        if folder_selected:
            self.entry_file_location.config(state="normal")  # Giriş alanını aktif yap
            self.entry_file_location.delete(0, tk.END)  # Mevcut metni temizle
            self.entry_file_location.insert(0, folder_selected)  # Seçilen yolu ekle
            self.entry_file_location.config(state="readonly")  # Giriş alanını tekrar salt okunur yap

    def start_scan(self):
        # Kullanıcıdan alınan veriler
        file_name = self.entry_file_name.get()
        file_owner = self.entry_file_owner.get()
        file_location = self.entry_file_location.get()

        # Dosya yolu kontrolü
        if not os.path.exists(file_location):
            self.status_label.config(text="Verilen dosya yolu mevcut değil!", foreground="red")
            return

        # Dosya yolu içinde bir klasör var mı kontrol et
        folder_path = os.path.join(file_location, file_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            self.status_label.config(text=f"Klasör oluşturuldu: {folder_path}", foreground="green")

        # Dosya sahibi adıyla dosya var mı kontrol et
        file_path = os.path.join(folder_path, f"{file_owner}.txt")
        if not os.path.exists(file_path):
            with open(file_path, 'w') as f:
                f.write(f"Dosya Sahibi: {file_owner}")
            self.status_label.config(text=f"Dosya oluşturuldu: {file_path}", foreground="green")
        else:
            self.status_label.config(text=f"{file_owner} isminde dosya zaten var.", foreground="yellow")

        # Tarama işlemini başlat
        self.is_scanning = True
        self.scan_button.config(state="disabled")
        self.stop_button.config(state="normal")

        # Tarama işlemi için ayrı bir thread başlat
        threading.Thread(target=self.scan_documents, args=(folder_path,)).start()

    def scan_documents(self, folder_path):
        # Seçilen cihazlardan ilkini al
        selected_device_index = self.device_listbox.curselection()
        if not selected_device_index:
            self.status_label.config(text="Lütfen bir cihaz seçin.", foreground="red")
            self.is_scanning = False
            self.stop_button.config(state="disabled")
            self.scan_button.config(state="normal")
            return

        device = self.devices[selected_device_index[0]]
        device_id = device['DeviceID']

        # Cihaz üzerinden tarama işlemi başlat
        wia = win32com.client.Dispatch("WIA.DeviceManager")
        device_info = wia.DeviceInfos.Item(device_id)
        device_obj = device_info.Connect()

        # Tarama işlemi
        image_count = 1
        while self.is_scanning:
            # Tarama işlemi burada simüle edilecek, gerçek tarama işlemi burada yapılabilir.
            self.status_label.config(text=f"Tarama işlemi devam ediyor... {image_count}. sayfa", foreground="green")
            time.sleep(2)  # Tarama süresi simülasyonu

            # Taranan belgeyi kaydet
            file_path = os.path.join(folder_path, f"{image_count}.png")
            with open(file_path, 'w') as f:
                f.write(f"Tarama {image_count} - {device_obj.Model}")
            image_count += 1

        self.status_label.config(text="Tarama tamamlandı.", foreground="green")
        self.stop_button.config(state="disabled")
        self.scan_button.config(state="normal")

    def stop_scan(self):
        # Tarama işlemini durdur
        self.is_scanning = False
        self.status_label.config(text="Tarama durduruldu.", foreground="red")
        self.stop_button.config(state="disabled")
        self.scan_button.config(state="normal")


# Ana pencereyi oluştur
root = tk.Tk()
app = ScannerApp(root)

# Uygulamayı başlat
root.mainloop()
