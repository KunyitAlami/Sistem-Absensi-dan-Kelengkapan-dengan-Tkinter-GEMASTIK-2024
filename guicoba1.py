import tkinter as tk
import datetime
import openpyxl
import pandas as pd
import util
import cv2
import os
import subprocess
from PIL import Image, ImageTk
import torch
import random
import numpy as np
import pathlib
import time
from pathlib import Path
from yolov5.utils.general import non_max_suppression, scale_boxes
from yolov5.utils.torch_utils import select_device
from tkinter import simpledialog
import time
import pandas as pd
from tkinter import simpledialog, messagebox
import sys
import serial



class App():
    def __init__(self):
        # aturan untuk jendela utama
        self.jendelaUtama = tk.Tk()
        lebar = self.jendelaUtama.winfo_screenwidth()               
        tinggi = self.jendelaUtama.winfo_screenheight()
        self.jendelaUtama.geometry("%dx%d" % (lebar, tinggi))
        self.jendelaUtama.state('zoomed')
        self.jendelaUtama.title("SALAMAT TEAM GEMASTIK 2024")

        # buat button
        self.buttonMulai = util.get_button(self.jendelaUtama, "MULAI", "green", self.menuAbsensi)
        self.buttonMulai.place(x=950, y=400)

        # buat button lagi
        self.buttonKeluar = util.get_button(self.jendelaUtama, "Keluar", "red", self.keluarLayarUtama)
        self.buttonKeluar.place(x=950, y=500)

        # kita buat teks 
        self.JudulSatu = util.get_text_label(self.jendelaUtama, "SALAMAT TEAM GEMASTIK 2024\n\n  Selamat Datang di Menu Absensi\n\n        Berbasis Face Recognition")
        self.JudulSatu.place(x=880, y=150)

        # kita membuat area untuk menampilkan webcam nantinya 
        self.areaWebcam = util.get_img_label(self.jendelaUtama)
        self.areaWebcam.place(x=0, y=0, width=700, height=lebar//2)
        self.tambahWebcam(self.areaWebcam)

        self.db_dir='./databaseFoto'
        if not os.path.exists(self.db_dir):
            os.mkdir(self.db_dir)

        self.log_path = './Absen.xlsx' 
        self.kelengkapan_path = './Kelengkapan.xlsx'

    def tambahWebcam(self, label):
        if 'cap' not in self.__dict__:
            self.cap = cv2.VideoCapture(0)

        self._label = label
        self.running = True
        self.webcamNyala()

    def webcamNyala(self):
        if not self.running:
            return

        ret, frame = self.cap.read()
        if ret:
            self.most_recent_capture_arr = frame
            img_ = cv2.cvtColor(self.most_recent_capture_arr, cv2.COLOR_BGR2RGB)
            self.most_recent_capture_pil = Image.fromarray(img_)
            imgtk = ImageTk.PhotoImage(image=self.most_recent_capture_pil)
            
            # Menghapus gambar lama
            self._label.configure(image=None)
            
            # Menambahkan gambar baru
            self._label.imgtk = imgtk
            self._label.configure(image=imgtk)

        self._label.after(20, self.webcamNyala)


    def menuAbsensi(self):
        self.running = False
        self.jendelaKedua = tk.Toplevel(self.jendelaUtama)
        lebar = self.jendelaUtama.winfo_screenwidth()               
        tinggi = self.jendelaUtama.winfo_screenheight()
        self.jendelaKedua.geometry("%dx%d" % (lebar, tinggi))
        self.jendelaKedua.state("zoomed")
        self.jendelaKedua.title("Menu Absensi")

        self.areaWebcamkedua = util.get_img_label(self.jendelaKedua)
        self.areaWebcamkedua.place(x=0, y=0, width=700, height=lebar//2)
        self.tambahWebcam(self.areaWebcamkedua)

        # kita buat teks 
        self.JudulSatuJendelaDua = util.get_text_label(self.jendelaKedua, "Anda Masuk Kedalam Menu Absensi\n")
        self.JudulSatuJendelaDua.place(x=880, y=150)

        self.JudulDuaJendelaDua = util.get_text_label(self.jendelaKedua, "Silahkan Menekan Tombol Untuk Melakukan Absensi\n")
        self.JudulDuaJendelaDua.place(x=780, y=180)

        # tombol login
        # urutannya adalah dimana wadah jendela tombol hndk diandak, teks, warna bg, command
        self.tombolLogin = util.get_button(self.jendelaKedua, "Absensi", "green", self.loginMenu)
        self.tombolLogin.place(x=950, y=300)

        self.tomboldaftar = util.get_button(self.jendelaKedua, "Daftar Baru", "gray", self.daftar)
        self.tomboldaftar.place(x=950, y=400)

        # Buat tombol kembali ke jendela utama
        self.buttonKembali = util.get_button(self.jendelaKedua, "Kembali", "blue", self.kembaliJendelaUtama)
        self.buttonKembali.place(x=950, y=500)

    def kembaliJendelaUtama(self):
        self.running = False
        self.jendelaKedua.destroy()
        self.running = True
        self.tambahWebcam(self.areaWebcam)

    def keluarLayarUtama(self):
        self.running = False
        self.jendelaUtama.destroy()
    
    def cekJabatan(self, nama):
        jabatan_dict = {
            "Ghani Mudzakir": "Engineer",
            "Randy Febrian": "Programmer",
            "Muhammad Rizky": "Lecture",
            "Aufa Fitrianda":"Boss"
        }
        return jabatan_dict.get(nama, "Unknown")

    def loginMenu(self):
        self.running = False
        self.jendelaLogin = tk.Toplevel(self.jendelaKedua)
        lebar = self.jendelaUtama.winfo_screenwidth()               
        tinggi = self.jendelaUtama.winfo_screenheight()
        self.jendelaLogin.geometry("%dx%d" % (lebar, tinggi))
        self.jendelaLogin.state("zoomed")
        self.jendelaLogin.title("Pilih Absensi")

        # kita buat teks 
        self.JudulSatuJendelalogin = util.get_text_label(self.jendelaLogin, "Pilihan Absensi\n")
        self.JudulSatuJendelalogin.place(x=(lebar//2)-110, y=200)

        # tombol login
        # urutannya adalah dimana wadah jendela tombol hndk diandak, teks, warna bg, command
        self.tombolLoginMuka = util.get_button(self.jendelaLogin, "Absensi Wajah", "green", self.menuLoginMuka)
        self.tombolLoginMuka.place(x=(lebar//2)-175, y=300)

        self.tombolLoginKelengkapan = util.get_button(self.jendelaLogin, "Absensi Kelengkapan", "green", self.menuLoginKelengkapan)
        self.tombolLoginKelengkapan.place(x=(lebar//2)-175, y=400)

        self.tombolbalikjendeladua = util.get_button(self.jendelaLogin, "Kembali", "blue", self.menubalikjendeladua)
        self.tombolbalikjendeladua.place(x=(lebar//2)-175, y=500)

    def menubalikjendeladua(self):
        self.running = False
        self.jendelaLogin.destroy()
        self.running = True
        self.tambahWebcam(self.areaWebcamkedua)

    def menubalikjendeladuadua(self):
        self.jendelaloginmuka.destroy()


    def menuLoginMuka(self):
        self.running = False
        self.jendelaloginmuka = tk.Toplevel(self.jendelaLogin)
        lebar = self.jendelaUtama.winfo_screenwidth()               
        tinggi = self.jendelaUtama.winfo_screenheight()
        self.jendelaloginmuka.geometry("%dx%d" % (lebar, tinggi))
        self.jendelaloginmuka.state("zoomed")
        self.jendelaloginmuka.title("Absensi Wajah")

        self.areaWebcamloginmuka = util.get_img_label(self.jendelaloginmuka)
        self.areaWebcamloginmuka.place(x=0, y=0, width=700, height=lebar//2)
        self.tambahWebcam(self.areaWebcamloginmuka)

        self.tombolbalikjendeladuadua = util.get_button(self.jendelaloginmuka, "Kembali", "blue", self.menubalikjendeladuadua)
        self.tombolbalikjendeladuadua.place(x=750, y=400)

        self.tombolambilfoto = util.get_button(self.jendelaloginmuka, "Ambil Foto", "green", self.login)
        self.tombolambilfoto.place(x=750, y=300)

    def menuLoginKelengkapan(self):
        self.running = False
        self.jendelaloginkelengakapan = tk.Toplevel(self.jendelaLogin)
        lebar = self.jendelaUtama.winfo_screenwidth()               
        tinggi = self.jendelaUtama.winfo_screenheight()
        self.jendelaloginkelengakapan.geometry("%dx%d" % (lebar, tinggi))
        self.jendelaloginkelengakapan.state("zoomed")
        self.jendelaloginkelengakapan.title("Absensi Kelengkapan")

        self.areaWebcamloginkelengkapan = util.get_img_label(self.jendelaloginkelengakapan)
        self.areaWebcamloginkelengkapan.place(x=0, y=0, width=700, height=lebar//2)
        self.tambahWebcam(self.areaWebcamloginkelengkapan)

        self.tombolbalikjendeladuadua = util.get_button(self.jendelaloginkelengakapan, "Kembali", "blue", self.menubalikjendeladuaduadua)
        self.tombolbalikjendeladuadua.place(x=750, y=400)

        self.tombolambilfoto = util.get_button(self.jendelaloginkelengakapan, "Mulai Cek Perlengkapan", "green", self.pemberitahuansebelum)
        self.tombolambilfoto.place(x=750, y=300)


    def pemberitahuansebelum(self):
        # Membuat jendela toplevel
        top = tk.Toplevel()
        top.title("Pemberitahuan Sebelum Memulai Cek Kelengkapan")
        
        # Mengatur ukuran dan posisi di tengah layar
        window_width, window_height = 600, 200
        screen_width = top.winfo_screenwidth()
        screen_height = top.winfo_screenheight()
        position_top = int(screen_height/2 - window_height/2)
        position_right = int(screen_width/2 - window_width/2)
        top.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
        
        # Menambahkan teks poin
        teks = [
            "\nPoin 1: Silahkan pasang seluruh perlengkapan kamu.",
            "Poin 2: Berdirilah cukup jauh dari kamera agar seluruh perlengkapan kami terlihat.",
            "Poin 3: Tunggu 5 detik sebelum mulai."
        ]
        for point in teks:
            label = tk.Label(top, text=point, pady=5)
            label.pack()

        # Menambahkan tombol "Mulai"
        start_button = tk.Button(top, text="Mulai", command=lambda: self.on_start_button_click(top, self.jendelaloginkelengakapan, self.jendelaKedua, self.jendelaLogin))
        start_button.pack(pady=20)

    def on_start_button_click(self, top, top2, top43, top4):
        # Fungsi untuk menutup jendela top
        top.destroy()
        top2.destroy()
        
        # Panggil fungsi cekkelengkapan() atau fungsi lain yang ingin Anda panggil setelah jendela ditutup
        self.cekkelengkapan()

    def cekkelengkapan(self):
        # ser = serial.Serial('COM3', 115200)
        # Ganti PosixPath dengan WindowsPath
        temp = pathlib.PosixPath
        pathlib.PosixPath = pathlib.WindowsPath

        # Inisialisasi koneksi serial ke ESP32
        ser = serial.Serial('COM3', 115200)  # Ganti 'COM3' dengan port serial yang sesuai

        def plot_one_box(x, img, color=None, label=None, line_thickness=3):
            # Menggambar satu kotak batas pada gambar img
            tl = line_thickness or round(0.002 * max(img.shape[0:2])) + 1  # ketebalan garis
            color = color or [random.randint(0, 255) for _ in range(3)]
            c1, c2 = (int(x[0]), int(x[1])), (int(x[2]), int(x[3]))
            cv2.rectangle(img, c1, c2, color, thickness=tl, lineType=cv2.LINE_AA)
            if label:
                tf = max(tl - 1, 1)  # ketebalan font
                t_size = cv2.getTextSize(label, 0, fontScale=tl / 3, thickness=tf)[0]
                c2 = c1[0] + t_size[0], c1[1] - t_size[1] - 3
                cv2.rectangle(img, c1, c2, color, -1, cv2.LINE_AA)  # terisi
                cv2.putText(img, label, (c1[0], c1[1] - 2), 0, tl / 3, [225, 255, 255], thickness=tf, lineType=cv2.LINE_AA)

        # Membuka jendela input untuk memasukkan nama
        root = tk.Tk()
        root.withdraw()  # menyembunyikan jendela utama
        user_name = ""

        while user_name == "":
            user_name = simpledialog.askstring(title="Input Nama", prompt="Masukkan Nama Anda:")

        # Menambahkan penundaan 5 detik sebelum kamera mulai
        print("Menunggu 5 detik sebelum kamera menyala...")
        time.sleep(5)

        # Memuat model
        model = torch.load('D:/GHANI/Gemastik Kode/Final Project/FIX SALAMAT TEAM/BEST.pt')['model'].float()
        model.eval()

        # Inisialisasi webcam
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("Error: Could not open webcam.")
            exit()

        # Pengaturan inferensi
        device = select_device('')
        half = device.type != 'cpu'  # presisi setengah hanya didukung di CUDA

        if half:
            model.half()  # ke FP16

        # Mendefinisikan nama dan warna
        names = model.module.names if hasattr(model, 'module') else model.names
        colors = [[random.randint(0, 255) for _ in range(3)] for _ in names]

        # Dictionary untuk menyimpan objek yang terdeteksi
        detected_objects = {}

        # Jumlah objek yang harus terdeteksi sebelum program berhenti
        target_object_count = 4

        while cap.isOpened():
            ret, frame = cap.read()
            if not ret:
                print("Error: Could not read frame.")
                break

            # Padded resize
            img = cv2.resize(frame, (640, 640))  # Ubah ukuran frame ke ukuran input model
            img = img[:, :, ::-1].transpose(2, 0, 1)  # BGR ke RGB, ke 3x640x640
            img = np.ascontiguousarray(img)
            
            img = torch.from_numpy(img).to(device)
            img = img.half() if half else img.float()  # uint8 ke fp16/32
            img /= 255.0  # 0 - 255 ke 0.0 - 1.0
            if img.ndimension() == 3:
                img = img.unsqueeze(0)

            # Inferensi
            with torch.no_grad():  # mempercepat
                pred = model(img)[0]

            # Terapkan NMS
            pred = non_max_suppression(pred, 0.25, 0.45, classes=None, agnostic=False)

            # Proses deteksi
            for i, det in enumerate(pred):  # deteksi per gambar
                if len(det):
                    # Skala ulang kotak dari img_size ke ukuran frame
                    det[:, :4] = scale_boxes(img.shape[2:], det[:, :4], frame.shape).round()

                    # Tulis hasil
                    for *xyxy, conf, cls in reversed(det):
                        label = names[int(cls)]
                        confidence = conf.item()
                        
                        # Jika objek belum terdeteksi dan confidence > 0.73
                        if confidence > 0.73 and label not in detected_objects:
                            detected_objects[label] = "ada"
                            print(f"Detected {label} with confidence {confidence:.2f}")
                        
                        # Gambar kotak hanya jika objek belum terdeteksi
                        if label not in detected_objects:
                            plot_one_box(xyxy, frame, label=f'{label} {confidence:.2f}', color=colors[int(cls)], line_thickness=3)

            # Tampilkan frame hasil
            cv2.imshow('frame', frame)
            
            
            # Cek jika sudah terdeteksi target_object_count objek
            if len(detected_objects) >= target_object_count:
                # Kirim sinyal ke ESP32
                a = "lengkap"
                ser.write(a.encode())
                print("lengkap")
                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                masukanDataKeExcel = {"Nama": [user_name], "Waktu Absensi": [current_time], "Perlengkapan" : [a]}
                if os.path.exists(self.kelengkapan_path):
                    # Baca file Excel yang ada
                    log_df = pd.read_excel(self.kelengkapan_path)
                    # Menambahkan data baru ke DataFrame yang sudah ada
                    log_df = pd.concat([log_df, pd.DataFrame(masukanDataKeExcel)], ignore_index=True)
                else:
                    log_df = pd.DataFrame(masukanDataKeExcel)

                # Menulis kembali data ke file Excel dengan pandas
                log_df.to_excel(self.kelengkapan_path, index=False)

                # Membuka file Excel dengan openpyxl untuk menjaga format
                wb = openpyxl.load_workbook(self.kelengkapan_path)
                ws = wb.active

                # Mengatur lebar kolom (misalnya untuk kolom "A" dan "B")
                ws.column_dimensions['A'].width = 30  # Sesuaikan dengan lebar yang diinginkan
                ws.column_dimensions['B'].width = 30  # Sesuaikan dengan lebar yang diinginkan            
                ws.column_dimensions['C'].width = 30  # Sesuaikan dengan lebar yang diinginkan
                ws.column_dimensions['D'].width = 30  # Sesuaikan dengan lebar yang diinginkan
                wb.save(self.kelengkapan_path)


                # Hentikan eksekusi program selama 5 detik
                time.sleep(5)
                
                # Tampilkan message box
                messagebox.showinfo("Terima Kasih", "TERIMA KASIH TELAH MEMAKAI PERLENGKAPAN SESUAI STANDARD, SILAHKAN MULAI BEKERJA")
                
                # Hentikan loop
                break
            
            # Tekan 'q' untuk keluar
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

        # Ketika semuanya selesai, lepaskan tangkapan dan tutup koneksi serial
        cap.release()
        cv2.destroyAllWindows()
        a="lengkap"
        ser.write(a.encode())
        ser.close()

        # Cetak objek yang terdeteksi
        print("Detected objects:", detected_objects)

    def menubalikjendeladuaduadua(self):
        self.jendelaloginkelengakapan.destroy()

    def login(self):
        fotoBelumDiketahui = './tmp.jpg'

        cv2.imwrite(fotoBelumDiketahui, self.most_recent_capture_arr)

        output = str(subprocess.check_output(['face_recognition', self.db_dir, fotoBelumDiketahui]))
        nama = output.split(',')[1][:-5]
        print(nama)

        cekJabatan = self.cekJabatan(nama)

        if nama in ['unknown_person', 'no_persons_found', 'unknown_per']:
            util.msg_box("Gagal Login", "Mohon Maaf Wajah Anda Tidak Terdeteksi\nDidalam Database Kami,\nMohon Ulangi atau Daftar Ulang")
        else:
            util.msg_box("Berhasil Login", "Selamat Datang {}\nTerima Kasih Sudah Melakukan Absensi Hari Ini\nSelamat Bekerja!".format(nama))
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            masukanDataKeExcel = {"Nama": [nama], "Waktu Absensi": [current_time], "Jabatan" : [cekJabatan]}


            if os.path.exists(self.log_path):
                # Baca file Excel yang ada
                log_df = pd.read_excel(self.log_path)
                # Menambahkan data baru ke DataFrame yang sudah ada
                log_df = pd.concat([log_df, pd.DataFrame(masukanDataKeExcel)], ignore_index=True)
            else:
                log_df = pd.DataFrame(masukanDataKeExcel)

            # Menulis kembali data ke file Excel dengan pandas
            log_df.to_excel(self.log_path, index=False)

            # Membuka file Excel dengan openpyxl untuk menjaga format
            wb = openpyxl.load_workbook(self.log_path)
            ws = wb.active

            # Mengatur lebar kolom (misalnya untuk kolom "A" dan "B")
            ws.column_dimensions['A'].width = 30  # Sesuaikan dengan lebar yang diinginkan
            ws.column_dimensions['B'].width = 30  # Sesuaikan dengan lebar yang diinginkan            
            ws.column_dimensions['C'].width = 30  # Sesuaikan dengan lebar yang diinginkan
            ws.column_dimensions['D'].width = 30  # Sesuaikan dengan lebar yang diinginkan
            wb.save(self.log_path)

        self.jendelaloginmuka.destroy()
        self.running = True
        self.tambahWebcam(self.areaWebcam)
        os.remove(fotoBelumDiketahui)

        # Menutup jendela login muka dan kembali ke jendela login


    def daftar(self):
        self.running = False
        self.jendelaDaftar = tk.Toplevel(self.jendelaKedua)
        lebar = self.jendelaKedua.winfo_screenwidth()               
        tinggi = self.jendelaKedua.winfo_screenheight()
        self.jendelaDaftar.geometry("%dx%d" % (lebar, tinggi))
        self.jendelaDaftar.state("zoomed")

        #kita meulah area hagan webcamnya
        self.fotoLogin = util.get_img_label(self.jendelaDaftar)
        self.fotoLogin.place(x=10, y=0, width=700, height=lebar//2)
        self.tambahFoto(self.fotoLogin)


        #kita meulah text lwn wadah text hagan user memasukan nama foto yang di ambil tadi
        self.labelMasukanNama = util.get_text_label(self.jendelaDaftar, "Silahkan Masukan\nNama Lengkap Anda")
        self.labelMasukanNama.place(x=950, y=150)

        self.wadahMasukanNama = util.get_entry_text(self.jendelaDaftar)
        self.wadahMasukanNama.place(x=950, y=250)

        self.tombolDaftar= util.get_button(self.jendelaDaftar, "Daftar",'green', 
                                           self.terimaPendaftaran)
        self.tombolDaftar.place(x=950, y=400)

        # Buat tombol kembali ke jendela utama
        self.UlangiFoto = util.get_button(self.jendelaDaftar, "Ulangi Foto", "gray", self.kembaliJendelaDaftar)
        self.UlangiFoto.place(x=950, y=500)

    def terimaPendaftaran(self):
        nama = self.wadahMasukanNama.get(1.0, "end-1c")
        cv2.imwrite(os.path.join(self.db_dir,'{}.jpg'.format(nama)), self.register_new_user_capture)

        util.msg_box("Pendaftaran Wajah Absensi Baru", "Pendaftaran Berhasil\n Silahkan Anda Mencoba Absensi Lagi!")

        self.running = False
        self.jendelaDaftar.destroy()
        self.tambahWebcam(self.areaWebcam)

    def tambahFoto(self,label):
        imgtk = ImageTk.PhotoImage(image=self.most_recent_capture_pil)
        label.imgtk = imgtk
        label.configure(image=imgtk)

        self.register_new_user_capture= self.most_recent_capture_arr.copy()

    def kembaliJendelaDaftar(self):
        self.running = False
        self.jendelaDaftar.destroy()
        self.tambahWebcam(self.areaWebcamkedua)

    def mulai(self):
        self.jendelaUtama.mainloop()

if __name__ == "__main__":
    app = App()
    app.mulai()


if __name__ == "__main__":
    p = App()
    p.jendelaUtama.mainloop()
