\# 📈 ASP ile Dinamik Koordinat ve Grafik Çizim Sistemi



Bu proje, \*\*Classic ASP (Active Server Pages)\*\* ve \*\*Microsoft Access\*\* veritabanı kullanılarak geliştirilmiş bir grafik çizim uygulamasıdır. Kullanıcıdan alınan 2D koordinat verilerini (X, Y) veritabanına kaydeder ve bu verileri işleyerek HTML5 \*\*SVG (Scalable Vector Graphics)\*\* teknolojisi ile dinamik olarak ekrana çizer.



\## 🚀 Proje Özellikleri



\* \*\*Veri Girişi:\*\* Kullanıcı X ve Y koordinatlarını ve çizim sırasını (Sıra No) sisteme girer.

\* \*\*Veritabanı Entegrasyonu:\*\* Girilen veriler Microsoft Access (.mdb) veritabanında kalıcı olarak saklanır.

\* \*\*Dinamik Çizim (SVG):\*\* `ciz.asp` sayfası, veritabanındaki noktaları sırasına göre okur ve `<polyline>` kullanarak şekli oluşturur.

\* \*\*Görselleştirme:\*\* Her koordinat noktası üzerinde görsel işaretleyici ve koordinat bilgisi yer alır.

\* \*\*Veri Yönetimi:\*\* Veritabanındaki çizim verileri tek tuşla temizlenebilir.



\## 🛠 Kullanılan Teknolojiler



\* \*\*Backend:\*\* Classic ASP (VBScript)

\* \*\*Veritabanı:\*\* Microsoft Access 2000-2003 (.mdb)

\* \*\*Veri Erişim:\*\* ADODB (Jet OLEDB 4.0)

\* \*\*Frontend:\*\* HTML5, CSS3

\* \*\*Grafik Motoru:\*\* SVG (Scalable Vector Graphics)



\## 📂 Dosya Yapısı



```text

Proje\_Klasoru/

│

├── db/

│   └── veritabani.mdb    <-- (Access Veritabanı Dosyası)

│

├── baglanti.asp          <-- (Veritabanı bağlantı ayarları)

├── index.asp             <-- (Veri giriş formu)

├── ekle.asp              <-- (Veriyi işleyen ve kaydeden arka plan kodu)

├── ciz.asp               <-- (Grafiği çizen ana sayfa)

├── temizle.asp           <-- (Veritabanını sıfırlayan kod)

└── README.md             <-- (Proje dökümantasyonu)


## ⚙️ Kurulum ve Çalıştırma 

Bu proje veritabanına yazma işlemi yaptığı için klasör izinlerinin doğru ayarlanması gerekmektedir.

    Sunucu: Baby Web Server veya IIS (Internet Information Services) kullanabilirsiniz.

    **Çalıştırma:**

        Sunucuyu başlattıktan sonra tarayıcınızda http://localhost/index.asp adresine gidin.

1. Veri Giriş Ekranı

![Giriş Ekranı](ekran1.png))

2. Dinamik Çizim Sonucu

![Çizim Ekranı](ekran2.png))
