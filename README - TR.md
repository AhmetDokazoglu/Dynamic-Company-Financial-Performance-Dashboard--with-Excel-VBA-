# 📊 Dynamic Company Financial Performance Dashboard (with Excel VBA)

### 💡 Otomatik Finansal Veri Görselleştirme ve Raporlama Sistemi
Bu proje, satış verilerini kullanarak bir şirketin finansal performansını çok boyutlu olarak analiz eden, **tamamen Excel tabanlı** bir dashboard sistemidir.  
Amaç, ürün, ülke ve dönem bazlı satış ve kârlılık analizlerini dinamik olarak sunmak ve karar alma süreçlerini kolaylaştırmaktır.

---

## ⚙️ Kullanılan Araçlar ve Teknolojiler
- **Microsoft Excel**
- **Pivot Table**
- **Power Query**
- **Grafikler (Çizgi, Pasta, Sütun, Dağılım)**
- **VBA (Visual Basic for Applications)**

Proje türü: 📈 Veri Analizi ve İş Zekâsı

---

## 🧩 Proje İçeriği

### 1️⃣ Proje Amacı
Şirketin satış verilerini analiz ederek performans göstergelerini (kâr, indirim, kâr marjı vb.) otomatik olarak hesaplayan bir **dinamik dashboard** oluşturmak.

### 2️⃣ Veri Seti
- Dönem: **2013–2014**
- Değişkenler: Ürün, Ülke, Satış Fiyatı, Brüt Satış, Net Satış, Kâr, İndirim, Tarih
- Kaynak: Örnek "Retail Financial Data Sample" veri tabanı  
- Kullanım amacı: Eğitim ve veri analizi pratikleri

### 3️⃣ Veri Hazırlama Süreci
Veri temizleme ve dönüştürme adımları:
- Boş ve hatalı hücrelerin kaldırılması  
- Farklı sayı biçimlerinin (`.`, `,`) birleştirilmesi  
- Tarih biçimlerinin standardizasyonu  
- Negatif veya eksik değerlerin sıfırlanması  
- Metin alanlarında büyük/küçük harf ve boşluk düzenlemesi  

---

## 📊 Dashboard Yapısı

### 🔸 Genel Finansal Performans Dashboard’u
- KPI Kartları: Toplam Satış, Toplam Kâr, Ortalama Kâr Marjı  
- Çizgi Grafik: Satış trendi  
- Sütun Grafik: Ürün bazlı satış  
- Pasta Grafik: Ülke bazlı satış payı  
- Gruplu Sütun: İndirim oranı ve kâr ilişkisi  

### 🔸 Ürün & Ülke Analizi Dashboard’u
- Sütun Grafik: Ürün bazlı satışlar  
- Pasta Grafik: Ülkelerin toplam satış payı  
- Çizgi Grafik: Aylık ülke trendleri  
- Filtre (Slicer): Ürün veya ülke bazlı dinamik filtreleme  

### 🔸 Kârlılık Analizi Dashboard’u
- Sütun Grafik: Ürün bazlı toplam kâr  
- Pasta Grafik: Ülke bazlı kâr dağılımı  
- Çizgi Grafik: Aylık kâr trendi  
- Dağılım Grafiği: İndirim oranı vs. kâr marjı  
- Filtre (Slicer): Ülke, ürün ve yıl seçimi  

---

## 🧠 VBA Otomasyon Sistemi

### 🔹 Dashboard Geçiş Mekanizması
Dashboardlar arası geçişler **VBA kodu** ile otomatikleştirilmiştir.  
Kullanıcı menüden seçim yaptığında yalnızca ilgili panel görünür olur.  
Sistem tek sayfa üzerinde çalışır ve **şekil görünürlüğü (Shape Visibility)** yöntemiyle optimize edilmiştir.

### 🔹 Kod Yapısı (Özet)
- Dashboard geçişleri: `Worksheet_SelectionChange`  
- Görünürlük yönetimi: `ShowDashboard` fonksiyonu  
- Hata kontrolü: `On Error Resume Next`  
- Dinamik grup yönetimi (Genel, Ürün & Ülke, Kârlılık)

### 🔹 Kodun Avantajları
- Kullanıcı dostu geçiş yapısı  
- Hatasız ve optimize edilmiş görünürlük kontrolü  
- Yeni dashboard eklemeye uygun modüler tasarım  

---

## ✅ Sonuç ve Değerlendirme

### 🔸 Dashboard’un Sağladığı Avantajlar
- Tüm finansal göstergelere tek ekrandan erişim  
- Grafiksel görselleştirme ile trendlerin hızlı analizi  
- Dinamik, filtrelenebilir yapı  
- Kod destekli otomasyon sistemi  
- Genişletilebilir modüler tasarım  


### 🔸 Genel Değerlendirme
Bu proje, Excel’in yalnızca bir hesaplama aracı değil, aynı zamanda güçlü bir **raporlama ve veri görselleştirme platformu** olarak kullanılabileceğini göstermektedir.  
VBA desteği sayesinde etkileşimli, dinamik ve sade bir kullanıcı deneyimi sağlanmıştır.

---

## 👨‍💻 Hazırlayan
**Ahmet Dokazoğlu**  
📍 Ankara, Türkiye  
🔗 [GitHub Profilim](https://github.com/AhmetDokazoglu)  
🔗 [LinkedIn Profilim](https://www.linkedin.com/in/ahmet-dokazo%C4%9Flu-9660b2346/)

---

## 📎 Ek Dökümanlar
📄 [Proje Raporunun Word Versiyonu (İndir)](https://github.com/AhmetDokazoglu/Dynamic-Company-Financial-Performance-Dashboard--with-Excel-VBA-/raw/refs/heads/main/Dynamic%20Company%20Financial%20Performance%20Dashboard%20(with%20Excel%20VBA)(TR).docx)
