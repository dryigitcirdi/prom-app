# PROM Takip Sistemi – Kurulum Kılavuzu

## Dosyalar

| Dosya | Açıklama |
|---|---|
| `form.html` | Hasta anketi (telefonda açılır) |
| `dashboard.html` | Doktor paneli (link üretme + sonuçları görme) |
| `appsscript.gs` | Google Sheets backend kodu |

---

## 1. Google Sheet + Apps Script Kurulumu (10 dakika)

### Adım 1 — Yeni Google Sheet oluşturun
1. [sheets.google.com](https://sheets.google.com) → Yeni boş tablo
2. Tabloya bir ad verin: **PROM Kayıtlar**

### Adım 2 — Apps Script editörünü açın
1. Sheet'te: **Uzantılar → Apps Script**
2. Açılan editörde `Code.gs` içindeki her şeyi silin
3. `appsscript.gs` dosyasının içeriğini buraya yapıştırın
4. Kaydedin (Ctrl+S / Cmd+S)

### Adım 3 — Deploy edin
1. Sağ üstte **Dağıt → Yeni dağıtım**
2. Tür: **Web uygulaması**
3. Ayarlar:
   - Şu kullanıcı olarak çalıştır: **Ben**
   - Erişim yetkisi: **Herkes** *(anonim POST için şart)*
4. **Dağıt** → Google hesabınızla izin verin
5. Çıkan URL'yi kopyalayın: `https://script.google.com/macros/s/XXXX/exec`

### Adım 4 — URL'yi her iki dosyaya girin
`form.html` ve `dashboard.html` içinde şu satırı bulun:

```js
const APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';
```

`'YOUR_APPS_SCRIPT_URL_HERE'` kısmını kopyaladığınız URL ile değiştirin.

---

## 2. Dosyaları Nerede Barındırmalısınız?

### En kolay yol — GitHub Pages (ücretsiz)
1. GitHub'da yeni repo oluşturun (public)
2. `form.html` ve `dashboard.html` dosyalarını yükleyin
3. Repo Ayarlar → Pages → Branch: main → /root → Kaydet
4. 2 dakika sonra siteniz hazır:
   - Hasta anketi: `https://KULLANICI.github.io/REPO/form.html`
   - Dr. paneli:   `https://KULLANICI.github.io/REPO/dashboard.html`

### Alternatif — Lokal kullanım
Dosyaları direkt tarayıcıda da açabilirsiniz:
```bash
cd prom-app
python3 -m http.server 8765
# http://localhost:8765/dashboard.html
```

---

## 3. Hasta Linki Gönderme

1. `dashboard.html` → **Link Oluştur** bölümü
2. Hasta adı + ameliyat tarihi girin → **Link Oluştur**
3. **WhatsApp** butonuna tıklayın → Hazır mesaj açılır
4. Hasta linke tıklar, anketi doldurur, veri otomatik Sheet'e düşer

### Link formatı
```
form.html?name=Ahmet Yılmaz&id=12345&surgery=2026-04-01&leg=sağ
```

---

## 4. Google Sheet'te Ne Görürsünüz?

| Sütun | İçerik |
|---|---|
| A-B | Tarih ve saat |
| C-H | Hasta bilgileri (ID, ad, tanı, ameliyat tarihi, **post-op gün**, taraf) |
| I-J | NRS ağrı (ortalama + en kötü) |
| K-L | IKDC skoru + ham veri |
| M-N | Lysholm skoru + ham veri |
| O | Tegner aktivite düzeyi |

IKDC ve Lysholm sütunları otomatik renklenir:
- 🟢 ≥ 80 → yeşil
- 🟡 60–79 → sarı  
- 🔴 < 60 → kırmızı

---

## 5. Sık Sorulan Sorular

**Hasta linki her seferinde farklı mı olmalı?**
Evet. Dashboard'dan her post-op değerlendirme için ayrı link oluşturun (2. hafta, 6. hafta, 3. ay vs.).

**Aynı hasta birden fazla doldurabilir mi?**
Evet, her doldurma ayrı satır olarak kaydedilir. `daysPostOp` sütunu hangi zaman diliminde olduğunu gösterir.

**Apps Script URL'si değişirse ne olur?**
Her yeni deployment yeni URL üretir. Sadece `form.html` ve `dashboard.html`'deki `APPS_SCRIPT_URL` satırını güncelleyin.
