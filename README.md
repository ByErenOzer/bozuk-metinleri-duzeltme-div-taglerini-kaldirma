# 🧹 Tetkik Sonucu Metin Temizleme

> Türkçe karakterleri bozmadan, HTML parçalarını sökerek okunur metin üretir. Çıktıda sayfa ve sütun düzeni birebir korunur; sadece `tetkik_sonucu` sütunu temizlenip `tetkik_sonucu_temiz` olarak güncellenir.

---

## 🎯 Amaç
- `tetkik_sonucu` sütunundaki bozuk metinleri HTML etiketlerinden arındırmak
- HTML entity’lerini (`&apos;`, `&#304;`, `&#305;`, `&#287;`, `&#252;`, `&nbsp;` …) doğru Türkçe karakterlere dönüştürmek
- Çalışma kitabının tüm sayfalarında düzeni ve diğer sütunları birebir korumak
- Temizlenmiş XLSX çıktısını üretmek

## 🧰 Kullanılan Teknolojiler ve Kütüphaneler
- Python 3.11
- `openpyxl`: Excel dosyalarını (XLSX) sayfa/sütun bazında düzenlerken yapıyı korumak için
- Standart kitaplıklar:
  - `html.unescape`: HTML entity’lerini çözer
  - `re`: HTML etiket temizliği ve boşluk/satır sonu normalizasyonu için düzenli ifadeler

## 🧠 Teknik Yaklaşım
- Çalışma kitabı `openpyxl.load_workbook` ile yüklenir; her sayfada başlık satırı taranır
- Başlığı `tetkik_sonucu` (tipoyu kapsamak için `tetkit_sonucu` da desteklenir) olan sütun bulunur ve başlık `tetkik_sonucu_temiz` yapılır
- İlgili sütunun tüm hücreleri aşağıdaki “temizlik kuralları”ndan geçirilir
- Kitaplık yeni dosya adına kaydedilir; diğer tüm içerik, sayfa adları ve sütun sırası aynen korunur

## 🧼 Temizlik Kuralları
- Satır sonu dönüşümleri: `<br>`, `</div>`, `</p>`, `</span>` → satır sonu (`\n`)
- Açılış etiketleri: `<div>`, `<p>`, `<span>` → kaldırma
- Kalan tüm HTML etiketleri: tamamen kaldırma (etiket gövdesi korunur)
- Entity çözümü: `html.unescape` ile en az iki tur çözüm (iç içe/çift kaçışlı metinler için)
- `\u00a0` (non‑breaking space) → normal boşluk
- Boşluk ve satır sonu normalizasyonu: birden fazla boşluk → tek boşluk; ardışık fazla satır sonu → en fazla iki satır sonu

## 🔁 Örnek Dönüşüm

**Girdi (Bozuk Metin):**
```html
<div>Spesimen T&#252;r&#252;: Lobektomi</div><div><br></div><div>Lateralite (Taraf): Sa&#287;</div><div><br></div><div>Spesimen A&#287;&#305;rl&#305;&#287;&#305;: 210 gr<br></div>
```

**Çıktı (Temizlenmiş Metin):**
```text
Spesimen Türü: Lobektomi
Lateralite (Taraf): Sağ
Spesimen Ağırlığı: 210 gr
```

**Diğer Örnekler:**
- `µl&apos;den` → `µl'den`
- `2 µg/ml&apos;den` → `2 µg/ml'den`
- `T&#252;m&#246;r&#252;n plevraya uzakl&#305;&#287;&#305;` → `Tümörün plevraya uzaklığı`

## ⚙️ Kurulum
```bash
python -V               # Python sürümünü doğrulayın
python -m pip install openpyxl
```

## 📌 Ek Script: Şifreli Excel (Password-Protected) için
Bu repoda ayrıca `2-clean_tektik_sonucu_password_excel.py` bulunur. Bu script, **parola korumalı** bir `.xlsx` dosyasını önce çözüp (decrypt), ardından tüm sayfalarda `tetkik_sonucu` / `tetkit_sonucu` sütununu temizleyerek çıktıyı yeni bir Excel dosyası olarak yazar.

Bu scriptte yaklaşım `pandas` + `openpyxl` üzerindendir:
- Excel dosyası bellek içine decrypt edilir
- Tüm sheet'ler tek tek okunur
- Hedef sütun temizlenir ve `tetkik_sonucu_temiz` olarak yeniden adlandırılır
- Her sheet çıktı dosyasına geri yazılır

Gerekli ek kütüphaneler:
```bash
python -m pip install pandas openpyxl msoffcrypto-tool
```

Notlar:
- Parola korumalı dosyalarda decrypt için script içinde parola kullanılır; kendi dosyanıza göre `password` değerini güncellemeniz gerekir.
- `src` ve `dst` dosya yolları scriptin en altındaki `__main__` bloğunda örnek olarak yer alır; kendi ortamınıza göre düzenleyin.

## 🔎 Doğrulama
- Yeni dosyayı açın ve her sayfada `tetkik_sonucu_temiz` başlığının bulunduğunu kontrol edin
- Metin içinde `<div>`, `&#NNN;`, `&apos;` gibi kalıntıların kalmadığını ve Türkçe karakterlerin doğru göründüğünü doğrulayın

## 📄 Kod Özeti
`clean_tetkik_sonucu.py` içindeki çekirdek fonksiyon:
```python
import re, html

def clean_text(s):
    if s is None:
        return s
    if not isinstance(s, str):
        s = str(s)
    t = s
    for _ in range(2):
        t2 = html.unescape(t)
        if t2 == t:
            break
        t = t2
    t = re.sub(r'(?i)<br\s*/?>', '\n', t)
    t = re.sub(r'(?i)</\s*(div|p|span)\s*>', '\n', t)
    t = re.sub(r'(?i)<\s*(div|p|span)[^>]*>', '', t)
    t = re.sub(r'(?i)<[^>]+>', '', t)
    t = t.replace('\u00a0', ' ')
    t = re.sub(r'\s+\n', '\n', t)
    t = re.sub(r'\n\s+', '\n', t)
    t = re.sub(r'\n{3,}', '\n\n', t)
    t = re.sub(r'[ \t]{2,}', ' ', t)
    t = t.strip()
    return t
```

## 🧪 Neden `openpyxl`?
- XLSX yapısını (sayfalar, hücre düzeni) korur; yalnızca hedef hücre değerleri değiştirilir
- Birden fazla sayfa ve farklı başlık sıralarına sahip dosyalarda güvenli çalışır
- Pandas yerine seçildi; çünkü biçim korunumu ve çok sayfalı kitaplarda başlık tarama/yenileme işleri için daha uygundur

## ✏️ Özelleştirme
- Başlık adı farklı olsun isterseniz `tetkik_sonucu_temiz` değerini script içinde değiştirebilirsiniz
- Yeni etiket türleri veya entity’ler eklemek için ilgili regex/temizlik adımlarına yeni kurallar ekleyebilirsiniz

## ✅ Sonuç
- Tüm sayfalarda `tetkik_sonucu`/`tetkit_sonucu` sütunları temizlenir ve çıktı dosyası oluşturulur
- Türkçe karakterler bozulmadan ve HTML kalıntıları olmadan okunabilir metin elde edilir

---

> İhtiyacınıza göre ek düzenlemeler (ek sütunlar, rapor üretimi, özel normalizasyon kuralları) hızlıca eklenebilir.
