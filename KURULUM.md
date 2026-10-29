# DABS Night – Kurulum Rehberi

Toplam süre yaklaşık 30–45 dakika. Bilgisayardan yapman daha kolay olur. Her şey ücretsiz.

Bu klasörde şunlar var:

| Dosya | Ne işe yarar |
|---|---|
| `index.html`, `app.js` | Uygulamanın kendisi |
| `sw.js`, `manifest.webmanifest`, `icons/` | Telefona kurulma ve sinyal yokken açılma |
| `firebase-config.js` | Firebase bilgilerini buraya yapıştıracaksın |
| `firestore.rules` | Veriyi sadece onayladığın ekibin görmesini sağlayan güvenlik kuralları |

---

## 1. Firebase projesi (veritabanı ve giriş sistemi)

1. https://console.firebase.google.com adresine Google hesabınla gir.
2. **Create a project** → ad: `dabs-night` → Google Analytics'i **kapat** → Create.
3. Sol menü **Build → Authentication → Get started**:
   - **Email/Password** → Enable → Save.
   - **Google** → Enable → destek e-postanı seç → Save.
4. Sol menü **Build → Firestore Database → Create database**:
   - Location: **europe-west2 (London)**.
   - **Start in production mode** → Create.
5. Firestore içinde **Rules** sekmesi:
   - İçindekini tamamen sil.
   - `firestore.rules` dosyasının içeriğini yapıştır.
   - **Publish** de.
6. Sol üstteki dişli → **Project settings → General**:
   - Aşağıda **Your apps → Web (`</>`)** simgesine bas.
   - Nickname: `DABS Night`. "Firebase Hosting" kutusunu işaretleme. **Register app** de.
   - Ekranda çıkan `firebaseConfig = { ... }` içindeki değerleri `firebase-config.js` dosyasına, `PASTE_...` yazan yerlere kopyala.

> **Yönetici e-postası:** `firebase-config.js` ve `firestore.rules` dosyalarında iki adresin de yönetici olarak yazılı: `fatihyilmaz.en@gmail.com` ve `yilmazfatih05@gmail.com`. Hangisiyle girersen gir ekibi onaylayabilirsin. Listeyi değiştirirsen iki dosyada da aynı yap.

## 2. GitHub Pages (uygulamanın internet adresi)

1. https://github.com adresinde ücretsiz hesap aç.
2. Sağ üstte **+ → New repository**:
   - Ad: `dabs-night`.
   - **Public** seç. Ücretsiz Pages için gerekiyor. Kod görünür olur ama **veriler görünmez**; veriler Firebase'de, kurallarla korunuyor.
   - **Create repository** de.
3. **uploading an existing file** linkine tıkla ve bu klasördeki **tüm dosyaları ve `icons` klasörünü** sürükle. `firebase-config.js` dosyasının doldurulmuş hali olmalı. **Commit changes** de.
4. Repo içinde **Settings → Pages**:
   - Source: **Deploy from a branch**.
   - Branch: **main**, klasör **/(root)** → Save.
   - 1–2 dakika sonra adresin görünür: `https://KULLANICIADIN.github.io/dabs-night/`

## 3. Firebase'e adresi tanıt

Firebase → **Authentication → Settings → Authorized domains → Add domain** → `KULLANICIADIN.github.io` yaz (sadece bu kısım, `https` ve `/dabs-night` olmadan).

## 4. İlk giriş (sen)

1. Telefonda adresi aç → **Create an account** → yönetici e-postanla hesap aç. Google ile girersen bu adımı atlarsın.
2. E-postana gelen doğrulama linkine tıkla → uygulamada **I've verified**.
3. Bugünkü DABS'i yükle.

## 5. Telefona kurma

- **Android (Chrome):** menü ⋮ → **Install app** (ya da "Ana ekrana ekle").
- **iPhone (Safari):** alttaki Paylaş butonu → **Add to Home Screen**.

Bundan sonra ikonundan açılır, tarayıcı çubuğu olmaz. Bir kez sinyalle açıldıktan sonra sinyal yokken de açılır.

> iPhone'da ana ekrandan açılan uygulamada **Google ile giriş** çalışmayabilir. Ekip arkadaşlarına **e-posta ve şifre** ile hesap açmalarını söyle.

## 6. Ekibi ekleme

1. Ekip arkadaşlarına adresi gönder. **Create an account** ile hesap açarlar ve "Waiting for approval" ekranında beklerler.
2. Sen uygulamada **Archive → Team & approvals** ekranından **Approve** dersin. Ekranları kendiliğinden açılır.
3. Roller:
   - **Contributor:** iş işaretler, aktivite ekler.
   - **Viewer:** sadece görür.
   - **Admin:** ekip onaylar ve gece siler.
4. Ayrılan birini **Remove** ile çıkarırsın. Tekrar hesap açsa bile giremez; istersen **Restore** ile geri alırsın.

## Sinyal yokken

- İşaretlemeler telefonda kaydedilir. Üstte **"Saved on this phone · sends when signal returns"** yazar.
- Sinyal gelince kendiliğinden gönderilir ve **"Synced"** olur.
- **Yeni DABS yüklemek**, **gece silmek** ve **ilk giriş** için sinyal gerekir.

## Güncelleme

Ben kodu değiştirdiğimde değişen dosyaları GitHub'a tekrar yüklersin (aynı yöntemle, üzerine yazarak). Telefonlar bir sonraki açılışta yeni sürümü alır.

## Ücret ve limitler

Firebase'in ücretsiz **Spark** planı yeterli. Günlük 50.000 okuma ve 20.000 yazma hakkı var; 10–15 kişilik bir gece ekibi bunun çok altında kalır. Kredi kartı gerekmez.

## Veri

Veriler senin Google hesabındaki Firebase projesinde, Londra bölgesinde duruyor. Proje verisi olduğu için ENKA'nın IT kurallarına uygun olduğundan emin ol.
