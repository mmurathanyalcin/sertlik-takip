# Sertlik Takip Uygulamasi

PDF sertlik raporlarini surukle-birak ile Excel tablosuna otomatik aktaran masaustu programi.

## Ne Yapar?

Operator bir PDF sertlik raporunu programa surukleyip birakinca, program otomatik olarak:
- **Ebat** (dosya adindan)
- **Ortalama Sertlik (HV)** (PDF icindeki tum olcumlerin ortalamasi)
- **Musteri / Parti No** (PDF icinde varsa)
- **Operator** (PDF icinde varsa)
- **Dosya Adi**
- **Dosya Degisiklik Tarihi ve Saati**

bilgilerini cikarip, belirlediginiz Excel dosyasina yeni bir satir olarak ekler.

---

## KURULUM SECENEK 1 (Onerilen): Hazir EXE ile

Eger elinizde zaten `SertlikTakip.exe` dosyasi varsa:
1. Bu dosyayi istediginiz klasore (orn. masaustu) kopyalayin.
2. Cift tiklayip calistirin. Kurulum gerekmez, Python gerekmez.
3. Program ilk acildiginda sizden Excel dosyasi konumu soracak:
   - **"Evet"** derseniz: var olan Sertlik_Tablosu.xlsx dosyanizi secersiniz.
   - **"Hayir"** derseniz: yeni bir Excel dosyasi nereye kaydedilecegini sorar.
4. Bundan sonra PDF dosyalarini pencereye surukleyip birakmaniz yeterli.

---

## KURULUM SECENEK 2: Kendiniz EXE Olusturmak

Eger elinizde sadece kaynak kod (.py dosyasi) varsa ve kendi EXE'nizi
olusturmak istiyorsaniz, Windows bilgisayarda asagidaki adimlari izleyin:

### Adim 1 - Python Kurulumu
1. https://www.python.org/downloads/ adresine gidin.
2. En guncel surumu indirip kurun.
3. **COK ONEMLI:** Kurulum ekraninda alt kisimdaki
   **"Add python.exe to PATH"** kutucugunu isaretleyin.

### Adim 2 - Programi Derleme
1. Bu klasordeki tum dosyalari (sertlik_takip.py, requirements.txt,
   EXE_OLUSTUR.bat) ayni klasore koyun.
2. **EXE_OLUSTUR.bat** dosyasina cift tiklayin.
3. Acilan siyah pencerede islemler otomatik yapilacak (1-2 dakika surer).
4. Islem bitince `dist` klasoru icinde **SertlikTakip.exe** dosyasini bulacaksiniz.
5. Bu exe dosyasini istediginiz bilgisayara kopyalayip kullanabilirsiniz.

---

## KURULUM SECENEK 3: GitHub Actions ile Otomatik Derleme (Kurulum Gerektirmez)

Eger Windows bilgisayarda hicbir sey kurmak istemiyorsaniz:

1. Bu klasoru bir GitHub deposuna yukleyin (`.github/workflows/build.yml` dahil).
2. GitHub deponun **Actions** sekmesine girin.
3. "Build Windows EXE" is akisini calistirin (push sonrasi otomatik calisir,
   ya da "Run workflow" ile manuel baslatabilirsiniz).
4. Is bitince **Artifacts** bolumunden **SertlikTakip-Windows** dosyasini
   indirin — icinde hazir SertlikTakip.exe olacak.

Bu yontemde Windows bilgisayara hicbir sey kurmadan, GitHub'in kendi
sunuculari uzerinde derleme yapilir.

---

## Kullanim

1. Programi acin.
2. Excel dosyasini secin/olusturun (sadece ilk acilista sorulur, sonrasinda hatirlanir).
3. PDF dosyalarini (tek tek veya birden fazla secerek) pencereye surukleyip birakin.
4. Asagidaki "Islem Gecmisi" kutusunda her dosya icin ne yapildigini gorebilirsiniz.
5. Ayni dosya tekrar surulenirse program uyarir: "Bu dosya zaten kayitli, yine de eklensin mi?"

### Excel Dosyasini Degistirmek
Ust kisimdaki **"Excel Dosyasi Degistir"** butonuna tiklayarak istediginiz
zaman farkli bir Excel dosyasina gecebilirsiniz.

---

## Onemli Notlar

- **MUSTERI** ve **PARTI NO** alanlari olcum cihazinin PDF ciktisinda her
  zaman doldurulmuyor. Doluysa otomatik cekilir, degilse bos birakilir.
- **Ebat** bilgisi, PDF icindeki EBAT alani guvenilir olmadigi durumlarda
  (bos veya hatali) dosya adindan turetilir — operatorlerin dosyayi
  adlandirirken yazdigi ebat baz alinir.
- Dosya Degisiklik Tarihi/Saati, dosyanin Windows'taki kendi ozelliginden
  (son degistirilme zamani) okunur. Bu yuzden PDF'yi olusturduktan sonra
  baska bir yere tasimak veya kopyalamak bu degeri degistirebilir.
