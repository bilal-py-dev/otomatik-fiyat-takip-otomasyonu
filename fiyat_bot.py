import requests
from bs4 import BeautifulSoup
import re
import json
import os
import time
from plyer import notification
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

DOSYA_ADI = "urunler.json"

def verileri_yukle():
    if os.path.exists(DOSYA_ADI):
        with open(DOSYA_ADI, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def verileri_kaydet(liste):
    with open(DOSYA_ADI, "w", encoding="utf-8") as f:
        json.dump(liste, f, ensure_ascii=False, indent=4)

def fiyat_cek(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    try:
        cevap = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(cevap.text, "html.parser")
        baslik = soup.title.string.strip()[:60] if soup.title else "Ürün"
        metin = soup.get_text(" ", strip=True)
        eslesmeler = re.findall(r'(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)\s*(?:TL|₺)', metin)
        if eslesmeler:
            ham = eslesmeler[0].replace(".", "").replace(",", ".")
            return baslik, float(ham)
    except Exception:
        pass
    return None, None

def excel_raporu_olustur(liste):
    wb = Workbook()
    ws = wb.active
    ws.title = "Fiyat Takip Raporu"

    # Başlık Satırı ve Tasarım
    basliklar = ["Sıra", "Ürün Adı", "İlk Fiyat (TL)", "Son Fiyat (TL)", "Hedef Fiyat (TL)", "Durum", "Ürün Linki"]
    ws.append(basliklar)

    baslik_dolgu = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    baslik_yazi = Font(color="FFFFFF", bold=True)
    for hucre in ws[1]:
        hucre.fill = baslik_dolgu
        hucre.font = baslik_yazi

    for idx, u in enumerate(liste, 1):
        ilk_fiyat = u.get("ilk_fiyat", u["son_fiyat"])
        durum = "HEDEF YAKALANDI!" if u["son_fiyat"] <= u["hedef"] else "Takip Ediliyor"
        ws.append([idx, u["baslik"], ilk_fiyat, u["son_fiyat"], u["hedef"], durum, u["url"]])

    dosya_adi = f"Fiyat_Raporu_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(dosya_adi)
    return dosya_adi

while True:
    print("\n==============================================")
    print("   🛒 OTOMATİK FİYAT TAKİP & ANALİZ SİSTEMİ")
    print("==============================================")
    print("1. Yeni Ürün Ekle")
    print("2. Takip Edilen Ürünleri Gör")
    print("3. Listeden Ürün Sil")
    print("4. Excel Raporu Oluştur (.xlsx)")
    print("5. Otomatik Takibi Başlat (Arka Plan Kontrol)")
    print("6. Çıkış")
    secim = input("Seçiminiz (1-6): ")

    if secim == "1":
        link = input("Ürün Linki: ").strip()
        try:
            hedef = float(input("Hedef Fiyat (TL): ").replace(",", "."))
        except ValueError:
            print("❌ Lütfen geçerli bir sayı girin.")
            continue

        print("🔍 Ürün bilgileri çekiliyor...")
        baslik, mevcut = fiyat_cek(link)
        if mevcut:
            liste = verileri_yukle()
            liste.append({
                "baslik": baslik,
                "url": link,
                "ilk_fiyat": mevcut,
                "son_fiyat": mevcut,
                "hedef": hedef,
                "bildirildi": False
            })
            verileri_kaydet(liste)
            print(f"✅ Eklendi: {baslik} | Şu Anki Fiyat: {mevcut} TL | Hedef: {hedef} TL")
        else:
            print("❌ Bu linkten fiyat okunamadı veya site korumalı.")

    elif secim == "2":
        liste = verileri_yukle()
        if not liste:
            print("📭 Henüz takip edilen ürün yok.")
        for i, u in enumerate(liste, 1):
            durum = "✅ Hedef Altında" if u["son_fiyat"] <= u["hedef"] else "⏳ Bekleniyor"
            print(f"{i}. {u['baslik']} | Son: {u['son_fiyat']} TL | Hedef: {u['hedef']} TL [{durum}]")

    elif secim == "3":
        liste = verileri_yukle()
        if not liste:
            print("📭 Silinecek ürün bulunmuyor.")
            continue
        for i, u in enumerate(liste, 1):
            print(f"{i}. {u['baslik']} ({u['son_fiyat']} TL)")
        try:
            sil_no = int(input("Silmek istediğiniz ürün numarası (İptal için 0): "))
            if 1 <= sil_no <= len(liste):
                silinen = liste.pop(sil_no - 1)
                verileri_kaydet(liste)
                print(f"🗑️ Silindi: {silinen['baslik']}")
        except ValueError:
            print("❌ Geçersiz numara.")

    elif secim == "4":
        liste = verileri_yukle()
        if not liste:
            print("📭 Raporlanacak ürün bulunmuyor.")
        else:
            rapor_dosyasi = excel_raporu_olustur(liste)
            print(f"📊 Excel raporu başarıyla oluşturuldu: {rapor_dosyasi}")

    elif secim == "5":
        print("🔄 Otomatik kontrol başladı! Durdurmak için Ctrl + C yapabilirsiniz.")
        try:
            while True:
                liste = verileri_yukle()
                degisiklik_var = False
                for u in liste:
                    _, guncel = fiyat_cek(u["url"])
                    if guncel:
                        u["son_fiyat"] = guncel
                        degisiklik_var = True
                        print(f"[{time.strftime('%H:%M:%S')}] {u['baslik'][:35]}... -> {guncel} TL (Hedef: {u['hedef']} TL)")
                        if guncel <= u["hedef"] and not u.get("bildirildi", False):
                            print(f"🚨 İNDİRİM YAKALANDI! {u['baslik']} -> {guncel} TL")
                            notification.notify(
                                title="🚨 FİYAT DÜŞTÜ!",
                                message=f"{u['baslik']} artık {guncel} TL!",
                                timeout=10
                            )
                            u["bildirildi"] = True
                if degisiklik_var:
                    verileri_kaydet(liste)
                time.sleep(60)
        except KeyboardInterrupt:
            print("\n⏹️ Otomatik takip durduruldu, ana menüye dönüldü.")

    elif secim == "6":
        print("👋 Programdan çıkılıyor...")
        break
