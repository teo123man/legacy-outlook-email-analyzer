"""Klasik Outlook için hazır komutlar: posta kutusu listeleme, arama, okuma, taslak.

Claude her oturumda Outlook COM kodunu yeniden yazmasın diye var. Mail GÖNDERMEZ,
taşımaz, silmez, okundu işaretlemez. Tek yazma işlemi `taslak` (Taslaklar'a kaydeder).

Örnekler (PowerShell, proje klasöründe):
  .venv\\Scripts\\python.exe outlook_araci.py kutular
  .venv\\Scripts\\python.exe outlook_araci.py ara --gonderen "Ahmet" --gun 14
  .venv\\Scripts\\python.exe outlook_araci.py ara --metin "vekalet" --klasor tum
  .venv\\Scripts\\python.exe outlook_araci.py oku <ID> --alintisiz
  .venv\\Scripts\\python.exe outlook_araci.py konusma <ID>
  .venv\\Scripts\\python.exe outlook_araci.py taslak --kime a@b.com --konu "X" --govde-dosya metin.txt
  .venv\\Scripts\\python.exe outlook_araci.py taslak --yanit <ID> --herkese --govde "..." --dene
"""
import argparse
import html
import os
import re
import sys
from datetime import datetime, timedelta

import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

KLASORLER = {"gelen": 6, "giden": 5, "taslak": 16, "silinen": 3}
OL_MAIL = 43
OL_DISCARD = 1

ALINTI = re.compile(
    "|".join([
        r"^\s*-{2,}\s*(?:original message|orijinal mesaj|özgün ileti|orijinal ileti)\s*-{2,}\s*$",
        r"^\s*_{10,}\s*$",
        r"^\s*(?:from|kimden)\s*:.*\n\s*(?:sent|date|gönderilen|gönderilme tarihi|tarih)\s*:",
        r"^.{0,200}\b(?:wrote|yazdı)\s*:\s*$",
    ]),
    re.MULTILINE | re.IGNORECASE,
)


def hata(mesaj):
    print("HATA: " + mesaj, file=sys.stderr)
    sys.exit(1)


# ---------------------------------------------------------------- bağlantı

def baglan(kutu_adi):
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    if not kutu_adi:
        return ns, ns.DefaultStore
    adlar = [s.DisplayName for s in ns.Stores]
    istenen = kutu_adi.lower()
    eslesen = [a for a in adlar if a.lower() == istenen] or [
        a for a in adlar if istenen in a.lower() and not a.lower().startswith(("online archive", "public folders"))
    ]
    if len(eslesen) != 1:
        hata(f"'{kutu_adi}' tek bir posta kutusuyla eşleşmedi. Mevcut kutular:\n" + "\n".join("  - " + a for a in adlar))
    return ns, next(s for s in ns.Stores if s.DisplayName == eslesen[0])


def hesap_bul(ns, store):
    for hesap in ns.Accounts:
        try:
            if hesap.DeliveryStore.StoreID == store.StoreID:
                return hesap
        except Exception:
            pass
    return None


def posta_klasorleri(kok):
    """Kök altındaki tüm posta klasörleri (DefaultItemType 0 = mail)."""
    yigin = [kok]
    while yigin:
        klasor = yigin.pop()
        try:
            yigin.extend(list(klasor.Folders))
            if klasor.DefaultItemType == 0:
                yield klasor
        except Exception:
            continue


def klasorleri_sec(store, ad):
    if ad == "tum":
        return list(posta_klasorleri(store.GetRootFolder()))
    if ad in KLASORLER:
        return [store.GetDefaultFolder(KLASORLER[ad])]
    if ad == "arsiv":
        for klasor in store.GetRootFolder().Folders:
            if klasor.Name.lower() in ("archive", "arşiv", "arsiv"):
                return [klasor]
        hata("Arşiv klasörü bulunamadı.")
    # Inbox altındaki ya da kökteki bir klasör adı (alt klasörleri dahil)
    for kok in (store.GetDefaultFolder(6), store.GetRootFolder()):
        for klasor in posta_klasorleri(kok):
            if klasor.Name.lower() == ad.lower():
                return [klasor]
    hata(f"'{ad}' klasörü bulunamadı. Kullanılabilir: {', '.join(KLASORLER)}, arsiv, tum veya klasör adı.")


# ---------------------------------------------------------------- yardımcılar

def gonderen_adresi(item):
    try:
        if getattr(item, "SenderEmailType", "") == "EX":
            kullanici = item.Sender.GetExchangeUser()
            if kullanici is not None:
                return kullanici.PrimarySmtpAddress or ""
    except Exception:
        pass
    try:
        return item.SenderEmailAddress or ""
    except Exception:
        return ""


def temiz_govde(metin):
    if not metin:
        return ""
    metin = metin.replace("\x00", "")
    metin = re.sub(r"[\x01-\x08\x0b\x0c\x0e-\x1f]", "", metin)
    metin = metin.replace("\r\n", "\n").replace("\r", "\n")
    metin = re.sub(r"[ \t]+\n", "\n", metin)
    metin = re.sub(r"\n{3,}", "\n\n", metin)
    return metin.strip()


def alintiyi_kes(metin):
    eslesme = ALINTI.search(metin)
    return (metin[: eslesme.start()].rstrip() + "\n[... önceki yazışma kesildi]") if eslesme else metin


def tarih(item):
    try:
        return item.ReceivedTime.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return "?"


def sql(metin):
    return metin.replace("'", "''")


def filtre_kur(a):
    """DASL filtresi: Outlook'un kendi aramasıyla çalışır, mailleri tek tek gezmez."""
    kosullar = []
    if a.gun:
        sinir = (datetime.now() - timedelta(days=a.gun)).strftime("%Y-%m-%d %H:%M")
        kosullar.append(f"\"urn:schemas:httpmail:datereceived\" >= '{sinir}'")
    if a.gonderen:
        g = sql(a.gonderen)
        kosullar.append(
            f"(\"urn:schemas:httpmail:fromname\" LIKE '%{g}%' OR \"urn:schemas:httpmail:fromemail\" LIKE '%{g}%'"
            f" OR \"http://schemas.microsoft.com/mapi/proptag/0x5D01001F\" LIKE '%{g}%')"
        )
    if a.kime:
        k = sql(a.kime)
        kosullar.append(
            f"(\"urn:schemas:httpmail:displayto\" LIKE '%{k}%' OR \"urn:schemas:httpmail:displaycc\" LIKE '%{k}%')"
        )
    if a.konu:
        kosullar.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{sql(a.konu)}%'")
    if a.metin:
        m = sql(a.metin)
        kosullar.append(
            f"(\"urn:schemas:httpmail:subject\" LIKE '%{m}%' OR \"urn:schemas:httpmail:textdescription\" LIKE '%{m}%')"
        )
    if a.okunmamis:
        kosullar.append("\"urn:schemas:httpmail:read\" = 0")
    if a.ekli:
        kosullar.append("\"urn:schemas:httpmail:hasattachment\" = 1")
    return "@SQL=" + " AND ".join(kosullar) if kosullar else None


# ---------------------------------------------------------------- komutlar

def kutular(a):
    ns, varsayilan = baglan(None)
    for s in ns.Stores:
        isaret = "  (varsayılan)" if s.StoreID == varsayilan.StoreID else ""
        print(f"- {s.DisplayName}{isaret}")


def ara(a):
    ns, store = baglan(a.kutu)
    filtre = filtre_kur(a)
    sonuclar = []
    for klasor in klasorleri_sec(store, a.klasor):
        try:
            items = klasor.Items.Restrict(filtre) if filtre else klasor.Items
            items.Sort("[ReceivedTime]", True)
        except Exception as exc:
            print(f"(atlandı: {klasor.Name}: {exc})", file=sys.stderr)
            continue
        item = items.GetFirst()
        alinan = 0
        while item is not None and alinan < a.limit:
            if getattr(item, "Class", None) == OL_MAIL:
                sonuclar.append((tarih(item), klasor.Name, item))
                alinan += 1
            item = items.GetNext()
    sonuclar.sort(key=lambda s: s[0], reverse=True)
    sonuclar = sonuclar[: a.limit]
    print(f"Kutu: {store.DisplayName} | Klasör: {a.klasor} | {len(sonuclar)} sonuç (en yeni önce, en fazla {a.limit})")
    for i, (t, klasor_adi, item) in enumerate(sonuclar, 1):
        try:
            okunmadi = " [okunmadı]" if item.UnRead else ""
            ek = f" [{item.Attachments.Count} ek]" if item.Attachments.Count else ""
            kategori = f" {{{item.Categories}}}" if item.Categories else ""
            klasor_notu = f" <{klasor_adi}>" if a.klasor == "tum" else ""
            print(f"\n{i}. {t} | {item.SenderName} <{gonderen_adresi(item)}>{okunmadi}{ek}{kategori}{klasor_notu}")
            print(f"   Konu: {item.Subject}")
            if a.onizleme:
                govde = alintiyi_kes(temiz_govde(item.Body)).replace("\n", " ")
                print(f"   Önizleme: {govde[: a.onizleme]}")
            print(f"   ID: {item.EntryID}")
        except Exception as exc:
            print(f"\n{i}. okunamadı: {exc}")


def item_al(ns, store, entry_id):
    try:
        return ns.GetItemFromID(entry_id, store.StoreID)
    except Exception:
        try:
            return ns.GetItemFromID(entry_id)
        except Exception as exc:
            hata(f"Mail bulunamadı: {exc}")


def mail_yaz(item, a):
    print(f"Tarih : {tarih(item)}")
    print(f"Kimden: {item.SenderName} <{gonderen_adresi(item)}>")
    print(f"Kime  : {item.To}")
    if item.CC:
        print(f"Bilgi : {item.CC}")
    print(f"Konu  : {item.Subject}")
    if item.Categories:
        print(f"Kategori: {item.Categories}")
    if item.Attachments.Count:
        print("Ekler : " + ", ".join(item.Attachments.Item(i).FileName for i in range(1, item.Attachments.Count + 1)))
    govde = temiz_govde(item.Body)
    if a.alintisiz:
        govde = alintiyi_kes(govde)
    if a.uzunluk and len(govde) > a.uzunluk:
        govde = govde[: a.uzunluk] + f"\n[... {len(govde) - a.uzunluk} karakter daha; tamamı için --uzunluk 0]"
    print("\n" + govde)


def oku(a):
    ns, store = baglan(a.kutu)
    mail_yaz(item_al(ns, store, a.id), a)


def konusma(a):
    ns, store = baglan(a.kutu)
    item = item_al(ns, store, a.id)
    mailler = []
    try:
        tablo = item.GetConversation().GetTable()
        while not tablo.EndOfTable:
            satir = tablo.GetNextRow()
            try:
                mailler.append(ns.GetItemFromID(satir["EntryID"]))
            except Exception:
                pass
    except Exception:
        mailler = [item]
    mailler = [m for m in mailler if getattr(m, "Class", None) == OL_MAIL]
    mailler.sort(key=tarih)
    print(f"Konuşma: {item.ConversationTopic} | {len(mailler)} mail (eskiden yeniye)")
    for i, m in enumerate(mailler, 1):
        print(f"\n{'=' * 70}\n#{i}  ID: {m.EntryID}")
        mail_yaz(m, a)


def govde_html(metin):
    paragraflar = html.escape(metin).replace("\r\n", "\n").split("\n")
    return "".join(f"<p style='margin:0'>{p or '&nbsp;'}</p>" for p in paragraflar)


def taslak(a):
    ns, store = baglan(a.kutu)
    if a.govde_dosya:
        with open(a.govde_dosya, encoding="utf-8") as f:
            govde = f.read().strip("\n")
    elif a.govde is not None:
        govde = a.govde.replace("\\n", "\n")
    else:
        hata("--govde ya da --govde-dosya gerekli.")

    if a.yanit:
        asil = item_al(ns, store, a.yanit)
        mail = asil.Forward() if a.ilet else asil.ReplyAll() if a.herkese else asil.Reply()
        eski_html = mail.HTMLBody or ""
        yeni = govde_html(govde)
        eslesme = re.search(r"<body[^>]*>", eski_html, re.IGNORECASE)
        mail.HTMLBody = (eski_html[: eslesme.end()] + yeni + eski_html[eslesme.end():]) if eslesme else yeni + eski_html
    else:
        if not a.kime:
            hata("Yeni taslak için --kime gerekli (yanıt için --yanit <ID>).")
        mail = ns.Application.CreateItem(0)
        mail.Body = govde.replace("\n", "\r\n")

    if a.kime:
        mail.To = a.kime if not a.yanit else "; ".join(x for x in (mail.To, a.kime) if x)
    if a.cc:
        mail.CC = "; ".join(x for x in (mail.CC, a.cc) if x)
    if a.konu:
        mail.Subject = a.konu
    hesap = hesap_bul(ns, store)
    if hesap is not None:
        mail.SendUsingAccount = hesap
    for yol in a.ek or []:
        mail.Attachments.Add(os.path.abspath(yol))

    print(f"Kutu : {store.DisplayName} | Hesap: {hesap.SmtpAddress if hesap else '?'}")
    print(f"Kime : {mail.To}\nBilgi: {mail.CC}\nKonu : {mail.Subject}\n\n{govde}")
    if a.dene:
        mail.Close(OL_DISCARD)
        print("\n[--dene: taslak KAYDEDİLMEDİ]")
        return
    mail.Save()  # sadece Taslaklar'a kaydeder; asla Send() çağrılmaz
    print(f"\nTaslak kaydedildi (gönderilmedi). ID: {mail.EntryID}")


# ---------------------------------------------------------------- argümanlar

def main():
    p = argparse.ArgumentParser(description="Klasik Outlook: ara / oku / taslak (göndermez).")
    p.add_argument("--kutu", help="Posta kutusu adı (boş = profilin varsayılan kutusu)")
    alt = p.add_subparsers(dest="komut", required=True)

    alt.add_parser("kutular", help="Profildeki posta kutularını listele")

    s = alt.add_parser("ara", help="Mail ara (Outlook filtresiyle, hızlı)")
    s.add_argument("--metin", help="Konu veya gövdede geçen ifade")
    s.add_argument("--konu", help="Konuda geçen ifade")
    s.add_argument("--gonderen", help="Gönderen adı ya da adresinin bir parçası")
    s.add_argument("--kime", help="Kime/Bilgi alanında geçen ad ya da adres")
    s.add_argument("--gun", type=int, help="Son N gün")
    s.add_argument("--klasor", default="gelen", help="gelen, giden, taslak, silinen, arsiv, tum veya klasör adı")
    s.add_argument("--okunmamis", action="store_true")
    s.add_argument("--ekli", action="store_true", help="Sadece ekli mailler")
    s.add_argument("--limit", type=int, default=20)
    s.add_argument("--onizleme", type=int, default=0, metavar="N", help="Gövdenin ilk N karakterini göster")

    for ad, yardim in (("oku", "Tek maili oku"), ("konusma", "Bir mailin tüm yazışmasını oku")):
        s = alt.add_parser(ad, help=yardim)
        s.add_argument("id", help="ara çıktısındaki ID")
        s.add_argument("--alintisiz", action="store_true", help="Alıntılanan eski yazışmayı kes")
        s.add_argument("--uzunluk", type=int, default=6000, help="Gövde karakter sınırı (0 = sınırsız)")

    s = alt.add_parser("taslak", help="Taslak oluştur (GÖNDERMEZ)")
    s.add_argument("--yanit", metavar="ID", help="Bu maile yanıt taslağı")
    s.add_argument("--herkese", action="store_true", help="Yanıtı tümüne yap")
    s.add_argument("--ilet", action="store_true", help="--yanit ID'li maili ilet (forward)")
    s.add_argument("--kime")
    s.add_argument("--cc")
    s.add_argument("--konu")
    s.add_argument("--govde", help="Metin (\\n satır sonu olur)")
    s.add_argument("--govde-dosya", help="UTF-8 metin dosyası (uzun/çok satırlı gövde için tercih edin)")
    s.add_argument("--ek", action="append", help="Eklenecek dosya (birden çok kez verilebilir)")
    s.add_argument("--dene", action="store_true", help="Göster ama kaydetme")

    a = p.parse_args()
    {"kutular": kutular, "ara": ara, "oku": oku, "konusma": konusma, "taslak": taslak}[a.komut](a)


if __name__ == "__main__":
    main()
