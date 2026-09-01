# 🐍 Automated Quarterly Financial Reporting & Email System

## 🎯 Cel projektu i Wartość Biznesowa (Business Value)
Projekt polega na automatyzacji procesu rozliczeń obrotów kwartalnych z producentami. Dotychczasowy proces wymagał ręcznego filtrowania danych, tworzenia osobnych plików Excel dla każdego kontrahenta, pilnowania porządku na dysku oraz ręcznego redagowania i wysyłania wiadomości e-mail z załącznikami.

Nowe autorskie rozwiązanie w Pythonie redukuje ten czas z kilku godzin do kilku sekund, całkowicie eliminuje błąd ludzki i dba o pełną archiwizację danych.

* **⏱️ Oszczędność czasu:** Przetworzenie wszystkich producentów, wyliczenie sum netto i wysyłka dzieją się natychmiastowo.
* **🛡️ Bezpieczeństwo danych:** Automatyczne parowanie adresu e-mail z odpowiednim kontrahentem i kwotą zakupów netto.
* **📁 Porządek w systemie:** Automatyczne tworzenie dedykowanego folderu archiwizacyjnego na pulpicie opatrzonego dzisiejszą datą.

---

## 🛠️ Wyzwania techniczne i próby pełnej automatyzacji systemowej
W trakcie pracy nad projektem podjęłam próbę pełnej integracji rozwiązania z konsolą kupca oraz automatycznego uruchamiania skryptu w tle (harmonogram zadań / PowerShell / RPA). 

Niestety, napotkałam na blokady wynikające ze sztywnego zabezpieczenia systemowego konsoli oraz braku dostępu do wykupionego modułu harmonogramów zautomatyzowanych w środowisku. System odrzucał próby zewnętrznej iniekcji procesów roboczych ze względów bezpieczeństwa.

**Rozwiązanie:** Stworzyłam stabilny, niezależny i w 100% bezpieczny mechanizm w Pythonie, który uruchamia się na żądanie z poziomu lokalnego środowiska. Eliminuje to ryzyko naruszenia polityki bezpieczeństwa systemu, a jednocześnie w ułamku sekundy realizuje cały proces podziału plików, archiwizacji i masowej wysyłki mailowej.

---

## 📸 Podgląd działania systemu

### 1. Struktura plików na pulpicie (Automatyczna archiwizacja)
*(Tutaj wgraj screen ze strukturą folderów na pulpicie)*

### 2. Przebieg procesu w konsoli
*(Tutaj wgraj screen z konsoli)*

### 3. Widok odebranej wiadomości e-mail przez kontrahenta
*(Tutaj wgraj screen maila)*

---

## 💻 Kod źródłowy (Python Core Logic)
Główny skrypt odpowiedzialny za przetwarzanie danych za pomocą biblioteki `pandas` oraz wysyłkę raportów przez serwer `SMTP`:

```python
from datetime import datetime
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import smtplib
import pandas as pd

# --- KONFIGURACJA GŁÓWNA ---
NADAWCA_EMAIL = "info.gepart@gmail.com"
HASLO_APLIKACJI = "*****************"

# 1. Ścieżka do pulpitu i automatyczne wyszukanie pliku źródłowego
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
pliki_na_pulpicie = os.listdir(desktop_path)
plik_zrodlowy = None
for f in pliki_na_pulpicie:
    if f.startswith("Zestawienie_obrotow") and f.endswith(".xlsx"):
        plik_zrodlowy = f
        break

if not plik_zrodlowy:
    raise FileNotFoundError("❌ Nie znaleziono pliku źródłowego z zestawieniem na pulpicie!")

file_path = os.path.join(desktop_path, plik_zrodlowy)
print(f"📂 [WCZYTANO PLIK ŹRÓDŁOWY]: {plik_zrodlowy}\n")

# 2. Utworzenie profesjonalnego folderu archiwizacyjnego z datą dzienną
dzis = datetime.now().strftime("%Y-%m-%d")
output_folder_name = f"Rozliczenia_Kwartalne_{dzis}"
output_folder_path = os.path.join(desktop_path, output_folder_name)
os.makedirs(output_folder_path, exist_ok=True)

# 3. Wczytanie danych z raportu
df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()
producenci = df[["Producent", "E-mail"]].drop_duplicates().values

print(f"🚀 [START SYSTEMU]: Znaleziono {len(producenci)} unikalnych producentów.")
print(f"📁 [ARCHIWIZACJA]: Folder docelowy: '{output_folder_name}'\n")
print("=" * 70)

# 4. Nawiązanie bezpiecznego połączenia z serwerem SMTP Gmail
try:
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(NADAWCA_EMAIL, HASLO_APLIKACJI)
    
    # 5. Główna pętla biznesowa procesująca każdego partnera
    for producent, email in producenci:
        if pd.isna(producent) or pd.isna(email):
            continue
            
        df_producent = df[df["Producent"] == producent]
        suma_netto = df_producent["Zakup netto"].sum()
        
        safe_name = str(producent).replace(" ", "_")
        output_file_name = f"Zestawienie_{safe_name}.xlsx"
        output_file_path = os.path.join(output_folder_path, output_file_name)
        df_producent.to_excel(output_file_path, index=False)
        
        msg = MIMEMultipart()
        msg["From"] = NADAWCA_EMAIL
        msg["To"] = email
        msg["Subject"] = f"Rozliczenie obrotów kwartalnych – {producent}"
        
        tresc = f"""Dzień dobry,

W załączeniu przesyłamy zestawienie obrotów za miniony kwartał dla firmy {producent} (Łączna wartość zakupów netto: {suma_netto:,.2f} zł).

Uprzejmie prosimy o weryfikację. Zgodnie z przyjętą procedurą, brak odpowiedzi lub zgłoszenia uwag do 8. dnia bieżącego miesiąca oznacza pełną akceptację obrotów do rozliczenia.

Pozdrawiamy,
Dział Handlowy"""
        
        msg.attach(MIMEText(tresc, "plain", "utf-8"))
        
        with open(output_file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment", filename=Header(output_file_name, "utf-8").encode())
        msg.attach(part)
        
        server.sendmail(NADAWCA_EMAIL, email, msg.as_string())
        print(f"✅ [WYSŁANO & UTWORZONO]: {producent} ({email})")
        print(f"   ↳ Kwota netto: {suma_netto:,.2f} zł | Plik: {output_file_name}")
        print("-" * 70)
        
    server.quit()
    print("\n✨ PROCES ZAKOŃCZONY SUKCESEM: Wszystkie raporty wygenerowane i rozesłane.")
    
except Exception as e:
    print(f"❌ [BŁĄD KRYTYCZNY]: {e}")


**[📥 Pobierz pełną dokumentację projektową oraz zrzuty ekranu (PDF)](https://raw.githubusercontent.com/katarzyna90kalinowska/Python_Quarterly_Reporting_Automation/main/Automatyzacja_wysylki_raportow.pdf)**


