# AUTOMATYZACJA ROZLICZEŃ KWARTALNYCH DLA PRODUCENTÓW
# Autor projektu / wdrożenie: Katarzyna Kalinowska

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

# 1. Ścieżka do pulpitu
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

# Automatyczne wyszukanie pliku źródłowego na pulpicie (zabezpieczenie przed błędem nazwy)
pliki_na_pulpicie = os.listdir(desktop_path)
plik_zrodlowy = None
for f in pliki_na_pulpicie:
  if f.startswith("Zestawienie_obrotow") and f.endswith(".xlsx"):
    plik_zrodlowy = f
    break

if not plik_zrodlowy:
  raise FileNotFoundError(
      "❌ Nie znaleziono pliku źródłowego z zestawieniem na pulpicie!"
  )

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

# Pobranie unikalnych par: Producent -> E-mail
producenci = df[["Producent", "E-mail"]].drop_duplicates().values

print(
    f"🚀 [START SYSTEMU]: Znaleziono {len(producenci)} unikalnych producentów."
)
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

    # Filtrowanie obrotów dedykowanych dla danego producenta
    df_producent = df[df["Producent"] == producent]
    suma_netto = df_producent["Zakup netto"].sum()

    # Generowanie dedykowanego pliku Excel dla kontrahenta
    safe_name = str(producent).replace(" ", "_")
    output_file_name = f"Zestawienie_{safe_name}.xlsx"
    output_file_path = os.path.join(output_folder_path, output_file_name)
    df_producent.to_excel(output_file_path, index=False)

    # Konstruowanie wiadomości e-mail
    msg = MIMEMultipart()
    msg["From"] = NADAWCA_EMAIL
    msg["To"] = email
    msg["Subject"] = f"Rozliczenie obrotów kwartalnych – {producent}"

    tresc = f"""Dzień dobry,

W załączniku przesyłamy zestawienie obrotów za miniony kwartał dla firmy {producent} (Łączna wartość zakupów netto: {suma_netto:,.2f} zł).

Uprzejmie prosimy o weryfikację. Zgodnie z przyjętą procedurą, brak odpowiedzi lub zgłoszenia uwag do 8. dnia bieżącego miesiąca oznacza pełną akceptację obrotów do rozliczenia.

Pozdrawiamy,
Dział Handlowy"""

    msg.attach(MIMEText(tresc, "plain", "utf-8"))

    # Dołączenie spersonalizowanego pliku raportu z poprawnym kodowaniem nazwy (brak "noname")
    with open(output_file_path, "rb") as attachment:
      part = MIMEBase("application", "octet-stream")
      part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        "attachment",
        filename=Header(output_file_name, "utf-8").encode(),
    )
    msg.attach(part)

    # Wysyłka wiadomości przez serwer
    server.sendmail(NADAWCA_EMAIL, email, msg.as_string())
    print(f"✅ [WYSŁANO & UTWORZONO]: {producent} ({email})")
    print(f"   ↳ Kwota netto: {suma_netto:,.2f} zł | Plik: {output_file_name}")
    print("-" * 70)

  server.quit()
  print(
      "\n✨ PROCES ZAKOŃCZONY SUKCESEM: Wszystkie raporty wygenerowane i"
      " rozesłane."
  )

except Exception as e:
  print(f"❌ [BŁĄD KRYTYCZNY]: {e}")
