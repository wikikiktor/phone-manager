# Baza Telefonów / Phone Fleet Manager

[PL](#instrukcja-obsługi-programu-język-polski) | [EN](#user-manual-english)

---

# Instrukcja Obsługi Programu (Język Polski)

**Baza Telefonów** to gotowy, samodzielny program dla systemu Windows (`BazaTelefonow.exe`) służący do ewidencji telefonów służbowych, kart SIM, historii zdarzeń, notatek serwisowych oraz błyskawicznego generowania dokumentów zdawczo-odbiorczych w formatach PDF i DOCX (Word).

Aplikacja **nie wymaga instalacji** środowiska deweloperskiego, żadnych bibliotek ani zewnętrznych baz danych – jest gotowa do uruchomienia od razu po pobraniu.

---

## 1. Uruchomienie i dane programu

### 1.1. Jak uruchomić aplikację?
* Pobierz plik `BazaTelefonow.exe`.
* Umieść go w dowolnym wygodnym folderze (np. na Pulpicie lub w folderze Dokumenty).
* Uruchom program dwuklikiem.

### 1.2. Gdzie przechowywane są Twoje dane?
Program korzysta z wbudowanej, autonomicznej bazy SQLite (`telefony.db`). Plik bazy tworzy się i aktualizuje automatycznie w bezpiecznym profilu użytkownika:

```text
%APPDATA%\BazaTelefonow\telefony.db
```
*Pełna ścieżka:*  
`C:\Użytkownicy\<Twoja_Nazwa_Użytkownika>\AppData\Roaming\BazaTelefonow\telefony.db`

> **Wskazówka:** Nawet jeśli przeniesiesz lub podmienisz sam plik `.exe` na nowszą wersję, Twoje dane pozostaną nienaruszone w profilu systemowym.

---

## 2. Podręcznik Użytkownika

### 2.1. Dodawanie nowego telefonu do bazy
1. Kliknij niebieski przycisk **`+ Nowy telefon`** w prawym górnym rogu okna.
2. Wypełnij formularz po prawej stronie. Pola obowiązkowe:
   * **Model tel** (np. *Samsung Galaxy A54*)
   * **Nr telefonu** (np. *500100200* lub *+48500100200* – program automatycznie sformatuje numer z myślnikami).
3. Uzupełnij pozostałe dane wedle uznania: *Nr IMEI*, *Nr SIM*, *Nr seryjny*, *Osoba użytkująca*, *Osoba odpowiedzialna*, *Rodzaj*.
4. Zaznacz dołączone akcesoria w sekcji **Wyposażenie** (*Ładowarka*, *Etui*, *Kabel USB*, *Szkło ochronne*).
5. Wybierz **Stan baterii** oraz opcjonalnie zaznacz **Wysłany protokół** (jeśli dokument został już odebrany).
6. Kliknij zielony przycisk **`💾 Zapisz zmiany`**. Telefon pojawi się na liście po lewej stronie, a w jego historii zostanie odnotowany fakt dodania.

### 2.2. Edycja i automatyczny rejestr zmian
* Kliknięcie na dowolny telefon z listy po lewej stronie ładuje jego dane oraz historię zdarzeń.
* Zmodyfikuj dowolne pole (np. zmiana użytkownika, numeru SIM, wymiana ładowarki) i kliknij **`💾 Zapisz zmiany`**.
* Program samoczynnie wykrywa różnice i rejestruje wpis w dzienniku zdarzeń (np. *Zmieniono Osoba użytkująca: 'Jan Kowalski' ➔ 'Anna Nowak'* wraz z dokładną datą i godziną).

### 2.3. Dziennik zdarzeń, serwis i notatki
W dolnej prawej części okna znajduje się dedykowana historia dla wybranego urządzenia:
* **Dodawanie wpisu:** Kliknij **`+ Dodaj wpis`**, wybierz kategorię (*Notatka / Uwaga*, *Zmiana użytkownika*, *Awaria / Serwis*, *Wydanie*, *Zwrot*, *Wymiana karty SIM*, *Inne*) i wprowadź treść.
* **Notatki/uwagi:** Zaznacz wybrany wpis w historii i kliknij **`✏️ Edytuj uwagę`** (lub kliknij dwukrotnie dany wiersz), aby dopisać dodatkowe adnotacje (np. numer zgłoszenia serwisowego, koszt naprawy).

### 2.4. Wyszukiwarka i sortowanie
* **Wyszukiwarka w czasie rzeczywistym:** Wpisz szukaną frazę w pole u góry. Program filtruje urządzenia po numerze telefonu (również wpisanym ciągiem bez spacji i myślników), modelu, nazwisku użytkownika lub numerze IMEI.
* **Przycisk `✕`:** Błyskawicznie czyści pole wyszukiwania.
* **Sortowanie:** Kliknij nagłówek kolumny na liście telefonów (*Nr telefonu*, *Użytkownik*, *Model*) lub w tabeli historii, aby posortować rekordy rosnąco lub malejąco (strzałki `▲` / `▼`).

### 2.5. Kosz i bezpieczne usuwanie (Kosz vs Usuwanie trwałe)
* **Przeniesienie do kosza:** Kliknij czerwony przycisk **`🗑️ Do kosza`**. Telefon zniknie z aktywnej listy, ale jego dane i pełna historia zostaną zachowane.
* **Podgląd kosza:** Zaznacz pole **`Pokaż kosz (usunięte)`** na górnym pasku.
* **Przywracanie urządzenia:** W widoku kosza wybierz telefon i kliknij **`↩️ Przywróć telefon`**.
* **Trwałe skasowanie z bazy:** W widoku kosza dostępny jest przycisk **`❌ Usuń trwale z bazy`** – bezpowrotnie usuwa urządzenie i całą jego historię z bazy danych.

### 2.6. Dane firmy do protokołów
Przed wydrukiem dokumentów ustaw dane pracodawcy:
1. Rozwiń menu **`⚙️ Opcje bazy ▾`** i kliknij **`🏢 Dane firmy / pracodawcy`**.
2. Wpisz nazwę firmy, adres oraz NIP.
3. Kliknij **`💾 Zapisz dane`**. Informacje zostaną zapamiętane i będą automatycznie wstawiane do nagłówków wszystkich protokołów.

### 2.7. Generowanie gotowych protokołów (PDF i Word DOCX)
1. Kliknij **`📋 Protokół`** na górnym pasku (jeśli wcześniej zaznaczono telefon na liście, zostanie wybrany automatycznie).
2. Wybierz typ: **Protokół przekazania** lub **Protokół zwrotu**.
3. Sprawdź podgląd danych (dane pracownika, numer IMEI, zaznaczone akcesoria, stan baterii).
4. Kliknij:
   * **`📄 Pobierz PDF`** – generuje gotowy do wydruku, estetyczny dokument A4 w formacie PDF.
   * **`📝 Pobierz DOCX`** – generuje edytowalny dokument w formacie programu Microsoft Word.
5. Wybierz folder zapisu na dysku.

### 2.8. Import i Eksport arkuszy Excel (.xlsx)
* **Eksport do Excela:** Kliknij **`📊 Eksport Excel`**. Program utworzy plik `.xlsx` zawierający dwa arkusze:
  * *Telefony* – kompletna, estetycznie ostylowana tabela aktualnie widocznych urządzeń z autofiltrem i zamrożonym nagłówkiem.
  * *Historia zdarzeń* – powiązany dziennik operacji i uwag.
* **Import z Excela:** Kliknij **`📥 Import Excel`** i wskaż plik z dysku. Kreator pozwoli Ci łatwo dopasować kolumny z Twojego dotychczasowego arkusza do pól w programie. Jeśli plik posiada zakładkę *Historia zdarzeń*, przeniesie również wpisy historyczne.

### 2.9. Kopia zapasowa i przywracanie bazy (Backup)
W rozwijanym menu **`⚙️️ Opcje bazy ▾`**:
* **`💾 Utwórz kopię zapasową bazy (.db)`** – zapisuje wierną kopię zapasową całej bazy danych w wybranym miejscu (np. dysk sieciowy, pendrive).
* **`🔄 Przywróć bazę z pliku (.db)`** – pozwala odtworzyć dane z wcześniej wykonanej kopii.

---

## 3. Rozwiązywanie częstych problemów

* **Komunikat „Błąd zapisu” przy imporcie / eksporcie / tworzeniu protokołu:**  
  Upewnij się, że plik o tej samej nazwie nie jest w tym momencie otwarty w programie Excel, Word lub Adobe Acrobat Reader. Zamknij otwarty dokument i spróbuj ponownie.
* **Przeniesienie programu na nowy komputer:**  
  Wystarczy na nowym komputerze skopiować plik wykonywalny programu oraz wczytać dane z kopii za pomocą opcji **`⚙️ Opcje bazy ▾` ➔ `🔄 Przywróć bazę z pliku (.db)`**.

---
---

# User Manual (English)

**Phone Fleet Manager** is a standalone, portable Windows application (`BazaTelefonow.exe`) engineered for corporate mobile device tracking, SIM card management, automated change auditing, and one-click PDF & DOCX handover/return protocol generation.

**No runtime setup, installer, or third-party database engines required.**

---

## 1. Launch & Data Storage

### 1.1. How to Run
* Download `BazaTelefonow.exe`.
* Place it in any directory (Desktop, Documents, or Shared Drive).
* Double-click the executable to start the application.

### 1.2. Database Location
The application maintains a local SQLite database (`telefony.db`) automatically created in your user directory:

```text
%APPDATA%\BazaTelefonow\telefony.db
```
*Full Path:*  
`C:\Users\<Your_Username>\AppData\Roaming\BazaTelefonow\telefony.db`

---

## 2. Feature Guide

### 2.1. Adding a Device
1. Click **`+ Nowy telefon`** (*+ New Phone*) on the top bar.
2. Complete the form fields. Mandatory fields: **Model tel** and **Nr telefonu**.
   * Phone numbers are formatted with standard hyphen notation automatically.
3. Check bundled accessories (*Ładowarka / Charger*, *Etui / Case*, *Kabel USB*, *Szkło ochronne / Screen protector*).
4. Select battery health condition and flag whether the signed protocol has been delivered.
5. Click **`💾 Zapisz zmiany`** (*Save changes*).

### 2.2. Editing & Automated Audit Trail
* Select any record from the left-side list to view its properties and event log.
* Modify values and click **`💾 Zapisz zmiany`**. The system detects changes and records an entry in the event log showing the previous and new values with timestamps.

### 2.3. Event Log & Notes
* **Add Event:** Click **`+ Dodaj wpis`** (*+ Add Entry*) to log events (e.g. Service/Repairs, SIM card swaps, User reassignments).
* **Notes:** Select an entry in the history table and click **`✏️ Edytuj uwagę`** (*Edit Note*) or double-click the row to add or edit notes.

### 2.4. Search & Sorting
* **Live Search:** Instant filtering across phone numbers (with or without dashes), device models, assigned users, and IMEI numbers.
* **Clear Search (`✕`):** Clears search query in one click.
* **Column Sorting:** Click any column header in the device list or history log to sort ascending or descending (`▲` / `▼`).

### 2.5. Trash Management (Soft vs Hard Delete)
* **Soft Delete:** Click **`🗑️ Do kosza`** (*To Trash*) to remove the device from the active view while preserving its audit history.
* **View Trash:** Check **`Pokaż kosz (usunięte)`** (*Show Trash*) on the top toolbar.
* **Restore:** Select an item in trash and click **`↩️ Przywróć telefon`** (*Restore Phone*).
* **Hard Delete:** Inside the trash view, click **`❌ Usuń trwale z bazy`** (*Delete Permanently*) to irreversibly purge the device and its history.

### 2.6. Employer Company Details
1. Open the dropdown **`⚙️ Opcje bazy ▾`** and select **`🏢 Dane firmy / pracodawcy`**.
2. Enter the Company Name, Address, and Tax ID (NIP).
3. Click **`💾 Zapisz dane`**. Information will automatically appear on generated legal protocols.

### 2.7. Protocol Generation (PDF & DOCX)
1. Click **`📋 Protokół`** on the top toolbar.
2. Choose **Protokół przekazania** (*Handover*) or **Protokół zwrotu** (*Return*).
3. Review the preview pane.
4. Export:
   * **`📄 Pobierz PDF`** – exports a formatted A4 PDF.
   * **`📝 Pobierz DOCX`** – exports an editable Microsoft Word document.

### 2.8. Excel Import & Export (.xlsx)
* **Export to Excel:** Click **`📊 Eksport Excel`** to generate a styled workbook containing active devices and their event logs on separate sheets.
* **Import from Excel:** Click **`📥 Import Excel`** to launch the interactive column mapper for spreadsheets.

### 2.9. Database Backup & Restore
Under **`⚙️ Opcje bazy ▾`**:
* **`💾 Utwórz kopię zapasową bazy (.db)`** – exports a point-in-time snapshot of the database.
* **`🔄 Przywróć bazę z pliku (.db)`** – restores all database tables from an existing backup file.
