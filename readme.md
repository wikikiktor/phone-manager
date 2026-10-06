# Baza Telefonów / Phone Fleet Manager

[PL](#instrukcja-obsługi-język-polski) | [EN](#user-manual-english)

---

# Instrukcja Obsługi (Język Polski)

**Baza Telefonów** to samodzielna aplikacja stacjonarna (`.exe`) przeznaczona do kompleksowego zarządzania flotą telefonów służbowych, kart SIM, historią zdarzeń oraz generowania dokumentów zdawczo-odbiorczych (PDF/DOCX).

---

## 1. Pierwsze uruchomienie i lokalizacja bazy danych

Program nie wymaga instalacji serwerów zewnętrznych — działa w oparciu o wbudowaną, autonomiczną bazę SQLite (`telefony.db`). 

### Gdzie znajduje się baza danych?
Plik bazy danych tworzony jest automatycznie podczas pierwszego uruchomienia programu w bezpiecznym katalogu profilu użytkownika:

* **System Windows (domyślnie dla pliku .exe):**
  ```text
  %APPDATA%\BazaTelefonow\telefony.db
  ```
  Pełna ścieżka:
  ```text
  C:\Użytkownicy\<Twoja_Nazwa_Użytkownika>\AppData\Roaming\BazaTelefonow\telefony.db
  ```
  *(Aby szybko otworzyć ten folder, wciśnij klawisze `Win + R`, wpisz `%APPDATA%\BazaTelefonow` i zatwierdź Enterem).*

* **Systemy Linux / macOS (uruchamianie awaryjne/środowisko bez zmiennej APPDATA):**
  ```text
  ~/.baza_telefonow/telefony.db
  ```

> **Wskazówka:** Aby zabezpieczyć swoje dane, wystarczy regularnie kopiować plik `telefony.db` lub korzystać z wbudowanej w aplikację funkcji tworzenia kopii zapasowej.

---

## 2. Główne moduły i funkcje programu

### 2.1. Dodawanie nowego telefonu
1. Kliknij przycisk **`+ Nowy telefon`** w górnym pasku.
2. Formularz po prawej stronie zostanie wyczyszczony.
3. Wypełnij pola (wymagane są: **Model tel** oraz **Nr telefonu**).
   * Numer telefonu zostanie automatycznie sformatowany do czytelnej postaci (np. `123-456-789` lub `+48 123-456-789`).
   * Zaznacz wyposażenie dodatkowe (ładowarka, etui, kabel USB, szkło ochronne).
   * Wybierz stan baterii oraz opcjonalnie odznacz/zaznacz „Wysłany protokół”.
4. Kliknij przycisk **`Zapisz zmiany`**. Telefon pojawi się na liście po lewej stronie, a w historii zostanie odnotowany fakt dodania urządzenia.

### 2.2. Edycja i automatyczny dziennik zmian
* Kliknięcie na telefon z listy wczytuje jego szczegóły oraz historię.
* Po zmianie dowolnego pola (np. zmiana użytkownika, numeru SIM, wyposażenia) i kliknięciu **`Zapisz zmiany`**, program automatycznie zarejestruje wpis w historii ze starą i nową wartością oraz dokładną datą modyfikacji.

### 2.3. Dziennik zdarzeń i notatki
* **Dodawanie zdarzenia:** Kliknij przycisk **`+ Dodaj wpis`** nad tabelą historii, wybierz kategorię (np. *Awaria / Serwis*, *Wymiana karty SIM*, *Zmiana użytkownika*) i wprowadź treść.
* **Dodawanie/edycja uwagi:** Kliknij dowolny wiersz w historii zdarzeń. Obok pojawi się przycisk **`+ Dodaj / edytuj uwagę`**, który pozwala dopisać adnotację do wybranego wpisu.

### 2.4. Wyszukiwanie i sortowanie
* **Szukanie:** Wpisz szukaną frazę w pole wyszukiwarki. Program przeszukuje bazę w czasie rzeczywistym pod kątem: numeru telefonu (z kreskami lub bez), modelu, użytkownika oraz numeru IMEI.
* **Sortowanie:** Kliknij nagłówek kolumny na liście telefonów lub w historii, aby posortować rekordy rosnąco lub malejąco (strzałki `▲` / `▼`).

### 2.5. Kosz i usuwanie (Soft Delete / Hard Delete)
* **Kosz (usuwanie miękkie):** Kliknij **`Przenieś do kosza`**. Telefon zniknie z głównej listy, ale jego dane i historia pozostaną zachowane.
* **Przeglądanie kosza:** Zaznacz opcję **`Pokaż kosz (usunięte)`** na górnym pasku.
* **Przywracanie:** Po zaznaczeniu telefonu w koszu kliknij **`Przywróć telefon`**.
* **Trwałe usunięcie (Hard Delete):** W widoku kosza dostępny jest przycisk **`Usuń trwale z bazy`** – bezpowrotnie usuwa urządzenie i całą powiązaną z nim historię.

### 2.6. Dane pracodawcy
Przed generowaniem protokołów uzupełnij dane firmy:
1. Kliknij **`Dane pracodawcy`**.
2. Wpisz nazwę firmy, adres oraz NIP.
3. Kliknij **`Zapisz`**. Dane zostaną zapamiętane na stałe dla wszystkich przyszłych protokołów.

### 2.7. Generowanie protokołów (PDF i DOCX)
Aplikacja pozwala generować gotowe do podpisu protokoły przekazania oraz zwrotu sprzętu służbowego:
1. Kliknij **`Stwórz protokół`** (lub najpierw zaznacz telefon na liście).
2. Wybierz typ: **Protokół przekazania** lub **Protokół zwrotu**.
3. Upewnij się, że wybrany jest właściwy telefon z listy rozwijanej.
4. Kliknij **`Eksportuj do PDF`** lub **`Eksportuj do DOCX`** (format programu Microsoft Word) i wskaż miejsce zapisu.

### 2.8. Eksport i Import Excel
* **Eksport do Excela:** Kliknij **`Eksportuj do Excela`**. Utworzy sformatowany plik `.xlsx` z dwoma arkuszami: listą aktualnie widocznych telefonów oraz pełną historią zdarzeń.
* **Import z Excela:** Kliknij **`Importuj z Excela`**, wskaż plik `.xlsx` / `.xls`. Pojawi się okno mapowania, które automatycznie dopasuje kolumny lub pozwoli Ci ręcznie przypisać nagłówki z pliku do pól bazy danych. Jeśli arkusz zawiera zakładkę *Historia zdarzeń*, możesz zaimportować również wpisy historyczne.

### 2.9. Kopia zapasowa (Backup i Przywracanie)
* **Utwórz kopię:** Kliknij **`Utwórz kopie`** i wskaż bezpieczną lokalizację (np. dysk sieciowy lub pendrive). Program wykona bezpieczny zrzut bazy w locie.
* **Przywróć kopię:** Kliknij **`Przywróć kopie`**, wybierz wcześniej zapisany plik `.db` i potwierdź ostrzeżenie. Aktualna baza zostanie zastąpiona danymi z kopii.

---

## 3. Rozwiązywanie problemów (FAQ)

1. **Błąd podczas zapisu protokołu / eksportu Excel:**
   * Upewnij się, że generowany plik o tej samej nazwie nie jest aktualnie otwarty w programie Word, Excel lub Adobe Acrobat Reader.
2. **Przenoszenie programu na inny komputer:**
   * Wystarczy skopiować plik wykonywalny `.exe` oraz plik bazy z `%APPDATA%\BazaTelefonow\telefony.db` do tej samej ścieżki na nowym komputerze (lub użyć opcji `Utwórz kopie` i `Przywróć kopie` w programie).

---
---

# User Manual (English)

**Phone Fleet Manager** is a standalone desktop application (`.exe`) designed for end-to-end management of corporate mobile devices, SIM cards, event auditing, and handover/return protocol generation (PDF/DOCX).

---

## 1. Initial Launch & Database Location

The application requires no external database server. It operates using a built-in, standalone SQLite database (`telefony.db`).

### Where is the database located?
The database file is created automatically on the first launch inside the user application directory:

* **Windows OS (standard for the .exe file):**
  ```text
  %APPDATA%\BazaTelefonow\telefony.db
  ```
  Full path:
  ```text
  C:\Users\<Your_Username>\AppData\Roaming\BazaTelefonow\telefony.db
  ```
  *(To access this folder directly, press `Win + R`, type `%APPDATA%\BazaTelefonow`, and press Enter).*

* **Linux / macOS (fallback path):**
  ```text
  ~/.baza_telefonow/telefony.db
  ```

> **Tip:** To back up your data manually, simply copy the `telefony.db` file to a secure location or use the built-in backup tool.

---

## 2. Key Modules & Usage Guide

### 2.1. Adding a New Phone
1. Click **`+ Nowy telefon`** (*+ New Phone*) on the top toolbar.
2. The form on the right panel will be cleared.
3. Fill in device details (required fields: **Model tel** and **Nr telefonu**).
   * Phone numbers are auto-formatted into clean standard notations (e.g., `123-456-789` or `+48 123-456-789`).
   * Select accessories provided with the phone (Charger, Case, USB Cable, Screen Protector).
   * Choose the battery health status and check if the protocol was delivered.
4. Click **`Zapisz zmiany`** (*Save Changes*). The record will appear in the table and an audit event will be logged.

### 2.2. Modifying Details & Automated Audit Log
* Selecting any phone from the list populates its current data and log history.
* Editing any field (e.g., changing the user, SIM card number, or accessories) and clicking **`Zapisz zmiany`** automatically inserts a detailed log entry indicating the previous and new values with timestamps.

### 2.3. Event Log and Notes
* **Adding custom event:** Click **`+ Dodaj wpis`** (*+ Add Entry*) above the history table, select the category (e.g., *Service/Repair*, *SIM Swap*, *User Change*), and submit the description.
* **Adding/Editing notes:** Click any entry in the event table. The button **`+ Dodaj / edytuj uwagę`** (*+ Add / Edit Note*) will appear, allowing you to attach remarks to that specific event.

### 2.4. Real-time Search & Column Sorting
* **Search:** Type in the top search bar. The list filters instantly by phone number (with or without dashes), model name, assigned user, or IMEI number.
* **Sorting:** Click any column header on the left list or the history table to toggle ascending/descending order (`▲` / `▼`).

### 2.5. Trash Bin (Soft Delete vs Hard Delete)
* **Soft Delete:** Click **`Przenieś do kosza`** (*Move to Trash*). The device is hidden from active searches but preserved in the archive.
* **View Trash:** Check **`Pokaż kosz (usunięte)`** (*Show Trash*) on the top bar.
* **Restore:** Select a trashed device and click **`Przywróć telefon`** (*Restore Phone*).
* **Hard Delete:** Inside the trash view, click **`Usuń trwale z bazy`** (*Delete Permanently*) to irreversibly purge the device and its history from the database.

### 2.6. Company / Employer Details
Before creating legal handover protocols, configure your company details:
1. Click **`Dane pracodawcy`** (*Employer Details*).
2. Enter the Company Name, Address, and Tax ID (NIP).
3. Click **`Zapisz`** (*Save*).

### 2.7. Handover & Return Protocol Generator (PDF & DOCX)
Create formal equipment handover or return documents:
1. Click **`Stwórz protokół`** (*Create Protocol*).
2. Choose the document type: **Protokół przekazania** (*Handover*) or **Protokół zwrotu** (*Return*).
3. Confirm the selected phone from the dropdown list.
4. Click **`Eksportuj do PDF`** or **`Eksportuj do DOCX`** and choose where to save the document.

### 2.8. Excel Import and Export
* **Export to Excel:** Click **`Eksportuj do Excela`**. Generates a styled `.xlsx` workbook containing active records and complete event logs on separate sheets.
* **Import from Excel:** Click **`Importuj z Excela`** and select your spreadsheet. An interactive mapper will automatically match headers and let you customize column bindings, including historical records.

### 2.9. Database Backup & Restoration
* **Create Backup:** Click **`Utwórz kopie`** (*Create Backup*) to generate a point-in-time snapshot of the database (`.db`).
* **Restore Backup:** Click **`Przywróć kopie`** (*Restore Backup*) and choose a previously saved `.db` file. The current database will be restored with the backup dataset.

---

## 3. Troubleshooting & FAQ

1. **Permission / File Lock Errors:**
   * If an error occurs when exporting protocols or Excel sheets, ensure that a file with the target name is not currently open in Microsoft Word, Excel, or a PDF reader.
2. **Moving the Application to Another Computer:**
   * Transfer the `.exe` file and your database file located at `%APPDATA%\BazaTelefonow\telefony.db` to the corresponding folder on the destination machine, or utilize the built-in **Backup** and **Restore** buttons.