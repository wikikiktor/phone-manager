# Baza Telefonów / Phone Fleet Manager

[PL](#instrukcja-obsługi-programu-język-polski) | [EN](#user-manual-english)

---

# Instrukcja Obsługi Programu (Język Polski)

**Baza Telefonów** to gotowy, samodzielny program dla systemu Windows (`BazaTelefonow.exe`) służący do ewidencji telefonów służbowych, kart SIM, historii zdarzeń, notatek serwisowych oraz błyskawicznego generowania dokumentów zdawczo-odbiorczych w formatach PDF i DOCX (Word).

Dane są **zaszyfrowane**, a dostęp do programu wymaga **zalogowania**.

Aplikacja **nie wymaga instalacji** środowiska deweloperskiego, żadnych bibliotek ani zewnętrznych baz danych – jest gotowa do uruchomienia od razu po pobraniu.

---

## 1. Uruchomienie, logowanie i dane programu

### 1.1. Jak uruchomić aplikację?
* Pobierz plik `BazaTelefonow.exe`.
* Umieść go w dowolnym wygodnym folderze (np. na Pulpicie lub w folderze Dokumenty).
* Uruchom program dwuklikiem.

### 1.2. Pierwsze uruchomienie – konto administratora
Przy pierwszym uruchomieniu program poprosi o utworzenie **konta administratora**:
1. Wpisz login oraz hasło (co najmniej 8 znaków) i powtórz hasło.
2. Kliknij **`Utwórz konto`**.
3. Program wyświetli jednorazowo **kod awaryjny** (np. `5N8C-QXRD-QA6W-9RV8-DHNZ-M8CS`). **Zapisz go i przechowuj w bezpiecznym miejscu** (np. na papierze w sejfie). Kod pozwala ustawić nowe hasło, gdy zostanie zapomniane. Kod można skopiować przyciskiem **`📋 Kopiuj do schowka`**. Okno zamkniesz dopiero po zaznaczeniu pola *„Zapisałem kod awaryjny w bezpiecznym miejscu”*.

> **Aktualizacja ze starszej wersji:** jeśli program znajdzie istniejącą, niezaszyfrowaną bazę, zaszyfruje ją automatycznie podczas tworzenia konta administratora – wszystkie dane zostaną zachowane. Przed pierwszym uruchomieniem nowej wersji zalecamy zrobić zwykłą kopię pliku `telefony.db`.

### 1.3. Logowanie
* Przy każdym uruchomieniu wpisz **login** i **hasło**, a następnie kliknij **`Zaloguj`** (lub naciśnij Enter). Wielkość liter w loginie nie ma znaczenia.
* Nazwa zalogowanego użytkownika jest widoczna na górnym pasku (👤).

### 1.4. Nie pamiętam hasła
1. W oknie logowania kliknij **`Nie pamiętam hasła`**.
2. Podaj login, **kod awaryjny** (wielkość liter, spacje i myślniki nie mają znaczenia) oraz dwukrotnie nowe hasło.
3. Kliknij **`Ustaw nowe hasło`** i zaloguj się nowym hasłem.

> ⚠️ **Bez hasła ani kodu awaryjnego nie da się odzyskać danych** – baza jest zaszyfrowana i nie istnieje „tylne wejście”.

### 1.5. Gdzie przechowywane są Twoje dane?
Program przechowuje dane w profilu użytkownika, w folderze:

```text
%APPDATA%\BazaTelefonow\
```
*Pełna ścieżka:*  
`C:\Użytkownicy\<Twoja_Nazwa_Użytkownika>\AppData\Roaming\BazaTelefonow\`

| Plik | Zawartość |
|---|---|
| `telefony.db` | Zaszyfrowana baza danych (SQLCipher, AES-256) |
| `keys.json` | Klucz do bazy zaszyfrowany hasłami użytkowników i kodem awaryjnym |

> **Ważne:** Oba pliki są potrzebne do otwarcia danych. Nie usuwaj i nie przenoś pliku `keys.json` osobno – bez niego zaszyfrowanej bazy nie da się otworzyć.

> **Wskazówka:** Nawet jeśli przeniesiesz lub podmienisz sam plik `.exe` na nowszą wersję, Twoje dane pozostaną nienaruszone w profilu systemowym.

---

## 2. Podręcznik Użytkownika

### 2.1. Dodawanie nowego telefonu do bazy
1. Kliknij niebieski przycisk **`+ Nowy telefon`** w prawym górnym rogu okna.
2. Wypełnij formularz po prawej stronie. Pola obowiązkowe:
   * **Model tel** (np. *Samsung Galaxy A54*)
   * **Nr telefonu** (np. *500100200* lub *+48500100200* – program automatycznie sformatuje numer z myślnikami).
3. Uzupełnij pozostałe dane wedle uznania: *Nr IMEI*, *Nr SIM*, *Nr seryjny*, *Osoba użytkująca*, *Osoba odpowiedzialna*, *Rodzaj użytkowania*.
4. Zaznacz dołączone akcesoria w sekcji **Wyposażenie** (*Ładowarka*, *Etui*, *Kabel USB*, *Szkło ochronne*).
5. Wybierz **Stan baterii** oraz opcjonalnie zaznacz **Wysłany protokół** (jeśli dokument został już odebrany).
6. Kliknij zielony przycisk **`💾 Zapisz zmiany`**. Telefon pojawi się na liście po lewej stronie, a w jego historii zostanie odnotowany fakt dodania.

### 2.2. Edycja i automatyczny rejestr zmian
* Kliknięcie na dowolny telefon z listy po lewej stronie ładuje jego dane oraz historię zdarzeń.
* Zmodyfikuj dowolne pole (np. zmiana użytkownika, numeru SIM, wymiana ładowarki) i kliknij **`💾 Zapisz zmiany`**.
* Program samoczynnie wykrywa różnice i rejestruje wpis w dzienniku zdarzeń (np. *Zmieniono Osoba użytkująca: 'Jan Kowalski' ➔ 'Anna Nowak'* wraz z dokładną datą i godziną).
* Przycisk **`Wyczyść`** czyści formularz bez zapisywania zmian.

### 2.3. Dziennik zdarzeń, serwis i notatki
W dolnej prawej części okna znajduje się dedykowana historia dla wybranego urządzenia:
* **Dodawanie wpisu:** Kliknij **`+ Dodaj wpis`**, wybierz kategorię (*Notatka / Uwaga*, *Zmiana użytkownika*, *Awaria / Serwis*, *Wydanie*, *Zwrot*, *Wymiana karty SIM*, *Inne*) i wprowadź treść.
* **Notatki/uwagi:** Zaznacz wybrany wpis w historii i kliknij **`✏️ Edytuj uwagę`** (lub kliknij dwukrotnie dany wiersz), aby dopisać dodatkowe adnotacje (np. numer zgłoszenia serwisowego, koszt naprawy).

### 2.4. Wyszukiwarka i sortowanie
* **Wyszukiwarka w czasie rzeczywistym:** Wpisz szukaną frazę w pole u góry. Program filtruje urządzenia po numerze telefonu (również wpisanym ciągiem bez spacji i myślników), modelu, nazwisku użytkownika oraz numerze IMEI. Licznik po prawej pokazuje, ile telefonów pasuje do wyszukiwania.
* **Przycisk `✕`:** Błyskawicznie czyści pole wyszukiwania.
* **Sortowanie:** Kliknij nagłówek kolumny na liście telefonów (*Nr telefonu*, *Użytkownik*, *Model*) lub w tabeli historii, aby posortować rekordy rosnąco lub malejąco (strzałki `▲` / `▼`). Sortowanie jest zachowywane po odświeżeniu listy.

### 2.5. Kosz i bezpieczne usuwanie (Kosz vs Usuwanie trwałe)
* **Przeniesienie do kosza:** Kliknij czerwony przycisk **`🗑️ Do kosza`**. Telefon zniknie z aktywnej listy, ale jego dane i pełna historia zostaną zachowane.
* **Podgląd kosza:** Zaznacz pole **`Pokaż kosz (usunięte)`** na górnym pasku.
* **Przywracanie urządzenia:** W widoku kosza wybierz telefon i kliknij **`↩️ Przywróć telefon`**.
* **Trwałe skasowanie z bazy:** W widoku kosza dostępny jest przycisk **`❌ Usuń trwale z bazy`** – bezpowrotnie usuwa urządzenie i całą jego historię z bazy danych.

### 2.6. Oznaczanie telefonu jako uszkodzony (Flaga ostrzegawcza)
* W przypadku awarii lub uszkodzenia sprzętu, wybierz telefon z listy i kliknij pomarańczowy przycisk **`⚠️ Oznacz telefon`**.
* Urządzenie zostanie wyraźnie podświetlone na **czerwono** na głównej liście (również po sortowaniu), co pozwala na błyskawiczną identyfikację sprzętu wymagającego naprawy.
* Po naprawie telefonu lub rozwiązaniu problemu wystarczy kliknąć przycisk **`✅ Odznacz telefon`**, aby przywrócić urządzeniu standardowy status.

### 2.7. Dane firmy do protokołów
Przed wydrukiem dokumentów ustaw dane pracodawcy:
1. Rozwiń menu **`⚙️ Opcje bazy ▾`** i kliknij **`🏢 Dane firmy / pracodawcy`**.
2. Wpisz nazwę firmy, adres oraz NIP.
3. Kliknij **`💾 Zapisz dane`**. Informacje zostaną zapamiętane i będą automatycznie wstawiane do nagłówków wszystkich protokołów.

### 2.8. Generowanie gotowych protokołów (PDF i Word DOCX)
1. Kliknij **`📋 Protokół`** na górnym pasku (jeśli wcześniej zaznaczono telefon na liście, zostanie wybrany automatycznie).
2. Wybierz typ: **Protokół przekazania** lub **Protokół zwrotu**.
3. Sprawdź podgląd danych (dane pracownika, numer IMEI, zaznaczone akcesoria, stan baterii).
4. Kliknij:
   * **`📄 Pobierz PDF`** – generuje gotowy do wydruku dokument A4 w formacie PDF (obsługuje pełne polskie znaki).
   * **`📝 Pobierz DOCX`** – generuje edytowalny dokument w formacie programu Microsoft Word.
5. Wybierz folder zapisu na dysku.

Brakujące dane (np. numer seryjny) są w protokole zastępowane linią kropek do ręcznego uzupełnienia.

### 2.9. Import i Eksport arkuszy Excel (.xlsx)
* **Eksport do Excela:** Kliknij **`📊 Eksport Excel`**. Program utworzy plik `.xlsx` zawierający dwa arkusze:
  * *Telefony* – kompletna, ostylowana tabela aktualnie widocznych urządzeń z autofiltrem i zamrożonym nagłówkiem (kolumny *Protokół* i *Czy uszkodzony* jako *Tak/Nie*).
  * *Historia zdarzeń* – powiązany dziennik operacji i uwag.

  > Plik Excel **nie jest szyfrowany** – przechowuj go ostrożnie.
* **Import z Excela:** Kliknij **`📥 Import Excel`** i wskaż plik `.xlsx` lub `.xlsm` (stary format `.xls` nie jest obsługiwany – zapisz go w Excelu jako `.xlsx`). Kreator pozwoli Ci łatwo dopasować kolumny z Twojego dotychczasowego arkusza do pól w programie. Jeśli plik posiada zakładkę *Historia zdarzeń*, przeniesie również wpisy historyczne.

### 2.10. Kopia zapasowa i przywracanie bazy (Backup)
W rozwijanym menu **`⚙️ Opcje bazy ▾`**:
* **`💾 Utwórz kopię zapasową bazy (.db)`** – zapisuje zaszyfrowaną kopię całej bazy danych w wybranym miejscu (np. dysk sieciowy, pendrive). Obok kopii zapisywany jest plik z kluczami, np.:
  ```text
  backup_telefony_20261105_1200.db
  backup_telefony_20261105_1200.keys.json
  ```
  **Przechowuj oba pliki razem** – bez pliku `.keys.json` kopii nie da się otworzyć na innym komputerze.
* **`🔄 Przywróć bazę z pliku (.db)`** – pozwala odtworzyć dane z wcześniej wykonanej kopii. Można przywracać kopie wykonane w tym programie oraz stare, niezaszyfrowane kopie ze starszych wersji.

### 2.11. Użytkownicy, hasła i kod awaryjny
W rozwijanym menu **`⚙️ Opcje bazy ▾`**:
* **`🔑 Zmień hasło`** – zmiana hasła zalogowanego użytkownika (wymaga podania obecnego hasła).
* **`👥 Dodaj użytkownika`** – tworzy nowe konto z dostępem do bazy.
* **`🚫 Usuń użytkownika`** – blokuje logowanie wybranemu użytkownikowi (nie można usunąć samego siebie).
* **`🆘 Wygeneruj nowy kod awaryjny`** – tworzy nowy kod awaryjny; dotychczasowy przestaje działać. Użyj tej opcji, jeśli kod został zgubiony lub mógł trafić w niepowołane ręce.

---

## 3. Rozwiązywanie częstych problemów

* **Komunikat „Błąd zapisu” przy imporcie / eksporcie / tworzeniu protokołu:**  
  Upewnij się, że plik o tej samej nazwie nie jest w tym momencie otwarty w programie Excel, Word lub Adobe Acrobat Reader. Zamknij otwarty dokument i spróbuj ponownie.
* **„Nieprawidłowy login lub hasło”:**  
  Sprawdź, czy nie jest włączony Caps Lock. Jeśli nie pamiętasz hasła, użyj opcji **`Nie pamiętam hasła`** i kodu awaryjnego (pkt 1.4).
* **„Plik bazy jest zablokowany” przy pierwszym uruchomieniu:**  
  Plik bazy jest używany przez inny program. Zamknij programy, które mogą mieć otwarty plik `telefony.db` (np. DB Browser for SQLite lub drugie okno aplikacji) i kliknij **`Utwórz konto`** ponownie – dotychczasowa baza pozostaje nienaruszona. Jeśli problem się powtarza, dodaj folder `%APPDATA%\BazaTelefonow` do wyjątków programu antywirusowego.
* **„Brak pliku kluczy” przy uruchomieniu:**  
  Plik `keys.json` został usunięty lub przeniesiony. Skopiuj plik `*.keys.json` z ostatniej kopii zapasowej do folderu `%APPDATA%\BazaTelefonow\` i zmień jego nazwę na `keys.json`.
* **Przeniesienie programu na nowy komputer:**  
  1. Na starym komputerze utwórz kopię zapasową (**`⚙️ Opcje bazy ▾` ➔ `💾 Utwórz kopię zapasową bazy (.db)`**).
  2. Na nowym komputerze utwórz folder `%APPDATA%\BazaTelefonow\` i **przed pierwszym uruchomieniem programu** skopiuj do niego:
     * plik kopii `*.db` – zmieniając nazwę na `telefony.db`,
     * plik `*.keys.json` – zmieniając nazwę na `keys.json`.
  3. Uruchom `BazaTelefonow.exe` i zaloguj się dotychczasowym loginem i hasłem.

  > Jeśli na nowym komputerze program został już uruchomiony i utworzono nowe konto, zamknij program, usuń z folderu `%APPDATA%\BazaTelefonow\` pliki `telefony.db` i `keys.json`, a następnie wykonaj krok 2.

---
---

# User Manual (English)

**Phone Fleet Manager** is a standalone, portable Windows application (`BazaTelefonow.exe`) engineered for corporate mobile device tracking, SIM card management, automated change auditing, and one-click PDF & DOCX handover/return protocol generation.

Data is **encrypted at rest** and access requires **signing in**.

**No runtime setup, installer, or third-party database engines required.**

---

## 1. Launch, Sign-in & Data Storage

### 1.1. How to Run
* Download `BazaTelefonow.exe`.
* Place it in any directory (Desktop, Documents, or Shared Drive).
* Double-click the executable to start the application.

### 1.2. First Run – Administrator Account
On first launch the application asks you to create an **administrator account**:
1. Enter a login and a password (at least 8 characters), then repeat the password.
2. Click **`Utwórz konto`** (*Create account*).
3. A one-time **recovery code** is displayed (e.g. `5N8C-QXRD-QA6W-9RV8-DHNZ-M8CS`). **Write it down and keep it somewhere safe.** It lets you set a new password if it is forgotten. The window can be closed only after confirming the code has been saved.

> **Upgrading from an older version:** an existing unencrypted database is encrypted automatically while the administrator account is being created – all data is preserved. Making a plain copy of `telefony.db` before the first run of the new version is recommended.

### 1.3. Signing In
* Enter your **login** and **password** and click **`Zaloguj`** (*Sign in*) or press Enter. Logins are case-insensitive.
* The signed-in user is shown on the top bar (👤).

### 1.4. Forgotten Password
1. In the sign-in window click **`Nie pamiętam hasła`** (*Forgot password*).
2. Enter your login, the **recovery code** (case, spaces and dashes are ignored) and the new password twice.
3. Click **`Ustaw nowe hasło`** (*Set new password*) and sign in with the new password.

> ⚠️ **Without a password or the recovery code the data cannot be recovered** – the database is encrypted and there is no backdoor.

### 1.5. Data Location
All data is stored in your user profile:

```text
%APPDATA%\BazaTelefonow\
```
*Full Path:*  
`C:\Users\<Your_Username>\AppData\Roaming\BazaTelefonow\`

| File | Contents |
|---|---|
| `telefony.db` | Encrypted database (SQLCipher, AES-256) |
| `keys.json` | Database key, encrypted with each user's password and the recovery code |

> **Important:** Both files are required to open the data. Never delete or move `keys.json` on its own.

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
* **Live Search:** Instant filtering across phone numbers (with or without dashes), device models, assigned users and IMEI numbers.
* **Clear Search (`✕`):** Clears search query in one click.
* **Column Sorting:** Click any column header in the device list or history log to sort ascending or descending (`▲` / `▼`). Sorting is kept when the list refreshes.

### 2.5. Trash Management (Soft vs Hard Delete)
* **Soft Delete:** Click **`🗑️ Do kosza`** (*To Trash*) to remove the device from the active view while preserving its audit history.
* **View Trash:** Check **`Pokaż kosz (usunięte)`** (*Show Trash*) on the top toolbar.
* **Restore:** Select an item in trash and click **`↩️ Przywróć telefon`** (*Restore Phone*).
* **Hard Delete:** Inside the trash view, click **`❌ Usuń trwale z bazy`** (*Delete Permanently*) to irreversibly purge the device and its history.

### 2.6. Flagging a Device as Damaged / Warning
* If a device is broken, damaged, or needs attention, select it and click the orange **`⚠️ Oznacz telefon`** (*Mark Phone*) button.
* The device row will be highlighted in **red** on the active list (also after sorting) for quick visual identification of equipment needing repair.
* Once resolved, click **`✅ Odznacz telefon`** (*Unmark Phone*) to revert to the standard status.

### 2.7. Employer Company Details
1. Open the dropdown **`⚙️ Opcje bazy ▾`** and select **`🏢 Dane firmy / pracodawcy`**.
2. Enter the Company Name, Address, and Tax ID (NIP).
3. Click **`💾 Zapisz dane`**. Information will automatically appear on generated protocols.

### 2.8. Protocol Generation (PDF & DOCX)
1. Click **`📋 Protokół`** on the top toolbar.
2. Choose **Protokół przekazania** (*Handover*) or **Protokół zwrotu** (*Return*).
3. Review the preview pane.
4. Export:
   * **`📄 Pobierz PDF`** – exports a formatted A4 PDF.
   * **`📝 Pobierz DOCX`** – exports an editable Microsoft Word document.

Missing values (e.g. serial number) are replaced with a dotted line to be filled in by hand.

### 2.9. Excel Import & Export (.xlsx)
* **Export to Excel:** Click **`📊 Eksport Excel`** to generate a styled workbook containing the visible devices and their event logs on separate sheets. **The Excel file is not encrypted** – store it with care.
* **Import from Excel:** Click **`📥 Import Excel`** to launch the interactive column mapper. Supported formats: `.xlsx` and `.xlsm` (legacy `.xls` files must be re-saved as `.xlsx` first).

### 2.10. Database Backup & Restore
Under **`⚙️ Opcje bazy ▾`**:
* **`💾 Utwórz kopię zapasową bazy (.db)`** – exports an encrypted point-in-time snapshot of the database, together with a key file saved next to it (`<name>.keys.json`). **Keep both files together** – the backup cannot be opened on another computer without the key file.
* **`🔄 Przywróć bazę z pliku (.db)`** – restores all database tables from a backup created by this application, or from an old unencrypted backup made by a previous version.

### 2.11. Users, Passwords & Recovery Code
Under **`⚙️ Opcje bazy ▾`**:
* **`🔑 Zmień hasło`** – change the signed-in user's password (current password required).
* **`👥 Dodaj użytkownika`** – create a new account with access to the database.
* **`🚫 Usuń użytkownika`** – revoke sign-in for another user (you cannot remove yourself).
* **`🆘 Wygeneruj nowy kod awaryjny`** – generate a new recovery code; the previous one stops working.

---

## 3. Troubleshooting

* **"Błąd zapisu" (Write error) on import / export / protocol generation:**  
  Make sure a file with the same name is not open in Excel, Word or Adobe Acrobat Reader.
* **"Nieprawidłowy login lub hasło" (Invalid login or password):**  
  Check Caps Lock. If the password is forgotten, use **`Nie pamiętam hasła`** with the recovery code (section 1.4).
* **"Plik bazy jest zablokowany" (Database file is locked) on first run:**  
  Another program is using `telefony.db`. Close it (e.g. DB Browser for SQLite or a second app window) and click **`Utwórz konto`** again – the existing database stays untouched. If it persists, add `%APPDATA%\BazaTelefonow` to your antivirus exclusions.
* **"Brak pliku kluczy" (Key file missing) on start:**  
  `keys.json` was deleted or moved. Copy the `*.keys.json` file from your latest backup into `%APPDATA%\BazaTelefonow\` and rename it to `keys.json`.
* **Moving to a new computer:**  
  1. On the old computer create a backup (**`⚙️ Opcje bazy ▾` ➔ `💾 Utwórz kopię zapasową bazy (.db)`**).
  2. On the new computer create the folder `%APPDATA%\BazaTelefonow\` and, **before the first launch**, copy into it the backup `*.db` file renamed to `telefony.db` and the `*.keys.json` file renamed to `keys.json`.
  3. Start `BazaTelefonow.exe` and sign in with your existing login and password.

  > If the application has already been started on the new computer and a new account was created, close it, delete `telefony.db` and `keys.json` from `%APPDATA%\BazaTelefonow\`, and repeat step 2.
