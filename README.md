# PAP MediaRoom Generator

Wewnętrzne narzędzie redakcyjne do tworzenia informacji prasowych na podstawie materiałów dostarczonych przez klientów. Oparte na zasadach PAP MediaRoom i wzbogacone o weryfikację wykorzystania materiału wejściowego.

## Co robi aplikacja

Operator wkleja materiał od klienta lub wgrywa plik Word/PDF. Aplikacja generuje informację prasową w stylu PAP MediaRoom, jednocześnie pokazując:

* procent wykorzystania materiału wejściowego
* które fragmenty materiału zostały pominięte i dlaczego
* które zdania w wygenerowanej informacji są dodaniami spoza materiału (jeśli włączony tryb kontekstowy)
* czerwone flagi przy zdaniach, które nie mają wyraźnej podstawy w materiale wejściowym

## Wdrożenie krok po kroku (dla nieprogramisty)

### Krok 1. Konto GitHub

1. Wejdź na https://github.com i załóż darmowe konto, jeśli go nie masz.
2. Po zalogowaniu kliknij zielony przycisk **New** (lub plus w prawym górnym rogu, potem **New repository**).
3. Nazwa repozytorium: na przykład `pap-mediaroom-generator`.
4. Ustaw **Private** (prywatne), żeby nikt poza Tobą nie miał dostępu.
5. Zaznacz **Add a README file** i kliknij **Create repository**.

### Krok 2. Wgranie plików aplikacji do GitHub

1. W swoim nowo utworzonym repozytorium kliknij **Add file** -> **Upload files**.
2. Przeciągnij wszystkie pliki z paczki (oprócz katalogu `.streamlit`, do niego zaraz wrócimy):
   * `streamlit_app.py`
   * `requirements.txt`
   * `.gitignore`
   * `README.md`
3. Kliknij **Commit changes**.
4. Teraz dodaj katalog `.streamlit`:
   * Wróć do głównego widoku repozytorium.
   * Kliknij **Add file** -> **Create new file**.
   * W polu nazwy wpisz: `.streamlit/secrets.toml.example` (slash w nazwie utworzy podkatalog).
   * Wklej zawartość pliku `secrets.toml.example` z paczki.
   * Kliknij **Commit changes**.

**WAŻNE:** Pliku `secrets.toml` (bez sufiksu `.example`) z prawdziwymi sekretami **nigdy** nie wgrywaj do GitHub. Sekrety ustawisz w Streamlit Cloud w kroku 5.

### Krok 3. Klucz API Anthropic

1. Wejdź na https://console.anthropic.com i załóż konto, jeśli go nie masz.
2. Doładuj konto kwotą minimum 5 do 10 dolarów (sekcja **Plans & Billing** -> **Pricing**).
3. Przejdź do sekcji **API Keys** i kliknij **Create Key**.
4. Nadaj kluczowi nazwę (np. `pap-mediaroom`).
5. Skopiuj klucz **natychmiast po wyświetleniu**. Zaczyna się od `sk-ant-`. Anthropic pokazuje go tylko raz.
6. Zachowaj klucz w bezpiecznym miejscu (np. menedżer haseł). Nie wysyłaj nikomu mailem ani Slackiem.

### Krok 4. Konto Streamlit Community Cloud

1. Wejdź na https://share.streamlit.io i kliknij **Sign up**.
2. Wybierz **Continue with GitHub** i autoryzuj dostęp.
3. Po zalogowaniu kliknij **Create app** -> **Deploy a public app from GitHub**.
4. W formularzu:
   * **Repository:** wybierz `<twoja-nazwa>/pap-mediaroom-generator`
   * **Branch:** `main`
   * **Main file path:** `streamlit_app.py`
   * **App URL:** wybierz nazwę, np. `pap-mediaroom`
5. **NIE klikaj jeszcze Deploy.** Najpierw skonfiguruj sekrety.

### Krok 5. Konfiguracja sekretów

1. W formularzu deploy kliknij **Advanced settings**.
2. W polu **Secrets** wklej dokładnie taką zawartość, podmieniając wartości na prawdziwe:

```toml
app_password = "wybrane_haslo_dla_zespolu"
anthropic_api_key = "sk-ant-twoj-prawdziwy-klucz-api"
```

3. Kliknij **Save**.
4. Kliknij **Deploy**.

### Krok 6. Pierwsze uruchomienie

1. Streamlit potrzebuje 1 do 3 minut na zainstalowanie zależności i uruchomienie aplikacji.
2. Po uruchomieniu zobaczysz ekran logowania.
3. Wpisz hasło, które ustawiłeś w sekretach (`app_password`).
4. Aplikacja jest gotowa do użycia.

## Codzienne użycie

1. Wejdź na URL Twojej aplikacji (znajdziesz go w panelu Streamlit Cloud).
2. Zaloguj się hasłem zespołu.
3. W panelu bocznym ustaw format (sztywny/elastyczny) i tryb uzupełnień (zerowa tolerancja/kontekstowe).
4. Zaznacz, które kategorie nie powinny być liczone do współczynnika wykorzystania.
5. Wklej materiał albo wgraj plik Word/PDF.
6. Kliknij **Generuj informację prasową**.
7. Zweryfikuj wynik:
   * Sprawdź współczynnik wykorzystania.
   * Zwróć uwagę na czerwone flagi (zdania bez podstawy w materiale).
   * Sprawdź dodane fragmenty (jeśli tryb kontekstowy).
   * Najedź kursorem na podświetlone fragmenty, by zobaczyć powód.
8. Skopiuj gotowy tekst lub pobierz jako plik.

## Aktualizacja aplikacji

Jeśli chcesz coś zmienić (np. dostosować zasady PAP w prompcie systemowym):

1. Wejdź do swojego repozytorium na GitHub.
2. Otwórz plik `streamlit_app.py` i kliknij ikonę ołówka (Edit).
3. Wprowadź zmiany i kliknij **Commit changes**.
4. Streamlit Cloud wykryje zmianę i automatycznie zrestartuje aplikację (zwykle w ciągu minuty).

## Zmiana hasła zespołu

1. Wejdź na https://share.streamlit.io.
2. Kliknij swoją aplikację.
3. Kliknij ikonę trzech kropek (lub menedżera) -> **Settings** -> **Secrets**.
4. Zmień wartość `app_password`.
5. Kliknij **Save**. Aplikacja zrestartuje się automatycznie.

## Ograniczenia obecnej wersji

* Pliki PDF muszą zawierać tekst (nie skany). Skanowane PDFy wymagają OCR, którego ta wersja nie obsługuje.
* Dzielenie tekstu na zdania jest oparte na regexie i radzi sobie z większością przypadków, ale przy nietypowej interpunkcji może popełniać błędy. W razie potrzeby tekst można edytować ręcznie po wyodrębnieniu z pliku.
* Aplikacja nie zapisuje historii zapytań (zgodnie z założeniem o poufności materiałów klientów).
* Generowanie trwa zwykle 15 do 30 sekund w zależności od długości materiału.

## Koszty

Streamlit Community Cloud jest darmowy. Płacisz tylko za wywołania API Anthropic. Szacunkowy koszt jednego wygenerowania informacji prasowej to 0,03 do 0,10 USD (zależnie od długości materiału). Doładowanie 10 USD wystarcza zwykle na 100 do 300 generowań.

## Wsparcie i rozwój

Aplikacja jest w wersji proof of concept. Plan rozwoju:

* Bibliotek par materiał wejściowy + gotowa informacja (jako wzorce dla modelu).
* Drugi, niezależny model jako weryfikator (zamiast samowywołania Claude).
* Eksport do Word z formatowaniem.
* Logowanie indywidualne zamiast wspólnego hasła.
* Słownik wewnętrznych skrótów PAP.
