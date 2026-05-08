# -*- coding: utf-8 -*-
"""
PAP MediaRoom Generator
Wewnętrzne narzędzie redakcyjne do tworzenia informacji prasowych
na podstawie materiałów dostarczonych przez klientów.
"""

import streamlit as st
import anthropic
import json
import re
import io
from docx import Document
import pdfplumber

# ============================================================
# KONFIGURACJA STRONY
# ============================================================

st.set_page_config(
    page_title="PAP MediaRoom Generator",
    page_icon="📰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# SYSTEM PROMPT (zasady PAP MediaRoom)
# ============================================================

SYSTEM_PROMPT = """Jesteś redaktorem PAP MediaRoom. Tworzysz informacje prasowe według standardów PAP MediaRoom.

ZASADY KLUCZOWE:
1. STRUKTURA: tytuł z oznaczeniem (MediaRoom) na końcu, pogrubiony lead (2 do 4 zdań, kto/co/gdzie/kiedy z atrybucją), akapit pomostowy, korpus z piramidą odwróconą, akapit kontekstowy, stopka.
2. TYTUŁ: format "Podmiot: konkretna teza" lub "Wydarzenie: konkretna teza". Bez sloganów, bez wykrzykników, bez wartościujących przymiotników.
3. JĘZYK: rzeczowy, neutralny, bez wartościowania poza cytatami. Bez metafor, sloganów, hipersuperlatywów. Bez "my/nasz" poza cytatami.
4. CYTATY: wprowadzane zróżnicowanymi czasownikami (powiedział, podkreślił, ocenił, zaznaczył, dodał, wyjaśnił, zauważył, wskazał). Atrybucja: imię nazwisko, funkcja. Cudzysłowy polskie „".
5. LICZBY: każda dana ma atrybucję. Format: "51 proc.", "12 mln zł", "5,7 mld zł", "19 kwietnia", "2026 r."
6. AKRONIMY: pełna nazwa + (skrót) przy pierwszym użyciu.
7. STOPKA STAŁA (zawsze kończy informację, bez modyfikacji):
"Źródło informacji: PAP MediaRoom

UWAGA: Za materiał opublikowany przez redakcję PAP MediaRoom odpowiedzialność ponosi jego nadawca, wskazany każdorazowo jako „źródło informacji"."

ZAKAZANE:
- Wprowadzanie informacji spoza materiału wejściowego (chyba że tryb uzupełnień = contextual).
- Wymyślanie cytatów lub modyfikowanie ich treści.
- Zmiana funkcji/tytułów rozmówców.
- Slogany, język marketingowy, wykrzykniki, pytania retoryczne.

ZADANIE:
Otrzymujesz materiał wejściowy od klienta podzielony na ponumerowane zdania. Twoim zadaniem jest:
1. Stworzyć informację prasową w stylu PAP MediaRoom.
2. Dla każdego zdania wejściowego sklasyfikować je: USED (wykorzystane), EXCLUDED (kategoria wykluczona z liczenia: kontakt prasowy, boilerplate, stopki, klauzule), SKIPPED (świadomie pominięte z krótkim uzasadnieniem).
3. Podzielić gotową informację prasową na zdania i dla każdego zdania wskazać, które zdania wejścia je wspierają (lista ID), oraz oznaczyć czy zdanie jest dodane (added=true) jako uzupełnienie kontekstowe spoza materiału (tylko w trybie contextual).
4. Wystawić ostrzeżenia jeśli materiał jest ubogi, sprzeczny lub nie nadaje się do publikacji.

ZWRÓĆ WYŁĄCZNIE JSON, bez markdown fences, bez preambuły:
{
  "press_release_text": "Pełny tekst informacji prasowej z łamaniami linii \\n\\n między akapitami",
  "press_release_sentences": [
    {"text": "Pierwsze zdanie wyjścia.", "added": false, "supported_by": [1, 3]},
    {"text": "Drugie zdanie wyjścia.", "added": false, "supported_by": [2]}
  ],
  "input_mapping": [
    {"id": 1, "status": "USED", "reason": null},
    {"id": 2, "status": "EXCLUDED", "reason": "kontakt prasowy"},
    {"id": 3, "status": "SKIPPED", "reason": "powtórzenie informacji"}
  ],
  "warnings": []
}"""


# ============================================================
# AUTORYZACJA
# ============================================================

def check_authentication():
    """Sprawdza czy użytkownik jest zalogowany. Zwraca True jeśli tak."""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    # Ekran logowania
    st.title("PAP MediaRoom Generator")
    st.caption("Wewnętrzne narzędzie redakcyjne. Logowanie wymagane.")

    col_left, col_center, col_right = st.columns([1, 2, 1])
    with col_center:
        with st.form("login_form", clear_on_submit=False):
            password = st.text_input("Hasło", type="password", placeholder="wpisz hasło")
            submitted = st.form_submit_button("Zaloguj", type="primary", use_container_width=True)
            if submitted:
                expected = st.secrets.get("app_password", "")
                if password == expected and expected != "":
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Niepoprawne hasło.")
    return False


# ============================================================
# EKSTRAKCJA TEKSTU Z PLIKÓW
# ============================================================

def extract_text_from_docx(file_bytes):
    """Wyciąga tekst z pliku Word (.docx)."""
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if text:
            paragraphs.append(text)
    # Wyciągnij też tekst z tabel
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                if text:
                    paragraphs.append(text)
    return "\n".join(paragraphs)


def extract_text_from_pdf(file_bytes):
    """Wyciąga tekst z pliku PDF."""
    text_parts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text_parts.append(page_text)
    return "\n".join(text_parts)


# ============================================================
# DZIELENIE NA ZDANIA
# ============================================================

POLISH_ABBREVIATIONS = [
    "r", "mln", "mld", "proc", "pkt", "tj", "np", "tzw", "tys",
    "str", "nr", "ul", "al", "woj", "godz", "min", "sek",
    "wg", "por", "art", "ust", "par", "dr", "prof", "inż", "mgr",
    "płk", "gen", "płd", "pn", "wsch", "zach", "ok"
]


def split_sentences(text):
    """Dzieli tekst na zdania, uwzględniając polskie skróty."""
    cleaned = re.sub(r"\s+", " ", text.strip())
    if not cleaned:
        return []

    # Chroń skróty przed traktowaniem ich jako końca zdania
    abbrev_pattern = r"\b(" + "|".join(POLISH_ABBREVIATIONS) + r")\.\s"
    protected = re.sub(abbrev_pattern, r"\1.\u00A0", cleaned)
    # Chroń "m.in."
    protected = re.sub(r"\bm\.in\.\s", "m.in.\u00A0", protected)
    # Chroń liczby z kropką (np. daty, numery)
    protected = re.sub(r"(\d)\.\s(\d)", r"\1.\u00A0\2", protected)

    parts = re.findall(r"[^.!?]+[.!?]+", protected)
    if not parts:
        parts = [protected]
    return [p.replace("\u00A0", " ").strip() for p in parts if p.strip()]


# ============================================================
# WYWOŁANIE CLAUDE
# ============================================================

def build_user_prompt(sentences, format_mode, supplement_mode, excluded):
    format_label = (
        "sztywny (tytuł, lead, korpus, kontakt, stopka)"
        if format_mode == "rigid"
        else "elastyczny (zachowane zasady języka i atrybucji, dopuszczona swoboda struktury)"
    )
    supplement_label = (
        "ZEROWA TOLERANCJA - wyłącznie materiał wejściowy, żadnych dodatków spoza"
        if supplement_mode == "zero"
        else "KONTEKSTOWE - można dodać podstawowy kontekst (np. czym jest dana instytucja), wszystkie dodatki oznaczyć added=true"
    )
    excl_label = ", ".join(excluded) if excluded else "brak"

    numbered = "\n".join(f"[{i+1}] {s}" for i, s in enumerate(sentences))

    return f"""KONFIGURACJA:
- Format: {format_label}
- Tryb uzupełnień: {supplement_label}
- Kategorie wykluczone z liczenia wykorzystania (sklasyfikuj odpowiednie zdania jako EXCLUDED): {excl_label}

MATERIAŁ WEJŚCIOWY (ponumerowane zdania):
{numbered}

Wykonaj zadanie i zwróć wyłącznie JSON według podanego formatu."""


def call_claude(sentences, format_mode, supplement_mode, excluded):
    """Wywołuje Claude API i zwraca sparsowany wynik."""
    client = anthropic.Anthropic(api_key=st.secrets["anthropic_api_key"])
    user_prompt = build_user_prompt(sentences, format_mode, supplement_mode, excluded)

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        temperature=0,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}]
    )

    raw_text = "".join(b.text for b in message.content if b.type == "text")
    cleaned = re.sub(r"^```(?:json)?\s*", "", raw_text)
    cleaned = re.sub(r"\s*```\s*$", "", cleaned).strip()

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as e:
        raise ValueError(
            f"Nie udało się sparsować odpowiedzi modelu jako JSON. "
            f"Błąd: {e}\n\nPoczątek odpowiedzi:\n{raw_text[:1500]}"
        )


# ============================================================
# OBLICZENIA I RENDEROWANIE
# ============================================================

def compute_usage(mapping):
    total = len(mapping)
    excluded = sum(1 for m in mapping if m.get("status") == "EXCLUDED")
    used = sum(1 for m in mapping if m.get("status") == "USED")
    denominator = total - excluded
    if denominator == 0:
        return {"percent": 0, "used": 0, "total": 0, "skipped": 0}
    skipped = denominator - used
    return {
        "percent": round((used / denominator) * 100),
        "used": used,
        "total": denominator,
        "skipped": skipped
    }


def escape_html(s):
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )


def render_input_with_highlights(sentences, mapping):
    """Renderuje materiał wejściowy z oznaczonymi statusami zdań."""
    map_dict = {m["id"]: m for m in mapping}
    parts = []
    for i, s in enumerate(sentences):
        sid = i + 1
        m = map_dict.get(sid, {"status": "USED", "reason": None})
        status = m.get("status", "USED")
        reason = m.get("reason") or ""

        if status == "USED":
            style = "background: transparent;"
            tooltip = ""
        elif status == "EXCLUDED":
            style = "background: #f1efe8; color: #888; text-decoration: line-through;"
            tooltip = f' title="EXCLUDED: {escape_html(reason)}"'
        else:  # SKIPPED
            style = "background: rgba(186, 117, 23, 0.15); color: #633806;"
            tooltip = f' title="SKIPPED: {escape_html(reason)}"'

        parts.append(
            f'<span style="{style} padding: 1px 3px; border-radius: 3px;"{tooltip}>'
            f'[{sid}] {escape_html(s)}</span>'
        )
    return " ".join(parts)


def render_output_with_flags(pr_sentences):
    """Renderuje informację prasową z oznaczonymi flagami."""
    parts = []
    for s in pr_sentences:
        text = s.get("text", "")
        added = s.get("added", False)
        supported_by = s.get("supported_by", []) or []

        if added:
            style = "background: rgba(186, 117, 23, 0.18); border-bottom: 1px dashed #BA7517;"
            tooltip = ' title="Uzupełnienie spoza materiału - zweryfikuj"'
        elif not supported_by:
            style = "background: rgba(226, 75, 74, 0.18); border-bottom: 1px solid #A32D2D;"
            tooltip = ' title="Brak podstawy w materiale wejściowym - zweryfikuj"'
        else:
            style = ""
            tooltip = f' title="Oparte na zdaniach: {", ".join(str(x) for x in supported_by)}"'

        parts.append(
            f'<span style="{style} padding: 1px 3px; border-radius: 3px;"{tooltip}>'
            f'{escape_html(text)}</span>'
        )
    return " ".join(parts)


def display_results(sentences, result):
    """Wyświetla wynik generowania."""
    usage = compute_usage(result.get("input_mapping", []))
    pr_sentences = result.get("press_release_sentences", [])
    flagged = sum(
        1 for s in pr_sentences
        if not s.get("added") and not (s.get("supported_by") or [])
    )
    added = sum(1 for s in pr_sentences if s.get("added"))

    # Statystyki
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(
            "Wykorzystanie materiału",
            f"{usage['percent']}%",
            help=f"{usage['used']} z {usage['total']} zdań kwalifikowanych (po wykluczeniu kategorii)"
        )
    with col2:
        st.metric(
            "Dodane fragmenty",
            added,
            help="Liczba zdań dodanych poza materiałem wejściowym (tryb kontekstowy)"
        )
    with col3:
        st.metric(
            "Czerwone flagi",
            flagged,
            help="Liczba zdań wyjścia bez wyraźnej podstawy w materiale - do weryfikacji"
        )

    # Ostrzeżenia
    warnings = result.get("warnings", [])
    if warnings:
        st.warning("**Ostrzeżenia operatora:**\n\n" + "\n".join(f"- {w}" for w in warnings))

    # Dwie kolumny: materiał i informacja
    left, right = st.columns(2)
    with left:
        st.subheader("Materiał wejściowy")
        st.caption("Najedź kursorem na fragment, aby zobaczyć status")
        html_input = render_input_with_highlights(sentences, result.get("input_mapping", []))
        st.markdown(
            f'<div style="font-size: 14px; line-height: 1.75;">{html_input}</div>',
            unsafe_allow_html=True
        )
        st.caption("🔘 Wykluczone (przekreślone)   🔶 Pominięte (bursztynowe)")

    with right:
        st.subheader("Informacja prasowa")
        st.caption("Najedź kursorem na fragment, aby zobaczyć źródło")
        html_output = render_output_with_flags(pr_sentences)
        st.markdown(
            f'<div style="font-size: 14px; line-height: 1.75;">{html_output}</div>',
            unsafe_allow_html=True
        )
        st.caption("🔶 Dodane spoza materiału   🔴 Bez podstawy w materiale")

    st.divider()

    # Surowy tekst do skopiowania
    st.subheader("Tekst gotowy do skopiowania")
    pr_text = result.get("press_release_text", "")
    st.code(pr_text, language=None)

    # Pobranie pliku
    st.download_button(
        label="Pobierz jako plik tekstowy",
        data=pr_text.encode("utf-8"),
        file_name="informacja_prasowa.txt",
        mime="text/plain"
    )


# ============================================================
# GŁÓWNY INTERFEJS
# ============================================================

def main_app():
    # Pasek górny
    col_title, col_logout = st.columns([5, 1])
    with col_title:
        st.title("PAP MediaRoom Generator")
    with col_logout:
        st.write("")
        if st.button("Wyloguj", use_container_width=True):
            st.session_state.authenticated = False
            st.rerun()

    # Sidebar z ustawieniami
    with st.sidebar:
        st.header("Ustawienia generowania")

        format_label = st.radio(
            "Format informacji",
            ["Sztywny", "Elastyczny"],
            help="Sztywny: tytuł + lead + korpus + kontakt + stopka. Elastyczny: dopuszczalne odstępstwa od struktury."
        )
        format_mode = "rigid" if format_label == "Sztywny" else "flexible"

        st.divider()

        supplement_label = st.radio(
            "Tryb uzupełnień",
            ["Zerowa tolerancja", "Kontekstowe"],
            help="Zerowa: tylko materiał klienta. Kontekstowe: model może dodać podstawowy kontekst, oznaczony w wyjściu."
        )
        supplement_mode = "zero" if supplement_label == "Zerowa tolerancja" else "contextual"

        st.divider()

        st.subheader("Wykluczone z liczenia")
        st.caption("Te kategorie nie są liczone do współczynnika wykorzystania")
        excluded = []
        if st.checkbox("Kontakt prasowy", value=True):
            excluded.append("kontakt prasowy")
        if st.checkbox("Nota o spółce / boilerplate", value=True):
            excluded.append("nota o spółce / boilerplate")
        if st.checkbox("Klauzule prawne / disclaimer", value=True):
            excluded.append("klauzule prawne / disclaimer")
        if st.checkbox("Stopki i podpisy", value=True):
            excluded.append("stopki i podpisy")

    # Główna część
    st.subheader("Materiał wejściowy od klienta")

    tab_text, tab_file = st.tabs(["Wklej tekst", "Wgraj plik (Word lub PDF)"])

    material_text = ""

    with tab_text:
        material_text = st.text_area(
            "Wklej tutaj materiał od klienta",
            height=300,
            label_visibility="collapsed",
            placeholder="Wklej tutaj materiał od klienta..."
        )

    with tab_file:
        uploaded = st.file_uploader(
            "Wybierz plik Word (.docx) lub PDF",
            type=["docx", "pdf"],
            label_visibility="collapsed"
        )
        if uploaded is not None:
            try:
                file_bytes = uploaded.read()
                if uploaded.name.lower().endswith(".docx"):
                    extracted = extract_text_from_docx(file_bytes)
                else:
                    extracted = extract_text_from_pdf(file_bytes)

                if not extracted.strip():
                    st.warning("Nie udało się wyciągnąć tekstu z pliku. Sprawdź czy plik nie jest skanem.")
                else:
                    st.success(f"Wyciągnięto {len(extracted)} znaków z pliku {uploaded.name}")
                    material_text = st.text_area(
                        "Wyodrębniony tekst (możesz go edytować przed wygenerowaniem)",
                        value=extracted,
                        height=300
                    )
            except Exception as e:
                st.error(f"Błąd odczytu pliku: {e}")

    # Przycisk generowania
    if st.button("Generuj informację prasową", type="primary", use_container_width=False):
        if not material_text or not material_text.strip():
            st.error("Wklej najpierw materiał albo wgraj plik.")
        else:
            sentences = split_sentences(material_text)
            if len(sentences) < 2:
                st.error("Materiał zbyt krótki. Wymagane minimum 2 zdania.")
            else:
                with st.spinner(f"Generuję informację prasową ({len(sentences)} zdań do analizy)..."):
                    try:
                        result = call_claude(sentences, format_mode, supplement_mode, excluded)
                        st.session_state.last_result = result
                        st.session_state.last_sentences = sentences
                    except Exception as e:
                        st.error(f"Błąd: {e}")
                        return

    # Wyświetl ostatni wynik (jeśli jest)
    if st.session_state.get("last_result") and st.session_state.get("last_sentences"):
        st.divider()
        display_results(
            st.session_state.last_sentences,
            st.session_state.last_result
        )


# ============================================================
# URUCHOMIENIE
# ============================================================

if check_authentication():
    main_app()
