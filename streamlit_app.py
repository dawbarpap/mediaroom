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

Wynik zwróć przez wywołanie narzędzia `generate_press_release` z odpowiednimi parametrami."""


# Schema narzędzia, które wymusza strukturę odpowiedzi modelu
GENERATE_TOOL = {
    "name": "generate_press_release",
    "description": "Zwraca wygenerowaną informację prasową PAP MediaRoom wraz z mapowaniem wykorzystania materiału wejściowego i listą ostrzeżeń.",
    "input_schema": {
        "type": "object",
        "properties": {
            "press_release_text": {
                "type": "string",
                "description": "Pełny tekst informacji prasowej w stylu PAP MediaRoom, z łamaniami linii między akapitami."
            },
            "press_release_sentences": {
                "type": "array",
                "description": "Wygenerowana informacja prasowa podzielona na pojedyncze zdania.",
                "items": {
                    "type": "object",
                    "properties": {
                        "text": {
                            "type": "string",
                            "description": "Treść pojedynczego zdania informacji prasowej."
                        },
                        "added": {
                            "type": "boolean",
                            "description": "Czy zdanie jest dodaniem spoza materiału wejściowego (tylko w trybie contextual)."
                        },
                        "supported_by": {
                            "type": "array",
                            "items": {"type": "integer"},
                            "description": "Lista ID zdań wejściowych, które wspierają to zdanie. Pusta lista oznacza brak podstawy w materiale (czerwona flaga)."
                        }
                    },
                    "required": ["text", "added", "supported_by"]
                }
            },
            "input_mapping": {
                "type": "array",
                "description": "Klasyfikacja każdego zdania wejściowego.",
                "items": {
                    "type": "object",
                    "properties": {
                        "id": {
                            "type": "integer",
                            "description": "Numer zdania wejściowego."
                        },
                        "status": {
                            "type": "string",
                            "enum": ["USED", "EXCLUDED", "SKIPPED"],
                            "description": "USED gdy wykorzystane, EXCLUDED gdy z kategorii wykluczonej, SKIPPED gdy świadomie pominięte."
                        },
                        "reason": {
                            "type": "string",
                            "description": "Krótkie uzasadnienie statusu (np. 'kontakt prasowy', 'powtórzenie informacji'). Dla USED może być pusty string."
                        }
                    },
                    "required": ["id", "status", "reason"]
                }
            },
            "warnings": {
                "type": "array",
                "items": {"type": "string"},
                "description": "Lista ostrzeżeń dla operatora (np. materiał ubogi, sprzeczność wewnętrzna, treść reklamowa). Pusta lista jeśli brak."
            }
        },
        "required": ["press_release_text", "press_release_sentences", "input_mapping", "warnings"]
    }
}


# ============================================================
# SYSTEM PROMPT I TOOL DLA TRYBU STAŻYSTY
# ============================================================

SYSTEM_PROMPT_TRAINEE = """Jesteś doświadczonym redaktorem PAP MediaRoom. Otrzymujesz informację prasową napisaną przez stażystę PAP. Twoim zadaniem jest ocenić ją i poprawić zgodnie ze standardami PAP MediaRoom, zachowując wszystkie fakty z oryginału stażysty.

ZASADY KLUCZOWE PAP MEDIAROOM:
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
- Wymyślanie faktów spoza tekstu stażysty. Pracujesz wyłącznie na tym, co napisał stażysta.
- Modyfikowanie treści cytatów (wolno tylko poprawiać atrybucję i cudzysłowy).
- Zmiana funkcji/tytułów rozmówców.
- Slogany, język marketingowy, wykrzykniki, pytania retoryczne.

ZADANIE:
1. Stwórz poprawioną wersję tekstu zgodną ze standardami PAP MediaRoom.
2. Wskaż konkretne uchybienia (co zostało źle zrobione i wymagało poprawki). Każde uchybienie sklasyfikuj według kategorii i wagi (poważne / drobne).
3. Wskaż konkretne zgodności ze stylebookiem (co stażysta zrobił dobrze, godne pochwały). To ważny element feedbacku motywacyjnego.
4. Podziel poprawiony tekst na zdania i oznacz, które zdania zostały zmienione względem oryginału.

KATEGORIE UCHYBIEŃ I ZGODNOŚCI:
- struktura (lead, korpus, kolejność, stopka)
- język (rzeczowość, neutralność, brak marketingu)
- cytaty (atrybucja, cudzysłowy, czasowniki wprowadzające)
- liczby (atrybucja, format)
- tytuł
- lead
- inne

Wynik zwróć przez wywołanie narzędzia review_trainee_text."""


TRAINEE_TOOL = {
    "name": "review_trainee_text",
    "description": "Ocenia i poprawia informację prasową napisaną przez stażystę PAP, zwracając poprawioną wersję, listę uchybień oraz listę elementów zgodnych ze stylebookiem.",
    "input_schema": {
        "type": "object",
        "properties": {
            "corrected_text": {
                "type": "string",
                "description": "Pełny poprawiony tekst informacji prasowej, gotowy do publikacji."
            },
            "corrected_sentences": {
                "type": "array",
                "description": "Poprawiona informacja prasowa podzielona na pojedyncze zdania.",
                "items": {
                    "type": "object",
                    "properties": {
                        "text": {
                            "type": "string",
                            "description": "Treść zdania w wersji poprawionej."
                        },
                        "changed": {
                            "type": "boolean",
                            "description": "Czy zdanie zostało zmienione względem oryginału stażysty (true) czy zachowane bez zmian (false)."
                        }
                    },
                    "required": ["text", "changed"]
                }
            },
            "compliances": {
                "type": "array",
                "description": "Lista konkretnych elementów, które stażysta zrobił zgodnie ze stylebookiem PAP MediaRoom (feedback pozytywny).",
                "items": {
                    "type": "object",
                    "properties": {
                        "category": {
                            "type": "string",
                            "enum": ["struktura", "język", "cytaty", "liczby", "tytuł", "lead", "inne"],
                            "description": "Kategoria zgodności."
                        },
                        "description": {
                            "type": "string",
                            "description": "Konkretny opis tego, co zostało zrobione dobrze."
                        }
                    },
                    "required": ["category", "description"]
                }
            },
            "issues": {
                "type": "array",
                "description": "Lista uchybień znalezionych w tekście stażysty (feedback wymagający poprawki).",
                "items": {
                    "type": "object",
                    "properties": {
                        "category": {
                            "type": "string",
                            "enum": ["struktura", "język", "cytaty", "liczby", "tytuł", "lead", "inne"],
                            "description": "Kategoria uchybienia."
                        },
                        "severity": {
                            "type": "string",
                            "enum": ["poważne", "drobne"],
                            "description": "Waga uchybienia: poważne (np. brak stopki, brak leadu, błędna atrybucja) lub drobne (np. niepolskie cudzysłowy, mała literówka stylistyczna)."
                        },
                        "description": {
                            "type": "string",
                            "description": "Konkretny opis uchybienia oraz tego, co zostało poprawione."
                        },
                        "fragment": {
                            "type": "string",
                            "description": "Fragment z oryginału stażysty, którego dotyczy uchybienie. Pusty string jeśli uchybienie dotyczy całości."
                        }
                    },
                    "required": ["category", "severity", "description", "fragment"]
                }
            }
        },
        "required": ["corrected_text", "corrected_sentences", "compliances", "issues"]
    }
}


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


DOT_PLACEHOLDER = "\u0001"  # tymczasowy znacznik chronionych kropek wewnątrz skrótów


def split_sentences(text):
    """Dzieli tekst na zdania, uwzględniając polskie skróty.

    Strategia: kropki wewnątrz skrótów (np. "r.", "proc.", "m.in.") oraz
    między cyframi (np. "12.04.2026") tymczasowo zastępujemy znacznikiem,
    dzielimy tekst, potem przywracamy oryginalne kropki.
    """
    cleaned = re.sub(r"\s+", " ", text.strip())
    if not cleaned:
        return []

    # Zastąp kropki w skrótach znacznikiem (case insensitive: "Dr." i "dr.")
    abbrev_pattern = r"\b(" + "|".join(POLISH_ABBREVIATIONS) + r")\."
    protected = re.sub(
        abbrev_pattern,
        lambda m: m.group(1) + DOT_PLACEHOLDER,
        cleaned,
        flags=re.IGNORECASE
    )
    # Specjalny przypadek "m.in."
    protected = re.sub(
        r"\bm\.in\.",
        "m" + DOT_PLACEHOLDER + "in" + DOT_PLACEHOLDER,
        protected,
        flags=re.IGNORECASE
    )
    # Liczby z kropką (daty, numery, dziesiętne)
    protected = re.sub(
        r"(\d)\.(\d)",
        lambda m: m.group(1) + DOT_PLACEHOLDER + m.group(2),
        protected
    )

    parts = re.findall(r"[^.!?]+[.!?]+", protected)
    if not parts:
        parts = [protected]
    return [p.replace(DOT_PLACEHOLDER, ".").strip() for p in parts if p.strip()]


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


_SENTINEL = object()


def _safe_get(obj, key, default=None):
    """Pobiera wartość spod klucza/atrybutu, niezależnie od typu obiektu.
    Działa na słownikach, modelach pydantic, namedtuples, dataclassach.
    """
    if obj is None:
        return default
    if isinstance(obj, dict):
        return obj.get(key, default)
    val = getattr(obj, key, _SENTINEL)
    if val is not _SENTINEL:
        return val
    try:
        return obj[key]
    except (TypeError, KeyError, IndexError):
        return default


def _to_plain_python(obj):
    """Konwertuje cokolwiek do czystych typów Pythona przez serializację JSON.
    Niezależnie od tego, czy SDK zwraca model pydantic, dict, MappingProxy,
    namedtuple czy inny obiekt, wynikiem jest gwarantowany dict/list/str/int/bool/None.
    """
    def _fallback(o):
        # Najpierw próbujemy model_dump (pydantic v2)
        dump = getattr(o, "model_dump", None)
        if callable(dump):
            try:
                return dump(mode="python")
            except Exception:
                pass
        # Potem .dict() (pydantic v1, jeśli ktoś tego używa)
        dict_method = getattr(o, "dict", None)
        if callable(dict_method) and not isinstance(o, dict):
            try:
                result = dict_method()
                if isinstance(result, dict):
                    return result
            except Exception:
                pass
        # Potem __dict__ (zwykłe obiekty Pythona)
        if hasattr(o, "__dict__"):
            d = vars(o)
            if d:
                return {k: v for k, v in d.items() if not k.startswith("_")}
        # Ostateczność: konwersja do stringa
        return str(o)

    # Krok 1: pre-konwersja, jeśli sam obiekt jest pydantic
    dump = getattr(obj, "model_dump", None)
    if callable(dump):
        try:
            obj = dump(mode="python")
        except Exception:
            pass

    # Krok 2: round-trip przez JSON wymusza czyste typy
    try:
        return json.loads(json.dumps(obj, default=_fallback, ensure_ascii=False))
    except Exception:
        # Awaryjny manualny przepływ rekursywny
        if isinstance(obj, dict):
            return {str(k): _to_plain_python(v) for k, v in obj.items()}
        if isinstance(obj, (list, tuple)):
            return [_to_plain_python(item) for item in obj]
        return obj


def call_claude(sentences, format_mode, supplement_mode, excluded):
    """Wywołuje Claude API używając tool use i zwraca strukturyzowany wynik."""
    client = anthropic.Anthropic(api_key=st.secrets["anthropic_api_key"])
    user_prompt = build_user_prompt(sentences, format_mode, supplement_mode, excluded)

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8000,
        temperature=0,
        system=SYSTEM_PROMPT,
        tools=[GENERATE_TOOL],
        tool_choice={"type": "tool", "name": "generate_press_release"},
        messages=[{"role": "user", "content": user_prompt}]
    )

    # Wynik znajduje się w bloku tool_use. Normalizujemy do czystego dict.
    for block in message.content:
        if block.type == "tool_use" and block.name == "generate_press_release":
            return _to_plain_python(block.input)

    raise ValueError(
        "Model nie wywołał oczekiwanego narzędzia. "
        f"Otrzymano typy bloków: {[b.type for b in message.content]}"
    )


def call_claude_trainee(text):
    """Wywołuje Claude API w trybie poprawiania tekstu stażysty."""
    client = anthropic.Anthropic(api_key=st.secrets["anthropic_api_key"])

    user_prompt = f"""TEKST STAŻYSTY DO OCENY I POPRAWY:

{text}

Wykonaj zadanie i wywołaj narzędzie review_trainee_text."""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8000,
        temperature=0,
        system=SYSTEM_PROMPT_TRAINEE,
        tools=[TRAINEE_TOOL],
        tool_choice={"type": "tool", "name": "review_trainee_text"},
        messages=[{"role": "user", "content": user_prompt}]
    )

    for block in message.content:
        if block.type == "tool_use" and block.name == "review_trainee_text":
            return _to_plain_python(block.input)

    raise ValueError(
        "Model nie wywołał oczekiwanego narzędzia. "
        f"Otrzymano typy bloków: {[b.type for b in message.content]}"
    )


# ============================================================
# OBLICZENIA I RENDEROWANIE
# ============================================================

def compute_usage(mapping):
    total = len(mapping)
    excluded = sum(1 for m in mapping if _safe_get(m, "status") == "EXCLUDED")
    used = sum(1 for m in mapping if _safe_get(m, "status") == "USED")
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
    map_dict = {_safe_get(m, "id"): m for m in mapping if _safe_get(m, "id") is not None}
    parts = []
    for i, s in enumerate(sentences):
        sid = i + 1
        m = map_dict.get(sid, {"status": "USED", "reason": ""})
        status = _safe_get(m, "status", "USED")
        reason = _safe_get(m, "reason", "") or ""

        if status == "USED":
            style = "background: transparent;"
            tooltip = ""
        elif status == "EXCLUDED":
            style = "background: #f1efe8; color: #888; text-decoration: line-through;"
            tooltip = f' title="EXCLUDED: {escape_html(reason)}"' if reason else ' title="EXCLUDED"'
        else:  # SKIPPED
            style = "background: rgba(186, 117, 23, 0.15); color: #633806;"
            tooltip = f' title="SKIPPED: {escape_html(reason)}"' if reason else ' title="SKIPPED"'

        parts.append(
            f'<span style="{style} padding: 1px 3px; border-radius: 3px;"{tooltip}>'
            f'[{sid}] {escape_html(s)}</span>'
        )
    return " ".join(parts)


def render_output_with_flags(pr_sentences):
    """Renderuje informację prasową z oznaczonymi flagami."""
    parts = []
    for s in pr_sentences:
        text = _safe_get(s, "text", "")
        added = _safe_get(s, "added", False)
        supported_by = _safe_get(s, "supported_by", []) or []

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
    input_mapping = _safe_get(result, "input_mapping", []) or []
    pr_sentences = _safe_get(result, "press_release_sentences", []) or []
    warnings = _safe_get(result, "warnings", []) or []
    pr_text = _safe_get(result, "press_release_text", "") or ""

    usage = compute_usage(input_mapping)
    flagged = sum(
        1 for s in pr_sentences
        if not _safe_get(s, "added") and not (_safe_get(s, "supported_by") or [])
    )
    added = sum(1 for s in pr_sentences if _safe_get(s, "added"))

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
    if warnings:
        st.warning("**Ostrzeżenia operatora:**\n\n" + "\n".join(f"- {w}" for w in warnings))

    # Dwie kolumny: materiał i informacja
    left, right = st.columns(2)
    with left:
        st.subheader("Materiał wejściowy")
        st.caption("Najedź kursorem na fragment, aby zobaczyć status")
        html_input = render_input_with_highlights(sentences, input_mapping)
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
    st.code(pr_text, language=None)

    # Pobranie pliku
    st.download_button(
        label="Pobierz jako plik tekstowy",
        data=pr_text.encode("utf-8"),
        file_name="informacja_prasowa.txt",
        mime="text/plain"
    )


def render_trainee_corrected(corrected_sentences):
    """Renderuje poprawiony tekst stażysty z podświetleniem zmienionych zdań."""
    parts = []
    for s in corrected_sentences:
        text = _safe_get(s, "text", "")
        changed = _safe_get(s, "changed", False)
        if changed:
            style = "background: rgba(99, 122, 145, 0.18); border-bottom: 1px solid #4a8db5;"
            tooltip = ' title="Zdanie poprawione przez redaktora"'
        else:
            style = ""
            tooltip = ' title="Zdanie zachowane bez zmian"'
        parts.append(
            f'<span style="{style} padding: 1px 3px; border-radius: 3px;"{tooltip}>'
            f'{escape_html(text)}</span>'
        )
    return " ".join(parts)


def render_trainee_original(text):
    """Renderuje oryginalny tekst stażysty z zachowaniem łamań linii."""
    return escape_html(text).replace("\n", "<br>")


def display_trainee_results(original_text, result):
    """Wyświetla wynik oceny i poprawy tekstu stażysty."""
    corrected_text = _safe_get(result, "corrected_text", "") or ""
    corrected_sentences = _safe_get(result, "corrected_sentences", []) or []
    compliances = _safe_get(result, "compliances", []) or []
    issues = _safe_get(result, "issues", []) or []

    major_issues = sum(1 for i in issues if _safe_get(i, "severity") == "poważne")
    minor_issues = sum(1 for i in issues if _safe_get(i, "severity") == "drobne")
    changed_count = sum(1 for s in corrected_sentences if _safe_get(s, "changed"))

    # Statystyki
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric(
            "Zgodne ze stylebookiem",
            len(compliances),
            help="Liczba elementów, które stażysta zrobił zgodnie ze standardami PAP."
        )
    with col2:
        st.metric(
            "Poważne uchybienia",
            major_issues,
            help="Uchybienia, które wymagają obowiązkowej poprawki przed publikacją."
        )
    with col3:
        st.metric(
            "Drobne uchybienia",
            minor_issues,
            help="Uchybienia stylistyczne lub formatowania, mniej krytyczne."
        )
    with col4:
        st.metric(
            "Zmienione zdania",
            changed_count,
            help="Liczba zdań poprawionych względem oryginału stażysty."
        )

    # Dwie kolumny: oryginał i poprawiona wersja
    left, right = st.columns(2)
    with left:
        st.subheader("Tekst stażysty (oryginał)")
        st.caption("Wersja przed redakcją")
        html_original = render_trainee_original(original_text)
        st.markdown(
            f'<div style="font-size: 14px; line-height: 1.75;">{html_original}</div>',
            unsafe_allow_html=True
        )

    with right:
        st.subheader("Wersja poprawiona")
        st.caption("Najedź kursorem na fragment, aby zobaczyć status zmiany")
        html_corrected = render_trainee_corrected(corrected_sentences)
        st.markdown(
            f'<div style="font-size: 14px; line-height: 1.75;">{html_corrected}</div>',
            unsafe_allow_html=True
        )
        st.caption("🔵 Zdanie zmienione przez redakcję")

    st.divider()

    # Dwie listy: zgodności i uchybienia
    col_good, col_bad = st.columns(2)
    with col_good:
        st.subheader("Zgodne ze stylebookiem")
        if compliances:
            for c in compliances:
                cat = _safe_get(c, "category", "inne")
                desc = _safe_get(c, "description", "")
                st.markdown(f"✓ **{cat.capitalize()}**: {desc}")
        else:
            st.caption("Brak wyróżnionych zgodności w tekście.")

    with col_bad:
        st.subheader("Uchybienia do poprawy")
        if issues:
            for i in issues:
                cat = _safe_get(i, "category", "inne")
                sev = _safe_get(i, "severity", "drobne")
                desc = _safe_get(i, "description", "")
                fragment = _safe_get(i, "fragment", "")
                badge = "🔴" if sev == "poważne" else "🟡"
                st.markdown(f"{badge} **{cat.capitalize()}** ({sev}): {desc}")
                if fragment:
                    st.caption(f'Fragment: „{fragment}"')
        else:
            st.caption("Brak uchybień. Tekst zgodny ze stylebookiem.")

    st.divider()

    # Surowy tekst do skopiowania
    st.subheader("Tekst gotowy do skopiowania")
    st.code(corrected_text, language=None)

    # Pobranie pliku
    st.download_button(
        label="Pobierz poprawiony tekst",
        data=corrected_text.encode("utf-8"),
        file_name="informacja_prasowa_poprawiona.txt",
        mime="text/plain",
        key="download_trainee"
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

    # Wybór trybu pracy
    mode = st.radio(
        "Tryb pracy",
        ["Pisanie z materiału klienta", "Poprawianie tekstu stażysty"],
        horizontal=True,
        key="work_mode",
        help="Pisanie: tworzysz informację prasową na podstawie materiału od klienta. Poprawianie: oceniasz i redagujesz gotową informację prasową napisaną przez stażystę."
    )

    st.divider()

    if mode == "Pisanie z materiału klienta":
        _client_material_flow()
    else:
        _trainee_correction_flow()


def _client_material_flow():
    """Tryb pisania informacji prasowej na podstawie materiału klienta."""
    # Sidebar z ustawieniami
    with st.sidebar:
        st.header("Ustawienia generowania")
        st.caption("Tryb: pisanie z materiału klienta")

        format_label = st.radio(
            "Format informacji",
            ["Sztywny", "Elastyczny"],
            help="Sztywny: tytuł + lead + korpus + kontakt + stopka. Elastyczny: dopuszczalne odstępstwa od struktury.",
            key="client_format"
        )
        format_mode = "rigid" if format_label == "Sztywny" else "flexible"

        st.divider()

        supplement_label = st.radio(
            "Tryb uzupełnień",
            ["Zerowa tolerancja", "Kontekstowe"],
            help="Zerowa: tylko materiał klienta. Kontekstowe: model może dodać podstawowy kontekst, oznaczony w wyjściu.",
            key="client_supplement"
        )
        supplement_mode = "zero" if supplement_label == "Zerowa tolerancja" else "contextual"

        st.divider()

        st.subheader("Wykluczone z liczenia")
        st.caption("Te kategorie nie są liczone do współczynnika wykorzystania")
        excluded = []
        if st.checkbox("Kontakt prasowy", value=True, key="client_excl_kontakt"):
            excluded.append("kontakt prasowy")
        if st.checkbox("Nota o spółce / boilerplate", value=True, key="client_excl_nota"):
            excluded.append("nota o spółce / boilerplate")
        if st.checkbox("Klauzule prawne / disclaimer", value=True, key="client_excl_klauzule"):
            excluded.append("klauzule prawne / disclaimer")
        if st.checkbox("Stopki i podpisy", value=True, key="client_excl_stopki"):
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
            placeholder="Wklej tutaj materiał od klienta...",
            key="client_material_input"
        )

    with tab_file:
        uploaded = st.file_uploader(
            "Wybierz plik Word (.docx) lub PDF",
            type=["docx", "pdf"],
            label_visibility="collapsed",
            key="client_file_upload"
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
                        height=300,
                        key="client_extracted_text"
                    )
            except Exception as e:
                st.error(f"Błąd odczytu pliku: {e}")

    # Przycisk generowania
    if st.button("Generuj informację prasową", type="primary", key="client_generate"):
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
                        st.session_state.last_client_result = result
                        st.session_state.last_client_sentences = sentences
                    except Exception as e:
                        st.error(f"Błąd: {e}")
                        return

    # Wyświetl ostatni wynik (jeśli jest)
    if st.session_state.get("last_client_result") and st.session_state.get("last_client_sentences"):
        st.divider()
        display_results(
            st.session_state.last_client_sentences,
            st.session_state.last_client_result
        )


def _trainee_correction_flow():
    """Tryb poprawiania informacji prasowej napisanej przez stażystę."""
    # Sidebar (informacyjnie)
    with st.sidebar:
        st.header("Tryb stażysty")
        st.caption(
            "Wklej lub wgraj informację prasową napisaną przez stażystę. "
            "Aplikacja oceni ją zgodnie ze standardami PAP MediaRoom, "
            "wskaże uchybienia oraz zgodności i zaproponuje poprawioną wersję."
        )

    # Główna część
    st.subheader("Tekst stażysty do oceny i poprawy")

    tab_text, tab_file = st.tabs(["Wklej tekst", "Wgraj plik (Word lub PDF)"])

    trainee_text = ""

    with tab_text:
        trainee_text = st.text_area(
            "Wklej tutaj tekst stażysty",
            height=300,
            label_visibility="collapsed",
            placeholder="Wklej tutaj informację prasową napisaną przez stażystę...",
            key="trainee_text_input"
        )

    with tab_file:
        uploaded = st.file_uploader(
            "Wybierz plik Word (.docx) lub PDF",
            type=["docx", "pdf"],
            label_visibility="collapsed",
            key="trainee_file_upload"
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
                    trainee_text = st.text_area(
                        "Wyodrębniony tekst (możesz go edytować przed sprawdzeniem)",
                        value=extracted,
                        height=300,
                        key="trainee_extracted_text"
                    )
            except Exception as e:
                st.error(f"Błąd odczytu pliku: {e}")

    # Przycisk
    if st.button("Sprawdź i popraw", type="primary", key="trainee_review"):
        if not trainee_text or not trainee_text.strip():
            st.error("Wklej najpierw tekst stażysty albo wgraj plik.")
        elif len(trainee_text.strip()) < 100:
            st.error("Tekst zbyt krótki. Wymagane minimum 100 znaków.")
        else:
            with st.spinner("Sprawdzam i poprawiam tekst..."):
                try:
                    result = call_claude_trainee(trainee_text)
                    st.session_state.last_trainee_result = result
                    st.session_state.last_trainee_text = trainee_text
                except Exception as e:
                    st.error(f"Błąd: {e}")
                    return

    # Wyświetl ostatni wynik
    if st.session_state.get("last_trainee_result") and st.session_state.get("last_trainee_text"):
        st.divider()
        display_trainee_results(
            st.session_state.last_trainee_text,
            st.session_state.last_trainee_result
        )


# ============================================================
# URUCHOMIENIE
# ============================================================

if check_authentication():
    main_app()
