#!/usr/bin/env python3
"""
compare_transcripts.py
------------------------
Primerja dva transkripta (NOV in TRENUTNI) z referenčnimi podnapisi VTT.
Izpiše odstotke podobnosti in kategorizirane razlike.

Uporaba:
    python3 compare_transcripts.py --nov pot/do/nov.docx --trenutni pot/do/trenutni.docx --vtt pot/do/podnapisi.vtt

Podprti formati vhodnih datotek:
    .docx  → Word dokument
    .txt   → navadni tekst (UTF-8)
    .vtt   → WebVTT podnapisi (za vse tri argumente)
"""

import re
import sys
import argparse
import difflib
import json
from pathlib import Path


# ─────────────────────────────────────────────
#  BRANJE DATOTEK
# ─────────────────────────────────────────────

def read_docx(path: str) -> str:
    try:
        import docx
    except ImportError:
        print("NAPAKA: python-docx ni nameščen. Namesti ga z: pip install python-docx")
        sys.exit(1)
    doc = docx.Document(path)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    return " ".join(paragraphs)


def read_vtt(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()
    # Odstrani WEBVTT glavo
    content = re.sub(r'^WEBVTT.*?\n', '', content, flags=re.MULTILINE)
    # Odstrani časovne oznake
    content = re.sub(r'\d{2}:\d{2}:\d{2}[\.,]\d{3}\s*-->\s*\d{2}:\d{2}:\d{2}[\.,]\d{3}[^\n]*\n', '', content)
    # Odstrani barvne in ostale HTML oznake
    content = re.sub(r'<[^>]+>', '', content)
    # Odstrani NOTE bloke
    content = re.sub(r'^NOTE\b.*', '', content, flags=re.MULTILINE)
    return content


def read_txt(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def read_file(path: str) -> str:
    """Avtomatsko prepozna format in prebere vsebino."""
    ext = Path(path).suffix.lower()
    if ext == ".docx":
        return read_docx(path)
    elif ext == ".vtt":
        return read_vtt(path)
    elif ext in (".txt", ".text"):
        return read_txt(path)
    else:
        # Poskusi kot TXT
        try:
            return read_txt(path)
        except Exception:
            print(f"NAPAKA: Nepodprt format datoteke: {ext}")
            sys.exit(1)


# ─────────────────────────────────────────────
#  ČIŠČENJE IN TOKENIZACIJA
# ─────────────────────────────────────────────

SPEAKER_PATTERN = re.compile(
    r'(\*{1,2})?'          # opcijski markdown bold (* ali **)
    r'speaker[_\s-]?\d+'   # speaker_1, speaker 1, speaker-1, SPEAKER_01 ...
    r'(\*{1,2})?'          # opcijski zaključni bold
    r'[:\s]*',             # opcijski dvopičje in presledki za oznako
    re.IGNORECASE
)

def clean_text(text: str) -> str:
    """Odstrani oznake govorcev, normalizira presledke."""
    text = SPEAKER_PATTERN.sub('', text)
    text = re.sub(r'\n+', ' ', text)
    text = re.sub(r' +', ' ', text)
    return text.strip()


def tokenize(text: str) -> list:
    """Vrne seznam besed in ločil kot ločenih tokenov."""
    return re.findall(r'\S+', text)


# ─────────────────────────────────────────────
#  KATEGORIZACIJA RAZLIK
# ─────────────────────────────────────────────

ABBREVIATIONS = {
    'oz.', 'itn.', 'ipd.', 'npr.', 'tj.', 'dr.', 'prof.',
    'mag.', 'itd.', 'apd.', 'sl.', 'st.', 'str.', 'let.',
}

# Zapis številk prepoznamo samo pri kratkih blokih — pri daljših je razlika
# skoraj zagotovo vsebinska in številka je le naključno zraven.
MAX_NUMFMT_BLOCK = 4

def has_abbreviation(part: str) -> bool:
    """True, če je katerikoli TOKEN v bloku okrajšava (ne podniz besede!)."""
    for tok in part.split():
        if tok.lower().strip('"\'()[]«»,;:') in ABBREVIATIONS:
            return True
    return False


def categorize_diff(a_tokens: list, b_tokens: list) -> dict:
    """
    Primerja dva seznama tokenov in vrne kategorizirane razlike.

    Kategorije:
      - punct_only   : razlika samo v ločilih (vejica/pika/velika začetnica)
      - number_format: število zapisano s cifro vs. z besedo
      - abbreviation : okrajšava vs. polna beseda
      - word_diff    : vsebinska razlika (napačna beseda / napačna prepoznava)

    Vsaka kategorija se šteje dvakrat:
      - <kategorija>        = število BLOKOV razlik (en blok = ena strnjena razlika,
                              lahko obsega 1 ali 30 besed)
      - <kategorija>_tokens = število BESED v teh blokih  ← edino primerljivo med
                              različnimi transkripti, ker je neodvisno od tega,
                              kako difflib združi razlike v bloke
    """
    sm = difflib.SequenceMatcher(None, a_tokens, b_tokens, autojunk=False)
    opcodes = sm.get_opcodes()

    stats = {
        "ratio":        sm.ratio(),
        "equal":        0,
        "total_diff_blocks": 0,
        "total_diff_tokens": 0,
        "punct_only":   0,
        "number_format": 0,
        "abbreviation": 0,
        "word_diff":    0,
        "punct_only_tokens":    0,
        "number_format_tokens": 0,
        "abbreviation_tokens":  0,
        "word_diff_tokens":     0,
        "replaced":     0,
        "inserted":     0,
        "deleted":      0,
        "a_len":        len(a_tokens),
        "b_len":        len(b_tokens),
        "examples": {
            "punct_only":    [],
            "number_format": [],
            "abbreviation":  [],
            "word_diff":     [],
        }
    }

    MAX_EXAMPLES = 15  # koliko primerov shranimo na kategorijo

    for op, i1, i2, j1, j2 in opcodes:
        if op == "equal":
            stats["equal"] += (i2 - i1)
            continue

        # Velikost bloka v besedah (daljša stran razlike)
        block_tokens = max(i2 - i1, j2 - j1)

        stats["total_diff_blocks"] += 1
        stats["total_diff_tokens"] += block_tokens

        if op == "replace":
            stats["replaced"] += max(i2 - i1, j2 - j1)
        elif op == "insert":
            stats["inserted"] += (j2 - j1)
        elif op == "delete":
            stats["deleted"] += (i2 - i1)

        a_part = " ".join(a_tokens[i1:i2]).strip()
        b_part = " ".join(b_tokens[j1:j2]).strip()

        example = (a_part, b_part)

        # Samo ločila: besedi sta enaki, ko odstranimo ločila in upoštevamo male/velike
        a_core = re.sub(r'[^\w\s]', '', a_part).strip().lower()
        b_core = re.sub(r'[^\w\s]', '', b_part).strip().lower()

        a_has_digit = bool(re.search(r'\d', a_part))
        b_has_digit = bool(re.search(r'\d', b_part))
        short_block = (i2 - i1) <= MAX_NUMFMT_BLOCK and (j2 - j1) <= MAX_NUMFMT_BLOCK

        if a_core == b_core:
            category = "punct_only"
        elif short_block and (a_has_digit != b_has_digit):
            category = "number_format"
        elif has_abbreviation(a_part) or has_abbreviation(b_part):
            category = "abbreviation"
        else:
            category = "word_diff"

        stats[category] += 1
        stats[category + "_tokens"] += block_tokens
        if len(stats["examples"][category]) < MAX_EXAMPLES:
            stats["examples"][category].append(example)

    return stats


# ─────────────────────────────────────────────
#  IZPIS
# ─────────────────────────────────────────────

RESET  = "\033[0m"
BOLD   = "\033[1m"
GREEN  = "\033[92m"
RED    = "\033[91m"
YELLOW = "\033[93m"
CYAN   = "\033[96m"
GRAY   = "\033[90m"

def bar(pct: float, width: int = 30) -> str:
    filled = round(pct / 100 * width)
    return "█" * filled + "░" * (width - filled)

def print_header(title: str):
    print(f"\n{BOLD}{CYAN}{'═'*60}{RESET}")
    print(f"{BOLD}{CYAN}  {title}{RESET}")
    print(f"{BOLD}{CYAN}{'═'*60}{RESET}")

def print_comparison(label_a: str, label_b: str, stats: dict, total_ref: int):
    similarity = stats["ratio"] * 100
    diff_pct   = 100 - similarity

    color = GREEN if similarity >= 88 else (YELLOW if similarity >= 80 else RED)

    print(f"\n{BOLD}── {label_a}  vs  {label_b} ──{RESET}")
    print(f"  Podobnost:  {color}{similarity:5.1f}%{RESET}  {bar(similarity)}")
    print(f"  Razlika:    {color}{diff_pct:5.1f}%{RESET}")
    print()
    print(f"  {'':<38} {'blokov':>8}  {'besed':>8}  {'% besed ref.':>13}")
    print(f"  {'Skupaj razlik:':<38} {stats['total_diff_blocks']:>8}  {stats['total_diff_tokens']:>8}  "
          f"{stats['total_diff_tokens']/max(total_ref,1)*100:>12.1f}%")
    for key, naslov, barva in (
        ("punct_only",    "└─ samo ločila (vejica, pika...):",     GRAY),
        ("number_format", "└─ zapis številk (38 % / odstotkov):",  GRAY),
        ("abbreviation",  "└─ okrajšave (oz., itn., ...):",        GRAY),
        ("word_diff",     "└─ vsebinske razlike (napačne besede):", RED),
    ):
        print(f"  {barva}{naslov:<38}{RESET} {barva}{stats[key]:>8}  {stats[key+'_tokens']:>8}  "
              f"{stats[key+'_tokens']/max(total_ref,1)*100:>12.1f}%{RESET}")

    print(f"\n  Enake besede: {stats['equal']}  |  Zamenjane: {stats['replaced']}  "
          f"|  Vstavljene: {stats['inserted']}  |  Izbrisane: {stats['deleted']}")
    print(f"  {GRAY}Opomba: 'blokov' ni primerljivo med transkripti (en blok = 1 ali 30 besed) — "
          f"primerjaj stolpec 'besed'.{RESET}")


def print_examples(stats: dict, label_a: str, label_b: str, show_n: int = 8):
    cats = [
        ("word_diff",    "🔴 Vsebinske razlike (napačne besede)"),
        ("punct_only",   "🟡 Samo ločila"),
        ("number_format","🔵 Zapis številk"),
        ("abbreviation", "🟢 Okrajšave"),
    ]
    for key, title in cats:
        examples = stats["examples"][key]
        if not examples:
            continue
        print(f"\n  {BOLD}{title}{RESET}  ({len(examples)} prikazanih od {stats[key]} skupaj)")
        print(f"  {'':>4}  {label_a:<35}  →  {label_b}")
        print(f"  {'─'*75}")
        for i, (a, b) in enumerate(examples[:show_n]):
            a_fmt = f'"{a}"' if a else "(prazno)"
            b_fmt = f'"{b}"' if b else "(prazno)"
            print(f"  {i+1:>3}.  {a_fmt:<35}  →  {b_fmt}")


def format_examples_for_md(stats: dict, label_a: str, label_b: str) -> str:
    """Formatira primere razlik za Markdown."""
    lines = []
    cats = [
        ("word_diff",    "🔴 Vsebinske razlike (napačne besede)"),
        ("punct_only",   "🟡 Samo ločila"),
        ("number_format","🔵 Zapis številk"),
        ("abbreviation", "🟢 Okrajšave"),
    ]
    for key, title in cats:
        examples = stats["examples"][key]
        if not examples:
            continue
        lines.append(f"\n#### {title}\n")
        lines.append(f"({len(examples)} prikazanih od {stats[key]} skupaj)\n")
        lines.append(f"| # | {label_a} | {label_b} |\n")
        lines.append(f"|---|---|---|\n")
        for i, (a, b) in enumerate(examples[:15], 1):
            a_fmt = f"`{a}`" if a else "*(prazno)*"
            b_fmt = f"`{b}`" if b else "*(prazno)*"
            lines.append(f"| {i} | {a_fmt} | {b_fmt} |\n")
    return "".join(lines)


def md_breakdown(stats: dict, total_ref: int) -> list:
    """Razčlenitev razlik po kategorijah — bloki IN besede."""
    lines = []
    lines.append("\n| Kategorija | Blokov | Besed | % besed reference |\n")
    lines.append("|---|---:|---:|---:|\n")
    for key, naslov in (
        ("punct_only",    "Samo ločila"),
        ("number_format", "Zapis številk"),
        ("abbreviation",  "Okrajšave"),
        ("word_diff",     "**Vsebinske razlike**"),
    ):
        lines.append(f"| {naslov} | {stats[key]} | {stats[key+'_tokens']} | "
                     f"{stats[key+'_tokens']/max(total_ref,1)*100:.1f}% |\n")
    lines.append(f"| _Skupaj_ | {stats['total_diff_blocks']} | {stats['total_diff_tokens']} | "
                 f"{stats['total_diff_tokens']/max(total_ref,1)*100:.1f}% |\n")
    return lines


def build_markdown_report(stats_nov: dict, stats_curr: dict, stats_cross: dict,
                          n_vtt: int, args) -> str:
    """Gradi Markdown poročilo s rezultati in primeri."""
    sim_nov = stats_nov["ratio"] * 100
    sim_curr = stats_curr["ratio"] * 100
    sim_cross = stats_cross["ratio"] * 100
    winner = "TRENUTNI" if sim_curr > sim_nov else "NOV"
    diff = abs(sim_nov - sim_curr)

    lines = []
    lines.append("# Primerjava transkriptov\n")
    lines.append(f"**Datoteka NOV:** `{Path(args.nov).name}`  \n")
    lines.append(f"**Datoteka TRENUTNI:** `{Path(args.trenutni).name}`  \n")
    lines.append(f"**Referenca VTT:** `{Path(args.vtt).name}` ({n_vtt} besed)\n")

    lines.append("\n> **Kako brati številke:** _blok_ je ena strnjena razlika in lahko obsega "
                 "1 ali 30 besed, zato število blokov med transkriptoma **ni primerljivo**. "
                 "Primerljiv je stolpec _besed_ oz. odstotek glede na referenco.\n")

    lines.append("\n## Rezultati primerjave\n")

    lines.append("### VTT vs NOV\n")
    lines.append(f"- **Podobnost:** {sim_nov:.1f}%\n")
    lines.append(f"- **Razlika:** {100-sim_nov:.1f}%\n")
    lines.extend(md_breakdown(stats_nov, n_vtt))

    lines.append("\n### VTT vs TRENUTNI\n")
    lines.append(f"- **Podobnost:** {sim_curr:.1f}%\n")
    lines.append(f"- **Razlika:** {100-sim_curr:.1f}%\n")
    lines.extend(md_breakdown(stats_curr, n_vtt))

    lines.append("\n## Medsebojna primerjava: TRENUTNI vs NOV\n")
    lines.append("_Brez reference — koliko se transkripta razlikujeta med seboj._\n")
    lines.append(f"\n- **Podobnost:** {sim_cross:.1f}%\n")
    lines.append(f"- **Razlika:** {100-sim_cross:.1f}%\n")
    lines.append(f"- Dolžina: TRENUTNI {stats_cross['a_len']} besed, NOV {stats_cross['b_len']} besed "
                 f"(razlika {stats_cross['b_len'] - stats_cross['a_len']:+d})\n")
    lines.extend(md_breakdown(stats_cross, stats_cross["a_len"]))

    lines.append("\n## Primeri razlik\n")

    lines.append("\n### VTT vs NOV\n")
    lines.append(format_examples_for_md(stats_nov, "VTT", "NOV"))

    lines.append("\n### VTT vs TRENUTNI\n")
    lines.append(format_examples_for_md(stats_curr, "VTT", "TRENUTNI"))

    lines.append("\n### TRENUTNI vs NOV\n")
    lines.append(format_examples_for_md(stats_cross, "TRENUTNI", "NOV"))

    lines.append("\n## Skupni rezultat\n")
    lines.append("| Transkript | Podobnost | Razlika | Vsebinske napake (besed) | % besed ref. | Blokov |\n")
    lines.append("|---|---|---|---:|---:|---:|\n")
    lines.append(f"| NOV | {sim_nov:.1f}% | {100-sim_nov:.1f}% | {stats_nov['word_diff_tokens']} | "
                 f"{stats_nov['word_diff_tokens']/max(n_vtt,1)*100:.1f}% | {stats_nov['word_diff']} |\n")
    lines.append(f"| TRENUTNI | {sim_curr:.1f}% | {100-sim_curr:.1f}% | {stats_curr['word_diff_tokens']} | "
                 f"{stats_curr['word_diff_tokens']/max(n_vtt,1)*100:.1f}% | {stats_curr['word_diff']} |\n")
    lines.append(f"\n_Medsebojna podobnost TRENUTNI ↔ NOV: {sim_cross:.1f}%_\n")

    lines.append(f"\n## 🏆 Zmagovalec\n")
    lines.append(f"**{winner}** (za {diff:.1f} točk boljši kot drugi)\n")

    return "".join(lines)


def print_summary(stats_nov: dict, stats_curr: dict, stats_cross: dict, n_vtt: int):
    sim_nov  = stats_nov["ratio"]  * 100
    sim_curr = stats_curr["ratio"] * 100

    print_header("SKUPNI REZULTAT")

    winner = "TRENUTNI" if sim_curr > sim_nov else "NOV"
    diff = abs(sim_nov - sim_curr)

    color_n = GREEN if sim_nov  > sim_curr else RED
    color_c = GREEN if sim_curr > sim_nov  else RED

    print(f"\n  {'Transkript':<12}  {'Podobnost z VTT':>16}  {'Razlika':>9}  "
          f"{'Vseb. napak (besed)':>21}  {'% ref.':>8}  {'blokov':>7}")
    print(f"  {'─'*82}")
    for label, s, c in (("NOV", stats_nov, color_n), ("TRENUTNI", stats_curr, color_c)):
        sim = s["ratio"] * 100
        print(f"  {c}{label:<12}{RESET}  {c}{sim:>15.1f}%{RESET}  {c}{100-sim:>8.1f}%{RESET}  "
              f"{c}{s['word_diff_tokens']:>21}{RESET}  "
              f"{c}{s['word_diff_tokens']/max(n_vtt,1)*100:>7.1f}%{RESET}  {c}{s['word_diff']:>7}{RESET}")

    print(f"\n  {BOLD}Zmagovalec (bližje VTT): {GREEN if winner=='TRENUTNI' else CYAN}{winner}{RESET}"
          f"  (za {diff:.1f} odstotnih točk)")
    print(f"  {GRAY}Medsebojna podobnost TRENUTNI ↔ NOV: {stats_cross['ratio']*100:.1f}%{RESET}")
    print()


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Primerja transkripta NOV in TRENUTNI z referenčnimi VTT podnapisi.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Primeri:
  python3 compare_transcripts.py --nov nov.docx --trenutni trenutni.docx --vtt podnapisi.vtt
  python3 compare_transcripts.py --nov nov.txt  --trenutni trenutni.txt  --vtt podnapisi.vtt

Podprti formati: .docx, .txt, .vtt
        """
    )
    parser.add_argument("--nov",      required=True, help="Pot do NOV transkripta (.docx ali .txt)")
    parser.add_argument("--trenutni", required=True, help="Pot do TRENUTNI transkripta (.docx ali .txt)")
    parser.add_argument("--vtt",      required=True, help="Pot do referenčnih VTT podnapisov")
    parser.add_argument("--primeri",  type=int, default=8, metavar="N",
                        help="Število primerov razlik na kategorijo (privzeto: 8)")
    parser.add_argument("--brez-barv", action="store_true",
                        help="Izpis brez ANSI barv (za redirect v datoteko)")
    args = parser.parse_args()

    # Preveri obstoj datotek
    for label, path in [("NOV", args.nov), ("TRENUTNI", args.trenutni), ("VTT", args.vtt)]:
        if not Path(path).exists():
            print(f"NAPAKA: Datoteka '{label}' ne obstaja: {path}")
            sys.exit(1)

    # Izklopi barve če zahtevano
    if args.brez_barv:
        global RESET, BOLD, GREEN, RED, YELLOW, CYAN, GRAY
        RESET = BOLD = GREEN = RED = YELLOW = CYAN = GRAY = ""

    print(f"\n{BOLD}Berem datoteke...{RESET}")
    nov_text  = clean_text(read_file(args.nov))
    curr_text = clean_text(read_file(args.trenutni))
    vtt_text  = clean_text(read_file(args.vtt))

    nov_tokens  = tokenize(nov_text)
    curr_tokens = tokenize(curr_text)
    vtt_tokens  = tokenize(vtt_text)

    print(f"  NOV:      {len(nov_tokens):>6} besed/tokenov  ({Path(args.nov).name})")
    print(f"  TRENUTNI: {len(curr_tokens):>6} besed/tokenov  ({Path(args.trenutni).name})")
    print(f"  VTT:      {len(vtt_tokens):>6} besed/tokenov  ({Path(args.vtt).name})")
    print(f"\n{BOLD}Primerjam...{RESET}")

    stats_nov   = categorize_diff(vtt_tokens, nov_tokens)
    stats_curr  = categorize_diff(vtt_tokens, curr_tokens)
    # Medsebojna primerjava brez reference: TRENUTNI → NOV
    stats_cross = categorize_diff(curr_tokens, nov_tokens)

    # ── REZULTATI ──────────────────────────────
    print_header("VTT (referenca) vs NOV")
    print_comparison("VTT", "NOV", stats_nov, len(vtt_tokens))
    print_header("VTT (referenca) vs TRENUTNI")
    print_comparison("VTT", "TRENUTNI", stats_curr, len(vtt_tokens))
    print_header("MEDSEBOJNA PRIMERJAVA: TRENUTNI vs NOV")
    print_comparison("TRENUTNI", "NOV", stats_cross, len(curr_tokens))

    # ── PRIMERI RAZLIK ─────────────────────────
    print_header("PRIMERI RAZLIK  —  VTT vs NOV")
    print_examples(stats_nov, "VTT", "NOV", show_n=args.primeri)

    print_header("PRIMERI RAZLIK  —  VTT vs TRENUTNI")
    print_examples(stats_curr, "VTT", "TRENUTNI", show_n=args.primeri)

    print_header("PRIMERI RAZLIK  —  TRENUTNI vs NOV")
    print_examples(stats_cross, "TRENUTNI", "NOV", show_n=args.primeri)

    # ── SKUPNI REZULTAT ────────────────────────
    print_summary(stats_nov, stats_curr, stats_cross, len(vtt_tokens))

    # ── ZAPIS REZULTATOV V DATOTEKE ───────────
    base_path = Path(args.nov).parent / Path(args.nov).stem

    # JSON datoteka
    json_path = Path(str(base_path) + ".json")

    def json_block(stats: dict, total_ref: int) -> dict:
        return {
            "similarity_percent": round(stats["ratio"] * 100, 2),
            "difference_percent": round(100 - (stats["ratio"] * 100), 2),
            "total_diff_blocks": stats["total_diff_blocks"],
            "total_diff_tokens": stats["total_diff_tokens"],
            # Bloki (odvisni od tega, kako difflib združi razlike — NEprimerljivi med transkripti)
            "word_diff": stats["word_diff"],
            "punct_only": stats["punct_only"],
            "number_format": stats["number_format"],
            "abbreviation": stats["abbreviation"],
            # Besede (primerljive med transkripti)
            "word_diff_tokens": stats["word_diff_tokens"],
            "punct_only_tokens": stats["punct_only_tokens"],
            "number_format_tokens": stats["number_format_tokens"],
            "abbreviation_tokens": stats["abbreviation_tokens"],
            "word_diff_percent_of_reference": round(
                stats["word_diff_tokens"] / max(total_ref, 1) * 100, 2),
            "equal_tokens": stats["equal"],
            "replaced": stats["replaced"],
            "inserted": stats["inserted"],
            "deleted": stats["deleted"],
        }

    results = {
        "files": {
            "nov": Path(args.nov).name,
            "trenutni": Path(args.trenutni).name,
            "vtt_reference": Path(args.vtt).name,
        },
        "token_counts": {
            "nov": len(nov_tokens),
            "trenutni": len(curr_tokens),
            "vtt_reference": len(vtt_tokens),
        },
        "comparison": {
            "nov_vs_vtt":      json_block(stats_nov,  len(vtt_tokens)),
            "trenutni_vs_vtt": json_block(stats_curr, len(vtt_tokens)),
            "nov_vs_trenutni": json_block(stats_cross, len(curr_tokens)),
        },
        "winner": "TRENUTNI" if stats_curr["ratio"] > stats_nov["ratio"] else "NOV",
        "similarity_difference_percent": round(abs((stats_curr["ratio"] - stats_nov["ratio"]) * 100), 2),
    }

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    # Markdown datoteka
    md_path = Path(str(base_path) + ".md")
    md_content = build_markdown_report(stats_nov, stats_curr, stats_cross, len(vtt_tokens), args)
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(md_content)

    print(f"\n{BOLD}✓ Rezultati shranjeni:{RESET}")
    print(f"  • JSON: {json_path}")
    print(f"  • Markdown: {md_path}")


if __name__ == "__main__":
    main()
