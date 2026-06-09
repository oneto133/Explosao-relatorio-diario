from __future__ import annotations

import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path


"""
Essa automação tem o intuito de pegar uma mensagem recebida de um movimentador de estoque
então ao passar ao pcp, é avaliado e então inserido em movimentador_injeção.txt para
padronização e separação no setor de almoxarifado, o intuito da aplicação é excluir a necessidade
de digitação manual e automatizar tarefas simples que demorariam horas para serem escritas com todos
os tratamentos que tem hoje, essa aplicação foi escrita de forma supervisionada usando codex com
pouca injeção, usando de técnicas de engenharia de contexto para poder se te o maior aproveitamento da 
ferramenta disponível com o menor uso de tokens possível.
"""
ROOT = Path(__file__).resolve().parent # pasta que o arquivo está
PAI = ROOT.parent #pasta acima de app
TXT_DIR = PAI / "txt"
HIST_MOV_PATH = TXT_DIR / "movimentador.txt"
HIST_PCP_PATH = TXT_DIR / "pcp.txt"
INPUT_PATH = TXT_DIR / "movimentador_injecao.txt"
OUTPUT_PATH = TXT_DIR / "pcp_saida.txt"

SKIP_PREFIXES = (
    "pra separar no almoxarifado",
    "bom dia",
    "segue abaixo",
    "necessidade de material",
    "lista de materiais",
)

QTY_UNIT_PATTERN = (
    r"caixas?|caixa|unid(?:ades?)?|unid\.?|pcs?|pc|pacotes?|"
    r"kgs?|kg|metros?|mt|rolos?|rolo"
)

STOP_WORDS = {
    "a",
    "as",
    "ao",
    "aos",
    "c",
    "com",
    "da",
    "das",
    "de",
    "do",
    "dos",
    "e",
    "em",
    "na",
    "nas",
    "no",
    "nos",
    "o",
    "os",
    "ou",
    "p",
    "para",
    "pra",
}

FORCE_TEMPLATE_QTY_CODES = {"911080", "100802", "916325", "103151", "103148"}


@dataclass(frozen=True)
class Template:
    raw: str
    base: str
    coded: bool
    search_key: str
    default_qty: str | None = None
    default_unit: str | None = None


def read_text_auto(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def remove_accents(text: str) -> str:
    normalized = unicodedata.normalize("NFD", text)
    return "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")


def normalize_text(text: str) -> str:
    normalized = remove_accents(text).lower()
    normalized = normalized.replace("/", " ")
    normalized = re.sub(r"[^\w\s\.\-]", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized


def split_messages(text: str) -> list[tuple[str, str]]:
    marker = re.compile(r"(?im)^\s*mensagem\s+(\d+)\s*$")
    matches = list(marker.finditer(text))

    if not matches:
        return [("mensagem 1", text)]

    result: list[tuple[str, str]] = []
    for idx, current in enumerate(matches):
        start = current.end()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
        body = text[start:end].strip()
        result.append((f"mensagem {current.group(1)}", body))
    return result


def is_ignored_line(line: str) -> bool:
    normalized = normalize_text(line)
    return any(normalized.startswith(prefix) for prefix in SKIP_PREFIXES)


def extract_item_lines(block: str) -> list[str]:
    items: list[str] = []
    for raw_line in block.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if is_ignored_line(line):
            continue
        items.append(line)
    return items


def split_qty_unit(text: str) -> tuple[str | None, str | None]:
    if not text:
        return None, None
    cleaned = text.strip()
    match = re.match(r"^\s*([\d\.,]+)\s*([A-Za-zÀ-ÿ\.]+)?\s*$", cleaned)
    if not match:
        return None, None
    qty = match.group(1).strip()
    unit = match.group(2).strip() if match.group(2) else None
    return qty, unit


def canonical_unit(unit: str | None) -> str | None:
    if not unit:
        return None
    normalized = normalize_text(unit).replace(".", "")
    if normalized.startswith("caix"):
        return "caixa"
    if normalized.startswith("unid"):
        return "unidades"
    if normalized in {"pc", "pcs"}:
        return "pc"
    if normalized.startswith("pacot"):
        return "pc"
    if normalized.startswith("kg"):
        return "KGS"
    if normalized.startswith("metr") or normalized == "mt":
        return "METROS"
    if normalized.startswith("rol"):
        return "rolo"
    return unit


def extract_quantity(line: str) -> tuple[str | None, str | None]:
    match = re.search(
        rf"\(\s*([\d\.,]+)\s*({QTY_UNIT_PATTERN})\s*\)", line, flags=re.IGNORECASE
    )
    if match:
        return match.group(1), canonical_unit(match.group(2))

    match = re.search(
        rf"\b([\d\.,]+)\s*({QTY_UNIT_PATTERN})\b", line, flags=re.IGNORECASE
    )
    if match:
        return match.group(1), canonical_unit(match.group(2))

    match = re.match(r"^\s*([\d\.,]+)\b", line)
    if match:
        return match.group(1), None

    return None, None


def strip_quantity_from_line(line: str) -> str:
    stripped = re.sub(
        rf"\(\s*[\d\.,]+\s*({QTY_UNIT_PATTERN})\s*\)",
        " ",
        line,
        flags=re.IGNORECASE,
    )
    stripped = re.sub(
        rf"\b[\d\.,]+\s*({QTY_UNIT_PATTERN})\b",
        " ",
        stripped,
        flags=re.IGNORECASE,
    )
    stripped = re.sub(r"^\s*[\d\.,]+\s+", " ", stripped)
    return re.sub(r"\s+", " ", stripped).strip()


def text_key_for_match(text: str, strip_qty: bool = True) -> str:
    base = strip_quantity_from_line(text) if strip_qty else text
    normalized = normalize_text(base)
    tokens = re.findall(r"[a-z0-9]+(?:\.[0-9]+)?", normalized)
    filtered = [tok for tok in tokens if tok not in STOP_WORDS]
    return " ".join(filtered)


def similarity(left: str, right: str) -> float:
    if not left or not right:
        return 0.0

    left_tokens = set(left.split())
    right_tokens = set(right.split())
    inter = len(left_tokens & right_tokens)
    union = len(left_tokens | right_tokens)
    jaccard = inter / union if union else 0.0

    seq = SequenceMatcher(None, left, right).ratio()

    left_nums = set(re.findall(r"\d+(?:\.\d+)?", left))
    right_nums = set(re.findall(r"\d+(?:\.\d+)?", right))
    num_bonus = 0.10 if left_nums and (left_nums & right_nums) else 0.0

    return (0.65 * jaccard) + (0.35 * seq) + num_bonus


def parse_template(line: str) -> Template:
    clean = line.strip()

    if re.match(r"^\s*\d{5,6}\s*-", clean):
        parts = [part.strip() for part in clean.split(" - ")]
        code = parts[0]
        tail = parts[-1] if len(parts) >= 2 else ""
        desc = " - ".join(parts[1:-1]).strip() if len(parts) >= 3 else ""
        qty, unit = split_qty_unit(tail)
        if not desc and len(parts) >= 2:
            desc = parts[1]
        return Template(
            raw=clean,
            base=f"{code} - {desc.strip()}",
            coded=True,
            search_key=text_key_for_match(desc or clean, strip_qty=False),
            default_qty=qty,
            default_unit=unit,
        )

    plain_with_qty = re.match(r"^\s*(.+?)\s*-\s*([\d\.,]+)\s*([A-Za-zÀ-ÿ\.]+)\s*$", clean)
    if plain_with_qty:
        base, qty, unit = plain_with_qty.groups()
        return Template(
            raw=clean,
            base=base.strip(),
            coded=False,
            search_key=text_key_for_match(base, strip_qty=False),
            default_qty=qty.strip(),
            default_unit=unit.strip(),
        )

    return Template(
        raw=clean,
        base=clean,
        coded=False,
        search_key=text_key_for_match(clean, strip_qty=False),
    )


def template_code(template: Template) -> str | None:
    if not template.coded:
        return None
    return template.base.split(" - ", 1)[0].strip()


def find_template_by_code(templates: dict[str, Template], code: str) -> str | None:
    for line, template in templates.items():
        if template_code(template) == code:
            return line
    return None


def find_plain_template(templates: dict[str, Template], text: str) -> str | None:
    needle = normalize_text(text)
    for line, template in templates.items():
        if template.coded:
            continue
        if normalize_text(template.raw) == needle:
            return line
    return None


def rule_based_template_line(line: str, templates: dict[str, Template]) -> str | None:
    norm = normalize_text(line)

    if "cap" in norm and re.search(r"\b12\b", norm):
        return find_template_by_code(templates, "101889")
    if "cap" in norm and re.search(r"\b25\b", norm):
        return find_template_by_code(templates, "101890")
    if "cap" in norm and re.search(r"\b30\b", norm):
        return find_template_by_code(templates, "101309")
    if "cap" in norm and re.search(r"\b40\b", norm):
        return find_template_by_code(templates, "101310")
    if any(token in norm for token in ("ampola", "reed")):
        return find_template_by_code(templates, "100817")
    if "mola" in norm and "amarela" in norm:
        return find_template_by_code(templates, "603519")
    if "mola" in norm and "verde" in norm:
        return find_template_by_code(templates, "603518")
    if "conector" in norm and "controle" in norm:
        return find_template_by_code(templates, "603517")
    if "central flex" in norm or "placa flex" in norm:
        return find_template_by_code(templates, "911080")
    if "fonte" in norm and re.search(r"\b24\b", norm):
        return find_template_by_code(templates, "603516")
    if "fonte" in norm and re.search(r"\b12\b", norm):
        return find_template_by_code(templates, "102815")
    if "central" in norm and "cancela" in norm:
        return find_template_by_code(templates, "603515")
    if "motor" in norm and "2.5" in norm:
        return find_template_by_code(templates, "916127")
    if "motor" in norm and "1.2" in norm:
        return find_template_by_code(templates, "916126")
    if "fio" in norm and ("1.5" in norm or "1 5" in norm or "paralelo" in norm):
        return find_template_by_code(templates, "103795")
    if "leitosa" in norm:
        return find_template_by_code(templates, "102573")
    if "canaleta" in norm:
        return find_template_by_code(templates, "916325")
    if "ima" in norm or "imã" in norm:
        return find_template_by_code(templates, "208003")
    if "borne" in norm:
        return find_template_by_code(templates, "103064")
    if "fim" in norm and "curso" in norm and re.search(r"\b3[.,]5\b", norm):
        return find_template_by_code(templates, "103148")
    if "fim" in norm and "curso" in norm and re.search(r"(?<![\d\.,])5(?![\d\.,])", norm):
        return find_template_by_code(templates, "103151")
    if "controle" in norm:
        return find_template_by_code(templates, "100802")
    if "adesivo" in norm and "vermelho" in norm:
        return find_template_by_code(templates, "811320")
    if "parafuso" in norm and "2.9" in norm:
        return find_template_by_code(templates, "401014")

    if "verniz" in norm:
        return find_plain_template(templates, "VERNIZ")
    if any(token in norm for token in ("thinner", "thiner", "tiner")):
        return find_plain_template(templates, "THINNER")
    if "ribbon" in norm:
        return find_plain_template(templates, "RIBBON")
    if "plastico" in norm and "bolha" in norm:
        return find_plain_template(templates, "PLÁSTICO BOLHA")
    if "prancheta" in norm:
        return find_plain_template(templates, "Prancheta")
    if "cola" in norm and "quente" in norm:
        return find_plain_template(templates, "COLA QUENTE - 3PC")
    if norm == "pu":
        return find_plain_template(templates, "PU")
    
    if norm == "tarugo":
        return find_plain_template(templates, "411039")

    return None


def build_mapping(
    historic_mov_messages: list[tuple[str, str]],
    historic_pcp_messages: list[tuple[str, str]],
) -> tuple[dict[str, str], dict[str, Template], list[str]]:
    votes: dict[str, Counter[str]] = defaultdict(Counter)
    templates: dict[str, Template] = {}
    catalog_keys: list[str] = []

    for _, pcp_body in historic_pcp_messages:
        for line in extract_item_lines(pcp_body):
            if line not in templates:
                templates[line] = parse_template(line)
                catalog_keys.append(line)

    total = min(len(historic_mov_messages), len(historic_pcp_messages))
    for idx in range(total):
        _, mov_body = historic_mov_messages[idx]
        _, pcp_body = historic_pcp_messages[idx]

        mov_items = extract_item_lines(mov_body)
        pcp_items = extract_item_lines(pcp_body)
        remaining = pcp_items[:]

        for mov_line in mov_items:
            mov_key = text_key_for_match(mov_line)
            if not mov_key:
                continue

            best_idx = -1
            best_score = 0.0
            for candidate_idx, pcp_line in enumerate(remaining):
                score = similarity(mov_key, templates[pcp_line].search_key)
                if score > best_score:
                    best_score = score
                    best_idx = candidate_idx

            if best_idx >= 0 and best_score >= 0.30:
                matched = remaining.pop(best_idx)
                votes[mov_key][matched] += best_score

    alias_map: dict[str, str] = {}
    for mov_key, counter in votes.items():
        alias_map[mov_key] = counter.most_common(1)[0][0]

    return alias_map, templates, catalog_keys


def choose_template_line(
    line: str,
    alias_map: dict[str, str],
    templates: dict[str, Template],
    catalog_lines: list[str],
) -> str | None:
    from_rules = rule_based_template_line(line, templates)
    if from_rules:
        return from_rules

    key = text_key_for_match(line)
    if not key:
        return None

    if key in alias_map:
        return alias_map[key]

    best_alias = None
    best_alias_score = 0.0
    for alias_key, template_line in alias_map.items():
        score = similarity(key, alias_key)
        if score > best_alias_score:
            best_alias_score = score
            best_alias = template_line
    if best_alias and best_alias_score >= 0.58:
        return best_alias

    best_catalog = None
    best_catalog_score = 0.0
    for template_line in catalog_lines:
        score = similarity(key, templates[template_line].search_key)
        if score > best_catalog_score:
            best_catalog_score = score
            best_catalog = template_line
    if best_catalog and best_catalog_score >= 0.45:
        return best_catalog

    return None


def format_unit_for_output(qty: str, input_unit: str | None, template: Template) -> str | None:
    if template.default_unit:
        return template.default_unit
    if not input_unit:
        return None

    canonical = canonical_unit(input_unit)
    if canonical == "caixa":
        return "caixa" if qty.replace(",", ".") == "1" else "caixas"
    if canonical == "unidades":
        return "UNIDADE" if qty.replace(",", ".") == "1" else "UNIDADES"
    if canonical == "pc":
        return "pc"
    if canonical in {"KGS", "METROS"}:
        return canonical
    if canonical == "rolo":
        return "rolo" if qty.replace(",", ".") == "1" else "rolos"
    return str(canonical)


def render_template(template: Template, qty: str | None, unit: str | None) -> str:
    if not qty:
        return template.raw

    code = template_code(template)
    if code in FORCE_TEMPLATE_QTY_CODES and template.default_qty:
        return template.raw

    output_unit = format_unit_for_output(qty, unit, template)
    suffix = qty if not output_unit else f"{qty} {output_unit}".strip()
    return f"{template.base} - {suffix}".strip()


def convert_message_body(
    body: str,
    alias_map: dict[str, str],
    templates: dict[str, Template],
    catalog_lines: list[str],
) -> str:
    lines = extract_item_lines(body)

    converted: list[str] = []
    for line in lines:
        qty, unit = extract_quantity(line)
        template_line = choose_template_line(line, alias_map, templates, catalog_lines)
        if template_line and template_line in templates:
            rendered = render_template(templates[template_line], qty, unit)
            converted.append(rendered)
        else:
            converted.append(line.strip().upper())

    coded = [line for line in converted if re.match(r"^\d{5,6}\s*-", line)]
    plain = [line for line in converted if not re.match(r"^\d{5,6}\s*-", line)]

    out: list[str] = ["Bom dia, segue abaixo a necessidade de material:", ""]
    out.extend(coded)
    if plain:
        if coded:
            out.append("")
        out.extend(plain)
    return "\n".join(out).rstrip()


def convert_input_text(
    input_text: str,
    alias_map: dict[str, str],
    templates: dict[str, Template],
    catalog_lines: list[str],
) -> str:
    messages = split_messages(input_text)
    if not input_text.strip():
        return ""

    blocks: list[str] = []
    for idx, (header, body) in enumerate(messages, start=1):
        fixed_header = header.lower().strip() or f"mensagem {idx}"
        blocks.append(fixed_header)
        blocks.append("")
        blocks.append(convert_message_body(body, alias_map, templates, catalog_lines))
        if idx < len(messages):
            blocks.append("")
    return "\n".join(blocks).rstrip() + "\n"


def main() -> None:
    historic_mov = read_text_auto(HIST_MOV_PATH)
    historic_pcp = read_text_auto(HIST_PCP_PATH)
    incoming = read_text_auto(INPUT_PATH) if INPUT_PATH.exists() else ""

    mov_messages = split_messages(historic_mov)
    pcp_messages = split_messages(historic_pcp)
    alias_map, templates, catalog_lines = build_mapping(mov_messages, pcp_messages)

    output = convert_input_text(incoming, alias_map, templates, catalog_lines)
    OUTPUT_PATH.write_text(output, encoding="utf-8")

    if incoming.strip():
        msg_count = len(split_messages(incoming))
        print(f"OK: {msg_count} mensagem(ns) processada(s) em {OUTPUT_PATH}.")
    else:
        print(
            "Entrada vazia em txt/movimentador_injecao.txt. "
            "Saída gerada vazia em txt/pcp_saida.txt."
        )


if __name__ == "__main__":
    main()
