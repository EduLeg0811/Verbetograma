"""Inspecao simples de paragrafos DOCX para apoio a auditoria de confor."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

try:
    from .analisar_verbete import read_docx
except ImportError:
    try:
        from scripts.analisar_verbete import read_docx
    except ImportError:
        from analisar_verbete import read_docx


def analisar_epigrafe_pontuacao(texto: str) -> dict:
    """Heuristica textual para indicar o realce esperado do ponto final."""
    head = texto.split(" ", 1)[0] if texto else ""
    abre_lista = ":" in head or ": " in texto[:120]
    return {
        "texto": texto,
        "abre_enumeracao": abre_lista,
        "ponto_final_esperado": "negrito" if abre_lista else "plano",
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("arquivo")
    parser.add_argument("--pontuacao", action="store_true")
    args = parser.parse_args()
    paragraphs = read_docx(Path(args.arquivo))
    if args.pontuacao:
        data = [analisar_epigrafe_pontuacao(p) for p in paragraphs]
    else:
        data = [{"n": i + 1, "texto": p} for i, p in enumerate(paragraphs)]
    print(json.dumps(data, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
