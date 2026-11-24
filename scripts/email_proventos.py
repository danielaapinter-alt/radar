#!/usr/bin/env python3
"""Gera o e-mail diário de proventos do Radar Corporativo.

Lê data/edicoes.json (mais recente primeiro) e escreve:
  out/email_proventos.html   corpo HTML (tabelas Brasil e EUA, novos no topo)
  out/email_proventos.txt    versão texto simples
  out/email_proventos.json   {"subject": ..., "n_br": ..., "n_us": ..., "novos": ...}

Uso: python3 scripts/email_proventos.py [caminho/do/edicoes.json] [pasta/de/saida]
"""
import html
import json
import os
import re
import sys
from datetime import datetime

SRC = sys.argv[1] if len(sys.argv) > 1 else "data/edicoes.json"
OUT = sys.argv[2] if len(sys.argv) > 2 else "out"

MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
DIAS = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]


def esc(x):
    return html.escape(str(x if x is not None else ""))


def p_class(t):
    x = (t or "").lower()
    if "jcp" in x:
        return "jcp"
    if "espec" in x or "extra" in x:
        return "esp"
    if "rend" in x:
        return "rend"
    return "div"


def ddmm(x):
    m = re.search(r"(\d{1,2})/(\d{1,2})", x or "")
    return f"{int(m[1]):02d}/{int(m[2]):02d}" if m else ""


def p_key(p):
    tk = re.split(r"[\s/,]+", str(p.get("ticker") or "").upper().strip())[0]
    return f"{tk}|{p_class(p.get('tipo'))}|{ddmm(p.get('dataCom')) or ddmm(p.get('pagamento'))}"


def is_us(p):
    return p.get("mercado") == "US" or str(p.get("valor") or "").strip().startswith("US$")


def date_key(x, ref):
    m = re.search(r"(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?", x or "")
    if not m:
        return 99999999
    d, mo = int(m[1]), int(m[2])
    if m[3]:
        y = int(m[3]); y = 2000 + y if y < 100 else y
    else:
        y = ref.year + 1 if mo < ref.month - 2 else ref.year
    return y * 10000 + mo * 100 + d


def short(s, n=110):
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    return s if len(s) <= n else s[: n - 1].rstrip() + "…"


def main():
    eds = json.load(open(SRC, encoding="utf-8"))
    eds.sort(key=lambda e: e["ts"], reverse=True)
    ed = eds[0]
    ref = datetime.fromisoformat(ed["ts"])
    prev = next((e for e in eds if e["dia"] < ed["dia"]), None)
    prev_keys = {p_key(p) for p in (prev or {}).get("proventos", [])} if prev else None

    def is_new(p):
        return prev_keys is not None and p_key(p) not in prev_keys

    provs = ed.get("proventos") or []
    br = [p for p in provs if not is_us(p)]
    us = [p for p in provs if is_us(p)]

    def ordered(arr):
        return sorted(arr, key=lambda p: (not is_new(p), date_key(p.get("pagamento"), ref), date_key(p.get("dataCom"), ref)))

    def table(arr):
        if not arr:
            return '<p class="muted">Nenhum provento vigente nesta edição.</p>'
        rows = []
        for p in ordered(arr):
            nv = is_new(p)
            obs = short(p.get("obs"), 120) if nv else ""  # observação só nos novos, para o e-mail ficar leve
            hoje = "hoje" in str(p.get("status") or "").lower()
            rows.append(
                f'<tr{" class=n" if nv else ""}>'
                f'<td class=t>{esc(p.get("ticker"))}{"<br><span class=nb>NOVO</span>" if nv else ""}</td>'
                f'<td><b>{esc(p.get("empresa"))}</b>{f"<br><span class=o>{esc(obs)}</span>" if obs else ""}</td>'
                f'<td>{esc(p.get("tipo"))}</td>'
                f'<td class=m>{esc(p.get("valor") or "—")}</td>'
                f'<td class=m>{esc(p.get("dataCom") or "—")}</td>'
                f'<td class=m>{esc(p.get("dataEx") or "—")}</td>'
                f'<td class=m>{esc(p.get("pagamento") or "—")}{"<br><span class=h>paga hoje</span>" if hoje else ""}</td>'
                "</tr>"
            )
        return (
            '<table cellpadding="0" cellspacing="0"><tr>'
            "<th>Ticker</th><th>Empresa</th><th>Tipo</th><th>Valor/ação</th><th>Data com</th><th>Data ex</th><th>Pagamento</th></tr>"
            + "".join(rows)
            + "</table>"
        )

    n_br, n_us = sum(map(is_new, br)), sum(map(is_new, us))
    dia_txt = f"{DIAS[ref.weekday()]}, {ref.day:02d}/{MESES[ref.month - 1]}"
    hora_txt = ref.strftime("%H:%M")
    novos_txt = n_br + n_us
    subject = f"Radar · Proventos vigentes {ref.day:02d}/{ref.month:02d}" + (
        f" · {novos_txt} novo{'s' if novos_txt != 1 else ''}" if novos_txt else ""
    )
    cmp_txt = (
        f"Novos = não estavam na última edição de {datetime.fromisoformat(prev['ts']).strftime('%d/%m')}."
        if prev
        else "Sem edição do dia anterior para comparar os novos."
    )

    style = (
        "body{margin:0;padding:16px;background:#EEF1F0;font-family:Segoe UI,Arial,sans-serif;color:#14201D;font-size:13px}"
        "h1{font-size:20px;margin:0 0 2px;text-transform:uppercase;letter-spacing:.02em}"
        "h2{font-size:15px;margin:22px 0 6px;text-transform:uppercase}"
        ".muted{color:#5B6B66;font-size:12px;margin:0}"
        "table{border-collapse:collapse;width:100%;background:#fff;border:1px solid #D5DDDA}"
        "th{background:#14201D;color:#fff;text-align:left;font-size:11px;font-weight:600;padding:7px 8px;text-transform:uppercase}"
        "td{padding:7px 8px;border-top:1px solid #D5DDDA;vertical-align:top}"
        ".t{font-family:Consolas,monospace;font-weight:700;white-space:nowrap}"
        ".m{font-family:Consolas,monospace;white-space:nowrap}"
        ".o{color:#5B6B66;font-size:11px}"
        "tr.n td{background:#FBEBD3}"
        ".nb{display:inline-block;background:#A35F00;color:#fff;font-size:10px;font-weight:700;padding:1px 5px;border-radius:3px}"
        ".h{color:#A35F00;font-size:11px;font-weight:700}"
    )
    body_html = (
        f'<!DOCTYPE html><html><head><meta charset="utf-8"><style>{style}</style></head><body>'
        f"<h1>Radar Corporativo · Proventos vigentes</h1>"
        f'<p class="muted">Edição de {dia_txt}, {hora_txt} (Brasília) · {len(br)} no Brasil, {len(us)} nos EUA'
        f"{f' · {novos_txt} novo(s) hoje' if novos_txt else ''}. Novos no topo, em laranja; depois por data de pagamento.</p>"
        f'<h2>Brasil (R$) · {len(br)}{f" · {n_br} novo(s)" if n_br else ""}</h2>{table(br)}'
        f'<h2>EUA (US$) · {len(us)}{f" · {n_us} novo(s)" if n_us else ""}</h2>{table(us)}'
        f'<p class="muted" style="margin-top:14px">{cmp_txt} Fonte: radar-corp.pages.dev. '
        "Dados de fontes públicas; confirme valores e datas no comunicado oficial da companhia.</p>"
        "</body></html>"
    )

    def txt_rows(arr):
        out = []
        for p in ordered(arr):
            out.append(
                f"{'[NOVO] ' if is_new(p) else ''}{p.get('ticker')} - {p.get('empresa')} - {p.get('tipo')} - "
                f"{p.get('valor') or '—'} - data com {p.get('dataCom') or '—'} - pagamento {p.get('pagamento') or '—'}"
            )
        return "\n".join(out) or "Nenhum."

    body_txt = (
        f"Radar Corporativo - Proventos vigentes - {dia_txt}, {hora_txt}\n\n"
        f"BRASIL (R$)\n{txt_rows(br)}\n\nEUA (US$)\n{txt_rows(us)}\n\n{cmp_txt}\n"
    )

    os.makedirs(OUT, exist_ok=True)
    open(os.path.join(OUT, "email_proventos.html"), "w", encoding="utf-8").write(body_html)
    open(os.path.join(OUT, "email_proventos.txt"), "w", encoding="utf-8").write(body_txt)
    meta = {"subject": subject, "edicao": ed["id"], "n_br": len(br), "n_us": len(us), "novos": novos_txt,
            "html_bytes": len(body_html.encode("utf-8"))}
    open(os.path.join(OUT, "email_proventos.json"), "w", encoding="utf-8").write(json.dumps(meta, ensure_ascii=False))
    print(json.dumps(meta, ensure_ascii=False))


if __name__ == "__main__":
    main()
