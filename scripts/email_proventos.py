#!/usr/bin/env python3
"""Gera o e-mail diário de proventos do Radar Corporativo.

Lê data/edicoes.json (mais recente primeiro) e escreve:
  out/email_proventos.html   corpo HTML (layout em tabelas, estilos inline: funciona no Outlook)
  out/email_proventos.txt    versão texto simples
  out/email_proventos.json   {"subject": ..., ...}

Uso: python3 scripts/email_proventos.py [caminho/do/edicoes.json] [pasta/de/saida]
"""
import html
import json
import os
import re
import sys
from datetime import date, datetime, timedelta

SRC = sys.argv[1] if len(sys.argv) > 1 else "data/edicoes.json"
OUT = sys.argv[2] if len(sys.argv) > 2 else "out"

MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
DIAS = ["segunda", "terça", "quarta", "quinta", "sexta", "sábado", "domingo"]

# Paleta (mesma do site)
INK, MUTED, LINE, BG, CARD = "#14201D", "#5B6B66", "#D5DDDA", "#EEF1F0", "#FFFFFF"
ZEBRA = "#F6F8F7"
NEW, NEW_BG = "#A35F00", "#FBEBD3"
TIPO_COR = {"jcp": ("#1F4FA8", "#E3EAF7"), "div": ("#0C7A4E", "#E0F1E8"), "esp": ("#A35F00", "#FBEBD3"), "rend": ("#5B3FA8", "#ECE6F7")}
FONT = "Segoe UI,Arial,Helvetica,sans-serif"
MONO = "Consolas,Courier New,monospace"


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


def parse_date(x, ref):
    """Primeira data dd/mm[/aa] do texto -> date (ano inferido pela data da edição)."""
    m = re.search(r"(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?", x or "")
    if not m:
        return None
    d, mo = int(m[1]), int(m[2])
    if m[3]:
        y = int(m[3]); y = 2000 + y if y < 100 else y
    else:
        y = ref.year + 1 if mo < ref.month - 2 else ref.year
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def short(s, n=140):
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    return s if len(s) <= n else s[: n - 1].rstrip() + "…"


def main():
    eds = json.load(open(SRC, encoding="utf-8"))
    eds.sort(key=lambda e: e["ts"], reverse=True)
    ed = eds[0]
    ts = datetime.fromisoformat(ed["ts"])
    ref = ts.date()
    prev = next((e for e in eds if e["dia"] < ed["dia"]), None)
    prev_keys = {p_key(p) for p in prev.get("proventos", [])} if prev else None

    # e-mails de 12h/18h: o que entrou desde a edicao usada no envio da manha (ate 8h30)
    base_manha = None
    if ts.hour >= 11:
        base_manha = next((e for e in eds if e["dia"] == ed["dia"] and e["ts"][11:16] <= "08:30"), None)
    keys_manha = {p_key(p) for p in base_manha.get("proventos", [])} if base_manha else None

    def is_new(p):
        return (prev_keys is not None and p_key(p) not in prev_keys) or \
               (keys_manha is not None and p_key(p) not in keys_manha)

    def pay(p):
        return parse_date(p.get("pagamento"), ref)

    def pays_today(p):
        return pay(p) == ref or "hoje" in str(p.get("status") or "").lower()

    provs = ed.get("proventos") or []
    br = [p for p in provs if not is_us(p)]
    us = [p for p in provs if is_us(p)]
    novos = [p for p in provs if prev_keys is not None and p_key(p) not in prev_keys]
    novos_manha = [p for p in provs if keys_manha is not None and p_key(p) not in keys_manha]
    hoje = [p for p in provs if pays_today(p)]
    fim_semana = ref + timedelta(days=6 - ref.weekday())
    semana = [p for p in provs if pay(p) and ref <= pay(p) <= fim_semana]

    def sort_key(p):
        d = pay(p)
        return (d is None, d or date.max, parse_date(p.get("dataCom"), ref) or date.max, str(p.get("ticker")))

    # ---------- blocos de HTML (tabelas + estilos inline, para o Outlook) ----------
    def pill(p):
        fg, bg = TIPO_COR[p_class(p.get("tipo"))]
        return (f'<span style="background:{bg};color:{fg};font-family:{FONT};font-size:11px;font-weight:bold;'
                f'padding:2px 7px;border-radius:10px;white-space:nowrap">{esc(p.get("tipo") or "—")}</span>')

    def section_title(txt, note=""):
        n = f'<span style="font-size:12px;font-weight:normal;color:{MUTED}">&nbsp; {note}</span>' if note else ""
        return (f'<tr><td style="padding:24px 2px 8px;font-family:{FONT};font-size:15px;font-weight:bold;'
                f'color:{INK};letter-spacing:.5px">{txt}{n}</td></tr>')

    def kpi(n, label, color=INK):
        return (f'<td width="25%" align="center" bgcolor="{CARD}" style="padding:14px 6px;border:1px solid {LINE}">'
                f'<div style="font-family:{FONT};font-size:26px;font-weight:bold;color:{color};line-height:1.15">{n}</div>'
                f'<div style="font-family:{FONT};font-size:11px;color:{MUTED};text-transform:uppercase;letter-spacing:.5px;padding-top:2px">{label}</div></td>')

    def card(p):
        mk = "EUA" if is_us(p) else "BR"
        datas = " &nbsp;·&nbsp; ".join(
            f'{lbl} <b style="font-family:{MONO}">{esc(v)}</b>'
            for lbl, v in (("Data com", p.get("dataCom")), ("Ex", p.get("dataEx")), ("Pagamento", p.get("pagamento")))
            if v and v != "—")
        obs = short(p.get("obs"))
        obs_html = f'<div style="padding-top:4px;font-size:12px;color:{MUTED}">{esc(obs)}</div>' if obs else ""
        return (
            f'<tr><td style="padding:0 0 8px"><table width="100%" cellpadding="0" cellspacing="0" bgcolor="{CARD}" style="border:1px solid {LINE}"><tr>'
            f'<td width="5" bgcolor="{NEW}" style="font-size:1px">&nbsp;</td>'
            f'<td style="padding:10px 14px;font-family:{FONT};font-size:13px;color:{INK}">'
            f'<table width="100%" cellpadding="0" cellspacing="0"><tr>'
            f'<td style="font-family:{FONT};color:{INK}"><span style="font-family:{MONO};font-size:15px;font-weight:bold">{esc(p.get("ticker"))}</span>'
            f'&nbsp; <span style="font-size:14px">{esc(p.get("empresa"))}</span>'
            f'&nbsp; <span style="font-size:10px;color:{MUTED};font-weight:bold">{mk}</span></td>'
            f'<td align="right" style="font-family:{MONO};font-size:15px;font-weight:bold;color:{INK};white-space:nowrap">{esc(p.get("valor") or "—")}</td></tr></table>'
            f'<div style="padding-top:6px">{pill(p)}&nbsp; <span style="font-size:12px;color:{MUTED}">{datas or "Datas ainda não divulgadas"}</span></div>'
            f'{obs_html}</td></tr></table></td></tr>')

    def mini_list(arr):
        items = "".join(
            f'<tr bgcolor="{ZEBRA if i % 2 else CARD}"><td style="font-family:{MONO};font-weight:bold;font-size:13px;color:{INK}">{esc(p.get("ticker"))}</td>'
            f'<td style="font-family:{FONT};font-size:13px;color:{INK}">{esc(p.get("empresa"))}</td><td>{pill(p)}</td>'
            f'<td align="right" style="font-family:{MONO};font-size:13px;color:{INK};white-space:nowrap">{esc(p.get("valor") or "—")}</td></tr>'
            for i, p in enumerate(sorted(arr, key=sort_key)))
        return f'<tr><td><table width="100%" cellpadding="7" cellspacing="0" style="border:1px solid {LINE}">{items}</table></td></tr>'

    HEAD = ("Ticker", "Empresa", "Tipo", "Valor/ação", "Data com", "Data ex", "Pagamento")

    TIPO_CURTO = {"jcp": "JCP", "div": "Div.", "esp": "Div. esp.", "rend": "Rend."}

    def tipo_txt(p):
        k = p_class(p.get("tipo"))
        return f'<font color="{TIPO_COR[k][0]}"><b>{TIPO_CURTO[k]}</b></font>'

    def market_table(arr):
        if not arr:
            return f'<tr><td style="font-family:{FONT};font-size:13px;color:{MUTED}">Nenhum provento vigente.</td></tr>'
        groups = [("Pagam nos próximos 7 dias", []), ("Pagam de 8 a 30 dias", []), ("Pagam em mais de 30 dias", []), ("Data de pagamento a definir", [])]
        for p in sorted(arr, key=sort_key):
            d = pay(p)
            if d is None:
                groups[3][1].append(p)
            elif (d - ref).days <= 7:
                groups[0][1].append(p)
            elif (d - ref).days <= 30:
                groups[1][1].append(p)
            else:
                groups[2][1].append(p)
        th = "".join(
            f'<td style="font-family:{FONT};font-size:10px;font-weight:bold;color:#FFFFFF;letter-spacing:.5px">{h.upper()}</td>'
            for h in HEAD)
        out = [f'<tr><td><table width="100%" cellpadding="6" cellspacing="0" style="border:1px solid {LINE};font-family:{FONT};font-size:12px;color:{INK}">'
               f'<tr bgcolor="{INK}">{th}</tr>']
        for name, rows in groups:
            if not rows:
                continue
            out.append(f'<tr bgcolor="{BG}"><td colspan="7"><b>{name}</b> <font color="{MUTED}">({len(rows)})</font></td></tr>')
            for i, p in enumerate(rows):
                nv, hj = is_new(p), pays_today(p)
                bg = NEW_BG if nv else (ZEBRA if i % 2 else CARD)
                tag = f' <font color="{NEW}" size="1"><b>NOVO</b></font>' if nv else ""
                pg = esc(p.get("pagamento") or "—") + (f' <font color="{NEW}" size="1"><b>HOJE</b></font>' if hj else "")
                out.append(
                    f'<tr bgcolor="{bg}"><td><b>{esc(p.get("ticker"))}</b>{tag}</td><td>{esc(p.get("empresa"))}</td>'
                    f'<td>{tipo_txt(p)}</td><td>{esc(p.get("valor") or "—")}</td><td>{esc(p.get("dataCom") or "—")}</td>'
                    f'<td>{esc(p.get("dataEx") or "—")}</td><td>{pg}</td></tr>')
        out.append("</table></td></tr>")
        return "".join(out)

    dia_txt = f"{DIAS[ref.weekday()]}, {ref.day} de {MESES[ref.month - 1]} de {ref.year}"
    n_novos = len(novos)
    subject = f"Proventos {ref.day:02d}/{ref.month:02d}/{ref.year}"
    cmp_txt = (f"Novo = não estava na última edição de {datetime.fromisoformat(prev['ts']).strftime('%d/%m')}."
               if prev else "Sem edição do dia anterior para comparar os novos.")
    vig = (f'{len(br)}<span style="font-size:13px;color:{MUTED}"> BR</span>'
           f' <span style="color:{LINE}">|</span> {len(us)}<span style="font-size:13px;color:{MUTED}"> EUA</span>')

    body = [
        f'<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width"></head>'
        f'<body style="margin:0;padding:0;background:{BG}">'
        f'<table width="100%" cellpadding="0" cellspacing="0" bgcolor="{BG}"><tr><td align="center" style="padding:16px 8px">'
        f'<table width="760" cellpadding="0" cellspacing="0" style="max-width:760px">',
        f'<tr><td bgcolor="{INK}" style="padding:20px 22px;font-family:{FONT}">'
        f'<div style="font-size:11px;color:#8FD3B6;letter-spacing:2px;font-weight:bold">RADAR CORPORATIVO</div>'
        f'<div style="font-size:24px;color:#FFFFFF;font-weight:bold;padding-top:4px">Proventos vigentes</div>'
        f'<div style="font-size:13px;color:#C9D3D0;padding-top:4px">{dia_txt} &nbsp;·&nbsp; edição das {ts.strftime("%H:%M")}</div>'
        f'</td></tr>',
        f'<tr><td style="padding-top:10px"><table width="100%" cellpadding="0" cellspacing="6"><tr>'
        + kpi(n_novos, "Novos hoje", NEW if n_novos else INK)
        + kpi(len(hoje), "Pagam hoje")
        + kpi(len(semana), "Pagam esta semana")
        + kpi(vig, "Vigentes")
        + "</tr></table></td></tr>",
    ]
    if novos_manha:
        body.append(section_title("ATUALIZAÇÃO · NOVOS DESDE O ENVIO DA MANHÃ",
                                  f"{len(novos_manha)} anúncio(s) depois da edição das {base_manha['ts'][11:16]}"))
        body.extend(card(p) for p in sorted(novos_manha, key=lambda p: (is_us(p), sort_key(p))))
    if novos:
        body.append(section_title("NOVOS HOJE", f"{n_novos} anúncio(s) desde a edição de ontem"))
        body.extend(card(p) for p in sorted(novos, key=lambda p: (is_us(p), sort_key(p))))
    if hoje:
        body.append(section_title("PAGAM HOJE", f"{len(hoje)} provento(s)"))
        body.append(mini_list(hoje))
    br_jcp = [p for p in br if p_class(p.get("tipo")) == "jcp"]
    br_div = [p for p in br if p_class(p.get("tipo")) != "jcp"]
    body.append(section_title("BRASIL · JCP · R$", f"{len(br_jcp)} vigentes, por data de pagamento"))
    body.append(market_table(br_jcp))
    body.append(section_title("BRASIL · DIVIDENDOS E DEMAIS · R$", f"{len(br_div)} vigentes, por data de pagamento"))
    body.append(market_table(br_div))
    body.append(section_title("EUA · US$", f"{len(us)} vigentes, por data de pagamento"))
    body.append(market_table(us))
    body.append(
        f'<tr><td style="padding:18px 2px 4px;font-family:{FONT};font-size:11px;color:{MUTED};line-height:1.5">'
        f'{cmp_txt} Linhas em laranja = novos. Painel completo: Radar.html na pasta da rede da mesa.<br>'
        f'Dados compilados de fontes públicas; confirme valores e datas no comunicado oficial da companhia antes de operar.</td></tr>'
        f'</table></td></tr></table>'
        # marcador invisivel lido pelo programa do Safra para saber se ha proventos novos desde o ultimo envio
        f'<div style="display:none;max-height:0;overflow:hidden;mso-hide:all;font-size:0">'
        f'PROVKEYS[{esc(";".join(sorted({p_key(p) for p in provs})))}]</div>'
        f'</body></html>')
    body_html = "".join(body)

    def txt_rows(arr):
        return "\n".join(
            f"{'[NOVO] ' if is_new(p) else ''}{p.get('ticker')} | {p.get('empresa')} | {p.get('tipo')} | "
            f"{p.get('valor') or '—'} | com {p.get('dataCom') or '—'} | pag {p.get('pagamento') or '—'}"
            for p in sorted(arr, key=sort_key)) or "Nenhum."

    body_txt = (f"RADAR CORPORATIVO - Proventos vigentes - {dia_txt}\n"
                f"Novos hoje: {n_novos} | Pagam hoje: {len(hoje)} | Vigentes: {len(br)} BR, {len(us)} EUA\n\n"
                + (f"NOVOS DESDE O ENVIO DA MANHA\n{txt_rows(novos_manha)}\n\n" if novos_manha else "")
                + (f"NOVOS HOJE\n{txt_rows(novos)}\n\n" if novos else "")
                + f"BRASIL - JCP (R$)\n{txt_rows([p for p in br if p_class(p.get('tipo')) == 'jcp'])}\n\n"
                + f"BRASIL - DIVIDENDOS E DEMAIS (R$)\n{txt_rows([p for p in br if p_class(p.get('tipo')) != 'jcp'])}\n\n"
                + f"EUA (US$)\n{txt_rows(us)}\n\n{cmp_txt}\n")

    os.makedirs(OUT, exist_ok=True)
    open(os.path.join(OUT, "email_proventos.html"), "w", encoding="utf-8").write(body_html)
    open(os.path.join(OUT, "email_proventos.txt"), "w", encoding="utf-8").write(body_txt)
    meta = {"subject": subject, "edicao": ed["id"], "n_br": len(br), "n_us": len(us), "novos": n_novos, "novos_desde_manha": len(novos_manha),
            "pagam_hoje": len(hoje), "html_bytes": len(body_html.encode("utf-8"))}
    open(os.path.join(OUT, "email_proventos.json"), "w", encoding="utf-8").write(json.dumps(meta, ensure_ascii=False))
    print(json.dumps(meta, ensure_ascii=False))


if __name__ == "__main__":
    main()
