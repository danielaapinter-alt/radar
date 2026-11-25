"""Empacota o Radar para o PC do Safra.

Uso: python3 scripts/radar_safra.py data/edicoes.json index.html OUT_DIR

Gera OUT_DIR/radar.txt: JSON {"v":1,"template":..., "edicoes":[...]} comprimido (gzip) e em base64,
com a edicao mais recente e a ultima edicao de cada um dos 2 dias anteriores (para a marcacao de
"novos"). O programa proventos.pyw no PC do Safra junta as edicoes num historico local e grava
Radar.html na pasta da rede.

Imprime um JSON com subject, id, path e sha256 do anexo.
"""
import base64
import gzip
import hashlib
import json
import os
import sys

PREFIXO = "Radar Corporativo - edicao"


def template_offline(html):
    antes = html
    html = html.replace(
        '<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wdth,wght@62..125,500..800&family=Public+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap">',
        '<meta http-equiv="refresh" content="300">')
    html = html.replace(
        '<a href="proventos.html" style="display:inline-block;margin-top:6px;font-size:13px;font-weight:600;color:var(--accent)">Página de proventos para e-mail →</a>',
        '<div style="margin-top:6px;font-size:12px;color:var(--muted)">Cópia na rede do Safra · atualiza sozinha a cada 5 min</div>')
    a = 'const r = await fetch("data/edicoes.json?t=" + Date.now(), { cache: "no-store" });'
    b = 'if (!r.ok) throw new Error(r.status);'
    c = 'const eds = await r.json();'
    assert a in html and b in html and c in html, "index.html mudou: ajuste radar_safra.py"
    html = html.replace(a, "").replace(b, "").replace(c, "const eds = window.RADAR_EDS || [];")
    html = html.replace("setInterval(load, 5 * 60 * 1000);", "")
    assert html != antes
    return html


MARCA_INI = "-----RADAR-INICIO-----"
MARCA_FIM = "-----RADAR-FIM-----"


def corpo(arq_ed, out):
    """Modo --corpo: a edicao mais recente (sem a lista de fontes) em JSON, para ir no CORPO do e-mail.
    As execucoes agendadas nao conseguem anexar arquivos no Gmail."""
    eds = json.load(open(arq_ed, encoding="utf-8"))
    eds.sort(key=lambda e: e["ts"], reverse=True)
    ult = dict(eds[0])
    ult.pop("fontes", None)
    c = lambda v: json.dumps(v, ensure_ascii=False, separators=(",", ":"))
    linhas = []
    for k, v in ult.items():
        if isinstance(v, list):
            linhas.append(c(k) + ":[\n" + ",\n".join(c(i) for i in v) + "\n]")
        else:
            linhas.append(c(k) + ":" + c(v))
    js = "{\n" + ",\n".join(linhas) + "\n}"
    assert json.loads(js) == ult
    texto = ("Edicao automatica do Radar Corporativo para a rede do Safra. "
             "O programa no PC do banco le este e-mail e o apaga.\n\n"
             f"{MARCA_INI}\n{js}\n{MARCA_FIM}\n")
    os.makedirs(out, exist_ok=True)
    caminho = os.path.abspath(os.path.join(out, "radar_corpo.txt"))
    with open(caminho, "w", encoding="utf-8", newline="\n") as f:
        f.write(texto)
    d, hhmm = ult["dia"], ult["ts"][11:16]
    print(json.dumps({"subject": f"{PREFIXO} {d[8:10]}/{d[5:7]}/{d[0:4]} {hhmm}", "id": ult["id"],
                      "path": caminho, "bytes": os.path.getsize(caminho)}, ensure_ascii=False))


def main():
    if "--corpo" in sys.argv:
        args = [a for a in sys.argv[1:] if a != "--corpo"]
        return corpo(args[0], args[-1])
    arq_ed, arq_html, out = sys.argv[1:4]
    eds = json.load(open(arq_ed, encoding="utf-8"))
    eds.sort(key=lambda e: e["ts"], reverse=True)
    ult = eds[0]
    pacote = [ult]
    dias = []
    for e in eds:
        if e["dia"] < ult["dia"] and e["dia"] not in dias:
            dias.append(e["dia"])
            pacote.append(e)
            if len(dias) == 2:
                break
    dados = {"v": 1, "template": template_offline(open(arq_html, encoding="utf-8").read()), "edicoes": pacote}
    raw = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    b64 = base64.b64encode(gzip.compress(raw, 9)).decode("ascii")
    linhas = "\r\n".join(b64[i:i + 76] for i in range(0, len(b64), 76)) + "\r\n"
    os.makedirs(out, exist_ok=True)
    caminho = os.path.abspath(os.path.join(out, "radar.txt"))
    with open(caminho, "w", encoding="ascii", newline="") as f:
        f.write(linhas)
    sha = hashlib.sha256(open(caminho, "rb").read()).hexdigest()
    hhmm = ult["ts"][11:16]
    d = ult["dia"]
    subject = f"{PREFIXO} {d[8:10]}/{d[5:7]}/{d[0:4]} {hhmm}"
    print(json.dumps({"subject": subject, "id": ult["id"], "path": caminho, "sha256": sha,
                      "bytes": os.path.getsize(caminho), "edicoes": len(pacote)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
