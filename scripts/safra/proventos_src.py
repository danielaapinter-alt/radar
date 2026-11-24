# -*- coding: utf-8 -*-
# Proventos e Radar - Outlook do Safra (Radar Corporativo)  - versao 4
# Le a caixa compartilhada mesa.aluguel (e a sua) e pode enviar em nome dela.
# --auto : envia os proventos do dia aos clientes (Agendador, 8h45)
# --rede : le os e-mails "Radar Corporativo - edicao" e regrava Radar.html nesta pasta (a cada 15 min)
#
# 1. Procura no Outlook o e-mail mais recente "Proventos DD/MM/AAAA" (enviado pelo Radar).
# 2. Listas Para / Copia / Copia oculta, salvas em proventos_destinatarios.json (mesma pasta).
# 3. "Enviar agora" manda do seu Outlook; "Abrir no Outlook" abre para conferir.
# 4. "Agendar envio 8h45" cria uma tarefa no Agendador do Windows que roda este programa
#    com --auto de segunda a sexta as 8h45: ele envia sozinho o e-mail de proventos DO DIA
#    para a lista cadastrada (uma vez por dia) e anota tudo em proventos_log.txt.
#
# Requisito (uma vez):  pip install pywin32
import datetime as dt
import json
import os
import re
import subprocess
import sys

try:
    import win32com.client  # pywin32
except ImportError:
    win32com = None

NOME_TAREFA = "Radar Proventos 8h45"
HORA_TAREFA = "08:45"


def caminho_universal(p):
    """Troca letra de unidade de rede (ex.: Y:) pelo caminho \\\\servidor\\... ."""
    try:
        import win32wnet
        return win32wnet.WNetGetUniversalName(p, 1)
    except Exception:
        return p


ARQ_PROG = caminho_universal(os.path.abspath(__file__))
PASTA = os.path.dirname(ARQ_PROG)
ARQ_DEST = os.path.join(PASTA, "proventos_destinatarios.json")
ARQ_ESTADO = os.path.join(PASTA, "proventos_estado.json")
ARQ_LOG = os.path.join(PASTA, "proventos_log.txt")
PADRAO = {
    "to": ["mesa.aluguel@safra.com.br"],
    "cc": ["daniela.pinter@safra.com.br", "jefferson.dias@safra.com.br", "william.parada@safra.com.br"],
    "bcc": [],
}
CHAVES = ("to", "cc", "bcc")
ARQ_CONFIG = os.path.join(PASTA, "proventos_config.json")
CONFIG_PADRAO = {"caixa": "mesa.aluguel@safra.com.br", "enviar_em_nome": True}
ASSUNTO = re.compile(r"^\s*(RES:|ENC:|FW:|RE:)?\s*Proventos (\d{2})/(\d{2})/(\d{4})\s*$", re.I)
EMAIL_OK = re.compile(r"^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$")


# ---------------- arquivos ----------------
def ler_json(caminho, padrao):
    try:
        with open(caminho, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return padrao


def gravar_json(caminho, dados):
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


def carregar_dest():
    d = ler_json(ARQ_DEST, PADRAO)
    return {k: list(d.get(k, PADRAO[k] if d is PADRAO else [])) for k in CHAVES}


def log(msg):
    linha = f"{dt.datetime.now():%d/%m/%Y %H:%M:%S}  {msg}\n"
    try:
        with open(ARQ_LOG, "a", encoding="utf-8") as f:
            f.write(linha)
    except Exception:
        pass


def marcar_enviado(origem, modo):
    gravar_json(ARQ_ESTADO, {
        "ultimo_envio": dt.date.today().isoformat(),
        "assunto": origem.Subject,
        "modo": modo,
        "quando": f"{dt.datetime.now():%d/%m/%Y %H:%M}",
    })


def ja_enviado_hoje():
    return ler_json(ARQ_ESTADO, {}).get("ultimo_envio") == dt.date.today().isoformat()


# ---------------- Outlook ----------------
def outlook():
    if win32com is None:
        raise RuntimeError("Falta o pacote pywin32. No Prompt de Comando rode:  pip install pywin32")
    return win32com.client.Dispatch("Outlook.Application")


def carregar_config():
    c = ler_json(ARQ_CONFIG, {})
    return {"caixa": c.get("caixa", CONFIG_PADRAO["caixa"]),
            "enviar_em_nome": bool(c.get("enviar_em_nome", CONFIG_PADRAO["enviar_em_nome"]))}


def pastas_caixa_mesa(ns, caixa):
    """Caixa de Entrada e Lixo Eletronico da caixa compartilhada (ex.: mesa.aluguel)."""
    if not caixa:
        return []
    # 1) caixa adicionada como conta/arquivo de dados no Outlook
    alvo = caixa.lower()
    for i in range(1, ns.Folders.Count + 1):
        try:
            raiz = ns.Folders.Item(i)
            if alvo in (raiz.Name or "").lower() or alvo.split("@")[0] in (raiz.Name or "").lower():
                st = raiz.Store
                return [st.GetDefaultFolder(6)] + [p for p in [_tenta(lambda: st.GetDefaultFolder(23))] if p]
        except Exception:
            continue
    # 2) caixa compartilhada aberta por permissao
    try:
        rec = ns.CreateRecipient(caixa)
        rec.Resolve()
        pastas = [ns.GetSharedDefaultFolder(rec, 6)]
        j = _tenta(lambda: ns.GetSharedDefaultFolder(rec, 23))
        return pastas + ([j] if j else [])
    except Exception:
        return []


def _tenta(f):
    try:
        return f()
    except Exception:
        return None


def pastas_busca(app):
    """Pastas onde procurar: caixa da mesa (se configurada) e a sua propria caixa."""
    ns = app.GetNamespace("MAPI")
    pastas = pastas_caixa_mesa(ns, carregar_config()["caixa"])
    for pid in (6, 23):  # 6 = Caixa de Entrada, 23 = Lixo Eletronico
        p = _tenta(lambda: ns.GetDefaultFolder(pid))
        if p:
            pastas.append(p)
    return pastas


def ultimo_email_proventos(app):
    melhor = None
    for pasta in pastas_busca(app):
        try:
            itens = pasta.Items
            itens.Sort("[ReceivedTime]", True)
        except Exception:
            continue
        n = 0
        for it in itens:
            n += 1
            if n > 400:
                break
            try:
                if it.Class != 43:
                    continue
                if ASSUNTO.match(it.Subject or ""):
                    if melhor is None or it.ReceivedTime > melhor.ReceivedTime:
                        melhor = it
                    break
            except Exception:
                continue
    return melhor


def eh_de_hoje(origem):
    m = ASSUNTO.match(origem.Subject or "")
    if not m:
        return False
    d, mes, a = int(m.group(2)), int(m.group(3)), int(m.group(4))
    return dt.date(a, mes, d) == dt.date.today()


def montar_email(app, origem, dest):
    m = app.CreateItem(0)
    m.To = "; ".join(dest["to"])
    m.CC = "; ".join(dest["cc"])
    m.BCC = "; ".join(dest["bcc"])
    m.Subject = re.sub(r"^\s*(RES:|ENC:|FW:|RE:)\s*", "", origem.Subject or "", flags=re.I).strip()
    m.HTMLBody = origem.HTMLBody
    cfg = carregar_config()
    if cfg["enviar_em_nome"] and cfg["caixa"]:
        m.SentOnBehalfOfName = cfg["caixa"]
    return m


def trava(nome, validade_min):
    """Trava simples na pasta da rede, para dois PCs nao fazerem a mesma coisa ao mesmo tempo."""
    caminho = os.path.join(PASTA, nome)
    try:
        if os.path.exists(caminho):
            idade = dt.datetime.now().timestamp() - os.path.getmtime(caminho)
            if idade < validade_min * 60:
                return None
            os.remove(caminho)
        fd = os.open(caminho, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.write(fd, os.environ.get("COMPUTERNAME", "?").encode())
        os.close(fd)
        return caminho
    except Exception:
        return None


def soltar(caminho):
    try:
        if caminho:
            os.remove(caminho)
    except Exception:
        pass


# ---------------- modo automatico (Agendador do Windows) ----------------
def rodar_auto():
    hoje = dt.date.today().isoformat()
    for f in os.listdir(PASTA):  # limpa travas de dias anteriores
        if f.startswith("proventos_envio_") and f.endswith(".lock") and hoje not in f:
            soltar(os.path.join(PASTA, f))
    lk = trava(f"proventos_envio_{hoje}.lock", 24 * 60)
    if not lk:
        log("AUTO: outro computador ja esta enviando ou ja enviou hoje, nada a fazer.")
        return
    ok = False
    try:
        ok = _rodar_auto()
    finally:
        if not ok:
            soltar(lk)


def _rodar_auto():
    try:
        if ja_enviado_hoje():
            log("AUTO: ja foi enviado hoje, nada a fazer.")
            return True
        dest = carregar_dest()
        if not any(dest[k] for k in CHAVES):
            log("AUTO: lista de destinatarios vazia, nao enviei.")
            return
        app = outlook()
        origem = ultimo_email_proventos(app)
        if origem is None:
            log("AUTO: nenhum e-mail 'Proventos' encontrado, nao enviei.")
            return
        if not eh_de_hoje(origem):
            log(f"AUTO: o mais recente e '{origem.Subject}' (nao e de hoje), nao enviei.")
            return
        montar_email(app, origem, dest).Send()
        marcar_enviado(origem, "automatico")
        n = sum(len(dest[k]) for k in CHAVES)
        log(f"AUTO: enviado '{origem.Subject}' para {n} destinatario(s).")
        return True
    except Exception as e:
        log(f"AUTO: ERRO {e!r}")


# ---------------- Radar na rede ----------------
RADAR_ASSUNTO = re.compile(r"^\s*(RES:|ENC:|FW:|RE:)?\s*Radar Corporativo - edicao ", re.I)
ARQ_RADAR_HTML = os.path.join(PASTA, "Radar.html")
ARQ_RADAR_HIST = os.path.join(PASTA, "radar_edicoes.json")
ARQ_RADAR_LOG = os.path.join(PASTA, "radar_log.txt")
NOME_TAREFA_RADAR = "Radar Corporativo rede"


def log_radar(msg):
    try:
        with open(ARQ_RADAR_LOG, "a", encoding="utf-8") as f:
            f.write(f"{dt.datetime.now():%d/%m/%Y %H:%M:%S}  {msg}\n")
    except Exception:
        pass


def emails_radar(app):
    achados = []
    for pasta in pastas_busca(app):
        try:
            itens = pasta.Items
            itens.Sort("[ReceivedTime]", True)
        except Exception:
            continue
        n = 0
        for it in itens:
            n += 1
            if n > 600:
                break
            try:
                if it.Class == 43 and RADAR_ASSUNTO.match(it.Subject or ""):
                    achados.append(it)
            except Exception:
                continue
    achados.sort(key=lambda it: it.ReceivedTime)
    return achados


def ler_pacote(item):
    import base64
    import gzip
    import tempfile
    for i in range(1, item.Attachments.Count + 1):
        att = item.Attachments.Item(i)
        if (att.FileName or "").lower().startswith("radar"):
            tmp = os.path.join(tempfile.gettempdir(), "radar_pacote.txt")
            att.SaveAsFile(tmp)
            with open(tmp, encoding="ascii") as f:
                b64 = "".join(f.read().split())
            os.remove(tmp)
            return json.loads(gzip.decompress(base64.b64decode(b64)).decode("utf-8"))
    raise RuntimeError("e-mail sem o anexo radar.txt")


def gravar_radar(template, eds):
    dados = json.dumps(eds, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = template.replace("<script>", "<script>window.RADAR_EDS = " + dados + ";</script>\n<script>", 1)
    tmp = ARQ_RADAR_HTML + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(html)
    os.replace(tmp, ARQ_RADAR_HTML)


def atualizar_radar(apagar_emails=True):
    """Junta as edicoes dos e-mails 'Radar Corporativo - edicao' no historico e regrava Radar.html."""
    lk = trava("radar.lock", 10)
    if not lk:
        return 0
    try:
        return _atualizar_radar(apagar_emails)
    finally:
        soltar(lk)


def _atualizar_radar(apagar_emails):
    app = outlook()
    itens = emails_radar(app)
    if not itens:
        return 0
    hist = ler_json(ARQ_RADAR_HIST, {"template": "", "edicoes": []})
    por_id = {e["id"]: e for e in hist.get("edicoes", [])}
    template = hist.get("template", "")
    novas = 0
    lidos = []
    for it in itens:
        try:
            pac = ler_pacote(it)
        except Exception as e:
            log_radar(f"REDE: nao consegui ler '{it.Subject}': {e!r}")
            continue
        template = pac.get("template") or template
        for e in pac.get("edicoes", []):
            if e["id"] not in por_id:
                novas += 1
            por_id[e["id"]] = e
        lidos.append(it)
    if not lidos or not template:
        return 0
    eds = sorted(por_id.values(), key=lambda e: e["ts"], reverse=True)
    limite = (dt.date.fromisoformat(eds[0]["dia"]) - dt.timedelta(days=3)).isoformat()
    eds = [e for e in eds if e["dia"] >= limite]
    gravar_json(ARQ_RADAR_HIST, {"template": template, "edicoes": eds})
    gravar_radar(template, eds)
    log_radar(f"REDE: {len(lidos)} e-mail(s) lido(s), {novas} edicao(oes) nova(s); mais recente {eds[0]['id']}.")
    if apagar_emails:
        for it in lidos:
            try:
                it.Delete()  # vai para Itens Excluidos
            except Exception:
                pass
    return novas


def rodar_rede():
    try:
        atualizar_radar()
    except Exception as e:
        log_radar(f"REDE: ERRO {e!r}")


# ---------------- Agendador do Windows ----------------
def pythonw():
    pasta = os.path.dirname(sys.executable)
    p = os.path.join(pasta, "pythonw.exe")
    return p if os.path.exists(p) else sys.executable


def _schtasks(args):
    r = subprocess.run(["schtasks"] + args, capture_output=True, text=True, creationflags=0x08000000)
    if r.returncode != 0:
        raise RuntimeError((r.stderr or r.stdout).strip())


def agendar():
    acao = f'"{pythonw()}" "{ARQ_PROG}" --auto'
    _schtasks(["/Create", "/F", "/TN", NOME_TAREFA, "/SC", "WEEKLY",
               "/D", "MON,TUE,WED,THU,FRI", "/ST", HORA_TAREFA, "/RI", "15", "/DU", "01:15", "/TR", acao])


def desagendar():
    _schtasks(["/Delete", "/F", "/TN", NOME_TAREFA])


def agendar_radar():
    acao = f'"{pythonw()}" "{ARQ_PROG}" --rede'
    _schtasks(["/Create", "/F", "/TN", NOME_TAREFA_RADAR, "/SC", "MINUTE", "/MO", "15", "/TR", acao])


def desagendar_radar():
    _schtasks(["/Delete", "/F", "/TN", NOME_TAREFA_RADAR])


def agendado(nome=NOME_TAREFA):
    r = subprocess.run(["schtasks", "/Query", "/TN", nome],
                       capture_output=True, text=True, creationflags=0x08000000)
    return r.returncode == 0


# ---------------- janela ----------------
def rodar_janela():
    import tkinter as tk
    from tkinter import messagebox, ttk

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("Proventos · envio pelo Outlook")
            self.geometry("680x820")
            self.minsize(580, 720)
            self.dest = carregar_dest()
            self.origem = None
            self._montar()
            self.atualizar_origem()
            self.atualizar_agenda()

        def _montar(self):
            pad = {"padx": 12, "pady": 5}
            topo = ttk.Frame(self)
            topo.pack(fill="x", **pad)
            ttk.Label(topo, text="E-mail de proventos mais recente:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
            self.lbl_origem = ttk.Label(topo, text="Procurando no Outlook…", font=("Segoe UI", 10))
            self.lbl_origem.pack(anchor="w")
            ttk.Button(topo, text="Procurar de novo", command=self.atualizar_origem).pack(anchor="w", pady=(4, 0))
            self.cfg = carregar_config()
            self.var_nome = tk.BooleanVar(value=self.cfg["enviar_em_nome"])
            ttk.Checkbutton(topo, text=f"Enviar em nome de {self.cfg['caixa']} (caixa da mesa)",
                            variable=self.var_nome, command=self.salvar_cfg).pack(anchor="w", pady=(6, 0))
            self.lbl_caixa = ttk.Label(topo, text="", font=("Segoe UI", 9), foreground="#5B6B66")
            self.lbl_caixa.pack(anchor="w")

            self.listas = {}
            for chave, titulo in (("to", "Para"), ("cc", "Cópia"),
                                  ("bcc", "Cópia oculta (CCO) — clientes não veem os e-mails uns dos outros")):
                fr = ttk.LabelFrame(self, text=titulo)
                fr.pack(fill="both", expand=True, **pad)
                lb = tk.Listbox(fr, height=3, font=("Segoe UI", 10), activestyle="none")
                lb.pack(side="left", fill="both", expand=True, padx=(8, 4), pady=6)
                lado = ttk.Frame(fr)
                lado.pack(side="right", fill="y", padx=(4, 8), pady=6)
                ent = ttk.Entry(lado, width=30)
                ent.pack(fill="x")
                ent.bind("<Return>", lambda e, k=chave: self.adicionar(k))
                ttk.Button(lado, text="Adicionar", command=lambda k=chave: self.adicionar(k)).pack(fill="x", pady=(4, 0))
                ttk.Button(lado, text="Remover selecionado", command=lambda k=chave: self.remover(k)).pack(fill="x", pady=(4, 0))
                self.listas[chave] = (lb, ent)
            self.desenhar_listas()

            rod = ttk.Frame(self)
            rod.pack(fill="x", **pad)
            ttk.Button(rod, text="Enviar agora", command=self.enviar).pack(side="left")
            ttk.Button(rod, text="Abrir no Outlook (conferir antes)", command=self.abrir).pack(side="left", padx=8)

            ag = ttk.LabelFrame(self, text="Envio automático (seg a sex, 8h45)")
            ag.pack(fill="x", **pad)
            self.lbl_agenda = ttk.Label(ag, text="", font=("Segoe UI", 10))
            self.lbl_agenda.pack(anchor="w", padx=8, pady=(6, 2))
            linha = ttk.Frame(ag)
            linha.pack(anchor="w", padx=8, pady=(2, 8))
            ttk.Button(linha, text="Ativar envio 8h45", command=self.ativar).pack(side="left")
            ttk.Button(linha, text="Desativar", command=self.desativar).pack(side="left", padx=8)
            ttk.Button(linha, text="Ver histórico", command=self.ver_log).pack(side="left")

            rd = ttk.LabelFrame(self, text="Radar completo na rede (Radar.html nesta pasta)")
            rd.pack(fill="x", **pad)
            self.lbl_radar = ttk.Label(rd, text="", font=("Segoe UI", 10))
            self.lbl_radar.pack(anchor="w", padx=8, pady=(6, 2))
            linha2 = ttk.Frame(rd)
            linha2.pack(anchor="w", padx=8, pady=(2, 8))
            ttk.Button(linha2, text="Atualizar agora", command=self.radar_agora).pack(side="left")
            ttk.Button(linha2, text="Ativar (a cada 15 min)", command=self.radar_ativar).pack(side="left", padx=8)
            ttk.Button(linha2, text="Desativar", command=self.radar_desativar).pack(side="left")
            ttk.Button(linha2, text="Abrir o Radar", command=self.radar_abrir).pack(side="left", padx=8)

            self.lbl_status = ttk.Label(self, text="", foreground="#0C7A4E", font=("Segoe UI", 9))
            self.lbl_status.pack(anchor="w", padx=12, pady=(0, 10))

        def salvar_cfg(self):
            self.cfg["enviar_em_nome"] = bool(self.var_nome.get())
            gravar_json(ARQ_CONFIG, self.cfg)

        def desenhar_listas(self):
            for k, (lb, _) in self.listas.items():
                lb.delete(0, "end")
                for e in self.dest[k]:
                    lb.insert("end", e)

        def adicionar(self, k):
            lb, ent = self.listas[k]
            novos = [x.strip().lower() for x in re.split(r"[;,\s]+", ent.get()) if x.strip()]
            ruins = [x for x in novos if not EMAIL_OK.match(x)]
            if ruins:
                messagebox.showwarning("E-mail inválido", "Confira: " + ", ".join(ruins))
                return
            for x in novos:
                if not any(x in self.dest[c] for c in CHAVES):
                    self.dest[k].append(x)
            gravar_json(ARQ_DEST, self.dest)
            ent.delete(0, "end")
            self.desenhar_listas()

        def remover(self, k):
            lb, _ = self.listas[k]
            sel = lb.curselection()
            if not sel:
                return
            del self.dest[k][sel[0]]
            gravar_json(ARQ_DEST, self.dest)
            self.desenhar_listas()

        def atualizar_origem(self):
            try:
                app = outlook()
                ok = bool(pastas_caixa_mesa(app.GetNamespace("MAPI"), self.cfg["caixa"]))
                self.lbl_caixa.config(text=(f"Lendo a caixa {self.cfg['caixa']} e a sua." if ok else
                                            f"Não achei a caixa {self.cfg['caixa']} no seu Outlook; lendo só a sua."))
                self.origem = ultimo_email_proventos(app)
                if self.origem is None:
                    self.lbl_origem.config(text="Nenhum e-mail \"Proventos DD/MM/AAAA\" encontrado.")
                else:
                    rec = self.origem.ReceivedTime.strftime("%d/%m %H:%M")
                    self.lbl_origem.config(text=f"{self.origem.Subject}  ·  recebido em {rec}")
            except Exception as e:
                self.origem = None
                self.lbl_origem.config(text=f"Não consegui ler o Outlook: {e}")

        def atualizar_agenda(self):
            est = ler_json(ARQ_ESTADO, {})
            ult = f"  ·  último envio: {est['quando']} ({est['modo']})" if est.get("quando") else ""
            try:
                txt = "ATIVO — envia sozinho às 8h45" if agendado() else "Desativado"
            except Exception:
                txt = "Não consegui consultar o Agendador"
            self.lbl_agenda.config(text=txt + ult)
            try:
                ativo = "ATIVO" if agendado(NOME_TAREFA_RADAR) else "Desativado"
            except Exception:
                ativo = "?"
            hist = ler_json(ARQ_RADAR_HIST, {}).get("edicoes", [])
            if hist:
                e = hist[0]
                quando = f"edição de {e['dia'][8:10]}/{e['dia'][5:7]} {e['ts'][11:16]}"
            else:
                quando = "ainda sem edições"
            self.lbl_radar.config(text=f"{ativo}  ·  {quando}")

        def radar_agora(self):
            try:
                n = atualizar_radar()
                self.lbl_status.config(text=f"Radar atualizado: {n} edição(ões) nova(s).")
            except Exception as e:
                messagebox.showerror("Radar", str(e))
            self.atualizar_agenda()

        def radar_ativar(self):
            try:
                agendar_radar()
                messagebox.showinfo("Radar", "Ativado. A cada 15 minutos o programa pega os e-mails do Radar, "
                                    "atualiza o Radar.html nesta pasta e apaga esses e-mails (vão para Itens Excluídos).")
            except Exception as e:
                messagebox.showerror("Não consegui agendar", str(e))
            self.atualizar_agenda()

        def radar_desativar(self):
            try:
                desagendar_radar()
            except Exception as e:
                messagebox.showerror("Erro", str(e))
            self.atualizar_agenda()

        def radar_abrir(self):
            if os.path.exists(ARQ_RADAR_HTML):
                os.startfile(ARQ_RADAR_HTML)
            else:
                messagebox.showinfo("Radar", "O Radar.html ainda não foi gerado. Clique em \"Atualizar agora\".")

        def _pronto(self):
            if self.origem is None:
                messagebox.showwarning("Proventos", "Não encontrei o e-mail de proventos. Clique em \"Procurar de novo\".")
                return False
            if not any(self.dest[k] for k in CHAVES):
                messagebox.showwarning("Proventos", "Cadastre pelo menos um destinatário.")
                return False
            return True

        def enviar(self):
            if not self._pronto():
                return
            n = sum(len(self.dest[k]) for k in CHAVES)
            if not messagebox.askyesno("Confirmar envio", f"Enviar \"{self.origem.Subject}\" para {n} destinatário(s)?"):
                return
            try:
                montar_email(outlook(), self.origem, self.dest).Send()
                if eh_de_hoje(self.origem):
                    marcar_enviado(self.origem, "manual")
                log(f"MANUAL: enviado '{self.origem.Subject}' para {n} destinatario(s).")
                self.lbl_status.config(text=f"Enviado: {self.origem.Subject} para {n} destinatário(s).")
                self.atualizar_agenda()
            except Exception as e:
                messagebox.showerror("Erro ao enviar", f"{e}\n\nTente \"Abrir no Outlook\" e envie por lá.")

        def abrir(self):
            if not self._pronto():
                return
            try:
                montar_email(outlook(), self.origem, self.dest).Display()
                self.lbl_status.config(text="E-mail aberto no Outlook: confira e clique em Enviar.")
            except Exception as e:
                messagebox.showerror("Erro", str(e))

        def ativar(self):
            try:
                agendar()
                messagebox.showinfo("Envio automático",
                                    "Ativado. De segunda a sexta, às 8h45, o e-mail de proventos do dia "
                                    "sai sozinho para a lista cadastrada. Se o e-mail das 8h30 atrasar, "
                                    "o programa tenta de novo a cada 15 min até as 10h.\n\n"
                                    "O computador precisa estar ligado e com você logada.")
            except Exception as e:
                messagebox.showerror("Não consegui agendar", str(e))
            self.atualizar_agenda()

        def desativar(self):
            try:
                desagendar()
            except Exception as e:
                messagebox.showerror("Erro", str(e))
            self.atualizar_agenda()

        def ver_log(self):
            if os.path.exists(ARQ_LOG):
                os.startfile(ARQ_LOG)
            else:
                messagebox.showinfo("Histórico", "Ainda não há nenhum envio registrado.")

    App().mainloop()


if __name__ == "__main__":
    if "--auto" in sys.argv:
        rodar_auto()
    elif "--rede" in sys.argv:
        rodar_rede()
    else:
        rodar_janela()
