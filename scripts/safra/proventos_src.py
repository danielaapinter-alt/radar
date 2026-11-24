# -*- coding: utf-8 -*-
# Proventos - envio pelo Outlook (Radar Corporativo)  - versao 2
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


def ultimo_email_proventos(app):
    ns = app.GetNamespace("MAPI")
    melhor = None
    for pasta_id in (6, 23):  # 6 = Caixa de Entrada, 23 = Lixo Eletronico
        try:
            itens = ns.GetDefaultFolder(pasta_id).Items
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
    return m


# ---------------- modo automatico (Agendador do Windows) ----------------
def rodar_auto():
    try:
        if ja_enviado_hoje():
            log("AUTO: ja foi enviado hoje, nada a fazer.")
            return
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
    except Exception as e:
        log(f"AUTO: ERRO {e!r}")


# ---------------- Agendador do Windows ----------------
def pythonw():
    pasta = os.path.dirname(sys.executable)
    p = os.path.join(pasta, "pythonw.exe")
    return p if os.path.exists(p) else sys.executable


def agendar():
    acao = f'"{pythonw()}" "{ARQ_PROG}" --auto'
    cmd = ["schtasks", "/Create", "/F", "/TN", NOME_TAREFA, "/SC", "WEEKLY",
           "/D", "MON,TUE,WED,THU,FRI", "/ST", HORA_TAREFA, "/TR", acao]
    r = subprocess.run(cmd, capture_output=True, text=True, creationflags=0x08000000)
    if r.returncode != 0:
        raise RuntimeError((r.stderr or r.stdout).strip())


def desagendar():
    r = subprocess.run(["schtasks", "/Delete", "/F", "/TN", NOME_TAREFA],
                       capture_output=True, text=True, creationflags=0x08000000)
    if r.returncode != 0:
        raise RuntimeError((r.stderr or r.stdout).strip())


def agendado():
    r = subprocess.run(["schtasks", "/Query", "/TN", NOME_TAREFA],
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
            self.geometry("660x720")
            self.minsize(560, 640)
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

            self.lbl_status = ttk.Label(self, text="", foreground="#0C7A4E", font=("Segoe UI", 9))
            self.lbl_status.pack(anchor="w", padx=12, pady=(0, 10))

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
                self.origem = ultimo_email_proventos(outlook())
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
                                    "sai sozinho do seu Outlook para a lista cadastrada.\n\n"
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
    else:
        rodar_janela()
