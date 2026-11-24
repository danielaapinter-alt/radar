# -*- coding: utf-8 -*-
"""
Proventos - envio pelo Outlook (Radar Corporativo)

O que faz:
  1. Procura na sua caixa do Outlook o e-mail mais recente com assunto "Proventos DD/MM/AAAA"
     (enviado pelo Radar às 8h30, 12h e 18h).
  2. Mostra a lista de destinatários (Para e Cópia), que você pode editar. A lista fica salva
     no arquivo proventos_destinatarios.json, na mesma pasta deste programa.
  3. "Enviar agora" manda um e-mail novo, do seu Outlook, com o mesmo conteúdo e assunto.
     "Abrir no Outlook" abre o e-mail pronto para você conferir e clicar em Enviar.

Requisitos (uma vez só, no Prompt de Comando):  pip install pywin32
Para abrir: dê dois cliques em proventos.pyw (ou rode:  pythonw proventos.pyw)
"""
import json
import os
import re
import tkinter as tk
from tkinter import messagebox, ttk

try:
    import win32com.client  # pacote pywin32
except ImportError:
    win32com = None

PASTA = os.path.dirname(os.path.abspath(__file__))
ARQ_DEST = os.path.join(PASTA, "proventos_destinatarios.json")
PADRAO = {
    "to": ["mesa.aluguel@safra.com.br"],
    "cc": ["daniela.pinter@safra.com.br", "jefferson.dias@safra.com.br", "william.parada@safra.com.br"],
}
ASSUNTO = re.compile(r"^\s*(RES:|ENC:|FW:|RE:)?\s*Proventos \d{2}/\d{2}/\d{4}\s*$", re.I)
EMAIL_OK = re.compile(r"^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$")


# ---------------- destinatários ----------------
def carregar_dest():
    try:
        with open(ARQ_DEST, encoding="utf-8") as f:
            d = json.load(f)
        return {"to": list(d.get("to", [])), "cc": list(d.get("cc", []))}
    except Exception:
        return {"to": list(PADRAO["to"]), "cc": list(PADRAO["cc"])}


def salvar_dest(d):
    with open(ARQ_DEST, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)


# ---------------- Outlook ----------------
def outlook():
    if win32com is None:
        raise RuntimeError("Falta o pacote pywin32. Abra o Prompt de Comando e rode:  pip install pywin32")
    return win32com.client.Dispatch("Outlook.Application")


def ultimo_email_proventos(app):
    """Devolve o e-mail 'Proventos DD/MM/AAAA' mais recente (Caixa de Entrada e Lixo Eletrônico)."""
    ns = app.GetNamespace("MAPI")
    melhor = None
    for pasta_id in (6, 23):  # 6 = Caixa de Entrada, 23 = Lixo Eletrônico
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
                if it.Class != 43:  # 43 = e-mail
                    continue
                if ASSUNTO.match(it.Subject or ""):
                    if melhor is None or it.ReceivedTime > melhor.ReceivedTime:
                        melhor = it
                    break
            except Exception:
                continue
    return melhor


def montar_email(app, origem, dest):
    m = app.CreateItem(0)  # 0 = e-mail
    m.To = "; ".join(dest["to"])
    m.CC = "; ".join(dest["cc"])
    assunto = re.sub(r"^\s*(RES:|ENC:|FW:|RE:)\s*", "", origem.Subject or "", flags=re.I).strip()
    m.Subject = assunto
    m.HTMLBody = origem.HTMLBody
    return m


# ---------------- janela ----------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Proventos · envio pelo Outlook")
        self.geometry("620x520")
        self.minsize(520, 460)
        self.dest = carregar_dest()
        self.origem = None
        self._montar()
        self.atualizar_origem()

    def _montar(self):
        pad = {"padx": 12, "pady": 6}
        topo = ttk.Frame(self)
        topo.pack(fill="x", **pad)
        ttk.Label(topo, text="E-mail de proventos mais recente:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.lbl_origem = ttk.Label(topo, text="Procurando no Outlook…", font=("Segoe UI", 10))
        self.lbl_origem.pack(anchor="w")
        ttk.Button(topo, text="Procurar de novo", command=self.atualizar_origem).pack(anchor="w", pady=(4, 0))

        self.listas = {}
        for chave, titulo in (("to", "Para"), ("cc", "Cópia")):
            fr = ttk.LabelFrame(self, text=titulo)
            fr.pack(fill="both", expand=True, **pad)
            lb = tk.Listbox(fr, height=4, font=("Segoe UI", 10), activestyle="none")
            lb.pack(side="left", fill="both", expand=True, padx=(8, 4), pady=8)
            lado = ttk.Frame(fr)
            lado.pack(side="right", fill="y", padx=(4, 8), pady=8)
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
            if x not in self.dest["to"] and x not in self.dest["cc"]:
                self.dest[k].append(x)
        salvar_dest(self.dest)
        ent.delete(0, "end")
        self.desenhar_listas()

    def remover(self, k):
        lb, _ = self.listas[k]
        sel = lb.curselection()
        if not sel:
            return
        del self.dest[k][sel[0]]
        salvar_dest(self.dest)
        self.desenhar_listas()

    def atualizar_origem(self):
        try:
            self.origem = ultimo_email_proventos(outlook())
            if self.origem is None:
                self.lbl_origem.config(text="Nenhum e-mail \"Proventos DD/MM/AAAA\" encontrado na Caixa de Entrada.")
            else:
                rec = self.origem.ReceivedTime.strftime("%d/%m %H:%M")
                self.lbl_origem.config(text=f"{self.origem.Subject}  ·  recebido em {rec}")
        except Exception as e:
            self.origem = None
            self.lbl_origem.config(text=f"Não consegui ler o Outlook: {e}")

    def _pronto(self):
        if self.origem is None:
            messagebox.showwarning("Proventos", "Não encontrei o e-mail de proventos. Clique em \"Procurar de novo\".")
            return False
        if not self.dest["to"] and not self.dest["cc"]:
            messagebox.showwarning("Proventos", "Cadastre pelo menos um destinatário.")
            return False
        return True

    def enviar(self):
        if not self._pronto():
            return
        n = len(self.dest["to"]) + len(self.dest["cc"])
        if not messagebox.askyesno("Confirmar envio", f"Enviar \"{self.origem.Subject}\" para {n} destinatário(s)?"):
            return
        try:
            montar_email(outlook(), self.origem, self.dest).Send()
            self.lbl_status.config(text=f"Enviado: {self.origem.Subject} para {n} destinatário(s).")
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


if __name__ == "__main__":
    App().mainloop()
