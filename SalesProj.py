from datetime import datetime
import os
import platform
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
import pandas as pd

ctk.set_appearance_mode("Light")

NOME_ARQUIVO = "VENDAS PIETRA ACELERANDO.xlsx"

MESES_PT = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}


class AppVendas(ctk.CTk):

  def __init__(self):
    super().__init__()

    self.title("🌸 Gerenciador de Vendas - Pietra Acelerando")
    self.geometry("980x620")
    self.configure(fg_color="#FFF5F7")

    self.venda_em_edicao_idx = None
    self.vendas = self.carregar_dados_excel()

    self.estilar_tabela()
    self.criar_interface()

  def estilar_tabela(self):
    style = ttk.Style()
    style.theme_use("default")

    style.configure(
        "Treeview",
        background="#FFFFFF",
        foreground="#4A3B40",
        rowheight=28,
        fieldbackground="#FFFFFF",
        font=("Segoe UI", 10),
    )
    style.configure(
        "Treeview.Heading",
        background="#FFD1DC",
        foreground="#6B4C57",
        font=("Segoe UI", 10, "bold"),
    )
    style.map("Treeview", background=[("selected", "#FFB6C1")])

  def carregar_dados_excel(self):
    vendas_carregadas = []
    if os.path.exists(NOME_ARQUIVO):
      try:
        df = pd.read_excel(NOME_ARQUIVO, sheet_name=0)
        for _, row in df.iterrows():
          lead = str(row.get("Nome do lead", "")).strip()
          cidade = str(row.get("Cidade", "")).strip()
          valor_raw = row.get("Valor da venda (R$)", 0)

          if lead and lead != "nan":
            try:
              valor = float(valor_raw)
            except (ValueError, TypeError):
              valor = 0.0

            data_str = str(row.get("Data", "2026-10-01")).strip()
            dt = self.converter_data(data_str)
            if not dt:
              dt = datetime(2026, 10, 1)

            mes_str = f"{MESES_PT[dt.month]}/{dt.year}"
            vendas_carregadas.append({
                "Data": dt.strftime("%Y-%m-%d"),
                "Data_Fmt": dt.strftime("%d/%m/%Y"),
                "Mes": mes_str,
                "Lead": lead,
                "Cidade": cidade,
                "Valor": valor,
            })
      except Exception as e:
        messagebox.showerror(
            "Erro ao carregar Excel", f"Não foi possível ler a planilha:\n{e}"
        )
    return vendas_carregadas

  def salvar_no_excel(self):
    try:
      df = pd.DataFrame(self.vendas)
      if not df.empty:
        df_export = pd.DataFrame({
            "Data": df["Data"],
            "Nome do lead": df["Lead"],
            "Cidade": df["Cidade"],
            "Valor da venda (R$)": df["Valor"],
        })
      else:
        df_export = pd.DataFrame(
            columns=["Data", "Nome do lead", "Cidade", "Valor da venda (R$)"]
        )

      df_export.to_excel(NOME_ARQUIVO, sheet_name="Minhas Vendas", index=False)
    except Exception as e:
      messagebox.showerror(
          "Erro ao Salvar",
          f"Erro ao gravar os dados no arquivo Excel:\n{e}\n\nFeche o arquivo Excel se ele estiver aberto.",
      )

  def abrir_planilha_excel(self):
    if not os.path.exists(NOME_ARQUIVO):
      self.salvar_no_excel()

    caminho_absoluto = os.path.abspath(NOME_ARQUIVO)
    sistema = platform.system()

    try:
      if sistema == "Darwin":
        subprocess.run(["open", caminho_absoluto])
      elif sistema == "Windows":
        os.startfile(caminho_absoluto)
      else:
        subprocess.run(["xdg-open", caminho_absoluto])
    except Exception as e:
      messagebox.showerror("Erro", f"Não foi possível abrir a planilha:\n{e}")

  def converter_data(self, data_str):
    for formato in ("%Y-%m-%d", "%d/%m/%Y"):
      try:
        return datetime.strptime(data_str.strip(), formato)
      except ValueError:
        pass
    return None

  def criar_interface(self):
    self.grid_columnconfigure(1, weight=1)
    self.grid_rowconfigure(0, weight=1)

    # --- PAINEL LATERAL ---
    self.frame_lateral = ctk.CTkFrame(
        self, width=280, corner_radius=15, fg_color="#FFF0F5"
    )
    self.frame_lateral.grid(row=0, column=0, padx=15, pady=15, sticky="nsew")

    self.lbl_titulo_form = ctk.CTkLabel(
        self.frame_lateral,
        text="✨ Nova Venda",
        font=ctk.CTkFont(size=18, weight="bold"),
        text_color="#8B4565",
    )
    self.lbl_titulo_form.pack(pady=15)

    self.ent_lead = ctk.CTkEntry(
        self.frame_lateral,
        placeholder_text="Nome do Lead",
        border_color="#FFB6C1",
    )
    self.ent_lead.pack(pady=8, padx=15, fill="x")

    self.ent_cidade = ctk.CTkEntry(
        self.frame_lateral,
        placeholder_text="Cidade / Unidade",
        border_color="#FFB6C1",
    )
    self.ent_cidade.pack(pady=8, padx=15, fill="x")

    self.ent_valor = ctk.CTkEntry(
        self.frame_lateral,
        placeholder_text="Valor da Venda (R$)",
        border_color="#FFB6C1",
    )
    self.ent_valor.pack(pady=8, padx=15, fill="x")

    self.ent_data = ctk.CTkEntry(
        self.frame_lateral,
        placeholder_text="Data (DD/MM/AAAA)",
        border_color="#FFB6C1",
    )
    self.ent_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
    self.ent_data.pack(pady=8, padx=15, fill="x")

    self.btn_salvar = ctk.CTkButton(
        self.frame_lateral,
        text="Registrar Venda 💖",
        command=self.registrar_ou_atualizar_venda,
        fg_color="#FFB6C1",
        hover_color="#FF99AC",
        text_color="#FFFFFF",
        font=ctk.CTkFont(weight="bold"),
    )
    self.btn_salvar.pack(pady=15, padx=15, fill="x")

    self.btn_cancelar_edicao = ctk.CTkButton(
        self.frame_lateral,
        text="Cancelar Edição",
        command=self.limpar_formulario,
        fg_color="#D8B4E2",
        hover_color="#C892D6",
        text_color="#FFFFFF",
    )

    # --- PAINEL PRINCIPAL ---
    self.frame_principal = ctk.CTkFrame(
        self, corner_radius=15, fg_color="#FFF0F5"
    )
    self.frame_principal.grid(row=0, column=1, padx=(0, 15), pady=15, sticky="nsew")

    # Barra Superior
    self.frame_topo = ctk.CTkFrame(self.frame_principal, fg_color="transparent")
    self.frame_topo.pack(fill="x", padx=15, pady=10)

    ctk.CTkLabel(
        self.frame_topo,
        text="📅 Mês:",
        font=ctk.CTkFont(weight="bold"),
        text_color="#8B4565",
    ).pack(side="left", padx=5)

    self.combo_mes = ctk.CTkOptionMenu(
        self.frame_topo,
        values=self.obter_lista_meses(),
        command=self.atualizar_tela,
        fg_color="#FFB6C1",
        button_color="#FF99AC",
        button_hover_color="#E07A93",
        text_color="#FFFFFF",
    )
    self.combo_mes.pack(side="left", padx=5)

    btn_abrir_excel = ctk.CTkButton(
        self.frame_topo,
        text="📊 Abrir Planilha Excel",
        command=self.abrir_planilha_excel,
        fg_color="#F3B0C3",
        hover_color="#E695AC",
        text_color="#FFFFFF",
    )
    btn_abrir_excel.pack(side="right", padx=5)

    # Card de Resumo
    self.lbl_total = ctk.CTkLabel(
        self.frame_principal,
        text="Total Vendas: R$ 0,00",
        font=ctk.CTkFont(size=16, weight="bold"),
        text_color="#B85B7A",
    )
    self.lbl_total.pack(pady=5)

    # Tabela Interativa
    frame_tabela = ctk.CTkFrame(self.frame_principal, fg_color="#FFFFFF")
    frame_tabela.pack(fill="both", expand=True, padx=15, pady=10)

    colunas = ("ID", "Data", "Lead", "Cidade", "Valor")
    self.tabela = ttk.Treeview(
        frame_tabela, columns=colunas, show="headings", height=10
    )

    self.tabela.heading("ID", text="#")
    self.tabela.heading("Data", text="Data")
    self.tabela.heading("Lead", text="Lead / Nome")
    self.tabela.heading("Cidade", text="Cidade / Unidade")
    self.tabela.heading("Valor", text="Valor (R$)")

    self.tabela.column("ID", width=40, anchor="center")
    self.tabela.column("Data", width=90, anchor="center")
    self.tabela.column("Lead", width=180, anchor="w")
    self.tabela.column("Cidade", width=180, anchor="w")
    self.tabela.column("Valor", width=110, anchor="e")

    scrollbar = ttk.Scrollbar(
        frame_tabela, orient="vertical", command=self.tabela.yview
    )
    self.tabela.configure(yscrollcommand=scrollbar.set)

    self.tabela.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Ações de Edição/Exclusão com Padrão de Cores (Amarelo / Vermelho)
    frame_acoes = ctk.CTkFrame(self.frame_principal, fg_color="transparent")
    frame_acoes.pack(fill="x", padx=15, pady=5)

    btn_editar = ctk.CTkButton(
        frame_acoes,
        text="✏️ Editar Selecionada",
        command=self.carregar_venda_para_edicao,
        fg_color="#EAB308",  # Amarelo Padrão de Edição
        hover_color="#CA8A04",
        text_color="#FFFFFF",
        font=ctk.CTkFont(weight="bold"),
    )
    btn_editar.pack(side="left", padx=5)

    btn_excluir = ctk.CTkButton(
        frame_acoes,
        text="🗑️ Excluir Selecionada",
        command=self.excluir_venda,
        fg_color="#EF4444",  # Vermelho Padrão de Exclusão
        hover_color="#DC2626",
        text_color="#FFFFFF",
        font=ctk.CTkFont(weight="bold"),
    )
    btn_excluir.pack(side="left", padx=5)

    self.tabela.bind("<Double-1>", lambda e: self.carregar_venda_para_edicao())

    mes_inicial = self.combo_mes.get()
    self.atualizar_tela(mes_inicial if mes_inicial else "Outubro/2026")

  def obter_lista_meses(self):
    meses = sorted(list(set(v["Mes"] for v in self.vendas)))
    return meses if meses else ["Outubro/2026"]

  def registrar_ou_atualizar_venda(self):
    lead = self.ent_lead.get().strip()
    cidade = self.ent_cidade.get().strip()
    valor_raw = self.ent_valor.get().strip()
    data_raw = self.ent_data.get().strip()

    if not lead or not cidade or not valor_raw:
      messagebox.showwarning(
          "Aviso", "Por favor, preencha todos os campos do formulário."
      )
      return

    try:
      valor = float(valor_raw.replace(".", "").replace(",", "."))
    except ValueError:
      messagebox.showerror(
          "Erro de Valor", "Digite um valor numérico válido (ex: 25000,00)."
      )
      return

    dt = self.converter_data(data_raw)
    if not dt:
      messagebox.showerror(
          "Erro de Data", "Formato de data inválido. Use DD/MM/AAAA."
      )
      return

    data_iso = dt.strftime("%Y-%m-%d")
    data_fmt = dt.strftime("%d/%m/%Y")
    mes_str = f"{MESES_PT[dt.month]}/{dt.year}"

    dados_venda = {
        "Data": data_iso,
        "Data_Fmt": data_fmt,
        "Mes": mes_str,
        "Lead": lead,
        "Cidade": cidade,
        "Valor": valor,
    }

    if self.venda_em_edicao_idx is not None:
      self.vendas[self.venda_em_edicao_idx] = dados_venda
      messagebox.showinfo("Sucesso", "Venda atualizada com sucesso!")
    else:
      self.vendas.append(dados_venda)

    self.salvar_no_excel()

    lista_meses = self.obter_lista_meses()
    self.combo_mes.configure(values=lista_meses)
    self.combo_mes.set(mes_str)

    self.limpar_formulario()
    self.atualizar_tela(mes_str)

  def carregar_venda_para_edicao(self):
    item_sel = self.tabela.selection()
    if not item_sel:
      messagebox.showwarning(
          "Aviso", "Selecione uma venda na tabela para editar."
      )
      return

    valores = self.tabela.item(item_sel[0], "values")
    idx_venda = int(valores[0])

    venda = self.vendas[idx_venda]
    self.venda_em_edicao_idx = idx_venda

    self.ent_lead.delete(0, tk.END)
    self.ent_lead.insert(0, venda["Lead"])

    self.ent_cidade.delete(0, tk.END)
    self.ent_cidade.insert(0, venda["Cidade"])

    self.ent_valor.delete(0, tk.END)
    self.ent_valor.insert(0, f"{venda['Valor']:.2f}".replace(".", ","))

    self.ent_data.delete(0, tk.END)
    self.ent_data.insert(0, venda["Data_Fmt"])

    self.lbl_titulo_form.configure(text="✏️ Editar Venda")
    self.btn_salvar.configure(
        text="Salvar Alterações ✨", fg_color="#EAB308", hover_color="#CA8A04"
    )
    self.btn_cancelar_edicao.pack(pady=5, padx=15, fill="x")

  def excluir_venda(self):
    item_sel = self.tabela.selection()
    if not item_sel:
      messagebox.showwarning(
          "Aviso", "Selecione uma venda na tabela para excluir."
      )
      return

    if messagebox.askyesno(
        "Confirmar", "Tem certeza de que deseja excluir esta venda?"
    ):
      valores = self.tabela.item(item_sel[0], "values")
      idx_venda = int(valores[0])

      del self.vendas[idx_venda]
      self.salvar_no_excel()

      mes_atual = self.combo_mes.get()
      self.combo_mes.configure(values=self.obter_lista_meses())
      self.atualizar_tela(mes_atual)

  def limpar_formulario(self):
    self.venda_em_edicao_idx = None
    self.ent_lead.delete(0, tk.END)
    self.ent_cidade.delete(0, tk.END)
    self.ent_valor.delete(0, tk.END)
    self.ent_data.delete(0, tk.END)
    self.ent_data.insert(0, datetime.now().strftime("%d/%m/%Y"))

    self.lbl_titulo_form.configure(text="✨ Nova Venda")
    self.btn_salvar.configure(
        text="Registrar Venda 💖", fg_color="#FFB6C1", hover_color="#FF99AC"
    )
    self.btn_cancelar_edicao.pack_forget()

  def atualizar_tela(self, mes_selecionado):
    for item in self.tabela.get_children():
      self.tabela.delete(item)

    vendas_mes = [
        (idx, v) for idx, v in enumerate(self.vendas) if v["Mes"] == mes_selecionado
    ]

    total_faturado = sum(v["Valor"] for _, v in vendas_mes)
    total_qtd = len(vendas_mes)

    self.lbl_total.configure(
        text=f"Mês: {mes_selecionado}  |  Qtd Vendas: {total_qtd}  |  Total: R$ {total_faturado:,.2f}".replace(
            ",", "X"
        )
        .replace(".", ",")
        .replace("X", ".")
    )

    for idx, v in vendas_mes:
      valor_fmt = (
          f"R$ {v['Valor']:,.2f}"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      self.tabela.insert(
          "",
          "end",
          values=(
              idx,
              v["Data_Fmt"],
              v["Lead"],
              v["Cidade"],
              valor_fmt,
          ),
      )


if __name__ == "__main__":
  app = AppVendas()
  app.mainloop()