"""Janelas modais de revisão e configuração por módulo (Canto L e Revisão de Peças)."""
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
import pandas as pd

from constantes import (
    COR_TEXTO_SECUNDARIO,
    COR_TITULO,
    COR_VERDE_PRIMARIO,
    COR_VERDE_PRIMARIO_HOVER,
    DICAS_COLUNAS,
    FORNECEDORES_FORMATO_PADRAO,
    FORNECEDORES_FORMATO_TMK,
    LIMITE_MAX_COMPRIMENTO,
    LIMITE_MAX_LARGURA,
    OPCOES_MARGEM_TERCEIRIZACAO_MM,
)
from regras_marcenaria import (
    calcular_deltas_ajuste_df,
    formatar_medida_mm,
)
from ui_componentes import TourGuiado, criar_tooltip


class JanelaModulosCantoL(ctk.CTkToplevel):
  """Pergunta, para cada módulo de canto em L encontrado no XML, se a
  peça mantém a medida original ou soma margem de 50/100mm nas peças de
  Base, Tampo e Prateleira — margem usada quando a peça vai ser cortada
  numa máquina que exige essa folga extra. Usa grid (não pack) pro botão
  de confirmar ficar sempre visível, mesmo com muitos módulos na lista —
  antes ele podia ficar escondido abaixo da borda da janela até
  maximizar."""

  OPCAO_SEM_MARGEM = "Medida Original"

  def __init__(self, parent, modulos: list):
    super().__init__(parent)
    self.title("Módulos de Canto L")
    self.geometry("640x500")
    self.minsize(560, 380)
    self.grab_set()

    self.resultado = {}
    self.vars_margem = {}

    self.grid_columnconfigure(0, weight=1)
    self.grid_rowconfigure(1, weight=1)

    frame_topo = ctk.CTkFrame(self, fg_color="transparent")
    frame_topo.grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 5))

    ctk.CTkLabel(
        frame_topo,
        text="📐 Módulos de Canto L encontrados",
        font=ctk.CTkFont(size=16, weight="bold"),
        text_color=COR_TITULO,
        anchor="w",
    ).pack(fill="x")

    ctk.CTkLabel(
        frame_topo,
        text=(
            "Módulos de canto L às vezes são cortados numa máquina que"
            " exige margem extra na peça. Para cada módulo, escolha a"
            " medida original (sem alteração) ou uma margem de"
            " +50mm/+100mm, somada na largura e no comprimento das peças"
            " de Base, Tampo e Prateleira."
        ),
        text_color=COR_TEXTO_SECUNDARIO,
        font=ctk.CTkFont(size=12),
        justify="left",
        wraplength=580,
        anchor="w",
    ).pack(fill="x", pady=(6, 0))

    frame_lista = ctk.CTkScrollableFrame(self)
    frame_lista.grid(row=1, column=0, sticky="nsew", padx=20, pady=5)

    self._opcoes_valores = {
        f"+{mm}mm": mm for mm in OPCOES_MARGEM_TERCEIRIZACAO_MM
    }
    opcoes = [self.OPCAO_SEM_MARGEM] + list(self._opcoes_valores.keys())
    for modulo in modulos:
      nome = modulo["nome"]
      linha = ctk.CTkFrame(frame_lista, fg_color="transparent")
      linha.pack(fill="x", pady=8, padx=4)

      ctk.CTkLabel(
          linha,
          text=nome,
          font=ctk.CTkFont(size=13, weight="bold"),
          anchor="w",
      ).pack(fill="x")

      medida = (
          f"Medida do módulo: {formatar_medida_mm(modulo['comprimento'])} x"
          f" {formatar_medida_mm(modulo['largura'])} mm"
      )
      ctk.CTkLabel(
          linha,
          text=medida,
          text_color=COR_TEXTO_SECUNDARIO,
          font=ctk.CTkFont(size=12),
          anchor="w",
      ).pack(fill="x", pady=(0, 6))

      var = ctk.StringVar(value=self.OPCAO_SEM_MARGEM)
      self.vars_margem[nome] = var
      ctk.CTkSegmentedButton(
          linha, values=opcoes, variable=var
      ).pack(fill="x")

    frame_rodape = ctk.CTkFrame(self, fg_color="transparent")
    frame_rodape.grid(row=2, column=0, sticky="ew", padx=20, pady=15)

    btn_confirmar = ctk.CTkButton(
        frame_rodape,
        text="✅ Confirmar",
        command=self._confirmar,
        fg_color=COR_VERDE_PRIMARIO,
        hover_color=COR_VERDE_PRIMARIO_HOVER,
        height=42,
    )
    btn_confirmar.pack(fill="x")

    self.protocol("WM_DELETE_WINDOW", self._confirmar)

  def _confirmar(self):
    self.resultado = {
        nome: self._opcoes_valores[var.get()]
        for nome, var in self.vars_margem.items()
        if var.get() in self._opcoes_valores
    }
    self.grab_release()
    self.destroy()


class JanelaRevisao(ctk.CTkToplevel):

  def __init__(
      self,
      parent,
      df: pd.DataFrame,
      callback_salvar,
      fornecedor_atual="Teletintas",
      iniciar_grab=True,
  ):
    super().__init__(parent)
    self.title("Revisão de Peças para Exportação")
    self.geometry("1120 x 660")
    self.minsize(940, 520)

    self.parent = parent
    self.df = df.copy()
    if hasattr(df, "attrs") and df.attrs:
      self.df.attrs = dict(df.attrs)
    self.callback_salvar = callback_salvar
    self._deltas_fundo, self._deltas_cf = calcular_deltas_ajuste_df(self.df)
    self._fundo_aplicado = {idx: False for idx in self.df.index}
    self._cf_aplicado = {idx: False for idx in self.df.index}

    # Sem transient(): no Windows, uma janela "transient" (presa à janela
    # dona) perde os botões de minimizar/maximizar — é comportamento nativo
    # de janela "owned", não uma opção separada. grab_set() sozinho já basta
    # para manter o comportamento modal (bloqueia a janela principal).
    if iniciar_grab:
      self.grab_set()

    self.frame_header = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_header.pack(pady=(12, 2), padx=20, fill="x")

    self.lbl_title = ctk.CTkLabel(
        self.frame_header,
        text="✏️ Revisão de Peças",
        font=ctk.CTkFont(size=18, weight="bold"),
        text_color=COR_TITULO,
    )
    self.lbl_title.pack(side="left")

    # Botão de ajuda discreto (ícone), fora do fluxo de ações principais.
    self.btn_ajuda = ctk.CTkButton(
        self.frame_header,
        text="❓",
        width=30,
        height=30,
        corner_radius=15,
        fg_color="transparent",
        text_color=COR_TEXTO_SECUNDARIO,
        hover_color="#e5e8ea",
        font=ctk.CTkFont(size=14, weight="bold"),
        command=lambda: TourGuiado(self.parent.controller, janela_revisao=self),
    )
    self.btn_ajuda.pack(side="right")

    self.lbl_inst = ctk.CTkLabel(
        self,
        text=(
            "💡 Dê dois cliques numa célula para editar direto na tabela"
            " (Enter confirma, Esc cancela). Clique no cabeçalho para"
            " ordenar. Passe o mouse nos cabeçalhos para ver a instrução da"
            " coluna."
        ),
        text_color=COR_TEXTO_SECUNDARIO,
        font=ctk.CTkFont(size=12),
    )
    self.lbl_inst.pack(pady=(0, 4))

    self.lbl_alerta_dim = ctk.CTkLabel(
        self,
        text="",
        font=ctk.CTkFont(size=12, weight="bold"),
        text_color="#e74c3c",
    )
    self.lbl_alerta_dim.pack(pady=(0, 6))

    # Barra de ferramentas de manipulação da lista, acima da tabela (a
    # seleção em massa fica no checkbox do cabeçalho da própria tabela).
    self.frame_ferramentas = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_ferramentas.pack(pady=(0, 6), padx=20, fill="x")

    self.btn_prox_excedente = ctk.CTkButton(
        self.frame_ferramentas,
        text="⚠️ Próxima Excedente",
        fg_color="transparent",
        text_color="#e67e22",
        border_width=1,
        border_color="#e67e22",
        hover_color="#fdf1e6",
        font=ctk.CTkFont(size=11),
        height=28,
        width=160,
        command=self.ir_para_proxima_excedente,
    )
    self.btn_prox_excedente.pack(side="left", padx=(0, 5))

    self.btn_excluir = ctk.CTkButton(
        self.frame_ferramentas,
        text="🗑️ Excluir Selecionadas",
        fg_color="transparent",
        text_color="#c0392b",
        border_width=1,
        border_color="#c0392b",
        hover_color="#fbeae8",
        font=ctk.CTkFont(size=11),
        height=28,
        width=170,
        command=self.excluir_selecionadas,
        state="disabled",
    )
    self.btn_excluir.pack(side="left", padx=5)

    self.btn_desfazer = ctk.CTkButton(
        self.frame_ferramentas,
        text="↩️ Desfazer Exclusão",
        fg_color="transparent",
        text_color=COR_TEXTO_SECUNDARIO,
        border_width=1,
        border_color="#bdc3c7",
        hover_color="#e5e8ea",
        font=ctk.CTkFont(size=11),
        height=28,
        width=160,
        command=self.desfazer_exclusao,
        state="disabled",
    )
    self.btn_desfazer.pack(side="left", padx=5)

    self.lbl_dica_selecao = ctk.CTkLabel(
        self.frame_ferramentas,
        text="☐ no canto da tabela marca/desmarca todas as linhas",
        text_color=COR_TEXTO_SECUNDARIO,
        font=ctk.CTkFont(size=12),
    )
    self.lbl_dica_selecao.pack(side="right")

    # Barra de ajustes rápidos de conferência (Fundo 6mm para avanço de 8mm
    # de cada lado = +6mm, e Contra Frente / Contra Fundo de Gaveta -3mm na
    # altura). Marcar ou desmarcar atualiza a tabela imediatamente.
    self.frame_ajustes = ctk.CTkFrame(self, fg_color="#2c3e50", corner_radius=8)
    self.frame_ajustes.pack(pady=(0, 6), padx=20, fill="x")

    self.lbl_ajustes_titulo = ctk.CTkLabel(
        self.frame_ajustes,
        text="🔧 Ajustes de Conferência:",
        font=ctk.CTkFont(size=12, weight="bold"),
        text_color=COR_TITULO,
    )
    self.lbl_ajustes_titulo.pack(side="left", padx=(12, 10), pady=8)

    self.var_fundo_8mm = ctk.BooleanVar(value=False)
    self.chk_fundo_8mm = ctk.CTkCheckBox(
        self.frame_ajustes,
        text="Avanço Fundo 6mm p/ 8mm de cada lado (+6mm)",
        variable=self.var_fundo_8mm,
        font=ctk.CTkFont(size=12, weight="bold"),
        text_color=COR_TITULO,
        command=self._ao_alternar_ajustes,
    )
    self.chk_fundo_8mm.pack(side="left", padx=(4, 16), pady=8)
    criar_tooltip(
        self.chk_fundo_8mm,
        "Passa o avanço do Fundo 6mm de 5mm para 8mm de cada lado (+6mm no"
        " total): soma +6mm nas duas medidas em armários com base e chapéu"
        " (roupeiros e armários aéreos) e +6mm na largura em balcões"
        " inferiores e gavetas.",
    )

    self.var_cf_menos_3mm = ctk.BooleanVar(value=False)
    self.chk_cf_menos_3mm = ctk.CTkCheckBox(
        self.frame_ajustes,
        text="Reduzir Contra Frente / Contra Fundo 3mm na altura (-3mm)",
        variable=self.var_cf_menos_3mm,
        font=ctk.CTkFont(size=12, weight="bold"),
        text_color=COR_TITULO,
        command=self._ao_alternar_ajustes,
    )
    self.chk_cf_menos_3mm.pack(side="left", padx=4, pady=8)
    criar_tooltip(
        self.chk_cf_menos_3mm,
        "Reduz 3mm na altura (menor medida) de todas as peças de Contra"
        " Frente / Contra Fundo de Gaveta.",
    )

    self.frame_tabla = ctk.CTkFrame(self)
    self.frame_tabla.pack(padx=15, pady=5, fill="both", expand=True)

    self.columns = list(self.df.columns)
    self.tree = ttk.Treeview(
        self.frame_tabla,
        columns=self.columns,
        show="tree headings",
        height=15,
    )

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", rowheight=26, font=("Arial", 10), indent=0)
    style.configure(
        "Treeview.Heading", font=("Arial", 10, "bold"), background="#34495e"
    )

    self.tree.tag_configure(
        "excedente", background="#ffcccc", foreground="#900c3f"
    )

    self._coluna_ordenada = None
    self._ordem_crescente = True

    # Coluna implícita "#0": usada só como checkbox de seleção em massa no
    # cabeçalho (marca/desmarca todas as linhas), sem editor de célula.
    self.tree.heading(
        "#0", text="☐", command=self._alternar_checkbox_cabecalho
    )
    self.tree.column("#0", width=34, minwidth=34, stretch=False, anchor="center")

    for col in self.columns:
      self.tree.heading(
          col, text=col, command=lambda c=col: self.ordenar_por_coluna(c)
      )
      largura = (
          160
          if col
          in [
              "Material",
              "COR MDF",
              "NOME DA PEÇA",
              "Complemento",
              "SERVIÇO ADICIONAIS",
          ]
          else (70 if col in ["Quantidade", "QUANT", "Girar"] else 95)
      )
      self.tree.column(col, width=largura, anchor="center")

    vsb = ttk.Scrollbar(
        self.frame_tabla, orient="vertical", command=self.tree.yview
    )
    hsb = ttk.Scrollbar(
        self.frame_tabla, orient="horizontal", command=self.tree.xview
    )
    self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    self.tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")

    self.frame_tabla.grid_columnconfigure(0, weight=1)
    self.frame_tabla.grid_rowconfigure(0, weight=1)

    self.tree.bind("<Button-1>", self._alternar_selecao_linha)
    self.tree.bind("<Double-1>", self.on_double_click)
    self.tree.bind("<Motion>", self.on_mouse_hover)
    self.tree.bind("<<TreeviewSelect>>", self.atualizar_contagem_selecionadas)

    self.tooltip_label = None
    self._editor_inline = None
    self._ultima_exclusao = None

    # Rodapé: reservado só para o resumo de status e o fechamento do
    # formulário (cancelar / destino de salvamento / confirmar). Nenhuma
    # ação de manipulação da lista fica aqui.
    self.frame_status = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_status.pack(pady=(6, 0), padx=20, fill="x")

    self.lbl_total_pecas = ctk.CTkLabel(
        self.frame_status,
        text="",
        font=ctk.CTkFont(size=12, weight="bold"),
    )
    self.lbl_total_pecas.pack(side="left")

    self.lbl_selecionadas = ctk.CTkLabel(
        self.frame_status,
        text="",
        font=ctk.CTkFont(size=12, weight="bold"),
        text_color="#2980b9",
    )
    self.lbl_selecionadas.pack(side="right")

    self.carregar_dados_tree()

    self.frame_btns = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_btns.pack(pady=12, padx=20, fill="x")

    # Estilo discreto (texto puro) para não competir com a ação primária.
    self.btn_cancelar = ctk.CTkButton(
        self.frame_btns,
        text="Cancelar",
        fg_color="transparent",
        text_color=COR_TEXTO_SECUNDARIO,
        hover_color="#e5e8ea",
        command=self.destroy,
        width=100,
    )
    self.btn_cancelar.pack(side="left", padx=5)

    # Seleção de fábrica(s) para salvar simultaneamente, como chips dentro
    # de um contêiner próprio. Só ficam disponíveis as fábricas que usam o
    # mesmo layout de colunas que já está em revisão (TMKPlanilha tem um
    # layout próprio, então aparece sozinho nesse caso).
    self.frame_sel_forn = ctk.CTkFrame(self.frame_btns, corner_radius=8)
    self.frame_sel_forn.pack(side="left", expand=True, padx=10)

    self.lbl_forn_escolha = ctk.CTkLabel(
        self.frame_sel_forn,
        text="💾 Salvar para:",
        font=ctk.CTkFont(size=12, weight="bold"),
    )
    self.lbl_forn_escolha.pack(side="left", padx=(10, 8), pady=8)

    grupo_compativel = (
        FORNECEDORES_FORMATO_TMK
        if fornecedor_atual in FORNECEDORES_FORMATO_TMK
        else FORNECEDORES_FORMATO_PADRAO
    )
    self.fornecedores_vars = {}
    self.fornecedores_chips = {}
    for nome_forn in grupo_compativel:
      selecionado = nome_forn == fornecedor_atual
      self.fornecedores_vars[nome_forn] = selecionado
      chip = ctk.CTkButton(
          self.frame_sel_forn,
          text=nome_forn,
          corner_radius=14,
          height=26,
          width=104,
          font=ctk.CTkFont(size=11, weight="bold"),
          command=lambda n=nome_forn: self._alternar_chip_fornecedor(n),
      )
      self._estilizar_chip_fornecedor(chip, selecionado)
      chip.pack(side="left", padx=4, pady=8)
      self.fornecedores_chips[nome_forn] = chip

    self.btn_confirmar = ctk.CTkButton(
        self.frame_btns,
        text="✅ Confirmar e Salvar",
        fg_color=COR_VERDE_PRIMARIO,
        hover_color=COR_VERDE_PRIMARIO_HOVER,
        font=ctk.CTkFont(size=13, weight="bold"),
        command=self.confirmar_e_salvar,
        height=40,
        width=170,
    )
    self.btn_confirmar.pack(side="right", padx=5)

  def _indice_col_quantidade(self):
    col_nome = "QUANT" if "QUANT" in self.columns else "Quantidade"
    return self.columns.index(col_nome)

  def atualizar_contagem_total(self):
    idx_qtd = self._indice_col_quantidade()
    linhas = self.tree.get_children()
    total_pecas = 0
    for row_id in linhas:
      vals = self.tree.item(row_id, "values")
      try:
        total_pecas += int(float(vals[idx_qtd]))
      except (ValueError, TypeError, IndexError):
        pass
    self.lbl_total_pecas.configure(
        text=f"📦 Total de peças: {total_pecas}  •  {len(linhas)} linha(s)"
    )

  def atualizar_contagem_selecionadas(self, event=None):
    idx_qtd = self._indice_col_quantidade()
    selecionadas = self.tree.selection()
    total_sel = 0
    for row_id in selecionadas:
      vals = self.tree.item(row_id, "values")
      try:
        total_sel += int(float(vals[idx_qtd]))
      except (ValueError, TypeError, IndexError):
        pass
    if selecionadas:
      self.lbl_selecionadas.configure(
          text=f"🖱️ Selecionadas: {len(selecionadas)} linha(s) / {total_sel} peça(s)"
      )
    else:
      self.lbl_selecionadas.configure(text="")

    self.btn_excluir.configure(state="normal" if selecionadas else "disabled")
    self._atualizar_glifo_checkbox_cabecalho()

  def _atualizar_glifo_checkbox_cabecalho(self):
    total = len(self.tree.get_children())
    marcado = total > 0 and len(self.tree.selection()) == total
    self.tree.heading("#0", text="☑" if marcado else "☐")

  def _alternar_checkbox_cabecalho(self):
    total = self.tree.get_children()
    if total and len(self.tree.selection()) == len(total):
      self.limpar_selecao()
    else:
      self.selecionar_todas()

  def _estilizar_chip_fornecedor(self, chip, selecionado):
    if selecionado:
      chip.configure(
          fg_color="#2980b9",
          text_color="white",
          hover_color="#3498db",
          border_width=0,
      )
    else:
      chip.configure(
          fg_color="transparent",
          text_color="#2980b9",
          hover_color="#eaf2f8",
          border_width=1,
          border_color="#2980b9",
      )

  def _alternar_chip_fornecedor(self, nome_forn):
    novo_estado = not self.fornecedores_vars[nome_forn]
    self.fornecedores_vars[nome_forn] = novo_estado
    self._estilizar_chip_fornecedor(
        self.fornecedores_chips[nome_forn], novo_estado
    )

  def _alternar_selecao_linha(self, event):
    if self.tree.identify("region", event.x, event.y) not in ("cell", "tree"):
      return None

    row_id = self.tree.identify_row(event.y)
    if not row_id:
      return None

    selecionadas = self.tree.selection()
    if row_id in selecionadas and len(selecionadas) == 1:
      self.tree.selection_remove(row_id)
      return "break"

    return None

  def selecionar_todas(self):
    self.tree.selection_set(self.tree.get_children())

  def limpar_selecao(self):
    self.tree.selection_remove(self.tree.get_children())

  def ir_para_proxima_excedente(self):
    excedentes = list(self.tree.tag_has("excedente"))
    if not excedentes:
      messagebox.showinfo(
          "Sem excedentes",
          "Não há peças que ultrapassem o tamanho máximo da chapa.",
          parent=self,
      )
      return

    atual = self.tree.selection()
    try:
      pos_atual = excedentes.index(atual[0]) if atual else -1
    except ValueError:
      pos_atual = -1
    proximo = excedentes[(pos_atual + 1) % len(excedentes)]

    self.tree.selection_set(proximo)
    self.tree.see(proximo)
    self.tree.focus(proximo)

  def ordenar_por_coluna(self, col_nome):
    if self._coluna_ordenada == col_nome:
      self._ordem_crescente = not self._ordem_crescente
    else:
      self._coluna_ordenada = col_nome
      self._ordem_crescente = True

    if pd.api.types.is_numeric_dtype(self.df[col_nome]):
      chave = self.df[col_nome]
    else:
      chave = self.df[col_nome].astype(str).str.lower()

    self.df = self.df.loc[
        chave.sort_values(
            ascending=self._ordem_crescente, kind="mergesort"
        ).index
    ]
    self.carregar_dados_tree()

  def _atualizar_indicadores_cabecalho(self):
    for col in self.columns:
      texto = col
      if col == self._coluna_ordenada:
        texto += " ▲" if self._ordem_crescente else " ▼"
      self.tree.heading(col, text=texto)

  def carregar_dados_tree(self):
    for row in self.tree.get_children():
      self.tree.delete(row)

    total_excedentes = 0
    col_comp = "COMP\n2750" if "COMP\n2750" in self.df.columns else "Comprimento"
    col_larg = "LARG\n1840" if "LARG\n1840" in self.df.columns else "Largura"

    for idx, row in self.df.iterrows():
      comp_val = 0.0
      larg_val = 0.0
      try:
        comp_val = float(row.get(col_comp, 0))
      except (ValueError, TypeError):
        comp_val = 0.0
      try:
        larg_val = float(row.get(col_larg, 0))
      except (ValueError, TypeError):
        larg_val = 0.0

      eh_excedente = (
          comp_val > LIMITE_MAX_COMPRIMENTO or larg_val > LIMITE_MAX_LARGURA
      )
      tags = ("excedente",) if eh_excedente else ()
      if eh_excedente:
        total_excedentes += 1

      self.tree.insert("", "end", iid=idx, values=list(row), tags=tags)

    if total_excedentes > 0:
      self.lbl_alerta_dim.configure(
          text=(
              f"⚠️ ATENÇÃO: Há {total_excedentes} peça(s) destacada(s) em"
              f" vermelho que ultrapassam o tamanho máximo da chapa ({int(LIMITE_MAX_COMPRIMENTO)}x{int(LIMITE_MAX_LARGURA)}mm)!"
          )
      )
    else:
      self.lbl_alerta_dim.configure(text="")

    self.atualizar_contagem_total()
    self.atualizar_contagem_selecionadas()
    self._atualizar_indicadores_cabecalho()

  def on_mouse_hover(self, event):
    region = self.tree.identify("region", event.x, event.y)
    if region == "heading":
      col_id = self.tree.identify_column(event.x)
      if col_id == "#0":
        col_label = "Selecionar tudo"
        dica = "Marca ou desmarca todas as linhas da tabela"
      else:
        col_idx = int(col_id.replace("#", "")) - 1
        col_label = self.columns[col_idx].replace(chr(10), " ")
        dica = DICAS_COLUNAS.get(
            self.columns[col_idx], "Clique para editar os dados"
        )

      if self.tooltip_label is None:
        self.tooltip_label = tk.Label(
            self,
            text="",
            bg="#f1c40f",
            fg="black",
            font=("Arial", 10, "bold"),
            relief="solid",
            bd=1,
            padx=8,
            pady=4,
        )

      x_pos = self.winfo_pointerx() - self.winfo_rootx() + 15
      y_pos = self.winfo_pointery() - self.winfo_rooty() + 15
      self.tooltip_label.config(text=f"{col_label}: {dica}")
      self.tooltip_label.place(x=x_pos, y=y_pos)
      self.tooltip_label.lift()
    else:
      if self.tooltip_label is not None:
        self.tooltip_label.place_forget()

  def on_double_click(self, event):
    self._fechar_editor_inline()

    region = self.tree.identify("region", event.x, event.y)
    if region != "cell":
      return

    col = self.tree.identify_column(event.x)
    item_id = self.tree.identify_row(event.y)
    if not item_id:
      return

    col_idx = int(col.replace("#", "")) - 1
    col_nome = self.columns[col_idx]

    bbox = self.tree.bbox(item_id, col)
    if not bbox:
      return
    x, y, largura, altura = bbox

    entry = tk.Entry(self.tree, font=("Arial", 10))
    entry.insert(0, self.tree.set(item_id, col_nome))
    entry.select_range(0, "end")
    entry.place(x=x, y=y, width=largura, height=altura)
    entry.focus_set()

    entry.bind(
        "<Return>", lambda e: self._salvar_edicao_inline(item_id, col_nome)
    )
    entry.bind(
        "<KP_Enter>", lambda e: self._salvar_edicao_inline(item_id, col_nome)
    )
    entry.bind("<Escape>", lambda e: self._fechar_editor_inline())
    entry.bind(
        "<FocusOut>", lambda e: self._salvar_edicao_inline(item_id, col_nome)
    )

    self._editor_inline = entry

  def _fechar_editor_inline(self):
    entry = self._editor_inline
    self._editor_inline = None
    if entry is not None:
      try:
        entry.destroy()
      except tk.TclError:
        pass

  def _salvar_edicao_inline(self, item_id, col_nome):
    entry = self._editor_inline
    if entry is None:
      return
    try:
      val_str = entry.get().strip()
    except tk.TclError:
      return
    self._fechar_editor_inline()

    if val_str == "":
      return

    idx = int(item_id)
    dtype_col = self.df[col_nome].dtype

    try:
      if pd.api.types.is_integer_dtype(dtype_col):
        novo_valor = int(float(val_str.replace(",", ".")))
      elif pd.api.types.is_float_dtype(dtype_col):
        novo_valor = float(val_str.replace(",", "."))
      else:
        novo_valor = val_str
    except (ValueError, TypeError):
      messagebox.showerror(
          "Valor inválido",
          f"'{val_str}' não é um número válido para a coluna"
          f" '{col_nome}'.\n\nA edição foi cancelada e o valor original foi"
          " mantido.",
          parent=self,
      )
      return

    if self.df.at[idx, col_nome] == novo_valor:
      return

    self.df.at[idx, col_nome] = novo_valor
    self.carregar_dados_tree()

  def excluir_selecionadas(self):
    selecionadas = self.tree.selection()
    if not selecionadas:
      messagebox.showwarning(
          "Nenhuma peça selecionada",
          "Selecione uma ou mais peças na tabela para excluir.",
          parent=self,
      )
      return

    quantidade = len(selecionadas)
    confirmar = messagebox.askyesno(
        "Confirmar exclusão",
        f"Excluir {quantidade} peça(s) selecionada(s) da exportação?\n\n"
        "Essa ação não altera o arquivo XML original.",
        parent=self,
    )
    if confirmar:
      indices_remover = [int(row_id) for row_id in selecionadas]
      self._ultima_exclusao = self.df.loc[indices_remover].copy()
      self.df = self.df.drop(index=indices_remover)
      self.carregar_dados_tree()
      self.btn_desfazer.configure(state="normal")

      if not self.tree.get_children():
        self.lbl_alerta_dim.configure(text="Nenhuma peça será exportada.")

  def _colunas_dimensoes(self):
    col_comp = "COMP\n2750" if "COMP\n2750" in self.df.columns else "Comprimento"
    col_larg = "LARG\n1840" if "LARG\n1840" in self.df.columns else "Largura"
    return col_comp, col_larg

  def _somar_nas_dimensoes(self, idx, d_comp: float, d_larg: float):
    col_comp, col_larg = self._colunas_dimensoes()
    for col, delta in ((col_comp, d_comp), (col_larg, d_larg)):
      if not delta or col not in self.df.columns:
        continue
      val_atual = float(self.df.at[idx, col])
      novo_val = round(val_atual + delta, 4)
      if pd.api.types.is_integer_dtype(self.df[col].dtype) and novo_val.is_integer():
        self.df.at[idx, col] = int(novo_val)
      else:
        self.df.at[idx, col] = novo_val

  def _sincronizar_ajustes_no_df(self):
    fundo_ativo = bool(self.var_fundo_8mm.get())
    cf_ativo = bool(self.var_cf_menos_3mm.get())

    for idx in self.df.index:
      d_comp_f, d_larg_f = self._deltas_fundo.get(idx, (0.0, 0.0))
      if d_comp_f or d_larg_f:
        aplicado_f = self._fundo_aplicado.get(idx, False)
        if fundo_ativo and not aplicado_f:
          self._somar_nas_dimensoes(idx, d_comp_f, d_larg_f)
          self._fundo_aplicado[idx] = True
        elif not fundo_ativo and aplicado_f:
          self._somar_nas_dimensoes(idx, -d_comp_f, -d_larg_f)
          self._fundo_aplicado[idx] = False

      d_comp_cf, d_larg_cf = self._deltas_cf.get(idx, (0.0, 0.0))
      if d_comp_cf or d_larg_cf:
        aplicado_cf = self._cf_aplicado.get(idx, False)
        if cf_ativo and not aplicado_cf:
          self._somar_nas_dimensoes(idx, d_comp_cf, d_larg_cf)
          self._cf_aplicado[idx] = True
        elif not cf_ativo and aplicado_cf:
          self._somar_nas_dimensoes(idx, -d_comp_cf, -d_larg_cf)
          self._cf_aplicado[idx] = False

  def _ao_alternar_ajustes(self):
    self._sincronizar_ajustes_no_df()
    self.carregar_dados_tree()

  def desfazer_exclusao(self):
    if self._ultima_exclusao is None or self._ultima_exclusao.empty:
      return
    self.df = pd.concat([self.df, self._ultima_exclusao]).sort_index()
    self._ultima_exclusao = None
    self.btn_desfazer.configure(state="disabled")
    self._sincronizar_ajustes_no_df()
    self.carregar_dados_tree()

  def confirmar_e_salvar(self):
    novos_dados = []
    for row_id in self.tree.get_children():
      vals = self.tree.item(row_id, "values")
      novos_dados.append(vals)

    df_editado = pd.DataFrame(novos_dados, columns=self.columns)
    if df_editado.empty:
      messagebox.showwarning(
          "Nenhuma peça para salvar",
          "Mantenha pelo menos uma peça na revisão antes de salvar.",
          parent=self,
      )
      return

    fornecedores_selecionados = [
        nome for nome, selecionado in self.fornecedores_vars.items() if selecionado
    ]
    if not fornecedores_selecionados:
      messagebox.showwarning(
          "Nenhuma fábrica selecionada",
          "Marque ao menos uma fábrica em '💾 Salvar para:' antes de salvar.",
          parent=self,
      )
      return

    self.destroy()
    self.callback_salvar(df_editado, fornecedores_selecionados)

