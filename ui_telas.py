"""Telas de navegação principal: TelaInicio, TelaNovoProjeto e TelaConfiguracaoProjeto."""
import os
import re
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
import zipfile
import customtkinter as ctk
import openpyxl
import pandas as pd

from constantes import (
    COR_BORDA_INPUT,
    COR_TEXTO_SECUNDARIO,
    COR_TITULO,
    COR_VERDE_PRIMARIO,
    COR_VERDE_PRIMARIO_HOVER,
    FORNECEDORES_DISPONIVEIS,
    FUNCOES_AFETADAS_TERCEIRIZACAO,
)
from regras_marcenaria import (
    calcular_delta_fundo_xml,
    deve_ignorar_item,
    extrair_cliente_ambiente_xml,
    extrair_espessura_fita,
    extrair_fabricantes_do_xml,
    extrair_modulo_pai,
    extrair_referencias,
    extrair_servicos_adicionais,
    formatar_material_com_espessura,
    listar_modulos_canto_l,
    mapear_contexto_modulos_xml,
    renomear_funcao,
)
from ui_componentes import aplicar_estilo_input, criar_tooltip
from ui_janelas import JanelaModulosCantoL, JanelaRevisao


class TelaInicio(ctk.CTkFrame):

  def __init__(self, parent, controller):
    super().__init__(parent, fg_color="transparent")
    self.controller = controller

    self.lbl_title = ctk.CTkLabel(
        self,
        text="Conversor de XML Promob / Start KNR",
        font=ctk.CTkFont(size=24, weight="bold"),
        text_color=COR_TITULO,
    )
    self.lbl_title.pack(pady=(20, 2))

    self.lbl_sub = ctk.CTkLabel(
        self,
        text="Gere planos de corte prontos pra fábrica a partir do XML.",
        font=ctk.CTkFont(size=13),
        text_color=COR_TEXTO_SECUNDARIO,
    )
    self.lbl_sub.pack(pady=(0, 10))

    self.btn_novo = ctk.CTkButton(
        self,
        text="➕ Criar Novo Projeto",
        font=ctk.CTkFont(size=16, weight="bold"),
        fg_color=COR_VERDE_PRIMARIO,
        hover_color=COR_VERDE_PRIMARIO_HOVER,
        height=48,
        width=320,
        command=self.criar_novo_projeto,
    )
    self.btn_novo.pack(pady=10)

    self.frame_recentes = ctk.CTkFrame(self)
    self.frame_recentes.pack(
        padx=20, pady=(10, 10), fill="both", expand=True
    )

    self.lbl_recentes_title = ctk.CTkLabel(
        self.frame_recentes,
        text="📋 Histórico de Projetos",
        font=ctk.CTkFont(size=14, weight="bold"),
        text_color=COR_TITULO,
        anchor="w",
    )
    self.lbl_recentes_title.pack(padx=15, pady=(10, 5), fill="x")

    # Cabeçalho de colunas do "grid" de histórico — as linhas em si são
    # montadas em atualizar_interface_historico(), usando grid() nesse
    # mesmo frame rolável pra alinhar as 3 colunas com o cabeçalho.
    self.frame_cabecalho_hist = ctk.CTkFrame(
        self.frame_recentes, fg_color="transparent"
    )
    self.frame_cabecalho_hist.pack(padx=15, fill="x")
    self.frame_cabecalho_hist.grid_columnconfigure(0, weight=3)
    self.frame_cabecalho_hist.grid_columnconfigure(1, weight=1)
    self.frame_cabecalho_hist.grid_columnconfigure(2, weight=0)

    for col, texto in enumerate(["Projeto / Arquivo", "Formato", "Ações"]):
      ctk.CTkLabel(
          self.frame_cabecalho_hist,
          text=texto,
          font=ctk.CTkFont(size=11, weight="bold"),
          text_color=COR_TEXTO_SECUNDARIO,
          anchor="w",
      ).grid(row=0, column=col, sticky="w", padx=6, pady=(2, 4))

    self.scroll_recentes = ctk.CTkScrollableFrame(self.frame_recentes)
    self.scroll_recentes.pack(padx=10, pady=(0, 10), fill="both", expand=True)
    self.scroll_recentes.grid_columnconfigure(0, weight=3)
    self.scroll_recentes.grid_columnconfigure(1, weight=1)
    self.scroll_recentes.grid_columnconfigure(2, weight=0)

    self.frame_backup_inferior = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_backup_inferior.pack(padx=20, pady=(5, 10), fill="x")

    # Backup/restaurar são ações de manutenção, não do fluxo principal —
    # estilo outlined pra não competir visualmente com "Criar Novo Projeto".
    self.btn_backup = ctk.CTkButton(
        self.frame_backup_inferior,
        text="☁️ Fazer Backup dos Dados",
        font=ctk.CTkFont(size=12, weight="bold"),
        fg_color="transparent",
        text_color="#5dade2",
        border_width=1,
        border_color="#5dade2",
        hover_color="#1b2631",
        height=38,
        command=self.fazer_backup,
    )
    self.btn_backup.pack(side="left", expand=True, fill="x", padx=5)

    self.btn_restaurar = ctk.CTkButton(
        self.frame_backup_inferior,
        text="📂 Restaurar Dados",
        font=ctk.CTkFont(size=12, weight="bold"),
        fg_color="transparent",
        text_color="#bb8fce",
        border_width=1,
        border_color="#bb8fce",
        hover_color="#241b2b",
        height=38,
        command=self.restaurar_backup,
    )
    self.btn_restaurar.pack(side="right", expand=True, fill="x", padx=5)

  def criar_novo_projeto(self):
    self.controller.cliente_nome = ""
    self.controller.ambiente_nome = ""
    self.controller.xml_path = None
    self.controller.show_frame("TelaNovoProjeto")

  def editar_projeto_historico(self, item_hist):
    self.controller.cliente_nome = item_hist.get("cliente", "")
    self.controller.ambiente_nome = item_hist.get("ambiente", "")
    self.controller.fornecedor_inicial = item_hist.get(
        "fornecedor", "Teletintas"
    )

    xml_antigo = item_hist.get("caminho_xml", "")
    if os.path.exists(xml_antigo):
      self.controller.xml_path = xml_antigo
      self.controller.pasta_destino = os.path.dirname(xml_antigo)
      self.controller.show_frame("TelaConfiguracaoProjeto")
      self.controller.frames["TelaConfiguracaoProjeto"].iniciar_revisao()
    else:
      self.controller.xml_path = None
      messagebox.showwarning(
          "Arquivo não encontrado",
          f"O arquivo XML original não foi encontrado em:\n{xml_antigo}\n\nPor"
          " favor, selecione o XML novamente.",
      )
      self.atualizar_interface_historico()
      self.controller.show_frame("TelaConfiguracaoProjeto")

  def fazer_backup(self):
    try:
      if hasattr(self.controller, "carregar_historico"):
        self.controller.historico = self.controller.carregar_historico()
        self.controller.salvar_historico()
      historico_path = self.controller.arquivo_historico
      if not os.path.exists(historico_path) or not self.controller.historico:
        messagebox.showwarning(
            "Aviso", "Não há dados salvos para fazer backup ainda."
        )
        return

      destino_zip = filedialog.asksaveasfilename(
          title="Salvar Backup do Sistema",
          initialfile="backup_conversor_promob.zip",
          defaultextension=".zip",
          filetypes=[("Arquivo Zip", "*.zip")],
      )

      if destino_zip:
        with zipfile.ZipFile(destino_zip, "w", zipfile.ZIP_DEFLATED) as zf:
          zf.write(historico_path, arcname="historico_recentes.json")
        messagebox.showinfo(
            "Sucesso",
            f"Backup gerado com sucesso!\n\nVocê pode enviar este arquivo"
            f" `.zip` para o seu e-mail ou nuvem:\n{destino_zip}",
        )
    except Exception as e:
      messagebox.showerror("Erro ao Gerar Backup", f"Erro:\n{e}")

  def restaurar_backup(self):
    try:
      arquivo_zip = filedialog.askopenfilename(
          title="Selecione o arquivo de Backup (.zip)",
          filetypes=[("Arquivo Zip", "*.zip")],
      )

      if arquivo_zip:
        pasta_temp = os.path.dirname(self.controller.arquivo_historico)
        with zipfile.ZipFile(arquivo_zip, "r") as zf:
          zf.extract("historico_recentes.json", path=pasta_temp)

        self.controller.historico = self.controller.carregar_historico()
        if hasattr(self.controller, "salvar_historico"):
          self.controller.salvar_historico()
        self.atualizar_interface_historico()
        messagebox.showinfo(
            "Sucesso", "Backup restaurado com sucesso! Seus projetos foram"
            " carregados."
        )
    except Exception as e:
      messagebox.showerror(
          "Erro ao Restaurar",
          f"O arquivo selecionado é inválido ou corrompido:\n{e}",
      )

  def ao_exibir_tela(self):
    self.atualizar_interface_historico()

  def atualizar_interface_historico(self):
    for widget in self.scroll_recentes.winfo_children():
      widget.destroy()

    if hasattr(self.controller, "carregar_historico"):
      self.controller.historico = self.controller.carregar_historico()

    # Um projeto só entra na lista se o XML original e a pasta do Excel
    # exportado ainda existirem — caso contrário "Abrir" e "Editar" não
    # teriam como funcionar, então a linha só confundiria.
    historico = [
        item
        for item in self.controller.historico
        if os.path.exists(item.get("caminho_xml", ""))
        and os.path.exists(os.path.dirname(item.get("caminho_excel", "")))
    ]
    if not historico:
      lbl_vazio = ctk.CTkLabel(
          self.scroll_recentes,
          text="Nenhum projeto registrado no histórico.",
          text_color=COR_TEXTO_SECUNDARIO,
      )
      lbl_vazio.pack(pady=20)
      return

    for linha, item in enumerate(historico):
      cli_str = f"{item.get('cliente', 'Cliente')} - {item.get('ambiente', 'Ambiente')}"
      nomes_arquivo = (
          f"📁 {item['nome_xml']}  ➔  📊"
          f" {os.path.basename(item['caminho_excel'])}"
      )
      data_hora = item.get("data_hora", "")

      frame_projeto = ctk.CTkFrame(self.scroll_recentes, fg_color="transparent")
      frame_projeto.grid(row=linha, column=0, sticky="ew", padx=6, pady=6)

      ctk.CTkLabel(
          frame_projeto,
          text=cli_str,
          font=ctk.CTkFont(size=12, weight="bold"),
          text_color=COR_TITULO,
          anchor="w",
      ).pack(fill="x")
      ctk.CTkLabel(
          frame_projeto,
          text=nomes_arquivo + (f"   🕒 {data_hora}" if data_hora else ""),
          font=ctk.CTkFont(size=11),
          text_color=COR_TEXTO_SECUNDARIO,
          anchor="w",
      ).pack(fill="x")

      ctk.CTkLabel(
          self.scroll_recentes,
          text=item.get("fornecedor", "Teletintas"),
          font=ctk.CTkFont(size=11),
          text_color=COR_TEXTO_SECUNDARIO,
          anchor="w",
      ).grid(row=linha, column=1, sticky="w", padx=6, pady=6)

      frame_acoes = ctk.CTkFrame(self.scroll_recentes, fg_color="transparent")
      frame_acoes.grid(row=linha, column=2, sticky="e", padx=6, pady=6)

      caminho_pasta = os.path.dirname(item["caminho_excel"])

      btn_abrir_pasta = ctk.CTkButton(
          frame_acoes,
          text="[ Abrir ]",
          width=80,
          height=26,
          font=ctk.CTkFont(size=11),
          fg_color="#4b6584",
          hover_color="#3867d6",
          command=lambda p=caminho_pasta: self._abrir_pasta_projeto(p),
      )
      btn_abrir_pasta.pack(side="left", padx=(0, 4))

      btn_editar = ctk.CTkButton(
          frame_acoes,
          text="[ Editar ]",
          width=80,
          height=26,
          font=ctk.CTkFont(size=11),
          fg_color="#2980b9",
          hover_color="#3498db",
          command=lambda it=item: self.editar_projeto_historico(it),
      )
      btn_editar.pack(side="left")

  def _abrir_pasta_projeto(self, caminho_pasta):
    if not os.path.exists(caminho_pasta):
      messagebox.showwarning(
          "Pasta não encontrada",
          f"A pasta do projeto não foi encontrada:\n{caminho_pasta}",
      )
      self.atualizar_interface_historico()
      return
    if os.name == "nt":
      os.startfile(caminho_pasta)
    else:
      os.system(f'open "{caminho_pasta}"')


class TelaNovoProjeto(ctk.CTkFrame):

  def __init__(self, parent, controller):
    super().__init__(parent, fg_color="transparent")
    self.controller = controller

    self.lbl_title = ctk.CTkLabel(
        self,
        text="📝 Cadastro do Projeto",
        font=ctk.CTkFont(size=20, weight="bold"),
        text_color=COR_TITULO,
    )
    self.lbl_title.pack(pady=(30, 5))

    self.lbl_sub = ctk.CTkLabel(
        self,
        text="Confirme ou altere o nome do cliente e do ambiente.",
        text_color=COR_TEXTO_SECUNDARIO,
    )
    self.lbl_sub.pack(pady=(0, 25))

    self.frame_form = ctk.CTkFrame(self, width=500)
    self.frame_form.pack(padx=30, pady=10, fill="x")

    self.lbl_cliente = ctk.CTkLabel(
        self.frame_form,
        text="Nome do Cliente:",
        font=ctk.CTkFont(size=13, weight="bold"),
    )
    self.lbl_cliente.pack(anchor="w", padx=20, pady=(20, 2))

    self.ent_cliente = ctk.CTkEntry(
        self.frame_form,
        placeholder_text="Ex: Raynnara Silveira Borges",
        height=40,
        font=ctk.CTkFont(size=13),
    )
    self.ent_cliente.pack(fill="x", padx=20, pady=(0, 15))
    aplicar_estilo_input(self.ent_cliente)

    self.lbl_ambiente = ctk.CTkLabel(
        self.frame_form,
        text="Nome do Ambiente:",
        font=ctk.CTkFont(size=13, weight="bold"),
    )
    self.lbl_ambiente.pack(anchor="w", padx=20, pady=(5, 2))

    self.ent_ambiente = ctk.CTkEntry(
        self.frame_form,
        placeholder_text="Ex: Escritório ou Cozinha",
        height=40,
        font=ctk.CTkFont(size=13),
    )
    self.ent_ambiente.pack(fill="x", padx=20, pady=(0, 25))
    aplicar_estilo_input(self.ent_ambiente)

    self.frame_nav = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_nav.pack(pady=25, fill="x", padx=30)

    self.btn_voltar = ctk.CTkButton(
        self.frame_nav,
        text="⬅️ Voltar",
        width=120,
        height=42,
        fg_color="transparent",
        text_color=COR_TEXTO_SECUNDARIO,
        border_width=1,
        border_color=COR_BORDA_INPUT,
        hover_color="#333333",
        command=lambda: controller.show_frame("TelaInicio"),
    )
    self.btn_voltar.pack(side="left")

    self.btn_avancar = ctk.CTkButton(
        self.frame_nav,
        text="Avançar ➔",
        width=160,
        height=42,
        font=ctk.CTkFont(size=14, weight="bold"),
        fg_color="#2980b9",
        hover_color="#3498db",
        command=self.avancar,
    )
    self.btn_avancar.pack(side="right")

  def ao_exibir_tela(self):
    self.ent_cliente.delete(0, "end")
    self.ent_ambiente.delete(0, "end")
    if self.controller.cliente_nome:
      self.ent_cliente.insert(0, self.controller.cliente_nome)
    if self.controller.ambiente_nome:
      self.ent_ambiente.insert(0, self.controller.ambiente_nome)

  def avancar(self):
    cliente = self.ent_cliente.get().strip()
    ambiente = self.ent_ambiente.get().strip()
    if not cliente or not ambiente:
      messagebox.showwarning(
          "Campos Vazios", "Por favor, preencha o Cliente e o Ambiente."
      )
      return
    self.controller.cliente_nome = cliente
    self.controller.ambiente_nome = ambiente
    self.controller.show_frame("TelaConfiguracaoProjeto")


class TelaConfiguracaoProjeto(ctk.CTkFrame):

  def __init__(self, parent, controller):
    super().__init__(parent, fg_color="transparent")
    self.controller = controller

    # Card de contexto: cliente/ambiente + edição, num único elemento.
    self.frame_header_info = ctk.CTkFrame(self, fg_color="#2c3e50")
    self.frame_header_info.pack(padx=20, pady=(10, 15), fill="x")

    # Rótulos ("Cliente:", "Ambiente:", a barra divisória) em negrito; os
    # valores em si (nome do cliente/ambiente) em peso regular — precisa de
    # labels separados porque um único CTkLabel não mistura pesos de fonte.
    self.frame_info_texto = ctk.CTkFrame(
        self.frame_header_info, fg_color="transparent"
    )
    self.frame_info_texto.pack(side="left", padx=15, pady=12)

    fonte_rotulo = ctk.CTkFont(size=14, weight="bold")
    fonte_valor = ctk.CTkFont(size=14, weight="normal")

    ctk.CTkLabel(
        self.frame_info_texto, text="Cliente:", font=fonte_rotulo,
        text_color=COR_TITULO,
    ).pack(side="left")
    self.lbl_cliente_valor = ctk.CTkLabel(
        self.frame_info_texto, text="-", font=fonte_valor,
        text_color=COR_TITULO,
    )
    self.lbl_cliente_valor.pack(side="left", padx=(4, 12))
    ctk.CTkLabel(
        self.frame_info_texto, text="|", font=fonte_rotulo,
        text_color=COR_TITULO,
    ).pack(side="left", padx=(0, 12))
    ctk.CTkLabel(
        self.frame_info_texto, text="Ambiente:", font=fonte_rotulo,
        text_color=COR_TITULO,
    ).pack(side="left")
    self.lbl_ambiente_valor = ctk.CTkLabel(
        self.frame_info_texto, text="-", font=fonte_valor,
        text_color=COR_TITULO,
    )
    self.lbl_ambiente_valor.pack(side="left", padx=(4, 0))

    self.btn_editar_cad = ctk.CTkButton(
        self.frame_header_info,
        text="Editar",
        width=70,
        height=34,
        corner_radius=17,
        fg_color="transparent",
        text_color=COR_TITULO,
        border_width=1,
        border_color="#5c6f80",
        hover_color="#34495e",
        font=ctk.CTkFont(size=13, weight="bold"),
        command=lambda: controller.show_frame("TelaNovoProjeto"),
    )
    self.btn_editar_cad.pack(side="right", padx=12, pady=8)
    criar_tooltip(self.btn_editar_cad, "Alterar Cliente / Ambiente")

    # Passo 1: XML — campo padronizado (entrada somente-leitura + botão).
    self.ent_xml_path = self._criar_campo_arquivo(
        "1. Arquivo XML", "Procurar...", self.selecionar_xml
    )
    self.btn_select_xml = self.ent_xml_path.botao_associado

    # Passo 2: Fábrica — a explicação de compatibilidade vira um ícone de
    # info ao lado do label, em vez de um parágrafo cinza sempre visível.
    self.frame_forn = ctk.CTkFrame(self)
    self.frame_forn.pack(padx=20, pady=6, fill="x")

    frame_forn_label = ctk.CTkFrame(self.frame_forn, fg_color="transparent")
    frame_forn_label.pack(side="left", padx=(12, 4), pady=10)

    self.lbl_forn = ctk.CTkLabel(
        frame_forn_label,
        text="2. Formato Inicial de Fábrica",
        font=ctk.CTkFont(size=13, weight="bold"),
    )
    self.lbl_forn.pack(side="left")

    self.lbl_info_forn = ctk.CTkLabel(
        frame_forn_label,
        text="ⓘ",
        font=ctk.CTkFont(size=14, weight="bold"),
        text_color="#2980b9",
        cursor="hand2",
    )
    self.lbl_info_forn.pack(side="left", padx=(6, 0))
    criar_tooltip(
        self.lbl_info_forn,
        "Teletintas, Madefer, Tobias e TMKCloud usam a mesma planilha e"
        " poderão ser marcadas juntas na revisão. TMKPlanilha usa um formato"
        " de colunas próprio e fica disponível sozinha.",
    )

    self.fornecedor_var = ctk.StringVar(value="Teletintas")
    self.cmb_fornecedor = ctk.CTkOptionMenu(
        self.frame_forn,
        values=FORNECEDORES_DISPONIVEIS,
        variable=self.fornecedor_var,
        width=180,
        height=34,
    )
    self.cmb_fornecedor.pack(side="left", padx=10, pady=10)

    # Passo 3: pasta de destino — mesmo padrão de campo do passo 1 — com a
    # opção de agrupar logo abaixo, coladas como um só grupo de decisão.
    self.ent_pasta_destino = self._criar_campo_arquivo(
        "3. Pasta de Destino", "Procurar...", self.selecionar_pasta_destino
    )
    self.btn_select_folder = self.ent_pasta_destino.botao_associado
    self._definir_texto_campo(
        self.ent_pasta_destino,
        "Salvar na mesma pasta do XML (Padrão)",
        cor=COR_TEXTO_SECUNDARIO,
    )

    self.chk_agrupar_var = ctk.BooleanVar(value=False)
    self.chk_agrupar = ctk.CTkCheckBox(
        self,
        text="Agrupar peças iguais (soma quantidades)",
        variable=self.chk_agrupar_var,
        font=ctk.CTkFont(size=12, weight="bold"),
    )
    self.chk_agrupar.pack(pady=(2, 22), padx=25, anchor="w")

    # Rodapé: só a navegação (Voltar, discreto) e a ação primária (verde).
    self.frame_rodape = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_rodape.pack(padx=20, pady=(0, 15), fill="x")

    self.btn_voltar = ctk.CTkButton(
        self.frame_rodape,
        text="⬅️ Voltar",
        width=120,
        height=42,
        fg_color="transparent",
        text_color=("gray20", "gray85"),
        border_width=1,
        border_color="#7f8c8d",
        hover_color="#e5e8ea",
        font=ctk.CTkFont(size=13, weight="bold"),
        command=lambda: controller.show_frame("TelaNovoProjeto"),
    )
    self.btn_voltar.pack(side="left")

    self.btn_convert = ctk.CTkButton(
        self.frame_rodape,
        text="Revisar e Exportar para Excel",
        command=self.iniciar_revisao,
        fg_color=COR_VERDE_PRIMARIO,
        hover_color=COR_VERDE_PRIMARIO_HOVER,
        # Sem isso, o CTk usa um cinza apagado (tema padrão) pro texto
        # quando state="disabled", fazendo o botão parecer de outra cor.
        text_color_disabled=COR_TITULO,
        height=42,
        font=ctk.CTkFont(size=14, weight="bold"),
        state="disabled",
    )
    self.btn_convert.pack(side="left", fill="x", expand=True, padx=(12, 0))

  def _criar_campo_arquivo(self, texto_passo, texto_botao, comando):
    """Linha padrão pra escolha de arquivo/pasta: legenda + campo
    somente-leitura mostrando o caminho atual + botão lateral pra trocar."""
    frame = ctk.CTkFrame(self)
    frame.pack(padx=20, pady=6, fill="x")

    ctk.CTkLabel(
        frame,
        text=texto_passo,
        font=ctk.CTkFont(size=12, weight="bold"),
        text_color=COR_TEXTO_SECUNDARIO,
        anchor="w",
    ).pack(anchor="w", padx=12, pady=(10, 0))

    frame_linha = ctk.CTkFrame(frame, fg_color="transparent")
    frame_linha.pack(fill="x", padx=12, pady=(4, 10))

    entrada = ctk.CTkEntry(frame_linha, state="disabled", height=36)
    entrada.pack(side="left", fill="x", expand=True, padx=(0, 8))

    botao = ctk.CTkButton(
        frame_linha,
        text=texto_botao,
        command=comando,
        width=120,
        height=36,
        fg_color="#34495e",
        hover_color="#2c3e50",
        font=ctk.CTkFont(size=12, weight="bold"),
    )
    botao.pack(side="left")
    entrada.botao_associado = botao
    return entrada

  @staticmethod
  def _definir_texto_campo(entrada, texto, cor=None):
    entrada.configure(state="normal")
    entrada.delete(0, "end")
    entrada.insert(0, texto)
    if cor is not None:
      entrada.configure(text_color=cor)
    entrada.configure(state="disabled")

  def ao_exibir_tela(self):
    cli = self.controller.cliente_nome or "Não informado"
    amb = self.controller.ambiente_nome or "Não informado"
    self.lbl_cliente_valor.configure(text=cli)
    self.lbl_ambiente_valor.configure(text=amb)
    self.fornecedor_var.set(
        getattr(self.controller, "fornecedor_inicial", "Teletintas")
    )

    if self.controller.xml_path and os.path.exists(self.controller.xml_path):
      filename = os.path.basename(self.controller.xml_path)
      self._definir_texto_campo(self.ent_xml_path, f"📄 {filename}", cor="#2980b9")
      self.ent_xml_path.botao_associado.configure(text="Alterar")
      self.btn_convert.configure(state="normal")
    else:
      self._definir_texto_campo(
          self.ent_xml_path, "Nenhum arquivo XML selecionado", cor=COR_TEXTO_SECUNDARIO
      )
      self.ent_xml_path.botao_associado.configure(text="Procurar...")
      self.btn_convert.configure(state="disabled")

    if self.controller.pasta_destino:
      self._definir_texto_campo(
          self.ent_pasta_destino, f"📁 {self.controller.pasta_destino}", cor=COR_TEXTO_SECUNDARIO
      )
      self.ent_pasta_destino.botao_associado.configure(text="Alterar")
    else:
      self._definir_texto_campo(
          self.ent_pasta_destino,
          "Salvar na mesma pasta do XML (Padrão)",
          cor=COR_TEXTO_SECUNDARIO,
      )
      self.ent_pasta_destino.botao_associado.configure(text="Procurar...")

  def selecionar_xml(self):
    filepath = filedialog.askopenfilename(
        title="Selecione o arquivo XML",
        filetypes=[("Arquivos XML", "*.xml"), ("Todos os Arquivos", "*.*")],
    )
    if filepath:
      self.controller.xml_path = filepath
      filename = os.path.basename(filepath)
      self._definir_texto_campo(self.ent_xml_path, f"📄 {filename}", cor="#2980b9")
      self.ent_xml_path.botao_associado.configure(text="Alterar")

      if not self.controller.pasta_destino:
        self.controller.pasta_destino = os.path.dirname(filepath)
        self._definir_texto_campo(
            self.ent_pasta_destino,
            f"📁 {self.controller.pasta_destino}",
            cor=COR_TEXTO_SECUNDARIO,
        )
        self.ent_pasta_destino.botao_associado.configure(text="Alterar")

      cli_xml, amb_xml = extrair_cliente_ambiente_xml(filepath)
      if not self.controller.cliente_nome and cli_xml:
        self.controller.cliente_nome = cli_xml
      if not self.controller.ambiente_nome and amb_xml:
        self.controller.ambiente_nome = amb_xml

      cli = self.controller.cliente_nome or "Não informado"
      amb = self.controller.ambiente_nome or "Não informado"
      self.lbl_cliente_valor.configure(text=cli)
      self.lbl_ambiente_valor.configure(text=amb)

      self.btn_convert.configure(state="normal")

  def selecionar_pasta_destino(self):
    folderpath = filedialog.askdirectory(
        title="Selecione a pasta onde deseja salvar o arquivo Excel"
    )
    if folderpath:
      self.controller.pasta_destino = folderpath
      self._definir_texto_campo(
          self.ent_pasta_destino, f"📁 {folderpath}", cor="#27ae60"
      )
      self.ent_pasta_destino.botao_associado.configure(text="Alterar")

  def processar_dados_para_revisao(
      self, xml_path, fornecedor_alvo="Teletintas", modulos_terceirizados=None
  ):
    modulos_terceirizados = modulos_terceirizados or {}
    tree = ET.parse(xml_path)
    root = tree.getroot()

    ambients_section = root.find("AMBIENTS")
    items_source = (
        ambients_section.findall(".//ITEM")
        if ambients_section is not None
        else root.findall(".//ITEM")
    )

    parent_map = {c: p for p in root.iter() for c in p}
    fab_map = extrair_fabricantes_do_xml(root)
    contexto_modulos = mapear_contexto_modulos_xml(items_source, parent_map)

    cli_xml, amb_xml = extrair_cliente_ambiente_xml(xml_path)
    nome_cliente = self.controller.cliente_nome or cli_xml or "Cliente"
    nome_ambiente = self.controller.ambiente_nome or amb_xml or "Ambiente"

    rows_padrao = []
    rows_tmk = []
    lista_deltas_fundo = []
    item_num_tmk = 1

    for item in items_source:
      refs = extrair_referencias(item)
      if deve_ignorar_item(item, refs):
        continue

      desc = item.attrib.get("DESCRIPTION", "").strip()
      material = refs.get("MATERIAL", "")

      funcao_formatada = renomear_funcao(desc)
      material_formatado = formatar_material_com_espessura(item, refs, fab_map)
      modulo_nome = extrair_modulo_pai(item, parent_map)
      servico_adic = extrair_servicos_adicionais(item, desc)

      bc1 = refs.get("BC1", "")
      bc2 = refs.get("BC2", "")
      bl1 = refs.get("BL1", "")
      bl2 = refs.get("BL2", "")

      comprimento = float(item.attrib.get("WIDTH", "0"))
      largura = float(item.attrib.get("DEPTH", "0"))
      margem_terceirizacao = modulos_terceirizados.get(modulo_nome, 0)
      if margem_terceirizacao and funcao_formatada in FUNCOES_AFETADAS_TERCEIRIZACAO:
        comprimento += margem_terceirizacao
        largura += margem_terceirizacao

      delta_fundo = calcular_delta_fundo_xml(
          item, refs, parent_map, contexto_modulos
      )
      lista_deltas_fundo.append(delta_fundo)

      row_p = {
          "Quantidade": int(float(item.attrib.get("QUANTITY", "1"))),
          "Comprimento": comprimento,
          "Largura": largura,
          "Função": funcao_formatada,
          "Fita C1": bc1,
          "Fita C2": bc2,
          "Fita L1": bl1,
          "Fita L2": bl2,
          "Material": material_formatado,
          "Complemento": modulo_nome,
          "Girar": "SIM"
          if item.attrib.get("PLATECUTTINGROTATE", "NONE") != "NONE"
          else "NÃO",
      }
      rows_padrao.append(row_p)

      cor_mdf = refs.get("MODEL", "").strip()
      if not cor_mdf or cor_mdf.lower() == "unica":
        cor_mdf = material

      height_str = item.attrib.get("HEIGHT", "").strip()
      esp_mdf = ""
      if height_str:
        try:
          h_val = float(height_str)
          esp_mdf = f"{int(h_val) if h_val.is_integer() else h_val}mm"
        except ValueError:
          esp_mdf = ""

      esp_fita = extrair_espessura_fita(bc1, bc2, bl1, bl2)

      row_t = {
          "ITEM": item_num_tmk,
          "QUANT": int(float(item.attrib.get("QUANTITY", "1"))),
          "COMP\n2750": comprimento,
          "LARG\n1840": largura,
          "NOME DA PEÇA": funcao_formatada,
          "SERVIÇO ADICIONAIS": servico_adic,
          "AMBIENTE": nome_ambiente,
          "CLIENTE": nome_cliente,
          "OBS 1": modulo_nome,
          "OBS 2": "",
          "FITA \n1ª COMP ": "x" if bc1 else "",
          "FITA \n2ª COMP ": "x" if bc2 else "",
          "FITA \n1ª LARG": "x" if bl1 else "",
          "FITA \n2ª LARG": "x" if bl2 else "",
          "COR MDF": cor_mdf,
          "ESP MDF": esp_mdf,
          "ESP FITA": esp_fita,
      }
      rows_tmk.append(row_t)
      item_num_tmk += 1

    if fornecedor_alvo == "TMKPlanilha":
      colunas_tmk = [
          "ITEM",
          "QUANT",
          "COMP\n2750",
          "LARG\n1840",
          "NOME DA PEÇA",
          "SERVIÇO ADICIONAIS",
          "AMBIENTE",
          "CLIENTE",
          "OBS 1",
          "OBS 2",
          "FITA \n1ª COMP ",
          "FITA \n2ª COMP ",
          "FITA \n1ª LARG",
          "FITA \n2ª LARG",
          "COR MDF",
          "ESP MDF",
          "ESP FITA",
      ]
      df = pd.DataFrame(rows_tmk, columns=colunas_tmk)
    else:
      colunas_p = [
          "Quantidade",
          "Comprimento",
          "Largura",
          "Função",
          "Fita C1",
          "Fita C2",
          "Fita L1",
          "Fita L2",
          "Material",
          "Complemento",
          "Girar",
      ]
      df = pd.DataFrame(rows_padrao, columns=colunas_p)

    deltas_por_idx = {
        idx: delta
        for idx, delta in enumerate(lista_deltas_fundo)
        if delta != (0.0, 0.0)
    }

    if self.chk_agrupar_var.get() and not df.empty:
      resposta = messagebox.askyesno(
          "Confirmar Agrupamento",
          "A opção de agrupar peças iguais está marcada.\n\nDeseja realmente"
          " agrupar as peças idênticas e somar suas quantidades?",
      )
      if resposta:
        df["_delta_comp_f"] = [d[0] for d in lista_deltas_fundo]
        df["_delta_larg_f"] = [d[1] for d in lista_deltas_fundo]
        if fornecedor_alvo == "TMKPlanilha":
          cols_agrupar = [
              "COMP\n2750",
              "LARG\n1840",
              "NOME DA PEÇA",
              "SERVIÇO ADICIONAIS",
              "AMBIENTE",
              "CLIENTE",
              "OBS 1",
              "OBS 2",
              "FITA \n1ª COMP ",
              "FITA \n2ª COMP ",
              "FITA \n1ª LARG",
              "FITA \n2ª LARG",
              "COR MDF",
              "ESP MDF",
              "ESP FITA",
              "_delta_comp_f",
              "_delta_larg_f",
          ]
          df = (
              df.groupby(cols_agrupar, as_index=False, dropna=False)[
                  "QUANT"
              ].sum()
          )
          deltas_por_idx = {
              idx: (float(row["_delta_comp_f"]), float(row["_delta_larg_f"]))
              for idx, row in df.iterrows()
              if (row["_delta_comp_f"] or row["_delta_larg_f"])
          }
          df["ITEM"] = range(1, len(df) + 1)
          df = df[colunas_tmk]
        else:
          cols_agrupar = [
              "Comprimento",
              "Largura",
              "Função",
              "Fita C1",
              "Fita C2",
              "Fita L1",
              "Fita L2",
              "Material",
              "Complemento",
              "Girar",
              "_delta_comp_f",
              "_delta_larg_f",
          ]
          df = (
              df.groupby(cols_agrupar, as_index=False, dropna=False)[
                  "Quantidade"
              ].sum()
          )
          deltas_por_idx = {
              idx: (float(row["_delta_comp_f"]), float(row["_delta_larg_f"]))
              for idx, row in df.iterrows()
              if (row["_delta_comp_f"] or row["_delta_larg_f"])
          }
          df = df[colunas_p]

    df.attrs["deltas_fundo"] = deltas_por_idx
    return df

  def iniciar_revisao(self):
    if not self.controller.xml_path:
      messagebox.showwarning(
          "Atenção", "Por favor, selecione um arquivo XML primeiro."
      )
      return

    modulos_terceirizados = {}
    try:
      modulos_canto_l = listar_modulos_canto_l(self.controller.xml_path)
    except Exception:
      modulos_canto_l = []

    if modulos_canto_l:
      janela_canto_l = JanelaModulosCantoL(self, modulos_canto_l)
      self.wait_window(janela_canto_l)
      modulos_terceirizados = janela_canto_l.resultado

    texto_original_btn = self.btn_convert.cget("text")
    self.btn_convert.configure(state="disabled", text="⏳ Processando XML...")
    self.update_idletasks()

    try:
      forn_inicial = self.fornecedor_var.get().strip()
      df = self.processar_dados_para_revisao(
          self.controller.xml_path,
          fornecedor_alvo=forn_inicial,
          modulos_terceirizados=modulos_terceirizados,
      )

      if df.empty:
        messagebox.showwarning(
            "Aviso", "Nenhuma peça foi encontrada no XML selecionado."
        )
        return

      JanelaRevisao(
          self, df, self.salvar_excel_final, fornecedor_atual=forn_inicial
      )

    except Exception as e:
      messagebox.showerror(
          "Erro ao Processar",
          f"Ocorreu um erro ao processar o arquivo XML:\n{str(e)}",
      )
    finally:
      self.btn_convert.configure(state="normal", text=texto_original_btn)

  def salvar_excel_final(self, df_final: pd.DataFrame, fornecedores_escolhidos):
    cli_xml, amb_xml = extrair_cliente_ambiente_xml(self.controller.xml_path)
    cliente = self.controller.cliente_nome or cli_xml or "Cliente"
    ambiente = self.controller.ambiente_nome or amb_xml or "Ambiente"

    def limpar_str(s):
      return re.sub(r'[\\/*?:"<>|]', "", s).strip().replace(" ", "_")

    pasta_destino = self.controller.pasta_destino or os.path.dirname(
        self.controller.xml_path
    )

    salvos = []
    for fornecedor_escolhido in fornecedores_escolhidos:
      try:
        nome_saida_sugerido = f"{limpar_str(cliente)}_{limpar_str(ambiente)}_{limpar_str(fornecedor_escolhido)}.xlsx"
        save_path = os.path.join(pasta_destino, nome_saida_sugerido)

        if os.path.exists(save_path):
          substituir = messagebox.askyesno(
              "Arquivo já existe",
              f"O arquivo já existe:\n{save_path}\n\nDeseja substituí-lo?",
          )
          if not substituir:
            continue

        if fornecedor_escolhido == "TMKPlanilha":
          wb = openpyxl.Workbook()
          ws = wb.active
          ws.title = "PREENCHA AQUI"
          ws.append(list(df_final.columns))
          for _, row_data in df_final.iterrows():
            ws.append(list(row_data))
          wb.save(save_path)
        else:
          df_final.to_excel(save_path, index=False)

        self.controller.adicionar_ao_historico(
            self.controller.xml_path, save_path, fornecedor_escolhido
        )
        salvos.append((fornecedor_escolhido, save_path))

      except Exception as e:
        messagebox.showerror(
            "Erro ao Salvar",
            f"Ocorreu um erro ao salvar para {fornecedor_escolhido}:\n{str(e)}",
        )

    if salvos:
      resumo = "\n".join(f"• {forn}: {path}" for forn, path in salvos)
      messagebox.showinfo(
          "Sucesso",
          f"Total de linhas no Excel: {len(df_final)}\n\nArquivo(s) salvo(s):\n{resumo}",
      )
      self.controller.show_frame("TelaInicio")
    else:
      messagebox.showwarning(
          "Nada salvo", "Nenhum arquivo foi salvo."
      )

