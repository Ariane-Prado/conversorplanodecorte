"""Componentes visuais reutilizáveis (tooltips, estilo de inputs) e Tour Guiado."""
import tkinter as tk
import customtkinter as ctk
import pandas as pd

from atualizador import salvar_config
from constantes import (
    COR_BORDA_INPUT,
    COR_FOCO_INPUT,
    COR_FUNDO_INPUT,
    COR_TEXTO_SECUNDARIO,
    COR_VERDE_PRIMARIO,
    COR_VERDE_PRIMARIO_HOVER,
)


def criar_tooltip(widget, texto: str, wraplength: int = 260):
  """Balão de texto simples que aparece ao passar o mouse sobre `widget` e
  some ao tirar o mouse — usado pra manter labels de apoio fora da tela
  (ex: ícone de info ao lado de um select, botões de ícone no cabeçalho)."""
  estado = {"janela": None}

  def mostrar(event=None):
    if estado["janela"] is not None:
      return
    x = widget.winfo_rootx() + widget.winfo_width() // 2
    y = widget.winfo_rooty() + widget.winfo_height() + 6
    janela = tk.Toplevel(widget)
    janela.wm_overrideredirect(True)
    janela.wm_geometry(f"+{x}+{y}")
    janela.attributes("-topmost", True)
    tk.Label(
        janela,
        text=texto,
        bg="#2c3e50",
        fg="white",
        font=("Arial", 10),
        relief="solid",
        bd=1,
        padx=8,
        pady=5,
        wraplength=wraplength,
        justify="left",
    ).pack()
    estado["janela"] = janela

  def esconder(event=None):
    janela = estado["janela"]
    if janela is not None:
      estado["janela"] = None
      try:
        janela.destroy()
      except tk.TclError:
        pass

  widget.bind("<Enter>", mostrar, add="+")
  widget.bind("<Leave>", esconder, add="+")
  widget.bind("<Destroy>", esconder, add="+")


def aplicar_estilo_input(entry: ctk.CTkEntry):
  """Estilo padrão dos campos de texto editáveis: fundo escuro (#2B2B2B) e
  borda neutra (#444444) em repouso, com a borda destacada na cor primária
  ao ganhar foco — dá pra quem está digitando saber onde o cursor está."""
  entry.configure(
      fg_color=COR_FUNDO_INPUT, border_color=COR_BORDA_INPUT, border_width=1
  )
  entry.bind(
      "<FocusIn>",
      lambda e: entry.configure(border_color=COR_FOCO_INPUT, border_width=2),
      add="+",
  )
  entry.bind(
      "<FocusOut>",
      lambda e: entry.configure(border_color=COR_BORDA_INPUT, border_width=1),
      add="+",
  )


class DestaqueWidget:
  """Moldura colorida ao redor de um ou mais widgets reais da tela.

  Tkinter não suporta uma janela com "buraco" no meio, então a moldura é
  simulada com 4 tiras finas (topo/base/esquerda/direita) — técnica padrão
  em Tk. Uma única instância vive durante todo o tour; `mover_para` só
  reposiciona as tiras existentes, evitando flicker e vazamento de janelas.
  """

  def __init__(self, master, cor="#e67e22", espessura=4, margem=4):
    self._cor = cor
    self._espessura = espessura
    self._margem = margem
    self._tiras = [self._criar_tira(master) for _ in range(4)]

  def _criar_tira(self, master) -> tk.Toplevel:
    t = tk.Toplevel(master)
    t.overrideredirect(True)
    t.attributes("-topmost", True)
    t.configure(bg=self._cor)
    t.withdraw()
    return t

  @staticmethod
  def _existe(widget) -> bool:
    try:
      return bool(widget.winfo_exists())
    except tk.TclError:
      return False

  def mover_para(self, widgets: list) -> bool:
    vivos = [w for w in (widgets or []) if w is not None and self._existe(w)]
    if not vivos:
      self.esconder()
      return False

    xs1, ys1, xs2, ys2 = [], [], [], []
    for w in vivos:
      w.update_idletasks()
      x, y = w.winfo_rootx(), w.winfo_rooty()
      xs1.append(x)
      ys1.append(y)
      xs2.append(x + w.winfo_width())
      ys2.append(y + w.winfo_height())

    m = self._margem
    x1, y1 = min(xs1) - m, min(ys1) - m
    x2, y2 = max(xs2) + m, max(ys2) + m
    largura, altura = x2 - x1, y2 - y1
    e = self._espessura

    geometrias = [
        (largura + 2 * e, e, x1 - e, y1 - e),  # topo
        (largura + 2 * e, e, x1 - e, y2),       # base
        (e, altura + 2 * e, x1 - e, y1 - e),    # esquerda
        (e, altura + 2 * e, x2, y1 - e),        # direita
    ]
    for tira, (lg, al, gx, gy) in zip(self._tiras, geometrias):
      try:
        tira.geometry(f"{max(1, lg)}x{max(1, al)}+{gx}+{gy}")
        tira.deiconify()
      except tk.TclError:
        pass
    return True

  def esconder(self):
    for t in self._tiras:
      try:
        if t.winfo_exists():
          t.withdraw()
      except tk.TclError:
        pass

  def destruir(self):
    for t in self._tiras:
      try:
        t.destroy()
      except tk.TclError:
        pass
    self._tiras = []


class TourGuiado:
  """Tour guiado: navega pelas telas reais do app e destaca, com uma
  moldura colorida, o widget real de cada passo — em vez de descrever tudo
  numa janela separada e desconectada da interface."""

  COR_DESTAQUE = "#e67e22"
  _ATRASO_REDESENHO_MS = 120
  _DEBOUNCE_CONFIGURE_MS = 80

  def __init__(self, controller, passo_inicial=0, janela_revisao=None):
    self.controller = controller
    self.janela_revisao_real = janela_revisao
    self._janela_revisao_demo = None
    self._grab_para_restaurar = False
    self._destruido = False
    self._after_id_redesenho = None
    self._after_id_configure = None
    self._bind_id_configure = None
    self._widget_configure_alvo = None
    self.callout = None

    self._passos = self._montar_passos()

    if janela_revisao is not None:
      indices_revisao = [i for i, p in enumerate(self._passos) if p.get("revisao")]
      self._passo_min = indices_revisao[0]
      self._passo_max = indices_revisao[-1]
      try:
        self.janela_revisao_real.grab_release()
        self._grab_para_restaurar = True
      except tk.TclError:
        pass
      self.janela_revisao_real.bind(
          "<Destroy>", self._ao_destruir_revisao_real, add="+"
      )
    else:
      self._passo_min = 0
      self._passo_max = len(self._passos) - 1

    self.passo_atual = max(self._passo_min, min(passo_inicial, self._passo_max))
    self._destaque = DestaqueWidget(controller, cor=self.COR_DESTAQUE)
    self.callout = self._criar_callout_base()
    self._mostrar_passo()

  # ---- construção dos passos ---------------------------------------------

  def _montar_passos(self) -> list:
    c = self.controller
    ti = c.frames["TelaInicio"]
    tn = c.frames["TelaNovoProjeto"]
    tc = c.frames["TelaConfiguracaoProjeto"]

    return [
        dict(
            tela="TelaInicio",
            widget=lambda: None,
            titulo="👋 Bem-vindo(a) ao Conversor XML Promob / Start KNR",
            texto=(
                "Vamos fazer um tour rápido pelas telas reais do sistema."
                " Use os botões abaixo pra navegar — pode fechar quando"
                " quiser, o '❓ Ajuda' sempre reabre este tour."
            ),
        ),
        dict(
            tela="TelaInicio",
            widget=lambda: ti.btn_novo,
            titulo="🏠 Começando um Projeto",
            texto="Clique aqui para iniciar um novo projeto do zero.",
        ),
        dict(
            tela="TelaInicio",
            widget=lambda: ti.frame_recentes,
            titulo="📋 Histórico de Projetos",
            texto=(
                "Seus projetos anteriores ficam aqui. Use o botão"
                " '[ Editar ]' pra reabrir um deles sem refazer tudo."
            ),
        ),
        dict(
            tela="TelaInicio",
            widget=lambda: ti.frame_backup_inferior,
            titulo="☁️ Backup e Restauração",
            texto=(
                "Salve seu histórico num arquivo .zip, ou restaure-o em"
                " outro computador."
            ),
        ),
        dict(
            tela="TelaNovoProjeto",
            widget=lambda: tn.ent_cliente,
            titulo="📝 Nome do Cliente",
            texto="Confirme ou digite o nome do cliente aqui.",
        ),
        dict(
            tela="TelaNovoProjeto",
            widget=lambda: tn.ent_ambiente,
            titulo="📝 Nome do Ambiente",
            texto="E o nome do ambiente aqui — ex: 'Cozinha', 'Quarto Bebê'.",
        ),
        dict(
            tela="TelaConfiguracaoProjeto",
            widget=lambda: tc.btn_select_xml,
            titulo="⚙️ Selecionar o XML",
            texto="Clique aqui e escolha o .xml exportado do Promob/Start KNR.",
        ),
        dict(
            tela="TelaConfiguracaoProjeto",
            widget=lambda: tc.cmb_fornecedor,
            titulo="🏭 Escolher a Fábrica",
            texto=(
                "Teletintas, Madefer, Tobias e TMKCloud usam o mesmo"
                " layout de planilha. TMKPlanilha usa um formato próprio."
            ),
        ),
        dict(
            tela="TelaConfiguracaoProjeto",
            widget=lambda: tc.btn_select_folder,
            titulo="📁 Pasta de Destino",
            texto="Escolha onde o Excel será salvo (por padrão, a pasta do XML).",
        ),
        dict(
            tela="TelaConfiguracaoProjeto",
            widget=lambda: tc.chk_agrupar,
            titulo="🔗 Agrupar Peças",
            texto="Marque pra somar automaticamente peças idênticas numa só linha.",
        ),
        dict(
            tela="TelaConfiguracaoProjeto",
            widget=lambda: tc.btn_convert,
            titulo="🔍 Ir para a Revisão",
            texto=(
                "Com tudo preenchido, clique aqui pra processar o XML e"
                " revisar. Se o projeto tiver módulos de Canto L, uma tela"
                " extra pergunta, por módulo, se a peça fica na medida"
                " original ou leva margem extra (+50mm/+100mm)."
            ),
        ),
        dict(
            revisao=True,
            widget=lambda: self._revisao_ativa().tree,
            titulo="✏️ Revisando as Peças",
            texto=(
                "Duplo-clique numa célula edita na hora (Enter confirma,"
                " Esc cancela). Clique no cabeçalho de uma coluna pra"
                " ordenar. O ☐ no canto superior esquerdo da tabela marca"
                " ou desmarca todas as linhas de uma vez."
            ),
        ),
        dict(
            revisao=True,
            widget=lambda: [
                self._revisao_ativa().btn_excluir,
                self._revisao_ativa().btn_desfazer,
            ],
            titulo="🗑️ Excluir e Desfazer",
            texto=(
                "Marque uma ou mais linhas (clicando nelas ou usando o ☐ do"
                " cabeçalho) e exclua da exportação. Errou? O botão"
                " '↩️ Desfazer Exclusão' reverte a última exclusão."
            ),
        ),
        dict(
            revisao=True,
            widget=lambda: self._revisao_ativa().btn_prox_excedente,
            titulo="⚠️ Peças Fora do Padrão",
            texto=(
                "Peças em vermelho ultrapassam o tamanho da chapa — este"
                " botão pula direto pra próxima delas."
            ),
        ),
        dict(
            revisao=True,
            widget=lambda: self._revisao_ativa().frame_sel_forn,
            titulo="💾 Salvar para Fábricas",
            texto=(
                "Marque uma ou mais fábricas compatíveis e gere todos os"
                " arquivos de uma vez."
            ),
        ),
        dict(
            revisao=True,
            widget=lambda: self._revisao_ativa().btn_confirmar,
            titulo="✅ Confirmar e Salvar",
            texto=(
                "Gera o Excel e o projeto entra no histórico. (Neste"
                " exemplo do tutorial o botão fica desativado.)"
            ),
        ),
        dict(
            tela="TelaInicio",
            widget=lambda: c.btn_ajuda,
            titulo="🎉 Pronto!",
            texto=(
                "Isso é tudo! Reabra este tour quando quiser clicando neste"
                " ❓ aqui em cima, no canto superior direito."
            ),
        ),
    ]

  # ---- janela de revisão (real ou demo) ----------------------------------

  def _revisao_ativa(self):
    if self.janela_revisao_real is not None and self._widget_vivo(
        self.janela_revisao_real
    ):
      return self.janela_revisao_real
    return self._garantir_janela_revisao_demo()

  def _garantir_janela_revisao_demo(self):
    if self._janela_revisao_demo is not None and self._widget_vivo(
        self._janela_revisao_demo
    ):
      return self._janela_revisao_demo

    from ui_janelas import JanelaRevisao

    df_demo = pd.DataFrame(
        [
            [2, 2740, 580, "Lateral", "0.45mm Branco", "", "0.45mm Branco", "",
             "Branco TX 15mm", "Módulo Demo", "NÃO"],
            [1, 2800, 400, "Prateleira", "0.45mm Branco", "0.45mm Branco", "",
             "", "Branco TX 15mm", "Módulo Demo", "SIM"],
        ],
        columns=[
            "Quantidade", "Comprimento", "Largura", "Função",
            "Fita C1", "Fita C2", "Fita L1", "Fita L2",
            "Material", "Complemento", "Girar",
        ],
    )
    self._janela_revisao_demo = JanelaRevisao(
        self.controller,
        df_demo,
        lambda *args: None,
        fornecedor_atual="Teletintas",
        iniciar_grab=False,
    )
    self._janela_revisao_demo.title("Revisão de Peças (exemplo do tutorial)")
    self._janela_revisao_demo.btn_confirmar.configure(state="disabled")
    return self._janela_revisao_demo

  def _fechar_demo_revisao(self):
    if self._janela_revisao_demo is not None:
      if self._widget_vivo(self._janela_revisao_demo):
        try:
          self._janela_revisao_demo.destroy()
        except tk.TclError:
          pass
      self._janela_revisao_demo = None

  @staticmethod
  def _widget_vivo(widget) -> bool:
    try:
      return bool(widget.winfo_exists())
    except tk.TclError:
      return False

  # ---- navegação / desenho ------------------------------------------------

  def _mostrar_passo(self):
    if self._destruido:
      return
    passo = self._passos[self.passo_atual]

    if passo.get("revisao"):
      janela = self._revisao_ativa()
      janela.lift()
      janela.focus_force()
    else:
      self._fechar_demo_revisao()
      self.controller.show_frame(passo["tela"])
      self.controller.lift()
      self.controller.focus_force()

    if self._after_id_redesenho is not None:
      try:
        self.controller.after_cancel(self._after_id_redesenho)
      except (tk.TclError, ValueError):
        pass
    self._after_id_redesenho = self.controller.after(
        self._ATRASO_REDESENHO_MS, self._desenhar_destaque_e_callout
    )

  def _desenhar_destaque_e_callout(self):
    self._after_id_redesenho = None
    if self._destruido:
      return

    passo = self._passos[self.passo_atual]
    try:
      alvo = passo["widget"]()
    except (tk.TclError, AttributeError):
      alvo = None

    brutos = alvo if isinstance(alvo, list) else ([alvo] if alvo is not None else [])
    widgets = [w for w in brutos if w is not None and self._widget_vivo(w)]

    if widgets:
      self._destaque.mover_para(widgets)
      topo = widgets[0].winfo_toplevel()
    else:
      self._destaque.esconder()
      topo = self.controller

    self._rebind_configure(topo)
    self._atualizar_callout(topo, passo)

  def _rebind_configure(self, topo):
    if self._widget_configure_alvo is not None and self._bind_id_configure is not None:
      try:
        self._widget_configure_alvo.unbind("<Configure>", self._bind_id_configure)
      except tk.TclError:
        pass
    self._widget_configure_alvo = topo
    self._bind_id_configure = topo.bind(
        "<Configure>", self._ao_configurar_hospedeiro, add="+"
    )

  def _ao_configurar_hospedeiro(self, event=None):
    if self._destruido:
      return
    if self._after_id_configure is not None:
      try:
        self.controller.after_cancel(self._after_id_configure)
      except (tk.TclError, ValueError):
        pass
    self._after_id_configure = self.controller.after(
        self._DEBOUNCE_CONFIGURE_MS, self._desenhar_destaque_e_callout
    )

  def _criar_callout_base(self) -> ctk.CTkToplevel:
    """Cria o cartão de instrução UMA vez só; os passos seguintes apenas
    reposicionam/reconfiguram os widgets aqui criados (ver
    `_atualizar_callout`). Destruir e recriar um CTkToplevel a cada passo
    dispararia um bug de callback assíncrono do customtkinter no Windows
    (`_revert_withdraw_after_windows_set_titlebar_color` tentando agir numa
    janela já destruída) — reaproveitar a mesma janela evita isso."""
    card = ctk.CTkToplevel(self.controller)
    card.title("Tutorial")
    card.resizable(False, False)
    card.attributes("-topmost", True)
    card.protocol("WM_DELETE_WINDOW", self.fechar)

    self._lbl_progresso = ctk.CTkLabel(
        card, text="", font=ctk.CTkFont(size=10), text_color=COR_TEXTO_SECUNDARIO
    )
    self._lbl_progresso.pack(pady=(10, 0))

    self._lbl_titulo = ctk.CTkLabel(
        card, text="", font=ctk.CTkFont(size=14, weight="bold"), wraplength=300
    )
    self._lbl_titulo.pack(pady=(2, 6), padx=15)

    self._lbl_texto = ctk.CTkLabel(
        card, text="", font=ctk.CTkFont(size=11), wraplength=300, justify="left"
    )
    self._lbl_texto.pack(padx=15, fill="both", expand=True)

    frame_nav = ctk.CTkFrame(card, fg_color="transparent")
    frame_nav.pack(pady=8, padx=12, fill="x")

    ctk.CTkButton(
        frame_nav,
        text="Pular",
        width=60,
        fg_color="#7f8c8d",
        hover_color="#95a5a6",
        command=self.fechar,
    ).pack(side="left")

    self._btn_anterior = ctk.CTkButton(
        frame_nav, text="⬅️", width=40, command=self._anterior
    )
    self._btn_anterior.pack(side="left", padx=(6, 0))

    self._btn_proximo = ctk.CTkButton(
        frame_nav,
        text="Próximo ➔",
        width=110,
        fg_color=COR_VERDE_PRIMARIO,
        hover_color=COR_VERDE_PRIMARIO_HOVER,
        command=self._proximo,
    )
    self._btn_proximo.pack(side="right")

    return card

  def _atualizar_callout(self, topo, passo):
    largura, altura = 340, 230
    try:
      topo.update_idletasks()
      x = topo.winfo_rootx() + max(20, topo.winfo_width() - largura - 30)
      y = topo.winfo_rooty() + max(20, topo.winfo_height() - altura - 60)
    except tk.TclError:
      x, y = 100, 100

    self.callout.geometry(f"{largura}x{altura}+{x}+{y}")

    total_passos = self._passo_max - self._passo_min + 1
    numero_passo = self.passo_atual - self._passo_min + 1
    self._lbl_progresso.configure(text=f"Passo {numero_passo} de {total_passos}")
    self._lbl_titulo.configure(text=passo["titulo"])
    self._lbl_texto.configure(text=passo["texto"])

    self._btn_anterior.configure(
        state="normal" if self.passo_atual > self._passo_min else "disabled"
    )
    self._btn_proximo.configure(
        text="Concluir ✅" if self.passo_atual == self._passo_max else "Próximo ➔"
    )

    self.callout.lift()

  def _anterior(self):
    if self.passo_atual > self._passo_min:
      self.passo_atual -= 1
      self._mostrar_passo()

  def _proximo(self):
    if self.passo_atual < self._passo_max:
      self.passo_atual += 1
      self._mostrar_passo()
    else:
      self.fechar()

  def _ao_destruir_revisao_real(self, event=None):
    if self._destruido:
      return
    self._grab_para_restaurar = False
    self.fechar()

  def fechar(self):
    if self._destruido:
      return
    self._destruido = True

    for after_id in (self._after_id_redesenho, self._after_id_configure):
      if after_id is not None:
        try:
          self.controller.after_cancel(after_id)
        except (tk.TclError, ValueError):
          pass

    if self._widget_configure_alvo is not None and self._bind_id_configure is not None:
      try:
        self._widget_configure_alvo.unbind("<Configure>", self._bind_id_configure)
      except tk.TclError:
        pass

    if self._grab_para_restaurar and self._widget_vivo(self.janela_revisao_real):
      try:
        self.janela_revisao_real.grab_set()
      except tk.TclError:
        pass

    self._fechar_demo_revisao()
    self._destaque.destruir()

    if self.callout is not None:
      try:
        self.callout.destroy()
      except tk.TclError:
        pass
      self.callout = None

    salvar_config({"tutorial_visto": True})

