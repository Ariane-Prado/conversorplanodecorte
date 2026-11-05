"""Ponto de entrada principal do Conversor XML Promob / Start KNR para Excel.

A arquitetura do projeto é dividida nos seguintes módulos:
- `constantes.py`: Paleta de cores (WCAG AA), limites de chapa, fornecedores e constantes de regras.
- `regras_marcenaria.py`: Regras de domínio de marcenaria (renomeação de peças, furos/rasgos, Canto L,
  ajustes de Fundo 6mm +6mm e Contra Frente/Fundo -3mm) e extração de dados do XML.
- `atualizador.py`: Persistência de configuração local e auto-updater via GitHub Releases.
- `ui_componentes.py`: Componentes visuais reutilizáveis (tooltips, estilo de input) e Tour Guiado.
- `ui_janelas.py`: Janelas modais de conferência (`JanelaModulosCantoL` e `JanelaRevisao`).
- `ui_telas.py`: Telas do fluxo principal (`TelaInicio`, `TelaNovoProjeto`, `TelaConfiguracaoProjeto`).
"""
import datetime
import json
import os
import sys
import threading
from tkinter import messagebox
import urllib.request
import customtkinter as ctk

from atualizador import (  # noqa: F401
    _versao_para_tupla,
    aplicar_atualizacao_e_reiniciar,
    baixar_atualizacao,
    carregar_config,
    obter_caminho_base,
    salvar_config,
    verificar_nova_versao,
)
from constantes import (  # noqa: F401
    _PADRAO_ESPESSURA_MM,
    _PADRAO_MODULO_CANTO_L,
    _TOKENS_ARMARIO_BASE_CHAPEU,
    _TOKENS_BALCAO_INFERIOR,
    AJUSTE_CONTRA_FRENTE_ALTURA_DELTA,
    AJUSTE_FUNDO_8MM_DELTA,
    AVANCO_PADRAO_TOTAL_MM,
    COR_BORDA_INPUT,
    COR_FOCO_INPUT,
    COR_FUNDO_INPUT,
    COR_TEXTO_SECUNDARIO,
    COR_TITULO,
    COR_VERDE_PRIMARIO,
    COR_VERDE_PRIMARIO_HOVER,
    DICAS_COLUNAS,
    FAMILIAS_IGNORADAS,
    FORNECEDORES_DISPONIVEIS,
    FORNECEDORES_FORMATO_PADRAO,
    FORNECEDORES_FORMATO_TMK,
    FUNCAO_CONTRA_FRENTE_FUNDO,
    FUNCAO_FUNDO,
    FUNCOES_AFETADAS_TERCEIRIZACAO,
    LIMITE_MAX_COMPRIMENTO,
    LIMITE_MAX_LARGURA,
    OPCOES_MARGEM_TERCEIRIZACAO_MM,
    URL_VERSAO_REMOTA,
    VERSAO_ATUAL,
)
from regras_marcenaria import (  # noqa: F401
    _eh_espessura_6mm_texto,
    _eh_item_gaveta_xml,
    _encontrar_item_ancestral,
    _normalizar_texto,
    _tokens_normalizados,
    aplicar_ajustes_medidas_df,
    calcular_delta_contra_frente,
    calcular_delta_fundo_xml,
    calcular_deltas_ajuste_df,
    deve_ignorar_item,
    eh_fundo_6mm,
    eh_modulo_canto_l,
    escolher_eixo_largura_fundo,
    extrair_cliente_ambiente_xml,
    extrair_espessura_fita,
    extrair_fabricantes_do_xml,
    extrair_modulo_pai,
    extrair_referencias,
    extrair_servicos_adicionais,
    formatar_material_com_espessura,
    formatar_medida_mm,
    listar_modulos_canto_l,
    mapear_contexto_modulos_xml,
    renomear_funcao,
)
from ui_componentes import (  # noqa: F401
    DestaqueWidget,
    TourGuiado,
    aplicar_estilo_input,
    criar_tooltip,
)
from ui_janelas import (  # noqa: F401
    JanelaModulosCantoL,
    JanelaRevisao,
)
from ui_telas import (  # noqa: F401
    TelaConfiguracaoProjeto,
    TelaInicio,
    TelaNovoProjeto,
)

# Configuração do tema visual da interface. Fixo em "dark": várias cores da
# interface (cards, botões, inputs) já são hardcoded pensando em fundo
# escuro — deixar em "System" fazia o contraste variar (e às vezes falhar)
# dependendo do tema do Windows do usuário.
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class ConversorXmlExcelApp(ctk.CTk):

  def __init__(self):
    super().__init__()

    self.title("Conversor XML Promob para Excel")
    self.geometry("860 x 680")
    self.minsize(800, 600)

    self.xml_path = None
    self.pasta_destino = None
    self.cliente_nome = ""
    self.ambiente_nome = ""
    self.fornecedor_inicial = "Teletintas"
    self.arquivo_historico = os.path.join(
        obter_caminho_base(), "historico_recentes.json"
    )
    self.arquivo_historico_usuario = os.path.join(
        os.path.expanduser("~"), ".conversor_promob_historico.json"
    )
    self.historico = self.carregar_historico()

    self.frame_topo = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_topo.pack(side="top", fill="x", padx=15, pady=(8, 0))

    # Botões-ícone com contorno visível: ajuda e verificação de atualização
    # não são ações do fluxo principal, mas precisam se ler como botões
    # clicáveis mesmo discretos — daí a borda e o ícone bem contrastado.
    self.btn_ajuda = ctk.CTkButton(
        self.frame_topo,
        text="❓",
        width=34,
        height=34,
        corner_radius=17,
        fg_color="transparent",
        text_color=COR_TEXTO_SECUNDARIO,
        border_width=1,
        border_color=COR_BORDA_INPUT,
        hover_color="#3a3f44",
        font=ctk.CTkFont(size=15, weight="bold"),
        command=self.abrir_tutorial,
    )
    self.btn_ajuda.pack(side="right")
    criar_tooltip(self.btn_ajuda, "Tutorial e Ajuda")

    self.btn_verificar_update = ctk.CTkButton(
        self.frame_topo,
        text="🔄",
        width=34,
        height=34,
        corner_radius=17,
        fg_color="transparent",
        text_color=COR_TEXTO_SECUNDARIO,
        border_width=1,
        border_color=COR_BORDA_INPUT,
        hover_color="#3a3f44",
        font=ctk.CTkFont(size=15, weight="bold"),
        command=self.verificar_atualizacoes_manual,
    )
    self.btn_verificar_update.pack(side="right", padx=(0, 8))
    criar_tooltip(self.btn_verificar_update, "Verificar atualizações")

    self.container = ctk.CTkFrame(self, fg_color="transparent")
    self.container.pack(fill="both", expand=True, padx=10, pady=10)
    self.container.grid_rowconfigure(0, weight=1)
    self.container.grid_columnconfigure(0, weight=1)

    self.lbl_footer = ctk.CTkLabel(
        self,
        text=f"Desenvolvido por Ariane Prado · v{VERSAO_ATUAL}",
        font=ctk.CTkFont(size=11, weight="bold"),
        text_color=COR_TEXTO_SECUNDARIO,
    )
    self.lbl_footer.pack(side="bottom", pady=6)

    self.frames = {}
    for F in (TelaInicio, TelaNovoProjeto, TelaConfiguracaoProjeto):
      page_name = F.__name__
      frame = F(parent=self.container, controller=self)
      self.frames[page_name] = frame
      frame.grid(row=0, column=0, sticky="nsew")

    self.show_frame("TelaInicio")
    self.after(400, self.mostrar_tutorial_se_necessario)
    self.after(1200, self.verificar_atualizacao_em_segundo_plano)

  def show_frame(self, page_name):
    frame = self.frames[page_name]
    if hasattr(frame, "ao_exibir_tela"):
      frame.ao_exibir_tela()
    frame.tkraise()

  def abrir_tutorial(self, passo_inicial=0):
    TourGuiado(self, passo_inicial=passo_inicial)

  def mostrar_tutorial_se_necessario(self):
    if not carregar_config().get("tutorial_visto", False):
      self.abrir_tutorial()

  # --- Atualização automática -------------------------------------------

  def verificar_atualizacao_em_segundo_plano(self):
    def worker():
      resultado = verificar_nova_versao()
      if resultado:
        self.after(0, lambda: self._notificar_nova_versao(*resultado))

    threading.Thread(target=worker, daemon=True).start()

  def verificar_atualizacoes_manual(self):
    self.btn_verificar_update.configure(state="disabled", text="⏳")

    def worker():
      resultado = verificar_nova_versao()
      self.after(0, lambda: self._resultado_verificacao_manual(resultado))

    threading.Thread(target=worker, daemon=True).start()

  def _resultado_verificacao_manual(self, resultado):
    self.btn_verificar_update.configure(state="normal", text="🔄")
    if resultado:
      self._notificar_nova_versao(*resultado)
    else:
      messagebox.showinfo(
          "Sem atualizações",
          f"Você já está usando a versão mais recente (v{VERSAO_ATUAL}).",
      )

  def _notificar_nova_versao(self, versao_nova, url_download):
    resposta = messagebox.askyesno(
      "Atualização disponível",
      f"Uma nova versão está disponível: v{versao_nova}\n"
      f"(você está usando v{VERSAO_ATUAL}).\n\n"
      "Deseja baixar e atualizar agora? O aplicativo vai fechar e reabrir"
      " sozinho ao terminar.",
    )
    if resposta:
      self._iniciar_download_atualizacao(versao_nova, url_download)

  def _iniciar_download_atualizacao(self, versao_nova, url_download):
    if not getattr(sys, "frozen", False):
      messagebox.showwarning(
          "Atualização automática indisponível",
          "A troca automática do executável só funciona no aplicativo"
          " instalado (.exe). Você está rodando a partir do código-fonte —"
          f" baixe e substitua manualmente pelo link:\n\n{url_download}",
      )
      return

    janela = ctk.CTkToplevel(self)
    janela.title("Atualizando")
    janela.geometry("380x130")
    janela.resizable(False, False)
    janela.transient(self)
    janela.grab_set()

    ctk.CTkLabel(
        janela,
        text=f"Baixando versão {versao_nova}...",
        font=ctk.CTkFont(size=13, weight="bold"),
    ).pack(pady=(20, 10))

    barra = ctk.CTkProgressBar(janela, width=300)
    barra.pack(pady=5)
    barra.set(0)

    lbl_pct = ctk.CTkLabel(
        janela, text="0%", font=ctk.CTkFont(size=11), text_color=COR_TEXTO_SECUNDARIO
    )
    lbl_pct.pack()

    def progresso(lido, total):
      if total > 0:
        frac = lido / total
        self.after(
            0,
            lambda: (barra.set(frac), lbl_pct.configure(text=f"{int(frac * 100)}%")),
        )

    def worker():
      pasta = os.path.dirname(sys.executable)
      destino = os.path.join(pasta, f"_app_novo_{versao_nova}.exe")
      try:
        baixar_atualizacao(url_download, destino, progresso_callback=progresso)
      except Exception as e:
        self.after(0, lambda: self._erro_download(janela, e))
        return
      self.after(0, lambda: self._finalizar_atualizacao(destino))

    threading.Thread(target=worker, daemon=True).start()

  def _erro_download(self, janela_progresso, erro):
    janela_progresso.destroy()
    messagebox.showerror(
        "Erro ao baixar atualização",
        f"Não foi possível baixar a atualização:\n{erro}",
    )

  def _finalizar_atualizacao(self, destino):
    try:
      aplicar_atualizacao_e_reiniciar(destino)
    except Exception as e:
      messagebox.showerror(
          "Erro ao aplicar atualização",
          "O download terminou, mas não foi possível trocar o executável"
          f" automaticamente:\n{e}\n\nO arquivo baixado está em:\n{destino}"
          "\n\nFeche o programa e substitua manualmente se quiser.",
      )

  @staticmethod
  def _normalizar_caminho_historico(caminho: str) -> str:
    if not caminho:
      return ""
    return os.path.normcase(os.path.normpath(str(caminho)))

  @classmethod
  def _chave_item_historico(cls, item: dict) -> tuple[str, str]:
    return (
        cls._normalizar_caminho_historico(item.get("caminho_xml", "")),
        cls._normalizar_caminho_historico(item.get("caminho_excel", "")),
    )

  @staticmethod
  def _timestamp_item_historico(item: dict) -> float:
    data_str = item.get("data_hora", "")
    if data_str:
      try:
        dt = datetime.datetime.strptime(data_str, "%d/%m/%Y %H:%M")
        return dt.timestamp()
      except ValueError:
        pass
    return 0.0

  def _caminhos_arquivos_historico(self) -> list[str]:
    caminhos = []
    for attr in ("arquivo_historico", "arquivo_historico_usuario"):
      caminho = getattr(self, attr, None)
      if caminho:
        norm = self._normalizar_caminho_historico(os.path.abspath(caminho))
        if all(
            self._normalizar_caminho_historico(os.path.abspath(c)) != norm
            for c in caminhos
        ):
          caminhos.append(caminho)
    return caminhos

  def carregar_historico(self):
    itens_por_chave = {}
    ordem_insercao = []
    for caminho in self._caminhos_arquivos_historico():
      if not os.path.exists(caminho):
        continue
      try:
        with open(caminho, "r", encoding="utf-8") as f:
          dados = json.load(f)
        if not isinstance(dados, list):
          continue
      except Exception:
        continue

      for item in dados:
        if not isinstance(item, dict):
          continue
        chave = self._chave_item_historico(item)
        if chave not in itens_por_chave:
          itens_por_chave[chave] = item
          ordem_insercao.append(chave)
        else:
          if self._timestamp_item_historico(item) > self._timestamp_item_historico(
              itens_por_chave[chave]
          ):
            itens_por_chave[chave] = item

    # Ordena do mais recente para o mais antigo mantendo a ordem relativa em caso de empate
    chaves_ordenadas = sorted(
         range(len(ordem_insercao)),
        key=lambda i: (
            -self._timestamp_item_historico(itens_por_chave[ordem_insercao[i]]),
            i,
        ),
    )
    return [itens_por_chave[ordem_insercao[i]] for i in chaves_ordenadas][:50]

  def salvar_historico(self):
    dados = self.historico[:50]
    for caminho in self._caminhos_arquivos_historico():
      try:
        with open(caminho, "w", encoding="utf-8") as f:
          json.dump(dados, f, ensure_ascii=False, indent=2)
      except Exception as e:
        print(f"Erro ao salvar histórico em {caminho}: {e}")

  def adicionar_ao_historico(self, caminho_xml, caminho_excel, fornecedor):
    # Recarrega antes de inserir para nunca sobrescrever projetos salvos por outra
    # janela aberta ou outra instância do aplicativo.
    self.historico = self.carregar_historico()
    novo_registro = {
        "nome_xml": os.path.basename(caminho_xml),
        "caminho_xml": caminho_xml,
        "caminho_excel": caminho_excel,
        "cliente": self.cliente_nome,
        "ambiente": self.ambiente_nome,
        "fornecedor": fornecedor,
        "data_hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    chave_nova = self._chave_item_historico(novo_registro)
    self.historico = [
        item
        for item in self.historico
        if self._chave_item_historico(item) != chave_nova
    ]
    self.historico.insert(0, novo_registro)
    self.historico = self.historico[:50]
    self.salvar_historico()

    if "TelaInicio" in self.frames:
      self.frames["TelaInicio"].atualizar_interface_historico()


if __name__ == "__main__":
  app = ConversorXmlExcelApp()
  app.mainloop()
