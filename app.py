import datetime
import json
import os
import re
import subprocess
import sys
import threading
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
import customtkinter as ctk
import openpyxl
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Configuração do tema visual da interface
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

VERSAO_ATUAL = "1.0.0"
URL_VERSAO_REMOTA = (
    "https://raw.githubusercontent.com/Ariane-Prado/conversorplanodecorte"
    "/main/versao.json"
)

LIMITE_MAX_COMPRIMENTO = 2740.0
LIMITE_MAX_LARGURA = 1840.0

DICAS_COLUNAS = {
    "ITEM": "(peça - número de ordem)",
    "QUANT": "(unidades físicas da peça)",
    "Quantidade": "(unidades físicas da peça)",
    "COMP\n2750": "(mm - medida do comprimento / sentido do veio)",
    "Comprimento": "(mm - medida do comprimento / sentido do veio)",
    "LARG\n1840": "(mm - medida da largura da peça)",
    "Largura": "(mm - medida da largura da peça)",
    "NOME DA PEÇA": "(descrição para constar na etiqueta da peça)",
    "Função": "(função ou nome da peça na marcenaria)",
    "SERVIÇO ADICIONAIS": (
        "(listar serviços como furo de dobradiça / prensa / usinagem / rasgo)"
    ),
    "AMBIENTE": (
        "(nome do ambiente para constar na etiqueta da peça - opcional)"
    ),
    "CLIENTE": "(nome do cliente para constar na etiqueta)",
    "OBS 1": "(informações complementares para sair na etiqueta - opcional)",
    "OBS 2": "(observações internas de fábrica)",
    "Complemento": "(nome do módulo ou caixa do móvel)",
    "FITA \n1ª COMP ": "(marcar 'x' para fitar 1º lado comprimento)",
    "FITA \n2ª COMP ": "(marcar 'x' para fitar 2º lado comprimento)",
    "FITA \n1ª LARG": "(marcar 'x' para fitar 1º lado largura)",
    "FITA \n2ª LARG": "(marcar 'x' para fitar 2º lado largura)",
    "Fita C1": "(fita do lado Comprimento 1)",
    "Fita C2": "(fita do lado Comprimento 2)",
    "Fita L1": "(fita do lado Largura 1)",
    "Fita L2": "(fita do lado Largura 2)",
    "COR MDF": "(colocar padrão e marca do MDF)",
    "Material": "(padrão, espessura e marca do MDF/MDP)",
    "ESP MDF": "colocar espessura do MDF (ex: 6mm, 15mm, 18mm ou 25mm)",
    "ESP FITA": "(colocar espessura da fita - ex: 0.45mm ou 1mm)",
    "Girar": "(marcar SIM se a peça puder ser girada no corte)",
}

FORNECEDORES_DISPONIVEIS = [
    "Teletintas",
    "Madefer",
    "Tobias",
    "TMKCloud",
    "TMKPlanilha",
]

# Fornecedores que compartilham a mesma estrutura de colunas (formato "padrão").
# TMKCloud usa um layout de colunas próprio (ver processar_dados_para_revisao),
# por isso fica de fora desse grupo.
FORNECEDORES_FORMATO_PADRAO = ["Teletintas", "Madefer", "Tobias", "TMKPlanilha"]
FORNECEDORES_FORMATO_TMK = ["TMKCloud"]

# Famílias de itens que não representam peças de MDF a serem cortadas
# (ferragens, roteiros de montagem, ou linhas de parceiros como a Alumistar,
# que fornece portas/painéis de vidro e alumínio prontos).
FAMILIAS_IGNORADAS = {
    "Roteiro Produtivo",
    "Sta_fer",
    "Ferragens",
    "Zen",
    "Acessórios",
    "Acessorios",
    "Alumistar",
}


def renomear_funcao(funcao_original: str) -> str:
  f = funcao_original.strip()
  if f in ["Contra Frente", "Contra Fundo"]:
    return "Contra Frente / Contra Fundo de Gaveta"
  if f.startswith("Porta") or "Porta " in f or "Porta" in f.split():
    return "Porta"
  if f in ["Lateral Direita Gaveta", "Lateral Esquerda Gaveta"]:
    return "Lateral de Gaveta"
  if f.lower().startswith("base"):
    return "Base"
  return f


def extrair_servicos_adicionais(item, desc_original: str) -> str:
  servicos = []
  is_porta = (
      desc_original.startswith("Porta")
      or "Porta " in desc_original
      or "POR_" in item.attrib.get("ID", "")
  )
  if is_porta:
    qtd_furos = 0
    for sub in item.findall(".//ITEM"):
      if sub.attrib.get("DESCRIPTION") == "Furar" or sub.attrib.get(
          "STRUCTUREKEY"
      ) in ["FURAR", "DOBRADICA"]:
        try:
          qtd_furos += int(float(sub.attrib.get("QUANTITY", "0")))
        except ValueError:
          pass
    desc_lower = desc_original.lower()
    lado = ""
    if "dir" in desc_lower or "direita" in desc_lower:
      lado = "Ld Dir"
    elif "esq" in desc_lower or "esquerda" in desc_lower:
      lado = "Ld Esq"
    if qtd_furos > 0:
      servicos.append(
          f"{qtd_furos} furos de dobr."
          + (f" {lado}" if lado else " Ld maior")
      )

  tem_rasgo = False
  for sub in item.findall(".//ITEM"):
    s_desc = sub.attrib.get("DESCRIPTION", "")
    s_key = sub.attrib.get("STRUCTUREKEY", "")
    s_ref = sub.attrib.get("REFERENCE", "")
    if (
        "Rasgo" in s_desc
        or "Rebaixo" in s_desc
        or s_key == "RASGO"
        or s_ref == "RASGAR"
    ):
      tem_rasgo = True
      break
  if tem_rasgo:
    servicos.append("C/ Rasgo de Fundo")

  return " / ".join(servicos)


def extrair_fabricantes_do_xml(root) -> dict:
  fab_map = {}
  for mi in root.findall(".//MODELTYPEINFORMATION"):
    t_desc = mi.attrib.get("DESCRIPTION", "").strip()
    if "–" in t_desc or "-" in t_desc:
      parts = re.split(r"[–\-]", t_desc)
      if len(parts) >= 2:
        fabricante = parts[-1].strip().upper()
        model_part = parts[0].strip()
        if ">" in model_part:
          model_part = model_part.split(">")[-1].strip()
        model_part = re.sub(r"^[A-Z]\s+", "", model_part)
        model_part = re.sub(r"\bBP\b", "", model_part).strip()
        if model_part:
          fab_map[model_part.lower()] = fabricante
  return fab_map


def formatar_material_com_espessura(
    item, refs: dict, fab_map: dict = None
) -> str:
  material = refs.get("MATERIAL", "").strip()
  modelo = refs.get("MODEL", "").strip()
  height_str = item.attrib.get("HEIGHT", "").strip()
  espessura_str = ""
  if height_str:
    try:
      h_val = float(height_str)
      espessura_str = f"{int(h_val) if h_val.is_integer() else h_val} mm"
    except ValueError:
      espessura_str = ""
  fabricante = (
      fab_map.get(modelo.lower(), "") if fab_map and modelo else ""
  )
  partes = []
  if material:
    partes.append(material)
  if modelo:
    partes.append(modelo)
  if espessura_str:
    partes.append(espessura_str)
  if fabricante:
    partes.append(fabricante)
  return " ".join(partes) if partes else "Não especificado"


def extrair_espessura_fita(bc1, bc2, bl1, bl2) -> str:
  for fita in [bc1, bc2, bl1, bl2]:
    if not fita:
      continue
    fita_upper = fita.upper()
    if "0.45" in fita_upper or "0,45" in fita_upper:
      return "0.45mm"
    elif (
        "1X" in fita_upper
        or "1.00" in fita_upper
        or "1MM" in fita_upper
        or "1,00" in fita_upper
    ):
      return "1mm"
    elif "2X" in fita_upper or "2MM" in fita_upper:
      return "2mm"
  return ""


def extrair_referencias(item) -> dict:
  refs = {}
  ref_elem = item.find("REFERENCES")
  if ref_elem is not None:
    for child in ref_elem:
      refs[child.tag] = child.attrib.get("REFERENCE", "")
  return refs


def deve_ignorar_item(item, refs: dict) -> bool:
  """Decide se um <ITEM> do XML deve ficar de fora do plano de corte.

  Cobre: itens sem descrição / que são apenas estrutura de montagem (STRUCTURE
  == 'Y'), famílias inteiras que não são MDF cortado (ver FAMILIAS_IGNORADAS),
  e sub-itens sem material/código que só existem para compor o preço (ex.
  bases de caixa e gavetas sem geometria própria).
  """
  desc = item.attrib.get("DESCRIPTION", "").strip()
  family = item.attrib.get("FAMILY", "").strip()
  structure = item.attrib.get("STRUCTURE", "").strip()
  is_geom = item.attrib.get("ISGEOMETRY", "").strip()

  if not desc or structure == "Y":
    return True

  if family in FAMILIAS_IGNORADAS:
    return True

  material = refs.get("MATERIAL", "")
  has_code_mat = bool(refs.get("CODE_MAT"))
  item_base = refs.get("ITEM_BASE", "")

  if (
      not has_code_mat
      and is_geom != "Y"
      and not material
      and (
          item_base.startswith("BAL")
          or item_base in ["CGAV", ""]
          or item.attrib.get("ID") == "GROUPENTITY"
      )
  ):
    return True

  return False


def extrair_modulo_pai(item, parent_map: dict) -> str:
  curr = parent_map.get(item)
  while curr is not None:
    tag = curr.tag
    desc = curr.attrib.get("DESCRIPTION", "").strip()
    if tag == "ITEM" and desc:
      return desc
    if tag == "CATEGORY" and desc:
      return desc
    curr = parent_map.get(curr)
  return ""


def extrair_cliente_ambiente_xml(xml_path: str):
  try:
    tree = ET.parse(xml_path)
    root = tree.getroot()
    nome_cliente = ""
    cust_data = root.find("CUSTOMERSDATA")
    if cust_data is not None:
      for data in cust_data.findall("DATA"):
        data_id = data.attrib.get("ID", "")
        val = data.attrib.get("VALUE", "").strip()
        if (
            data_id in ["nomecliente", "corporateName", "nickName"]
            and val
            and not nome_cliente
        ):
          nome_cliente = val
    nome_ambiente = ""
    ambients_sec = root.find("AMBIENTS")
    if ambients_sec is not None:
      ambient_node = ambients_sec.find("AMBIENT")
      if ambient_node is not None:
        desc = ambient_node.attrib.get("DESCRIPTION", "").strip()
        if desc.startswith("Projeto - "):
          desc = desc.replace("Projeto - ", "").strip()
        nome_ambiente = desc
    if not nome_ambiente and cust_data is not None:
      for data in cust_data.findall("DATA"):
        if data.attrib.get("ID") == "Environment":
          nome_ambiente = data.attrib.get("VALUE", "").strip()
    return nome_cliente, nome_ambiente
  except Exception as e:
    print(f"Aviso: falha ao ler cliente/ambiente de '{xml_path}': {e}")
    return "", ""


def obter_caminho_base() -> str:
  """Pasta do executável empacotado (PyInstaller) ou do script .py."""
  if getattr(sys, "frozen", False):
    return os.path.dirname(sys.executable)
  return os.path.dirname(os.path.abspath(__file__))


def carregar_config() -> dict:
  caminho = os.path.join(obter_caminho_base(), "config_app.json")
  if os.path.exists(caminho):
    try:
      with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)
    except Exception:
      return {}
  return {}


def salvar_config(novos_valores: dict) -> None:
  caminho = os.path.join(obter_caminho_base(), "config_app.json")
  try:
    atual = carregar_config()
    atual.update(novos_valores)
    with open(caminho, "w", encoding="utf-8") as f:
      json.dump(atual, f, ensure_ascii=False, indent=2)
  except Exception as e:
    print(f"Erro ao salvar configuração: {e}")


def _versao_para_tupla(versao_str: str) -> tuple:
  try:
    return tuple(int(p) for p in versao_str.strip().split("."))
  except (ValueError, AttributeError):
    return (0,)


def verificar_nova_versao():
  """Consulta o manifesto remoto (versao.json).

  Retorna (versao_nova, url_download) se houver versão mais nova que
  VERSAO_ATUAL, ou None se estiver atualizado / a checagem falhar (sem
  internet, repositório fora do ar etc. — nunca levanta exceção).
  """
  try:
    with urllib.request.urlopen(URL_VERSAO_REMOTA, timeout=6) as resp:
      dados = json.loads(resp.read().decode("utf-8"))
    versao_remota = str(dados.get("versao", "")).strip()
    url_download = str(dados.get("url_download", "")).strip()
    if not versao_remota or not url_download:
      return None
    if _versao_para_tupla(versao_remota) > _versao_para_tupla(VERSAO_ATUAL):
      return versao_remota, url_download
    return None
  except Exception as e:
    print(f"Aviso: falha ao verificar atualização: {e}")
    return None


def baixar_atualizacao(url_download: str, destino: str, progresso_callback=None):
  """Baixa o novo executável para 'destino'. Levanta exceção se falhar."""
  req = urllib.request.Request(url_download, headers={"User-Agent": "Mozilla/5.0"})
  with urllib.request.urlopen(req, timeout=30) as resp:
    total = int(resp.headers.get("Content-Length", 0))
    lido = 0
    with open(destino, "wb") as f:
      while True:
        bloco = resp.read(65536)
        if not bloco:
          break
        f.write(bloco)
        lido += len(bloco)
        if progresso_callback:
          progresso_callback(lido, total)

  if os.path.getsize(destino) < 100_000:
    with open(destino, "rb") as f:
      inicio = f.read(300).lower()
    os.remove(destino)
    if b"<html" in inicio or b"<!doctype" in inicio:
      raise RuntimeError(
          "O link de download retornou uma página HTML em vez do arquivo"
          " .exe. Verifique se 'url_download' no versao.json é um link"
          " direto de download."
      )
    raise RuntimeError("Arquivo baixado parece inválido (muito pequeno).")


def aplicar_atualizacao_e_reiniciar(caminho_novo_exe: str):
  """Substitui o .exe em execução pelo baixado e reabre o app (Windows).

  Um .exe não pode se auto-sobrescrever enquanto roda, então isso grava um
  .bat que espera este processo terminar, move o arquivo novo por cima do
  antigo, reabre o app e se autodestrói.
  """
  exe_atual = sys.executable
  pasta = os.path.dirname(exe_atual)
  bat_path = os.path.join(pasta, "_atualizar.bat")
  pid_atual = os.getpid()

  conteudo_bat = (
      "@echo off\n"
      ":espera\n"
      f'tasklist /FI "PID eq {pid_atual}" 2>NUL | find /I "{pid_atual}" >NUL\n'
      "if not errorlevel 1 (\n"
      "    timeout /t 1 /nobreak >NUL\n"
      "    goto espera\n"
      ")\n"
      f'move /Y "{caminho_novo_exe}" "{exe_atual}" >NUL\n'
      f'start "" "{exe_atual}"\n'
      'del "%~f0"\n'
  )
  with open(bat_path, "w", encoding="utf-8") as f:
    f.write(conteudo_bat)

  subprocess.Popen(
      ["cmd", "/c", bat_path],
      creationflags=subprocess.CREATE_NO_WINDOW,
      close_fds=True,
  )
  os._exit(0)


class JanelaTutorial(ctk.CTkToplevel):
  """Wizard de ajuda passo a passo, reaberto pelo botão '❓ Ajuda'."""

  PASSOS = [
      {
          "titulo": "👋 Bem-vindo(a) ao Conversor XML Promob / Start KNR",
          "texto": (
              "Este aplicativo transforma o orçamento exportado do"
              " Promob/Start KNR (arquivo .xml) em uma planilha de corte"
              " pronta para enviar para a fábrica.\n\n"
              "Este tutorial rápido mostra o passo a passo. Você pode"
              " reabri-lo quando quiser clicando no botão '❓ Ajuda'."
          ),
      },
      {
          "titulo": "🏠 Tela Início",
          "texto": (
              "Aqui você:\n\n"
              "• Clica em '➕ Criar Novo Projeto' para começar um projeto"
              " novo.\n"
              "• Vê o histórico dos últimos projetos convertidos e pode"
              " '✏️ Editar / Reutilizar' um deles sem refazer tudo.\n"
              "• Pode '☁️ Fazer Backup' ou '📂 Restaurar' o histórico em um"
              " arquivo .zip — útil ao trocar de computador."
          ),
      },
      {
          "titulo": "📝 Cadastro do Projeto",
          "texto": (
              "Confirme (ou digite) o nome do Cliente e do Ambiente"
              " (ex: 'Cozinha', 'Quarto Bebê'). Esses nomes entram no nome"
              " do arquivo Excel gerado e, dependendo da fábrica, na"
              " etiqueta das peças."
          ),
      },
      {
          "titulo": "⚙️ Configuração do Projeto",
          "texto": (
              "1. Selecione o arquivo .xml exportado do Promob/Start KNR.\n"
              "2. Escolha a fábrica: Teletintas, Madefer, Tobias e"
              " TMKPlanilha usam a mesma planilha e podem ser salvas juntas"
              " depois; TMKCloud usa um formato próprio e fica sozinha.\n"
              "3. Escolha a pasta onde o Excel será salvo (por padrão, a"
              " mesma pasta do XML).\n\n"
              "Marque 'Agrupar peças iguais' se quiser somar automaticamente"
              " peças idênticas em uma só linha."
          ),
      },
      {
          "titulo": "🔍 Revisão de Peças",
          "texto": (
              "Antes de gerar o Excel, você revisa todas as peças:\n\n"
              "• Duplo-clique numa célula edita na hora (Enter confirma, Esc"
              " cancela).\n"
              "• Clique no cabeçalho de uma coluna para ordenar.\n"
              "• Selecione linhas e use 'Excluir selecionadas' para tirá-las"
              " da exportação (dá pra 'Desfazer' depois).\n"
              "• Peças em vermelho ultrapassam o tamanho da chapa — use"
              " '⚠️ Próxima Excedente' para encontrá-las rápido.\n"
              "• Em '💾 Salvar para:', marque uma ou mais fábricas"
              " compatíveis para gerar todos os arquivos de uma vez."
          ),
      },
      {
          "titulo": "✅ Pronto!",
          "texto": (
              "Depois de salvar, o projeto entra no histórico da Tela"
              " Início, pronto para reabrir e ajustar quando precisar.\n\n"
              "Ficou com dúvida em qualquer tela? Clique no botão"
              " '❓ Ajuda' a qualquer momento para rever este tutorial."
          ),
      },
  ]

  def __init__(self, parent, passo_inicial=0):
    super().__init__(parent)
    self.title("Tutorial")
    self.geometry("560x460")
    self.resizable(False, False)
    self.transient(parent)
    self.grab_set()

    self.passo_atual = max(0, min(passo_inicial, len(self.PASSOS) - 1))

    self.lbl_progresso = ctk.CTkLabel(
        self, text="", font=ctk.CTkFont(size=11), text_color="gray"
    )
    self.lbl_progresso.pack(pady=(16, 0))

    self.lbl_titulo = ctk.CTkLabel(
        self,
        text="",
        font=ctk.CTkFont(size=17, weight="bold"),
        wraplength=500,
        justify="left",
    )
    self.lbl_titulo.pack(pady=(6, 10), padx=25, anchor="w")

    self.lbl_texto = ctk.CTkLabel(
        self,
        text="",
        font=ctk.CTkFont(size=13),
        wraplength=500,
        justify="left",
        anchor="nw",
    )
    self.lbl_texto.pack(pady=(0, 10), padx=30, fill="both", expand=True)

    self.frame_pontos = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_pontos.pack(pady=(0, 8))
    self.pontos = []
    for _ in self.PASSOS:
      pt = ctk.CTkLabel(self.frame_pontos, text="●", font=ctk.CTkFont(size=14))
      pt.pack(side="left", padx=3)
      self.pontos.append(pt)

    self.var_nao_mostrar = ctk.BooleanVar(
        value=carregar_config().get("tutorial_visto", False)
    )
    self.chk_nao_mostrar = ctk.CTkCheckBox(
        self,
        text="Não mostrar automaticamente ao abrir o app",
        variable=self.var_nao_mostrar,
        font=ctk.CTkFont(size=11),
    )
    self.chk_nao_mostrar.pack(pady=(0, 10))

    self.frame_nav = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_nav.pack(pady=(0, 16), padx=20, fill="x")

    self.btn_pular = ctk.CTkButton(
        self.frame_nav,
        text="Pular",
        width=90,
        fg_color="#7f8c8d",
        hover_color="#95a5a6",
        command=self._fechar,
    )
    self.btn_pular.pack(side="left")

    self.btn_anterior = ctk.CTkButton(
        self.frame_nav, text="⬅️ Anterior", width=110, command=self._anterior
    )
    self.btn_anterior.pack(side="left", padx=(10, 0))

    self.btn_proximo = ctk.CTkButton(
        self.frame_nav,
        text="Próximo ➔",
        width=110,
        fg_color="#27ae60",
        hover_color="#219150",
        command=self._proximo,
    )
    self.btn_proximo.pack(side="right")

    self._atualizar_passo()

  def _atualizar_passo(self):
    passo = self.PASSOS[self.passo_atual]
    self.lbl_progresso.configure(
        text=f"Passo {self.passo_atual + 1} de {len(self.PASSOS)}"
    )
    self.lbl_titulo.configure(text=passo["titulo"])
    self.lbl_texto.configure(text=passo["texto"])

    for i, pt in enumerate(self.pontos):
      pt.configure(text_color="#2980b9" if i == self.passo_atual else "gray")

    self.btn_anterior.configure(
        state="normal" if self.passo_atual > 0 else "disabled"
    )
    ultimo = self.passo_atual == len(self.PASSOS) - 1
    self.btn_proximo.configure(text="Concluir ✅" if ultimo else "Próximo ➔")

  def _anterior(self):
    if self.passo_atual > 0:
      self.passo_atual -= 1
      self._atualizar_passo()

  def _proximo(self):
    if self.passo_atual < len(self.PASSOS) - 1:
      self.passo_atual += 1
      self._atualizar_passo()
    else:
      self._fechar()

  def _fechar(self):
    if self.var_nao_mostrar.get():
      salvar_config({"tutorial_visto": True})
    self.destroy()


class JanelaRevisao(ctk.CTkToplevel):

  def __init__(self, parent, df: pd.DataFrame, callback_salvar, fornecedor_atual="Teletintas"):
    super().__init__(parent)
    self.title("Revisão de Peças para Exportação")
    self.geometry("1120 x 660")
    self.minsize(940, 520)

    self.parent = parent
    self.df = df.copy()
    self.callback_salvar = callback_salvar

    self.transient(parent)
    self.grab_set()

    self.lbl_title = ctk.CTkLabel(
        self,
        text="✏️ Tela de Revisão dos Dados",
        font=ctk.CTkFont(size=18, weight="bold"),
    )
    self.lbl_title.pack(pady=(12, 2))

    self.lbl_inst = ctk.CTkLabel(
        self,
        text=(
            "💡 Dê dois cliques numa célula para editar direto na tabela"
            " (Enter confirma, Esc cancela). Clique no cabeçalho para"
            " ordenar. Passe o mouse nos cabeçalhos para ver a instrução da"
            " coluna."
        ),
        text_color="gray",
        font=ctk.CTkFont(size=11),
    )
    self.lbl_inst.pack(pady=(0, 4))

    self.lbl_alerta_dim = ctk.CTkLabel(
        self,
        text="",
        font=ctk.CTkFont(size=12, weight="bold"),
        text_color="#e74c3c",
    )
    self.lbl_alerta_dim.pack(pady=(0, 6))

    self.frame_tabla = ctk.CTkFrame(self)
    self.frame_tabla.pack(padx=15, pady=5, fill="both", expand=True)

    self.columns = list(self.df.columns)
    self.tree = ttk.Treeview(
        self.frame_tabla, columns=self.columns, show="headings", height=15
    )

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", rowheight=26, font=("Arial", 10))
    style.configure(
        "Treeview.Heading", font=("Arial", 10, "bold"), background="#34495e"
    )

    self.tree.tag_configure(
        "excedente", background="#ffcccc", foreground="#900c3f"
    )

    self._coluna_ordenada = None
    self._ordem_crescente = True

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

    self.tree.bind("<Double-1>", self.on_double_click)
    self.tree.bind("<Motion>", self.on_mouse_hover)
    self.tree.bind("<<TreeviewSelect>>", self.atualizar_contagem_selecionadas)

    self.tooltip_label = None
    self._editor_inline = None
    self._ultima_exclusao = None

    self.frame_status = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_status.pack(pady=(2, 0), padx=20, fill="x")

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

    self.frame_ferramentas = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_ferramentas.pack(pady=(6, 0), padx=20, fill="x")

    self.btn_selecionar_todas = ctk.CTkButton(
        self.frame_ferramentas,
        text="Selecionar Todas",
        fg_color="#7f8c8d",
        hover_color="#95a5a6",
        font=ctk.CTkFont(size=11),
        height=28,
        width=130,
        command=self.selecionar_todas,
    )
    self.btn_selecionar_todas.pack(side="left", padx=(0, 5))

    self.btn_limpar_selecao = ctk.CTkButton(
        self.frame_ferramentas,
        text="Limpar Seleção",
        fg_color="#7f8c8d",
        hover_color="#95a5a6",
        font=ctk.CTkFont(size=11),
        height=28,
        width=130,
        command=self.limpar_selecao,
    )
    self.btn_limpar_selecao.pack(side="left", padx=5)

    self.btn_prox_excedente = ctk.CTkButton(
        self.frame_ferramentas,
        text="⚠️ Próxima Excedente",
        fg_color="#e67e22",
        hover_color="#d35400",
        font=ctk.CTkFont(size=11),
        height=28,
        width=160,
        command=self.ir_para_proxima_excedente,
    )
    self.btn_prox_excedente.pack(side="left", padx=5)

    self.btn_ajuda = ctk.CTkButton(
        self.frame_ferramentas,
        text="❓ Ajuda",
        fg_color="#8e44ad",
        hover_color="#9b59b6",
        font=ctk.CTkFont(size=11),
        height=28,
        width=90,
        command=lambda: JanelaTutorial(self, passo_inicial=4),
    )
    self.btn_ajuda.pack(side="right")

    self.carregar_dados_tree()

    self.frame_btns = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_btns.pack(pady=12, padx=20, fill="x")

    self.btn_cancelar = ctk.CTkButton(
        self.frame_btns,
        text="Cancelar",
        fg_color="#e74c3c",
        hover_color="#c0392b",
        command=self.destroy,
        width=110,
    )
    self.btn_cancelar.pack(side="left", padx=5)

    self.btn_excluir = ctk.CTkButton(
      self.frame_btns,
      text="Excluir selecionadas",
      fg_color="#c0392b",
      hover_color="#a93226",
      command=self.excluir_selecionadas,
      width=150,
    )
    self.btn_excluir.pack(side="left", padx=5)

    self.btn_desfazer = ctk.CTkButton(
        self.frame_btns,
        text="↩️ Desfazer Exclusão",
        fg_color="#7f8c8d",
        hover_color="#95a5a6",
        command=self.desfazer_exclusao,
        width=160,
        state="disabled",
    )
    self.btn_desfazer.pack(side="left", padx=5)

    # Seleção de fábrica(s) para salvar simultaneamente. Só ficam disponíveis
    # as fábricas que usam o mesmo layout de colunas que já está em revisão
    # (TMKCloud tem um layout próprio, então aparece sozinho nesse caso).
    self.frame_sel_forn = ctk.CTkFrame(self.frame_btns, fg_color="transparent")
    self.frame_sel_forn.pack(side="left", expand=True)

    self.lbl_forn_escolha = ctk.CTkLabel(
        self.frame_sel_forn,
        text="💾 Salvar para:",
        font=ctk.CTkFont(size=12, weight="bold"),
    )
    self.lbl_forn_escolha.pack(side="left", padx=(5, 8))

    grupo_compativel = (
        FORNECEDORES_FORMATO_TMK
        if fornecedor_atual in FORNECEDORES_FORMATO_TMK
        else FORNECEDORES_FORMATO_PADRAO
    )
    self.fornecedores_vars = {}
    for nome_forn in grupo_compativel:
      var = ctk.BooleanVar(value=(nome_forn == fornecedor_atual))
      chk = ctk.CTkCheckBox(
          self.frame_sel_forn,
          text=nome_forn,
          variable=var,
          font=ctk.CTkFont(size=11),
          checkbox_width=18,
          checkbox_height=18,
      )
      chk.pack(side="left", padx=4)
      self.fornecedores_vars[nome_forn] = var

    self.btn_confirmar = ctk.CTkButton(
        self.frame_btns,
        text="✅ Confirmar e Salvar",
        fg_color="#27ae60",
        hover_color="#219150",
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
        text=f"📦 Total de peças: {total_pecas}  ({len(linhas)} linha(s))"
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
      col_idx = int(col_id.replace("#", "")) - 1
      col_nome = self.columns[col_idx]
      dica = DICAS_COLUNAS.get(col_nome, "Clique para editar os dados")

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
      self.tooltip_label.config(text=f"{col_nome.replace(chr(10), ' ')}: {dica}")
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

  def desfazer_exclusao(self):
    if self._ultima_exclusao is None or self._ultima_exclusao.empty:
      return
    self.df = pd.concat([self.df, self._ultima_exclusao]).sort_index()
    self._ultima_exclusao = None
    self.btn_desfazer.configure(state="disabled")
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
        nome for nome, var in self.fornecedores_vars.items() if var.get()
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
    self.historico = self.carregar_historico()

    self.frame_topo = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_topo.pack(side="top", fill="x", padx=15, pady=(8, 0))

    self.btn_ajuda = ctk.CTkButton(
        self.frame_topo,
        text="❓ Ajuda / Tutorial",
        width=150,
        height=30,
        fg_color="#8e44ad",
        hover_color="#9b59b6",
        font=ctk.CTkFont(size=12, weight="bold"),
        command=self.abrir_tutorial,
    )
    self.btn_ajuda.pack(side="right")

    self.btn_verificar_update = ctk.CTkButton(
        self.frame_topo,
        text="🔄 Verificar Atualizações",
        width=180,
        height=30,
        fg_color="#2980b9",
        hover_color="#3498db",
        font=ctk.CTkFont(size=12, weight="bold"),
        command=self.verificar_atualizacoes_manual,
    )
    self.btn_verificar_update.pack(side="right", padx=(0, 8))

    self.container = ctk.CTkFrame(self, fg_color="transparent")
    self.container.pack(fill="both", expand=True, padx=10, pady=10)
    self.container.grid_rowconfigure(0, weight=1)
    self.container.grid_columnconfigure(0, weight=1)

    self.lbl_footer = ctk.CTkLabel(
        self,
        text=f"Desenvolvido por Ariane Prado · v{VERSAO_ATUAL}",
        font=ctk.CTkFont(size=11, weight="bold"),
        text_color="gray",
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
    JanelaTutorial(self, passo_inicial=passo_inicial)

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
    self.btn_verificar_update.configure(
        state="disabled", text="🔄 Verificando..."
    )

    def worker():
      resultado = verificar_nova_versao()
      self.after(0, lambda: self._resultado_verificacao_manual(resultado))

    threading.Thread(target=worker, daemon=True).start()

  def _resultado_verificacao_manual(self, resultado):
    self.btn_verificar_update.configure(
        state="normal", text="🔄 Verificar Atualizações"
    )
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
        janela, text="0%", font=ctk.CTkFont(size=11), text_color="gray"
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

  def carregar_historico(self):
    if os.path.exists(self.arquivo_historico):
      try:
        with open(self.arquivo_historico, "r", encoding="utf-8") as f:
          return json.load(f)
      except Exception:
        return []
    return []

  def salvar_historico(self):
    try:
      with open(self.arquivo_historico, "w", encoding="utf-8") as f:
        json.dump(self.historico[:15], f, ensure_ascii=False, indent=2)
    except Exception as e:
      print(f"Erro ao salvar histórico: {e}")

  def adicionar_ao_historico(self, caminho_xml, caminho_excel, fornecedor):
    novo_registro = {
        "nome_xml": os.path.basename(caminho_xml),
        "caminho_xml": caminho_xml,
        "caminho_excel": caminho_excel,
        "cliente": self.cliente_nome,
        "ambiente": self.ambiente_nome,
        "fornecedor": fornecedor,
        "data_hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    self.historico = [
        item
        for item in self.historico
        if item["caminho_xml"] != caminho_xml
        or item["caminho_excel"] != caminho_excel
    ]
    self.historico.insert(0, novo_registro)
    self.historico = self.historico[:15]
    self.salvar_historico()

    if "TelaInicio" in self.frames:
      self.frames["TelaInicio"].atualizar_interface_historico()


class TelaInicio(ctk.CTkFrame):

  def __init__(self, parent, controller):
    super().__init__(parent, fg_color="transparent")
    self.controller = controller

    self.lbl_title = ctk.CTkLabel(
        self,
        text="Conversor de XML Promob / Start KNR",
        font=ctk.CTkFont(size=24, weight="bold"),
    )
    self.lbl_title.pack(pady=(20, 2))

    self.lbl_sub = ctk.CTkLabel(
        self,
        text=(
            "Sistema de conversão automatizada de arquivos XML para planos de"
            " corte."
        ),
        font=ctk.CTkFont(size=13),
        text_color="gray",
    )
    self.lbl_sub.pack(pady=(0, 10))

    self.btn_novo = ctk.CTkButton(
        self,
        text="➕ Criar Novo Projeto",
        font=ctk.CTkFont(size=16, weight="bold"),
        fg_color="#27ae60",
        hover_color="#219150",
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
        text="📋 Projetos Anteriores & Histórico de Clientes",
        font=ctk.CTkFont(size=14, weight="bold"),
        anchor="w",
    )
    self.lbl_recentes_title.pack(padx=15, pady=(10, 5), fill="x")

    self.scroll_recentes = ctk.CTkScrollableFrame(self.frame_recentes)
    self.scroll_recentes.pack(padx=10, pady=(0, 10), fill="both", expand=True)

    self.frame_backup_inferior = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_backup_inferior.pack(padx=20, pady=(5, 10), fill="x")

    self.btn_backup = ctk.CTkButton(
        self.frame_backup_inferior,
        text="☁️ Fazer Backup dos Dados",
        font=ctk.CTkFont(size=12, weight="bold"),
        fg_color="#2980b9",
        hover_color="#3498db",
        height=38,
        command=self.fazer_backup,
    )
    self.btn_backup.pack(side="left", expand=True, fill="x", padx=5)

    self.btn_restaurar = ctk.CTkButton(
        self.frame_backup_inferior,
        text="📂 Restaurar Dados",
        font=ctk.CTkFont(size=12, weight="bold"),
        fg_color="#8e44ad",
        hover_color="#9b59b6",
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
      self.controller.show_frame("TelaConfiguracaoProjeto")

  def fazer_backup(self):
    try:
      historico_path = self.controller.arquivo_historico
      if not os.path.exists(historico_path):
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

    historico = self.controller.historico
    if not historico:
      lbl_vazio = ctk.CTkLabel(
          self.scroll_recentes,
          text="Nenhum projeto registrado no histórico.",
          text_color="gray",
      )
      lbl_vazio.pack(pady=20)
      return

    for item in historico:
      item_frame = ctk.CTkFrame(self.scroll_recentes, fg_color="transparent")
      item_frame.pack(fill="x", pady=4, padx=5)

      cli_str = f"[{item.get('cliente', 'Cliente')} - {item.get('ambiente', 'Ambiente')}]"
      txt_label = f"📁 {cli_str} {item['nome_xml']} ➔ 📊 {os.path.basename(item['caminho_excel'])} ({item.get('fornecedor', 'Teletintas')})"
      data_hora = item.get("data_hora", "")
      if data_hora:
        txt_label += f"  🕒 {data_hora}"

      lbl_info = ctk.CTkLabel(
          item_frame, text=txt_label, font=ctk.CTkFont(size=12), anchor="w"
      )
      lbl_info.pack(side="left", fill="x", expand=True)

      caminho_pasta = os.path.dirname(item["caminho_excel"])

      btn_abrir_pasta = ctk.CTkButton(
          item_frame,
          text="Abrir Pasta",
          width=85,
          height=26,
          font=ctk.CTkFont(size=11),
          fg_color="#4b6584",
          hover_color="#3867d6",
          command=lambda p=caminho_pasta: os.startfile(p)
          if os.name == "nt"
          else os.system(f'open "{p}"'),
      )
      btn_abrir_pasta.pack(side="right", padx=4)

      btn_editar = ctk.CTkButton(
          item_frame,
          text="✏️ Editar / Reutilizar",
          width=120,
          height=26,
          font=ctk.CTkFont(size=11),
          fg_color="#2980b9",
          hover_color="#3498db",
          command=lambda it=item: self.editar_projeto_historico(it),
      )
      btn_editar.pack(side="right", padx=4)


class TelaNovoProjeto(ctk.CTkFrame):

  def __init__(self, parent, controller):
    super().__init__(parent, fg_color="transparent")
    self.controller = controller

    self.lbl_title = ctk.CTkLabel(
        self,
        text="📝 Cadastro do Projeto",
        font=ctk.CTkFont(size=20, weight="bold"),
    )
    self.lbl_title.pack(pady=(30, 5))

    self.lbl_sub = ctk.CTkLabel(
        self,
        text="Confirme ou altere o nome do cliente e do ambiente.",
        text_color="gray",
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

    self.frame_nav = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_nav.pack(pady=25, fill="x", padx=30)

    self.btn_voltar = ctk.CTkButton(
        self.frame_nav,
        text="⬅️ Voltar",
        width=120,
        height=42,
        fg_color="#7f8c8d",
        hover_color="#95a5a6",
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

    self.frame_header_info = ctk.CTkFrame(self, fg_color="#2c3e50")
    self.frame_header_info.pack(padx=20, pady=(10, 15), fill="x")

    self.lbl_info_header = ctk.CTkLabel(
        self.frame_header_info,
        text="👤 Cliente: -   |   🏠 Ambiente: -",
        font=ctk.CTkFont(size=14, weight="bold"),
        text_color="white",
    )
    self.lbl_info_header.pack(pady=12, padx=15)

    self.frame_xml = ctk.CTkFrame(self)
    self.frame_xml.pack(padx=20, pady=6, fill="x")

    self.btn_select_xml = ctk.CTkButton(
        self.frame_xml,
        text="1. Selecionar XML",
        command=self.selecionar_xml,
        width=170,
        height=38,
        font=ctk.CTkFont(size=13, weight="bold"),
    )
    self.btn_select_xml.pack(side="left", padx=12, pady=10)

    self.lbl_xml_status = ctk.CTkLabel(
        self.frame_xml, text="Nenhum arquivo XML selecionado", anchor="w"
    )
    self.lbl_xml_status.pack(
        side="left", padx=10, pady=10, fill="x", expand=True
    )

    self.frame_forn = ctk.CTkFrame(self)
    self.frame_forn.pack(padx=20, pady=6, fill="x")

    self.lbl_forn = ctk.CTkLabel(
        self.frame_forn,
        text="2. Formato Inicial de Fábrica:",
        font=ctk.CTkFont(size=13, weight="bold"),
    )
    self.lbl_forn.pack(side="left", padx=12, pady=10)

    self.fornecedor_var = ctk.StringVar(value="Teletintas")
    self.cmb_fornecedor = ctk.CTkOptionMenu(
        self.frame_forn,
        values=FORNECEDORES_DISPONIVEIS,
        variable=self.fornecedor_var,
        width=180,
        height=34,
    )
    self.cmb_fornecedor.pack(side="left", padx=10, pady=10)

    self.lbl_forn_dica = ctk.CTkLabel(
        self,
        text=(
            "💡 Teletintas, Madefer, Tobias e TMKPlanilha usam a mesma"
            " planilha e poderão ser marcadas juntas na revisão. TMKCloud"
            " usa um formato de colunas próprio e fica disponível sozinha."
        ),
        text_color="gray",
        font=ctk.CTkFont(size=11),
        justify="left",
        wraplength=760,
        anchor="w",
    )
    self.lbl_forn_dica.pack(padx=25, pady=(0, 8), fill="x")

    self.frame_folder = ctk.CTkFrame(self)
    self.frame_folder.pack(padx=20, pady=6, fill="x")

    self.btn_select_folder = ctk.CTkButton(
        self.frame_folder,
        text="3. Escolher Onde Salvar",
        command=self.selecionar_pasta_destino,
        width=170,
        height=38,
        font=ctk.CTkFont(size=13, weight="bold"),
        fg_color="#34495e",
        hover_color="#2c3e50",
    )
    self.btn_select_folder.pack(side="left", padx=12, pady=10)

    self.lbl_folder_status = ctk.CTkLabel(
        self.frame_folder,
        text="Salvar na mesma pasta do XML (Padrão)",
        anchor="w",
    )
    self.lbl_folder_status.pack(
        side="left", padx=10, pady=10, fill="x", expand=True
    )

    self.chk_agrupar_var = ctk.BooleanVar(value=False)
    self.chk_agrupar = ctk.CTkCheckBox(
        self,
        text="Agrupar peças iguais (soma quantidades)",
        variable=self.chk_agrupar_var,
        font=ctk.CTkFont(size=12, weight="bold"),
    )
    self.chk_agrupar.pack(pady=10, padx=25, anchor="w")

    self.btn_convert = ctk.CTkButton(
        self,
        text="🔍 Revisar e Exportar para Excel",
        command=self.iniciar_revisao,
        fg_color="#27ae60",
        hover_color="#219150",
        height=50,
        font=ctk.CTkFont(size=15, weight="bold"),
        state="disabled",
    )
    self.btn_convert.pack(pady=15, padx=20, fill="x")

    self.frame_botoes_nav = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_botoes_nav.pack(padx=20, pady=5, fill="x")

    self.btn_voltar = ctk.CTkButton(
        self.frame_botoes_nav,
        text="⬅️ Voltar",
        width=120,
        height=36,
        fg_color="#7f8c8d",
        hover_color="#95a5a6",
        command=lambda: controller.show_frame("TelaNovoProjeto"),
    )
    self.btn_voltar.pack(side="left")

    self.btn_editar_cad = ctk.CTkButton(
        self.frame_botoes_nav,
        text="✏️ Alterar Cliente/Ambiente",
        width=180,
        height=36,
        fg_color="#2980b9",
        hover_color="#3498db",
        command=lambda: controller.show_frame("TelaNovoProjeto"),
    )
    self.btn_editar_cad.pack(side="right")

  def ao_exibir_tela(self):
    cli = self.controller.cliente_nome or "Não informado"
    amb = self.controller.ambiente_nome or "Não informado"
    self.lbl_info_header.configure(
        text=f"👤 Cliente: {cli}   |   🏠 Ambiente: {amb}"
    )
    self.fornecedor_var.set(
        getattr(self.controller, "fornecedor_inicial", "Teletintas")
    )

    if self.controller.xml_path and os.path.exists(self.controller.xml_path):
      filename = os.path.basename(self.controller.xml_path)
      self.lbl_xml_status.configure(text=f"📄 {filename}", text_color="#2980b9")
      self.btn_convert.configure(state="normal")
    else:
      self.lbl_xml_status.configure(text="Nenhum arquivo XML selecionado")
      self.btn_convert.configure(state="disabled")

  def selecionar_xml(self):
    filepath = filedialog.askopenfilename(
        title="Selecione o arquivo XML",
        filetypes=[("Arquivos XML", "*.xml"), ("Todos os Arquivos", "*.*")],
    )
    if filepath:
      self.controller.xml_path = filepath
      filename = os.path.basename(filepath)
      self.lbl_xml_status.configure(
          text=f"📄 {filename}", text_color="#2980b9"
      )

      if not self.controller.pasta_destino:
        self.controller.pasta_destino = os.path.dirname(filepath)
        self.lbl_folder_status.configure(
            text=f"📁 {self.controller.pasta_destino}", text_color="#7f8c8d"
        )

      cli_xml, amb_xml = extrair_cliente_ambiente_xml(filepath)
      if not self.controller.cliente_nome and cli_xml:
        self.controller.cliente_nome = cli_xml
      if not self.controller.ambiente_nome and amb_xml:
        self.controller.ambiente_nome = amb_xml

      cli = self.controller.cliente_nome or "Não informado"
      amb = self.controller.ambiente_nome or "Não informado"
      self.lbl_info_header.configure(
          text=f"👤 Cliente: {cli}   |   🏠 Ambiente: {amb}"
      )

      self.btn_convert.configure(state="normal")

  def selecionar_pasta_destino(self):
    folderpath = filedialog.askdirectory(
        title="Selecione a pasta onde deseja salvar o arquivo Excel"
    )
    if folderpath:
      self.controller.pasta_destino = folderpath
      self.lbl_folder_status.configure(
          text=f"📁 {folderpath}", text_color="#27ae60"
      )

  def processar_dados_para_revisao(
      self, xml_path, fornecedor_alvo="Teletintas"
  ):
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

    cli_xml, amb_xml = extrair_cliente_ambiente_xml(xml_path)
    nome_cliente = self.controller.cliente_nome or cli_xml or "Cliente"
    nome_ambiente = self.controller.ambiente_nome or amb_xml or "Ambiente"

    rows_padrao = []
    rows_tmk = []
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

      row_p = {
          "Quantidade": int(float(item.attrib.get("QUANTITY", "1"))),
          "Comprimento": float(item.attrib.get("WIDTH", "0")),
          "Largura": float(item.attrib.get("DEPTH", "0")),
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
          "COMP\n2750": float(item.attrib.get("WIDTH", "0")),
          "LARG\n1840": float(item.attrib.get("DEPTH", "0")),
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

    # CORREÇÃO: "TMK" foi alterado para "TMKCloud" para ativar a lógica específica
    if fornecedor_alvo == "TMKCloud":
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

    if self.chk_agrupar_var.get() and not df.empty:
      resposta = messagebox.askyesno(
          "Confirmar Agrupamento",
          "A opção de agrupar peças iguais está marcada.\n\nDeseja realmente"
          " agrupar as peças idênticas e somar suas quantidades?",
      )
      if resposta:
        # CORREÇÃO: "TMK" alterado para "TMKCloud"
        if fornecedor_alvo == "TMKCloud":
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
          ]
          df = (
              df.groupby(cols_agrupar, as_index=False, dropna=False)[
                  "QUANT"
              ].sum()
          )
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
          ]
          df = (
              df.groupby(cols_agrupar, as_index=False, dropna=False)[
                  "Quantidade"
              ].sum()
          )
          df = df[colunas_p]
    return df

  def iniciar_revisao(self):
    if not self.controller.xml_path:
      messagebox.showwarning(
          "Atenção", "Por favor, selecione um arquivo XML primeiro."
      )
      return

    texto_original_btn = self.btn_convert.cget("text")
    self.btn_convert.configure(state="disabled", text="⏳ Processando XML...")
    self.update_idletasks()

    try:
      forn_inicial = self.fornecedor_var.get().strip()
      df = self.processar_dados_para_revisao(
          self.controller.xml_path, fornecedor_alvo=forn_inicial
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

        # CORREÇÃO: "TMK" alterado para "TMKCloud"
        if fornecedor_escolhido == "TMKCloud":
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


if __name__ == "__main__":
  app = ConversorXmlExcelApp()
  app.mainloop()