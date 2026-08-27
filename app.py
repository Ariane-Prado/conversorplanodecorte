import datetime
import json
import os
import re
import subprocess
import sys
import threading
import unicodedata
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
import customtkinter as ctk
import openpyxl
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Configuração do tema visual da interface. Fixo em "dark": várias cores da
# interface (cards, botões, inputs) já são hardcoded pensando em fundo
# escuro — deixar em "System" fazia o contraste variar (e às vezes falhar)
# dependendo do tema do Windows do usuário.
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Paleta de contraste (WCAG AA) usada nos textos e inputs da interface.
COR_TITULO = "#FFFFFF"
COR_TEXTO_SECUNDARIO = "#E0E0E0"
COR_VERDE_PRIMARIO = "#1e7e34"  # mais escuro que o #27ae60 original: texto
# branco sobre esse verde passa de ~2.9:1 para ~5.1:1 de contraste (WCAG AA
# exige 4.5:1 pra texto normal).
COR_VERDE_PRIMARIO_HOVER = "#17692a"
COR_FUNDO_INPUT = "#2B2B2B"
COR_BORDA_INPUT = "#444444"
COR_FOCO_INPUT = "#2980b9"

VERSAO_ATUAL = "1.2.0"
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

# Módulos curvos em L (identificados pela expressão "Canto L" em algum
# ponto do nome do módulo — ex: "Mod. Inferior Canto L", "Mod. Canto L s/
# Rod Dormitorio") podem precisar ser cortados numa máquina terceirizada,
# que exige uma margem extra de medida. Quando terceirizado, soma-se essa
# margem na largura e no comprimento das peças de Base, Tampo e Prateleira
# do módulo. O usuário escolhe, por módulo, se a margem é de 50mm ou 100mm.
OPCOES_MARGEM_TERCEIRIZACAO_MM = [50, 100]
FUNCOES_AFETADAS_TERCEIRIZACAO = {"Base", "Tampo", "Prateleira"}

# "L" como palavra isolada (não início de "Lavatório" etc.) logo após
# "canto", com qualquer quantidade de espaço entre os dois.
_PADRAO_MODULO_CANTO_L = re.compile(r"\bcanto\s*l\b")


def _normalizar_texto(texto: str) -> str:
  nfkd = unicodedata.normalize("NFKD", texto or "")
  return "".join(c for c in nfkd if not unicodedata.combining(c)).strip().lower()


def eh_modulo_canto_l(nome_modulo: str) -> bool:
  return bool(_PADRAO_MODULO_CANTO_L.search(_normalizar_texto(nome_modulo)))


def formatar_medida_mm(valor: float) -> str:
  try:
    v = float(valor)
  except (TypeError, ValueError):
    return "-"
  return f"{int(v)}" if v.is_integer() else f"{v:.1f}"


def _encontrar_item_ancestral(item, parent_map: dict):
  """Retorna o <ITEM> ancestral mais próximo com DESCRIPTION (o módulo/caixa
  que contém a peça), ou None se só houver uma CATEGORY como ancestral."""
  curr = parent_map.get(item)
  while curr is not None:
    if curr.tag == "ITEM" and curr.attrib.get("DESCRIPTION", "").strip():
      return curr
    curr = parent_map.get(curr)
  return None


def listar_modulos_canto_l(xml_path: str) -> list:
  """Retorna, na ordem de aparição, os módulos de canto em L (nome contendo
  "Canto L") encontrados no XML, como dicts {"nome", "comprimento",
  "largura"} — comprimento/largura são a medida do módulo (WIDTH/DEPTH do
  próprio item), só como referência de tamanho para a decisão do usuário."""
  tree = ET.parse(xml_path)
  root = tree.getroot()
  ambients_section = root.find("AMBIENTS")
  items_source = (
      ambients_section.findall(".//ITEM")
      if ambients_section is not None
      else root.findall(".//ITEM")
  )
  parent_map = {c: p for p in root.iter() for c in p}
  vistos = set()
  modulos = []
  for item in items_source:
    refs = extrair_referencias(item)
    if deve_ignorar_item(item, refs):
      continue
    modulo_item = _encontrar_item_ancestral(item, parent_map)
    if modulo_item is None:
      continue
    modulo_nome = modulo_item.attrib.get("DESCRIPTION", "").strip()
    if not eh_modulo_canto_l(modulo_nome) or modulo_nome in vistos:
      continue
    vistos.add(modulo_nome)
    try:
      comprimento = float(modulo_item.attrib.get("WIDTH", "0"))
    except ValueError:
      comprimento = 0.0
    try:
      largura = float(modulo_item.attrib.get("DEPTH", "0"))
    except ValueError:
      largura = 0.0
    modulos.append(
        {"nome": modulo_nome, "comprimento": comprimento, "largura": largura}
    )
  return modulos


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
  if f.lower().startswith("divis"):
    return "Divisória"
  primeira_palavra = f.split()[0].lower() if f.split() else ""
  if primeira_palavra in ("chapeu", "chapéu", "tampo"):
    return "Tampo"
  if f.startswith("Lateral"):
    tokens = f.split()
    if "Dir" in tokens or "Direita" in tokens:
      return "Lateral Direita"
    if "Esq" in tokens or "Esquerda" in tokens:
      return "Lateral Esquerda"
  # Regra geral: qualquer peça sem regra específica acima mantém só a
  # primeira palavra, removendo sufixos de módulo/lado/ambiente/espessura
  # ("Fundo Mod Sup" -> "Fundo", "Painel 15mm c/ fita 1mm" -> "Painel").
  # Peças que precisarem manter o sufixo entram como exceção explícita.
  return f.split()[0] if f.split() else f


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
                "Teletintas, Madefer, Tobias e TMKPlanilha usam o mesmo"
                " layout de planilha. TMKCloud usa um formato próprio."
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
    self.callback_salvar = callback_salvar

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
    # mesmo layout de colunas que já está em revisão (TMKCloud tem um
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
        "Teletintas, Madefer, Tobias e TMKPlanilha usam a mesma planilha e"
        " poderão ser marcadas juntas na revisão. TMKCloud usa um formato"
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

      comprimento = float(item.attrib.get("WIDTH", "0"))
      largura = float(item.attrib.get("DEPTH", "0"))
      margem_terceirizacao = modulos_terceirizados.get(modulo_nome, 0)
      if margem_terceirizacao and funcao_formatada in FUNCOES_AFETADAS_TERCEIRIZACAO:
        comprimento += margem_terceirizacao
        largura += margem_terceirizacao

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