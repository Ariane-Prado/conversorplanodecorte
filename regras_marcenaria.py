"""Regras de negócio de marcenaria e extração de dados do XML Promob / Start KNR.

Concentra toda a inteligência de domínio:
- Renomeação padronizada das funções de peças (`renomear_funcao`)
- Extração de serviços adicionais como furos de dobradiça e rasgos (`extrair_servicos_adicionais`)
- Filtro de itens ignorados como ferragens, estruturas e parceiros (`deve_ignorar_item`)
- Identificação de módulos de Canto L para margem de terceirização (`listar_modulos_canto_l`)
- Regras de ajuste de conferência para Fundo 6mm (+6mm) e Contra Frente/Fundo (-3mm na altura)
- Extração de fabricantes, materiais, fitas de borda e cliente/ambiente do XML
"""
import re
import unicodedata
import xml.etree.ElementTree as ET
import pandas as pd

from constantes import (
    _PADRAO_ESPESSURA_MM,
    _PADRAO_MODULO_CANTO_L,
    _TOKENS_ARMARIO_BASE_CHAPEU,
    _TOKENS_BALCAO_INFERIOR,
    AJUSTE_CONTRA_FRENTE_ALTURA_DELTA,
    AJUSTE_FUNDO_8MM_DELTA,
    AVANCO_PADRAO_TOTAL_MM,
    FAMILIAS_IGNORADAS,
    FUNCAO_CONTRA_FRENTE_FUNDO,
    FUNCAO_FUNDO,
)


def _normalizar_texto(texto: str) -> str:
  nfkd = unicodedata.normalize("NFKD", texto or "")
  return "".join(c for c in nfkd if not unicodedata.combining(c)).strip().lower()


def _tokens_normalizados(texto: str) -> set:
  return set(re.findall(r"\w+", _normalizar_texto(texto)))


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


def _eh_espessura_6mm_texto(texto: str) -> bool:
  """Retorna True se `texto` indica 6mm (ou não especifica espessura)."""
  matches = _PADRAO_ESPESSURA_MM.findall(texto or "")
  if not matches:
    return True
  for m in matches:
    try:
      if abs(float(m.replace(",", ".")) - 6.0) < 1.0:
        return True
    except ValueError:
      pass
  return False


def eh_fundo_6mm(item, refs: dict, funcao_formatada: str = None) -> bool:
  """Verifica se um <ITEM> do XML representa uma peça de Fundo de 6mm."""
  desc = item.attrib.get("DESCRIPTION", "").strip()
  func = funcao_formatada if funcao_formatada is not None else renomear_funcao(desc)
  if func != FUNCAO_FUNDO:
    return False

  height_str = item.attrib.get("HEIGHT", "").strip()
  if height_str:
    try:
      return abs(float(height_str.replace(",", ".")) - 6.0) < 1.0
    except ValueError:
      pass

  material_txt = f"{refs.get('MATERIAL', '')} {desc}"
  return _eh_espessura_6mm_texto(material_txt)


def escolher_eixo_largura_fundo(
    comp_fundo: float,
    larg_fundo: float,
    medidas_referencia: list = None,
    delta: float = AJUSTE_FUNDO_8MM_DELTA,
) -> tuple:
  """Escolhe qual medida do Fundo (Comprimento ou Largura) corresponde à
  largura do módulo/gaveta (que já tem avanço de 5mm de cada lado = +10mm em
  relação à largura da Base ou da Contra Frente).

  Se `larg_fundo` bate com `medida_ref + 10mm` (±1.5mm) e `comp_fundo` não,
  aplica `delta` na Largura; caso contrário aplica no Comprimento (padrão
  WIDTH do XML no Promob).
  """
  refs = [float(r) for r in (medidas_referencia or []) if r]
  if refs:
    comp_bate = any(
        abs(comp_fundo - (r + AVANCO_PADRAO_TOTAL_MM)) <= 1.5 for r in refs
    )
    larg_bate = any(
        abs(larg_fundo - (r + AVANCO_PADRAO_TOTAL_MM)) <= 1.5 for r in refs
    )
    if larg_bate and not comp_bate:
      return (0.0, float(delta))
  return (float(delta), 0.0)


def calcular_delta_contra_frente(
    comp: float,
    larg: float,
    delta: float = AJUSTE_CONTRA_FRENTE_ALTURA_DELTA,
) -> tuple:
  """Retorna (delta_comp, delta_larg) para reduzir a altura (menor medida)
  da Contra Frente / Contra Fundo de Gaveta."""
  if larg <= comp:
    return (0.0, float(delta))
  return (float(delta), 0.0)


def mapear_contexto_modulos_xml(items_source: list, parent_map: dict) -> dict:
  """Agrupa, por módulo ancestral, as funções e medidas de Base e Contra Frente
  para apoiar a classificação dos Fundos de 6mm do projeto."""
  contexto = {}
  for item in items_source:
    refs = extrair_referencias(item)
    if deve_ignorar_item(item, refs):
      continue
    anc = _encontrar_item_ancestral(item, parent_map)
    chave_mod = (
        anc if anc is not None else extrair_modulo_pai(item, parent_map)
    )
    info = contexto.setdefault(
        chave_mod,
        {
            "funcoes_caixa": set(),
            "medidas_base": [],
            "larguras_cf": [],
        },
    )
    desc = item.attrib.get("DESCRIPTION", "").strip()
    func = renomear_funcao(desc)
    try:
      comp = float(item.attrib.get("WIDTH", "0"))
    except ValueError:
      comp = 0.0
    try:
      larg = float(item.attrib.get("DEPTH", "0"))
    except ValueError:
      larg = 0.0

    if "gav" not in _normalizar_texto(desc):
      info["funcoes_caixa"].add(func)
    if func == "Base" and (comp or larg):
      info["medidas_base"].extend([m for m in (comp, larg) if m > 0])
    elif func == FUNCAO_CONTRA_FRENTE_FUNDO and (comp or larg):
      info["larguras_cf"].append(max(comp, larg))
  return contexto


def _eh_item_gaveta_xml(
    item, parent_map: dict, info_mod: dict, comp: float, larg: float
) -> bool:
  desc_norm = _normalizar_texto(item.attrib.get("DESCRIPTION", ""))
  if "gav" in desc_norm:
    return True

  # Verifica ancestrais intermediários (ex: sub-item "Gaveta" ou CGAV)
  curr = parent_map.get(item)
  while curr is not None:
    if curr.tag == "ITEM":
      refs_anc = extrair_referencias(curr)
      desc_anc = _normalizar_texto(curr.attrib.get("DESCRIPTION", ""))
      if (
          refs_anc.get("ITEM_BASE") == "CGAV"
          or "gaveta" in _tokens_normalizados(desc_anc)
      ):
        return True
    curr = parent_map.get(curr)

  larguras_cf = info_mod.get("larguras_cf", [])
  medidas_base = info_mod.get("medidas_base", [])
  funcoes_caixa = info_mod.get("funcoes_caixa", set())

  if larguras_cf and "Base" not in funcoes_caixa:
    return True

  if larguras_cf:
    bate_cf = any(
        abs(comp - (r + AVANCO_PADRAO_TOTAL_MM)) <= 1.5
        or abs(larg - (r + AVANCO_PADRAO_TOTAL_MM)) <= 1.5
        for r in larguras_cf
    )
    bate_base = any(
        abs(comp - (r + AVANCO_PADRAO_TOTAL_MM)) <= 1.5
        or abs(larg - (r + AVANCO_PADRAO_TOTAL_MM)) <= 1.5
        for r in medidas_base
    )
    if bate_cf and not bate_base:
      return True

  return False


def calcular_delta_fundo_xml(
    item,
    refs: dict,
    parent_map: dict,
    contexto_modulos: dict,
    delta: float = AJUSTE_FUNDO_8MM_DELTA,
) -> tuple:
  """Retorna (delta_comp, delta_larg) para um <ITEM> de Fundo 6mm:
  - Roupeiro / Armário Aéreo (armário com Base e Chapéu): (+6, +6)
  - Balcão Inferior ou Gaveta: (+6, 0) ou (0, +6) apenas na medida da largura
  - Demais peças (ou Fundo que não seja 6mm): (0, 0)
  """
  if not eh_fundo_6mm(item, refs):
    return (0.0, 0.0)

  try:
    comp = float(item.attrib.get("WIDTH", "0"))
  except ValueError:
    comp = 0.0
  try:
    larg = float(item.attrib.get("DEPTH", "0"))
  except ValueError:
    larg = 0.0

  anc = _encontrar_item_ancestral(item, parent_map)
  chave_mod = (
      anc if anc is not None else extrair_modulo_pai(item, parent_map)
  )
  info_mod = contexto_modulos.get(chave_mod, {})

  if _eh_item_gaveta_xml(item, parent_map, info_mod, comp, larg):
    return escolher_eixo_largura_fundo(
        comp, larg, info_mod.get("larguras_cf", []), delta=delta
    )

  tokens_peca = _tokens_normalizados(item.attrib.get("DESCRIPTION", ""))
  modulo_nome = extrair_modulo_pai(item, parent_map)
  tokens_mod = _tokens_normalizados(modulo_nome)
  funcoes_caixa = info_mod.get("funcoes_caixa", set())
  medidas_base = info_mod.get("medidas_base", [])

  if tokens_peca & _TOKENS_BALCAO_INFERIOR:
    return escolher_eixo_largura_fundo(comp, larg, medidas_base, delta=delta)
  if tokens_peca & _TOKENS_ARMARIO_BASE_CHAPEU:
    return (float(delta), float(delta))

  if "Base" in funcoes_caixa and "Tampo" in funcoes_caixa:
    return (float(delta), float(delta))

  if "Base" in funcoes_caixa and "Tampo" not in funcoes_caixa:
    if (tokens_mod & _TOKENS_ARMARIO_BASE_CHAPEU) and not (
        tokens_mod & _TOKENS_BALCAO_INFERIOR
    ):
      return (float(delta), float(delta))
    return escolher_eixo_largura_fundo(comp, larg, medidas_base, delta=delta)

  if tokens_mod & _TOKENS_BALCAO_INFERIOR:
    return escolher_eixo_largura_fundo(comp, larg, medidas_base, delta=delta)
  if tokens_mod & _TOKENS_ARMARIO_BASE_CHAPEU:
    return (float(delta), float(delta))

  return escolher_eixo_largura_fundo(comp, larg, medidas_base, delta=delta)


def calcular_deltas_ajuste_df(df: pd.DataFrame) -> tuple:
  """Retorna (deltas_fundo, deltas_cf) mapeados pelo índice do DataFrame.

  Se `df.attrs["deltas_fundo"]` já tiver sido calculado na leitura do XML,
  reaproveita esses valores; caso contrário, infere diretamente das linhas do
  DataFrame usando as mesmas regras de módulo (Base + Tampo vs Balcão/Gaveta).
  """
  col_comp = "COMP\n2750" if "COMP\n2750" in df.columns else "Comprimento"
  col_larg = "LARG\n1840" if "LARG\n1840" in df.columns else "Largura"
  col_func = "NOME DA PEÇA" if "NOME DA PEÇA" in df.columns else "Função"
  col_mod = "OBS 1" if "OBS 1" in df.columns else "Complemento"
  col_mat = "ESP MDF" if "ESP MDF" in df.columns else "Material"

  deltas_cf = {}
  for idx, row in df.iterrows():
    func = str(row.get(col_func, "")).strip()
    if func == FUNCAO_CONTRA_FRENTE_FUNDO:
      try:
        comp = float(row.get(col_comp, 0) or 0)
      except (TypeError, ValueError):
        comp = 0.0
      try:
        larg = float(row.get(col_larg, 0) or 0)
      except (TypeError, ValueError):
        larg = 0.0
      deltas_cf[idx] = calcular_delta_contra_frente(comp, larg)

  deltas_xml = df.attrs.get("deltas_fundo") if hasattr(df, "attrs") else None
  if isinstance(deltas_xml, dict) and deltas_xml:
    deltas_fundo = {
        idx: tuple(deltas_xml.get(idx, (0.0, 0.0))) for idx in df.index
    }
    return deltas_fundo, deltas_cf

  # Inferência direta a partir do DataFrame quando não há metadados do XML
  info_por_mod = {}
  for idx, row in df.iterrows():
    mod = str(row.get(col_mod, "")).strip()
    func = str(row.get(col_func, "")).strip()
    try:
      comp = float(row.get(col_comp, 0) or 0)
    except (TypeError, ValueError):
      comp = 0.0
    try:
      larg = float(row.get(col_larg, 0) or 0)
    except (TypeError, ValueError):
      larg = 0.0
    info = info_por_mod.setdefault(
        mod, {"funcoes": set(), "medidas_base": [], "larguras_cf": []}
    )
    info["funcoes"].add(func)
    if func == "Base" and (comp or larg):
      info["medidas_base"].extend([m for m in (comp, larg) if m > 0])
    elif func == FUNCAO_CONTRA_FRENTE_FUNDO and (comp or larg):
      info["larguras_cf"].append(max(comp, larg))

  deltas_fundo = {}
  for idx, row in df.iterrows():
    func = str(row.get(col_func, "")).strip()
    if func != FUNCAO_FUNDO:
      continue
    mat_txt = str(row.get(col_mat, "") or "")
    if not _eh_espessura_6mm_texto(mat_txt):
      continue
    try:
      comp = float(row.get(col_comp, 0) or 0)
    except (TypeError, ValueError):
      comp = 0.0
    try:
      larg = float(row.get(col_larg, 0) or 0)
    except (TypeError, ValueError):
      larg = 0.0
    mod = str(row.get(col_mod, "")).strip()
    info = info_por_mod.get(mod, {})
    funcoes = info.get("funcoes", set())
    medidas_base = info.get("medidas_base", [])
    larguras_cf = info.get("larguras_cf", [])
    tokens_mod = _tokens_normalizados(mod)

    eh_gav = "gaveta" in tokens_mod or (
        bool(larguras_cf) and "Base" not in funcoes
    )
    if not eh_gav and larguras_cf:
      bate_cf = any(
          abs(comp - (r + AVANCO_PADRAO_TOTAL_MM)) <= 1.5
          or abs(larg - (r + AVANCO_PADRAO_TOTAL_MM)) <= 1.5
          for r in larguras_cf
      )
      bate_base = any(
          abs(comp - (r + AVANCO_PADRAO_TOTAL_MM)) <= 1.5
          or abs(larg - (r + AVANCO_PADRAO_TOTAL_MM)) <= 1.5
          for r in medidas_base
      )
      if bate_cf and not bate_base:
        eh_gav = True

    if eh_gav:
      deltas_fundo[idx] = escolher_eixo_largura_fundo(comp, larg, larguras_cf)
    elif "Base" in funcoes and "Tampo" in funcoes:
      deltas_fundo[idx] = (AJUSTE_FUNDO_8MM_DELTA, AJUSTE_FUNDO_8MM_DELTA)
    elif "Base" in funcoes and "Tampo" not in funcoes:
      if (tokens_mod & _TOKENS_ARMARIO_BASE_CHAPEU) and not (
          tokens_mod & _TOKENS_BALCAO_INFERIOR
      ):
        deltas_fundo[idx] = (AJUSTE_FUNDO_8MM_DELTA, AJUSTE_FUNDO_8MM_DELTA)
      else:
        deltas_fundo[idx] = escolher_eixo_largura_fundo(comp, larg, medidas_base)
    elif tokens_mod & _TOKENS_BALCAO_INFERIOR:
      deltas_fundo[idx] = escolher_eixo_largura_fundo(comp, larg, medidas_base)
    elif tokens_mod & _TOKENS_ARMARIO_BASE_CHAPEU:
      deltas_fundo[idx] = (AJUSTE_FUNDO_8MM_DELTA, AJUSTE_FUNDO_8MM_DELTA)
    else:
      deltas_fundo[idx] = escolher_eixo_largura_fundo(comp, larg, medidas_base)

  return deltas_fundo, deltas_cf


def aplicar_ajustes_medidas_df(
    df: pd.DataFrame,
    avanco_fundo_8mm: bool = False,
    reduzir_contra_frente_3mm: bool = False,
) -> pd.DataFrame:
  """Retorna uma cópia de `df` aplicando os ajustes solicitados nas medidas."""
  resultado = df.copy()
  deltas_fundo, deltas_cf = calcular_deltas_ajuste_df(df)
  col_comp = "COMP\n2750" if "COMP\n2750" in resultado.columns else "Comprimento"
  col_larg = "LARG\n1840" if "LARG\n1840" in resultado.columns else "Largura"

  def _somar(idx, col, delta):
    if not delta or col not in resultado.columns:
      return
    novo = round(float(resultado.at[idx, col]) + delta, 4)
    if pd.api.types.is_integer_dtype(resultado[col].dtype) and novo.is_integer():
      resultado.at[idx, col] = int(novo)
    else:
      resultado.at[idx, col] = novo

  if avanco_fundo_8mm:
    for idx, (dc, dl) in deltas_fundo.items():
      if idx in resultado.index:
        _somar(idx, col_comp, dc)
        _somar(idx, col_larg, dl)

  if reduzir_contra_frente_3mm:
    for idx, (dc, dl) in deltas_cf.items():
      if idx in resultado.index:
        _somar(idx, col_comp, dc)
        _somar(idx, col_larg, dl)

  return resultado


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

