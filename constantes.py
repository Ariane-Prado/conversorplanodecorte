"""Constantes globais de configuração, cores (WCAG AA), fornecedores e regras."""
import re

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

VERSAO_ATUAL = "1.3.0"
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
# TMKPlanilha usa um layout de colunas próprio (ver processar_dados_para_revisao),
# por isso fica de fora desse grupo.
FORNECEDORES_FORMATO_PADRAO = ["Teletintas", "Madefer", "Tobias", "TMKCloud"]
FORNECEDORES_FORMATO_TMK = ["TMKPlanilha"]

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

# Ajustes opcionais de conferência:
# 1) Fundo 6mm: passar o avanço lateral de 5mm para 8mm de cada lado (+3mm de
#    cada lado = +6mm no total). Em armários com base e chapéu (roupeiros e
#    armários aéreos) soma +6mm nas duas medidas (Comprimento e Largura); em
#    balcões inferiores e gavetas soma +6mm apenas na medida da largura (em
#    relação à Base ou à Contra Frente).
# 2) Contra Frente / Contra Fundo de Gaveta: reduzir 3mm na altura (a menor
#    medida da tira).
AVANCO_PADRAO_TOTAL_MM = 10.0  # 5mm de cada lado no XML original
AJUSTE_FUNDO_8MM_DELTA = 6.0   # +3mm de cada lado (passando de 5mm para 8mm)
AJUSTE_CONTRA_FRENTE_ALTURA_DELTA = -3.0
FUNCAO_CONTRA_FRENTE_FUNDO = "Contra Frente / Contra Fundo de Gaveta"
FUNCAO_FUNDO = "Fundo"

_TOKENS_BALCAO_INFERIOR = {"inf", "inferior", "balcao", "gaveteiro"}
_TOKENS_ARMARIO_BASE_CHAPEU = {
    "sup",
    "superior",
    "aereo",
    "roup",
    "roupeiro",
    "dor",
    "dorm",
    "dormitorio",
    "armario",
}
_PADRAO_ESPESSURA_MM = re.compile(r"(\d+(?:[.,]\d+)?)\s*mm\b", re.IGNORECASE)

