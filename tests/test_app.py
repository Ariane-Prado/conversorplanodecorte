"""Testes das funções puras de extração/formatação do conversor XML -> Excel.

Roda sem precisar abrir a interface gráfica (usa apenas xml.etree, sem Tk).
Execução: python -m unittest discover -s tests -v
"""
import io
import os
import sys
import unittest
from unittest import mock
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app  # noqa: E402


XML_EXEMPLO = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "Conf - Cozinha.xml",
)


def item_de(xml_str: str):
  """Cria um <ITEM> ElementTree a partir de um trecho de XML para teste."""
  return ET.fromstring(xml_str)


class TestRenomearFuncao(unittest.TestCase):

  def test_contra_frente_e_contra_fundo_viram_um_so_nome(self):
    self.assertEqual(
        app.renomear_funcao("Contra Frente"),
        "Contra Frente / Contra Fundo de Gaveta",
    )
    self.assertEqual(
        app.renomear_funcao("Contra Fundo"),
        "Contra Frente / Contra Fundo de Gaveta",
    )

  def test_qualquer_porta_vira_porta(self):
    self.assertEqual(app.renomear_funcao("Porta Direita"), "Porta")

  def test_laterais_de_gaveta_sao_unificadas(self):
    self.assertEqual(
        app.renomear_funcao("Lateral Direita Gaveta"), "Lateral de Gaveta"
    )
    self.assertEqual(
        app.renomear_funcao("Lateral Esquerda Gaveta"), "Lateral de Gaveta"
    )

  def test_laterais_de_movel_viram_direita_ou_esquerda(self):
    for nome in ["Lateral Inf Dir", "Lateral Sup Dir", "Lateral Dir Dor"]:
      self.assertEqual(app.renomear_funcao(nome), "Lateral Direita")
    for nome in ["Lateral Inf Esq", "Lateral Sup Esq", "Lateral Esq Dor"]:
      self.assertEqual(app.renomear_funcao(nome), "Lateral Esquerda")

  def test_base_vira_base(self):
    for nome in ["Base Inf", "Base Sup", "Base Dor", "Base Dir Canto Curvo"]:
      self.assertEqual(app.renomear_funcao(nome), "Base")

  def test_divisoria_vira_divisoria(self):
    self.assertEqual(app.renomear_funcao("Divisória Vertical"), "Divisória")

  def test_chapeu_e_tampo_viram_tampo(self):
    for nome in ["Chapeu Sup", "Chapeu Dor", "Tampo Dir Canto Curvo"]:
      self.assertEqual(app.renomear_funcao(nome), "Tampo")

  def test_tamponamento_mantem_primeira_palavra(self):
    self.assertEqual(
        app.renomear_funcao("Tamponamento 15mm c/ fita 1mm"), "Tamponamento"
    )
    self.assertEqual(
        app.renomear_funcao("Tamponamento 18mm c/ fita 2mm"), "Tamponamento"
    )

  def test_painel_com_sufixo_vira_painel(self):
    self.assertEqual(app.renomear_funcao("Painel 15mm c/ fita 1mm"), "Painel")

  def test_pecas_genericas_mantem_so_a_primeira_palavra(self):
    self.assertEqual(app.renomear_funcao("Fundo Mod Sup"), "Fundo")
    self.assertEqual(app.renomear_funcao("Prateleira Dormitorio"), "Prateleira")
    self.assertEqual(app.renomear_funcao("Prateleira Furada"), "Prateleira")

  def test_nome_sem_regra_especial_fica_inalterado(self):
    self.assertEqual(app.renomear_funcao("Prateleira"), "Prateleira")


class TestExtrairEspessuraFita(unittest.TestCase):

  def test_reconhece_045mm(self):
    self.assertEqual(app.extrair_espessura_fita("Fita 0.45mm Branco", "", "", ""), "0.45mm")

  def test_reconhece_1mm(self):
    self.assertEqual(app.extrair_espessura_fita("", "ABS 1MM", "", ""), "1mm")

  def test_sem_fita_retorna_vazio(self):
    self.assertEqual(app.extrair_espessura_fita("", "", "", ""), "")


class TestDeveIgnorarItem(unittest.TestCase):

  def test_item_sem_descricao_e_ignorado(self):
    item = item_de('<ITEM DESCRIPTION="" STRUCTURE="N" FAMILY="COZ" />')
    self.assertTrue(app.deve_ignorar_item(item, {}))

  def test_item_estrutural_e_ignorado(self):
    item = item_de(
        '<ITEM DESCRIPTION="Estrutura" STRUCTURE="Y" FAMILY="COZ" />'
    )
    self.assertTrue(app.deve_ignorar_item(item, {}))

  def test_familia_alumistar_e_sempre_ignorada(self):
    # Cobre a peça de vidro/alumínio da linha parceira Alumistar (não é MDF
    # cortado na fábrica, é fornecido pronto pela Alumistar).
    item = item_de(
        '<ITEM DESCRIPTION="Painel Shine Stop Sol (Posição 1)" '
        'STRUCTURE="N" FAMILY="Alumistar" ISGEOMETRY="N" />'
    )
    refs = {"PAINEL": "Vidro"}
    self.assertTrue(app.deve_ignorar_item(item, refs))

  def test_familia_ferragens_e_ignorada(self):
    item = item_de(
        '<ITEM DESCRIPTION="Dobradiça" STRUCTURE="N" FAMILY="Ferragens" />'
    )
    self.assertTrue(app.deve_ignorar_item(item, {}))

  def test_peca_sem_material_nem_geometria_e_ignorada(self):
    item = item_de(
        '<ITEM DESCRIPTION="Base Gaveta" STRUCTURE="N" FAMILY="COZ" '
        'ISGEOMETRY="N" ID="X" />'
    )
    refs = {"ITEM_BASE": "CGAV"}
    self.assertTrue(app.deve_ignorar_item(item, refs))

  def test_peca_de_mdf_normal_nao_e_ignorada(self):
    item = item_de(
        '<ITEM DESCRIPTION="Lateral Direita" STRUCTURE="N" FAMILY="COZ" '
        'ISGEOMETRY="Y" WIDTH="600" DEPTH="560" />'
    )
    refs = {"MATERIAL": "MDF Branco", "ITEM_BASE": "1234.5678"}
    self.assertFalse(app.deve_ignorar_item(item, refs))


class TestFormatarMaterialComEspessura(unittest.TestCase):

  def test_junta_material_modelo_espessura_e_fabricante(self):
    item = item_de('<ITEM HEIGHT="18" />')
    refs = {"MATERIAL": "MDF", "MODEL": "Branco TX"}
    fab_map = {"branco tx": "ARAUCO"}
    resultado = app.formatar_material_com_espessura(item, refs, fab_map)
    self.assertIn("MDF", resultado)
    self.assertIn("Branco TX", resultado)
    self.assertIn("18 mm", resultado)
    self.assertIn("ARAUCO", resultado)

  def test_sem_nenhum_dado_retorna_nao_especificado(self):
    item = item_de("<ITEM />")
    self.assertEqual(
        app.formatar_material_com_espessura(item, {}, {}), "Não especificado"
    )


class TestExtrairServicosAdicionais(unittest.TestCase):

  def test_porta_com_furos_de_dobradica(self):
    item = item_de(
        '<ITEM DESCRIPTION="Porta Direita" ID="POR_1">'
        '<ITEM DESCRIPTION="Furar" QUANTITY="2" />'
        "</ITEM>"
    )
    resultado = app.extrair_servicos_adicionais(item, "Porta Direita")
    self.assertIn("2 furos de dobr.", resultado)
    self.assertIn("Ld Dir", resultado)

  def test_peca_com_rasgo_de_fundo(self):
    item = item_de(
        '<ITEM DESCRIPTION="Lateral">'
        '<ITEM DESCRIPTION="Rasgo" />'
        "</ITEM>"
    )
    resultado = app.extrair_servicos_adicionais(item, "Lateral")
    self.assertEqual(resultado, "C/ Rasgo de Fundo")

  def test_peca_sem_servicos_retorna_vazio(self):
    item = item_de('<ITEM DESCRIPTION="Prateleira" />')
    self.assertEqual(app.extrair_servicos_adicionais(item, "Prateleira"), "")


@unittest.skipUnless(
    os.path.exists(XML_EXEMPLO), "XML de exemplo não encontrado no repositório"
)
class TestComXmlReal(unittest.TestCase):
  """Testes de integração usando o XML de exemplo real do projeto."""

  @classmethod
  def setUpClass(cls):
    cls.root = ET.parse(XML_EXEMPLO).getroot()
    ambients = cls.root.find("AMBIENTS")
    cls.items = ambients.findall(".//ITEM") if ambients is not None else []

  def test_nenhum_item_alumistar_sobrevive_ao_filtro(self):
    sobreviventes_alumistar = [
        item
        for item in self.items
        if item.attrib.get("FAMILY", "").strip() == "Alumistar"
        and not app.deve_ignorar_item(item, app.extrair_referencias(item))
    ]
    self.assertEqual(
        sobreviventes_alumistar,
        [],
        "Item(ns) da família Alumistar (vidro/alumínio do parceiro) não"
        " deveriam passar no filtro de exportação.",
    )

  def test_ha_pecas_de_mdf_que_sobrevivem_ao_filtro(self):
    sobreviventes = [
        item
        for item in self.items
        if not app.deve_ignorar_item(item, app.extrair_referencias(item))
    ]
    self.assertGreater(
        len(sobreviventes), 0, "O filtro não deveria remover todas as peças."
    )

  def test_extrai_cliente_e_ambiente_do_xml(self):
    cliente, ambiente = app.extrair_cliente_ambiente_xml(XML_EXEMPLO)
    self.assertTrue(cliente)
    self.assertTrue(ambiente)


class _FakeResponse(io.BytesIO):
  """Simula o objeto retornado por urllib.request.urlopen (usado com 'with')."""

  def __init__(self, data: bytes, headers: dict = None):
    super().__init__(data)
    self.headers = headers or {}

  def __enter__(self):
    return self

  def __exit__(self, *args):
    self.close()


class TestVersaoParaTupla(unittest.TestCase):

  def test_versao_simples(self):
    self.assertEqual(app._versao_para_tupla("1.2.3"), (1, 2, 3))

  def test_versao_invalida_vira_zero(self):
    self.assertEqual(app._versao_para_tupla("abc"), (0,))

  def test_versao_vazia_vira_zero(self):
    self.assertEqual(app._versao_para_tupla(""), (0,))

  def test_comparacao_semantica(self):
    self.assertGreater(app._versao_para_tupla("1.10.0"), app._versao_para_tupla("1.9.0"))


class TestVerificarNovaVersao(unittest.TestCase):

  def setUp(self):
    self._versao_original = app.VERSAO_ATUAL

  def tearDown(self):
    app.VERSAO_ATUAL = self._versao_original

  def test_detecta_versao_mais_nova(self):
    app.VERSAO_ATUAL = "1.0.0"
    payload = b'{"versao": "1.1.0", "url_download": "https://exemplo.com/app.exe"}'
    with mock.patch("app.urllib.request.urlopen", return_value=_FakeResponse(payload)):
      resultado = app.verificar_nova_versao()
    self.assertEqual(resultado, ("1.1.0", "https://exemplo.com/app.exe"))

  def test_nao_notifica_se_ja_atualizado(self):
    app.VERSAO_ATUAL = "1.1.0"
    payload = b'{"versao": "1.0.0", "url_download": "https://exemplo.com/app.exe"}'
    with mock.patch("app.urllib.request.urlopen", return_value=_FakeResponse(payload)):
      resultado = app.verificar_nova_versao()
    self.assertIsNone(resultado)

  def test_falha_de_rede_nao_propaga_excecao(self):
    with mock.patch("app.urllib.request.urlopen", side_effect=OSError("sem rede")):
      resultado = app.verificar_nova_versao()
    self.assertIsNone(resultado)

  def test_manifesto_incompleto_retorna_none(self):
    payload = b'{"versao": "9.9.9"}'  # sem url_download
    with mock.patch("app.urllib.request.urlopen", return_value=_FakeResponse(payload)):
      resultado = app.verificar_nova_versao()
    self.assertIsNone(resultado)


class TestBaixarAtualizacao(unittest.TestCase):

  def test_conteudo_html_e_rejeitado(self):
    import tempfile
    pagina_html = b"<html><body>Google Drive scan warning</body></html>"
    with tempfile.TemporaryDirectory() as tmp:
      destino = os.path.join(tmp, "app_novo.exe")
      with mock.patch(
          "app.urllib.request.urlopen", return_value=_FakeResponse(pagina_html)
      ):
        with self.assertRaises(RuntimeError):
          app.baixar_atualizacao("https://exemplo.com/app.exe", destino)
      self.assertFalse(os.path.exists(destino), "arquivo inválido deveria ser removido")

  def test_download_valido_grava_arquivo_e_reporta_progresso(self):
    import tempfile
    conteudo = b"X" * 200_000  # simula um .exe pequeno, acima do limiar de 100KB
    chamadas_progresso = []
    with tempfile.TemporaryDirectory() as tmp:
      destino = os.path.join(tmp, "app_novo.exe")
      headers = {"Content-Length": str(len(conteudo))}
      with mock.patch(
          "app.urllib.request.urlopen",
          return_value=_FakeResponse(conteudo, headers=headers),
      ):
        app.baixar_atualizacao(
            "https://exemplo.com/app.exe",
            destino,
            progresso_callback=lambda lido, total: chamadas_progresso.append(
                (lido, total)
            ),
        )
      self.assertTrue(os.path.exists(destino))
      self.assertEqual(os.path.getsize(destino), len(conteudo))
      self.assertTrue(chamadas_progresso)
      self.assertEqual(chamadas_progresso[-1], (len(conteudo), len(conteudo)))


class TestAjustesConferenciaFundoEContraFrente(unittest.TestCase):
  """Testa os ajustes de conferência:
  1) Fundo 6mm: avanço de 5mm -> 8mm de cada lado (+6mm no total):
     - Armários com Base e Chapéu (roupeiros e aéreos): +6mm nas duas medidas
     - Balcões inferiores e gavetas: +6mm apenas na largura (em relação à
       Base ou à Contra Frente)
  2) Contra Frente / Contra Fundo de Gaveta: -3mm na altura (menor medida)
  """

  def _montar_xml_completo(self) -> ET.Element:
    xml_str = """
    <ROOT>
      <AMBIENTS>
        <AMBIENT DESCRIPTION="Projeto - Cozinha e Quarto">
          <!-- 1) Armário Aéreo (tem Base Sup e Chapeu Sup) -->
          <ITEM DESCRIPTION="Armário Aéreo 2P" STRUCTURE="Y">
            <ITEM DESCRIPTION="Base Sup" STRUCTURE="N" FAMILY="COZ" ISGEOMETRY="Y" WIDTH="770" DEPTH="330" HEIGHT="15">
              <REFERENCES><MATERIAL REFERENCE="MDF Branco" /></REFERENCES>
            </ITEM>
            <ITEM DESCRIPTION="Chapeu Sup" STRUCTURE="N" FAMILY="COZ" ISGEOMETRY="Y" WIDTH="770" DEPTH="330" HEIGHT="15">
              <REFERENCES><MATERIAL REFERENCE="MDF Branco" /></REFERENCES>
            </ITEM>
            <ITEM DESCRIPTION="Fundo Mod Sup" STRUCTURE="N" FAMILY="COZ" ISGEOMETRY="Y" WIDTH="780" DEPTH="680" HEIGHT="6">
              <REFERENCES><MATERIAL REFERENCE="MDF Branco" /></REFERENCES>
            </ITEM>
          </ITEM>

          <!-- 2) Roupeiro com Gaveta interna (tem Base Dor, Chapeu Dor, Fundo Dor e Fundo Gaveta) -->
          <ITEM DESCRIPTION="Roupeiro 2 Portas" STRUCTURE="Y">
            <ITEM DESCRIPTION="Base Dor" STRUCTURE="N" FAMILY="DOR" ISGEOMETRY="Y" WIDTH="1570" DEPTH="550" HEIGHT="15">
              <REFERENCES><MATERIAL REFERENCE="MDF Branco" /></REFERENCES>
            </ITEM>
            <ITEM DESCRIPTION="Chapeu Dor" STRUCTURE="N" FAMILY="DOR" ISGEOMETRY="Y" WIDTH="1570" DEPTH="550" HEIGHT="15">
              <REFERENCES><MATERIAL REFERENCE="MDF Branco" /></REFERENCES>
            </ITEM>
            <ITEM DESCRIPTION="Fundo Dor" STRUCTURE="N" FAMILY="DOR" ISGEOMETRY="Y" WIDTH="1580" DEPTH="2180" HEIGHT="6">
              <REFERENCES><MATERIAL REFERENCE="MDF Branco" /></REFERENCES>
            </ITEM>
            <ITEM DESCRIPTION="Contra Frente" STRUCTURE="N" FAMILY="DOR" ISGEOMETRY="Y" WIDTH="700" DEPTH="120" HEIGHT="15">
              <REFERENCES><MATERIAL REFERENCE="MDF Branco" /></REFERENCES>
            </ITEM>
            <ITEM DESCRIPTION="Fundo Gaveta" STRUCTURE="N" FAMILY="DOR" ISGEOMETRY="Y" WIDTH="710" DEPTH="450" HEIGHT="6">
              <REFERENCES><MATERIAL REFERENCE="MDF Branco" /></REFERENCES>
            </ITEM>
          </ITEM>

          <!-- 3) Balcão Inferior (tem Base Inf, sem Chapeu) -->
          <ITEM DESCRIPTION="Mod. Inferior 2 Portas" STRUCTURE="Y">
            <ITEM DESCRIPTION="Base Inf" STRUCTURE="N" FAMILY="COZ" ISGEOMETRY="Y" WIDTH="770" DEPTH="550" HEIGHT="15">
              <REFERENCES><MATERIAL REFERENCE="MDF Branco" /></REFERENCES>
            </ITEM>
            <ITEM DESCRIPTION="Fundo Inf" STRUCTURE="N" FAMILY="COZ" ISGEOMETRY="Y" WIDTH="780" DEPTH="680" HEIGHT="6">
              <REFERENCES><MATERIAL REFERENCE="MDF Branco" /></REFERENCES>
            </ITEM>
          </ITEM>
        </AMBIENT>
      </AMBIENTS>
    </ROOT>
    """
    return ET.fromstring(xml_str)

  def test_fundo_armario_aereo_e_roupeiro_soma_6mm_nas_duas_medidas(self):
    root = self._montar_xml_completo()
    items = root.findall(".//ITEM")
    parent_map = {c: p for p in root.iter() for c in p}
    ctx = app.mapear_contexto_modulos_xml(items, parent_map)

    fundo_aereo = [i for i in items if i.attrib.get("DESCRIPTION") == "Fundo Mod Sup"][0]
    fundo_roup = [i for i in items if i.attrib.get("DESCRIPTION") == "Fundo Dor"][0]

    self.assertEqual(
        app.calcular_delta_fundo_xml(
            fundo_aereo, app.extrair_referencias(fundo_aereo), parent_map, ctx
        ),
        (6.0, 6.0),
    )
    self.assertEqual(
        app.calcular_delta_fundo_xml(
            fundo_roup, app.extrair_referencias(fundo_roup), parent_map, ctx
        ),
        (6.0, 6.0),
    )

  def test_fundo_balcao_inferior_e_gaveta_soma_6mm_apenas_na_largura(self):
    root = self._montar_xml_completo()
    items = root.findall(".//ITEM")
    parent_map = {c: p for p in root.iter() for c in p}
    ctx = app.mapear_contexto_modulos_xml(items, parent_map)

    fundo_balcao = [i for i in items if i.attrib.get("DESCRIPTION") == "Fundo Inf"][0]
    fundo_gaveta = [i for i in items if i.attrib.get("DESCRIPTION") == "Fundo Gaveta"][0]

    self.assertEqual(
        app.calcular_delta_fundo_xml(
            fundo_balcao, app.extrair_referencias(fundo_balcao), parent_map, ctx
        ),
        (6.0, 0.0),
    )
    self.assertEqual(
        app.calcular_delta_fundo_xml(
            fundo_gaveta, app.extrair_referencias(fundo_gaveta), parent_map, ctx
        ),
        (6.0, 0.0),
    )

  def test_contra_frente_reduz_3mm_na_menor_medida_altura(self):
    self.assertEqual(app.calcular_delta_contra_frente(500, 120), (0.0, -3.0))
    self.assertEqual(app.calcular_delta_contra_frente(120, 500), (-3.0, 0.0))

  def test_aplicar_ajustes_medidas_df_padrao_e_tmk(self):
    import pandas as pd

    df = pd.DataFrame(
        [
            [1, 770.0, 330.0, "Base", "MDF 15 mm", "Armário Aéreo"],
            [1, 770.0, 330.0, "Tampo", "MDF 15 mm", "Armário Aéreo"],
            [1, 780.0, 680.0, "Fundo", "MDF 6 mm", "Armário Aéreo"],
            [1, 770.0, 550.0, "Base", "MDF 15 mm", "Balcão Inferior"],
            [1, 780.0, 680.0, "Fundo", "MDF 6 mm", "Balcão Inferior"],
            [2, 500.0, 120.0, "Contra Frente / Contra Fundo de Gaveta", "MDF 15 mm", "Gaveteiro"],
            [1, 510.0, 450.0, "Fundo", "MDF 6 mm", "Gaveteiro"],
            [1, 800.0, 600.0, "Fundo", "MDF 15 mm", "Armário Aéreo"],  # Fundo 15mm não altera
        ],
        columns=["Quantidade", "Comprimento", "Largura", "Função", "Material", "Complemento"],
    )

    df_ajustado = app.aplicar_ajustes_medidas_df(
        df, avanco_fundo_8mm=True, reduzir_contra_frente_3mm=True
    )

    # Armário aéreo (Base + Tampo): Fundo 6mm soma +6 nas duas medidas
    self.assertEqual(df_ajustado.loc[2, "Comprimento"], 786.0)
    self.assertEqual(df_ajustado.loc[2, "Largura"], 686.0)

    # Balcão inferior (só Base): Fundo 6mm soma +6 apenas na largura (780 -> 786)
    self.assertEqual(df_ajustado.loc[4, "Comprimento"], 786.0)
    self.assertEqual(df_ajustado.loc[4, "Largura"], 680.0)

    # Contra Frente / Contra Fundo: reduz 3mm na altura (menor medida: 120 -> 117)
    self.assertEqual(df_ajustado.loc[5, "Comprimento"], 500.0)
    self.assertEqual(df_ajustado.loc[5, "Largura"], 117.0)

    # Fundo de gaveta: soma +6 apenas na largura em relação à Contra Frente (510 -> 516)
    self.assertEqual(df_ajustado.loc[6, "Comprimento"], 516.0)
    self.assertEqual(df_ajustado.loc[6, "Largura"], 450.0)

    # Fundo de 15mm: permanece inalterado
    self.assertEqual(df_ajustado.loc[7, "Comprimento"], 800.0)
    self.assertEqual(df_ajustado.loc[7, "Largura"], 600.0)


class TestPersistenciaHistorico(unittest.TestCase):
  """Garante que o histórico é mesclado entre pastas e nunca sobrescrito por outra janela."""

  def _criar_instancia_fake(self, arq_local, arq_usuario, cliente="Vitor", ambiente="Suite 1"):
    inst = object.__new__(app.ConversorXmlExcelApp)
    inst.arquivo_historico = arq_local
    inst.arquivo_historico_usuario = arq_usuario
    inst.cliente_nome = cliente
    inst.ambiente_nome = ambiente
    inst.frames = {}
    inst.historico = inst.carregar_historico()
    return inst

  def test_duas_instancias_nao_sobrescrevem_historico_uma_da_outra(self):
    import tempfile

    with tempfile.TemporaryDirectory() as tmpdir:
      arq_local_1 = os.path.join(tmpdir, "pasta1", "historico_recentes.json")
      arq_local_2 = os.path.join(tmpdir, "pasta2", "historico_recentes.json")
      os.makedirs(os.path.dirname(arq_local_1), exist_ok=True)
      os.makedirs(os.path.dirname(arq_local_2), exist_ok=True)
      arq_usr = os.path.join(tmpdir, ".conversor_promob_historico.json")

      # Instância 1 e Instância 2 abertas ao mesmo tempo em pastas diferentes
      inst1 = self._criar_instancia_fake(arq_local_1, arq_usr, "Vitor", "Suite 1")
      inst2 = self._criar_instancia_fake(arq_local_2, arq_usr, "Vitor", "Suite Casal")

      inst1.adicionar_ao_historico(
          "C:/vitor/Conf - Suite 1.xml",
          "C:/vitor/Vitor_Suite_1_Tobias.xlsx",
          "Tobias",
      )
      inst2.adicionar_ao_historico(
          "C:/vitor/Conf - Suite Casal.xml",
          "C:/vitor/Vitor_Suite_Casal_Tobias.xlsx",
          "Tobias",
      )

      inst2.ambiente_nome = "Escritorio"
      inst2.adicionar_ao_historico(
          "C:/vitor/Conf - Escritorio.xml",
          "C:/vitor/Vitor_Escritorio_Tobias.xlsx",
          "Tobias",
      )

      ambientes_salvos = [item["ambiente"] for item in inst2.carregar_historico()]
      self.assertEqual(
          ambientes_salvos,
          ["Escritorio", "Suite Casal", "Suite 1"],
      )


class TestSubstituicaoInPlaceAtualizacao(unittest.TestCase):
  """Garante que a atualização substitui o .exe no mesmo lugar sem criar outro."""

  def test_substitui_executavel_no_mesmo_lugar_e_limpa_meipass(self):
    import tempfile

    with tempfile.TemporaryDirectory() as tmpdir:
      exe_existente = os.path.join(tmpdir, "app.exe")
      novo_baixado = os.path.join(tempfile.gettempdir(), "_teste_update_novo.tmp")
      with open(exe_existente, "wb") as f:
        f.write(b"VERSAO_ANTIGA")
      with open(novo_baixado, "wb") as f:
        f.write(b"VERSAO_NOVA_ATUALIZADA")

      with (
          mock.patch.object(sys, "executable", exe_existente),
          mock.patch.dict(os.environ, {"_MEIPASS2": "C:/Temp/_MEI12345"}),
          mock.patch("atualizador.subprocess.Popen") as mock_popen,
          mock.patch("atualizador.os._exit") as mock_exit,
      ):
        app.aplicar_atualizacao_e_reiniciar(novo_baixado)

      # O arquivo original deve ter sido substituído no mesmo caminho
      self.assertTrue(os.path.exists(exe_existente))
      with open(exe_existente, "rb") as f:
        self.assertEqual(f.read(), b"VERSAO_NOVA_ATUALIZADA")

      # Nenhum arquivo extra (_app_novo_*.exe ou _atualizar.bat) deve existir na pasta
      self.assertEqual(os.listdir(tmpdir), ["app.exe"])

      # Deve ter reiniciado o próprio exe_existente sem _MEIPASS2 no ambiente
      mock_popen.assert_called_once()
      args_chamada, kwargs_chamada = mock_popen.call_args
      self.assertEqual(args_chamada[0], [os.path.abspath(exe_existente)])
      self.assertNotIn("_MEIPASS2", kwargs_chamada["env"])
      mock_exit.assert_called_once_with(0)


if __name__ == "__main__":
  unittest.main()
