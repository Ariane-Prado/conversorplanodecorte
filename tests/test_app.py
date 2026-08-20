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


if __name__ == "__main__":
  unittest.main()
