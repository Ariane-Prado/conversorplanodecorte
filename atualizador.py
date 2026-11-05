"""Persistência de configurações locais e sistema de atualização automática."""
import json
import os
import subprocess
import sys
import urllib.request

from constantes import URL_VERSAO_REMOTA, VERSAO_ATUAL


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


def _obter_contexto_app():
  """Permite que testes unitários façam mock em `app.urllib.request.urlopen`
  e `app.VERSAO_ATUAL` mantendo compatibilidade com o módulo principal."""
  mod_app = sys.modules.get("app")
  urllib_mod = getattr(mod_app, "urllib", urllib) if mod_app else urllib
  versao_atual = getattr(mod_app, "VERSAO_ATUAL", VERSAO_ATUAL) if mod_app else VERSAO_ATUAL
  url_remota = getattr(mod_app, "URL_VERSAO_REMOTA", URL_VERSAO_REMOTA) if mod_app else URL_VERSAO_REMOTA
  return urllib_mod, versao_atual, url_remota


def verificar_nova_versao():
  """Consulta o manifesto remoto (versao.json).

  Retorna (versao_nova, url_download) se houver versão mais nova que
  VERSAO_ATUAL, ou None se estiver atualizado / a checagem falhar (sem
  internet, repositório fora do ar etc. — nunca levanta exceção).
  """
  urllib_mod, versao_atual, url_remota = _obter_contexto_app()
  try:
    with urllib_mod.request.urlopen(url_remota, timeout=6) as resp:
      dados = json.loads(resp.read().decode("utf-8"))
    versao_remota = str(dados.get("versao", "")).strip()
    url_download = str(dados.get("url_download", "")).strip()
    if not versao_remota or not url_download:
      return None
    if _versao_para_tupla(versao_remota) > _versao_para_tupla(versao_atual):
      return versao_remota, url_download
    return None
  except Exception as e:
    print(f"Aviso: falha ao verificar atualização: {e}")
    return None


def baixar_atualizacao(url_download: str, destino: str, progresso_callback=None):
  """Baixa o novo executável para 'destino'. Levanta exceção se falhar."""
  urllib_mod, _, _ = _obter_contexto_app()
  req = urllib_mod.request.Request(
      url_download, headers={"User-Agent": "Mozilla/5.0"}
  )
  with urllib_mod.request.urlopen(req, timeout=30) as resp:
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
  antigo e reabre o app.

  A janela do .bat fica visível (com título e mensagem) e o script não se
  autodestrói mais no final — um script oculto que troca um executável e
  apaga a si mesmo é exatamente o padrão que antivírus/EDR costumam
  bloquear por heurística de auto-replicação, mesmo sendo legítimo aqui.
  O arquivo `_atualizar.bat` fica na pasta do app e é sobrescrito na
  próxima atualização.

  O `move` tem retry: mesmo depois do processo antigo sumir do tasklist,
  o Windows (ou o antivírus fazendo scan em tempo real do .exe recém-
  baixado) pode segurar o arquivo travado por mais um instante — sem
  retry isso falhava com "Acesso negado" e reabria o app antigo sem
  trocar nada.

  O limite de tentativas é de 5 minutos: em máquinas mais lentas o
  antivírus pode demorar bem mais que alguns segundos pra liberar um
  .exe desconhecido recém-baixado (visto na prática — 60 tentativas de
  1s, ou seja 1 minuto, não foi suficiente e a troca falhava mesmo com
  o arquivo íntegro).
  """
  exe_atual = sys.executable
  pasta = os.path.dirname(exe_atual)
  bat_path = os.path.join(pasta, "_atualizar.bat")
  pid_atual = os.getpid()

  conteudo_bat = (
      "@echo off\n"
      "title Atualizando Conversor XML Promob\n"
      "echo Aplicando atualizacao, aguarde...\n"
      ":espera\n"
      f'tasklist /FI "PID eq {pid_atual}" 2>NUL | find /I "{pid_atual}" >NUL\n'
      "if not errorlevel 1 (\n"
      "    timeout /t 1 /nobreak >NUL\n"
      "    goto espera\n"
      ")\n"
      "set tentativas=0\n"
      ":mover\n"
      f'move /Y "{caminho_novo_exe}" "{exe_atual}" >NUL 2>&1\n'
      "if errorlevel 1 (\n"
      "    set /a tentativas+=1\n"
      "    if %tentativas% GEQ 300 (\n"
      "        echo Nao foi possivel substituir o executavel apos 5 minutos"
      " (arquivo ainda em uso ou bloqueado pelo antivirus).\n"
      "        echo Tente novamente ou substitua manualmente pelo arquivo"
      " baixado em:\n"
      f"        echo {caminho_novo_exe}\n"
      "        pause\n"
      "        exit /b 1\n"
      "    )\n"
      "    timeout /t 1 /nobreak >NUL\n"
      "    goto mover\n"
      ")\n"
      f'start "" "{exe_atual}"\n'
      "exit\n"
  )
  with open(bat_path, "w", encoding="utf-8") as f:
    f.write(conteudo_bat)

  subprocess.Popen(["cmd", "/c", bat_path], close_fds=True)
  os._exit(0)

