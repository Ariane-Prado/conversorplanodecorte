"""Persistência de configurações locais e sistema de atualização automática."""
import glob
import json
import os
import shutil
import subprocess
import sys
import tempfile
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


def _ambiente_limpo_pyinstaller() -> dict:
  """Remove variáveis internas do bootloader PyInstaller (_MEIPASS2 etc.).

  Sem isso, o processo filho herda `_MEIPASS2` apontando para a pasta `_MEI...`
  da versão antiga e reabre o código antigo em vez de descompactar o novo `.exe`.
  """
  env = os.environ.copy()
  for chave in list(env.keys()):
    if chave == "_MEIPASS2" or chave.startswith("_PYI_") or chave.startswith("_MEI"):
      env.pop(chave, None)
  return env


def limpar_residuos_atualizacao() -> None:
  """Remove arquivos temporários de atualizações anteriores (`_app_novo_*.exe`,
  `_atualizar.bat`, `_app_antigo*.tmp`, `_conversor_old_*.tmp`)."""
  pasta_base = obter_caminho_base()
  padroes = [
      os.path.join(pasta_base, "_atualizar.bat"),
      os.path.join(pasta_base, "_app_novo_*.exe"),
      os.path.join(pasta_base, "_app_antigo*.tmp"),
      os.path.join(tempfile.gettempdir(), "_conversor_old_*.tmp"),
      os.path.join(tempfile.gettempdir(), "_conversor_update_*.tmp"),
  ]
  for padrao in padroes:
    for caminho in glob.glob(padrao):
      try:
        os.remove(caminho)
      except OSError:
        pass


def aplicar_atualizacao_e_reiniciar(caminho_novo_exe: str):
  """Substitui o .exe em execução no próprio lugar (in-place) e reabre o app.

  No Windows (NTFS), um executável em execução não pode ser sobrescrito
  diretamente, mas pode ser renomeado/movido para `%TEMP%` enquanto ainda roda.
  Assim:
  1. Renomeamos o `app.exe` atual para um arquivo temporário em `%TEMP%`.
  2. Movemos o novo executável baixado diretamente para o caminho original do
     `app.exe` (sem deixar cópias `_app_novo_*.exe` nem `_atualizar.bat` na pasta).
  3. Iniciamos o novo `app.exe` com o ambiente limpo (sem `_MEIPASS2`) e
     encerramos o processo antigo.
  """
  exe_atual = os.path.abspath(sys.executable)
  pasta = os.path.dirname(exe_atual)
  pid_atual = os.getpid()
  pasta_temp = tempfile.gettempdir()
  env_limpo = _ambiente_limpo_pyinstaller()

  caminho_old = os.path.join(pasta_temp, f"_conversor_old_{pid_atual}.tmp")
  substituido_in_place = False

  try:
    try:
      os.replace(exe_atual, caminho_old)
    except OSError:
      # Fallback caso %TEMP% esteja em outro disco (ex.: D:\ vs C:\)
      caminho_old = os.path.join(pasta, f"_app_antigo_{pid_atual}.tmp")
      os.replace(exe_atual, caminho_old)

    try:
      shutil.move(caminho_novo_exe, exe_atual)
      substituido_in_place = True
    except Exception:
      # Se falhar ao colocar o novo no lugar, restaura o executável original
      if os.path.exists(caminho_old) and not os.path.exists(exe_atual):
        os.replace(caminho_old, exe_atual)
      raise
  except Exception:
    substituido_in_place = False

  if substituido_in_place:
    limpar_residuos_atualizacao()
    flags = 0
    if os.name == "nt":
      flags = (
           getattr(subprocess, "DETACHED_PROCESS", 0x00000008)
          | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
      )
    subprocess.Popen(
        [exe_atual],
        cwd=pasta,
        env=env_limpo,
        close_fds=True,
        creationflags=flags,
    )
    os._exit(0)
    return

  # Fallback via script em %TEMP% caso o antivírus impeça o rename imediato
  bat_path = os.path.join(pasta_temp, f"_conversor_atualizar_{pid_atual}.bat")
  conteudo_bat = (
      "@echo off\n"
      "set tentativas=0\n"
      ":mover\n"
      f'move /Y "{caminho_novo_exe}" "{exe_atual}" >NUL 2>&1\n'
      "if errorlevel 1 (\n"
      "    set /a tentativas+=1\n"
      "    if %tentativas% GEQ 120 exit /b 1\n"
      "    ping 127.0.0.1 -n 2 >NUL\n"
      "    goto mover\n"
      ")\n"
      f'start "" "{exe_atual}"\n'
      f'del /Q "{bat_path}" >NUL 2>&1\n'
      "exit\n"
  )
  with open(bat_path, "w", encoding="utf-8") as f:
    f.write(conteudo_bat)

  flags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000) if os.name == "nt" else 0
  subprocess.Popen(
      ["cmd", "/c", bat_path],
      cwd=pasta_temp,
      env=env_limpo,
      close_fds=True,
      creationflags=flags,
  )
  os._exit(0)

