import os
import sys
import time
import glob
import pyautogui
import pygetwindow as gw

DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
WINDOW_TITLE_PART = "SISACOB"
USUARIO = "gabriel.fantin"
SENHA = "050106"
TEMPO_ABERTURA = 1
EXTENSOES_VALIDAS = (".lnk", ".exe", ".url")

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0


def encontrar_atalho():
    padrao = os.path.join(DESKTOP, "*SISACOB*")
    arquivos = glob.glob(padrao)
    for arq in arquivos:
        if arq.lower().endswith(EXTENSOES_VALIDAS):
            return arq
    return None


def focar_janela_por_titulo(parte_titulo, tentativas=30, espera=0.3):
    for i in range(tentativas):
        janelas = gw.getWindowsWithTitle(parte_titulo)
        if janelas:
            janela = janelas[0]
            if janela.isMinimized:
                janela.restore()
            janela.activate()
            time.sleep(1)
            return True
        print(f"Aguardando janela... ({i+1}/{tentativas})")
        time.sleep(espera)
    return False


def clicar_relativo_janela(janela, offset_x, offset_y):
    """Clica em uma posicao relativa ao canto superior esquerdo da janela."""
    x = janela.left + offset_x
    y = janela.top + offset_y
    pyautogui.click(x, y)


def navegar_monitoria_bsc(tentativas=30, espera=0.3):
    """Clica em Monitoria no menu e depois em BSC Prioridade."""
    janelas = gw.getWindowsWithTitle(WINDOW_TITLE_PART)
    if not janelas:
        raise RuntimeError("Janela do SISACOB nao encontrada.")
    janela = janelas[0]

    # Clica em "Monitoria" no menu (ajuste os offsets se necessario)
    print("Clicando em 'Monitoria'...")
    clicar_relativo_janela(janela, 224, 27)
    time.sleep(1)

    # Procura e clica em "BSC Prioridade" no submenu
    print("Procurando 'BSC Prioridade' no submenu...")
    encontrou = False
    for i in range(tentativas):
        pos = pyautogui.locateOnScreen(
            os.path.join(DESKTOP, "codigos", "img", "bsc_prioridade.png"),
            confidence=0.4,
        )
        if pos:
            pyautogui.click(pyautogui.center(pos))
            encontrou = True
            print("Clicou em 'BSC Prioridade'.")
            break
        time.sleep(espera)

    if not encontrou:
        raise RuntimeError(
            "Nao encontrou 'BSC Prioridade'. Salve um recorte do botao como "
            "img/bsc_prioridade.png na pasta codigos."
        )


def main():
    if not SENHA:
        raise RuntimeError("Senha vazia.")

    atalho = encontrar_atalho()
    if not atalho:
        raise RuntimeError(f"Nao encontrei atalho do SISACOB na area de trabalho: {DESKTOP}")

    print(f"Abrindo: {atalho}")
    os.startfile(atalho)
    print(f"Aguardando {TEMPO_ABERTURA}s para o ERP carregar...")
    time.sleep(TEMPO_ABERTURA)

    focar_janela_por_titulo(WINDOW_TITLE_PART)

    time.sleep(0.2)
    pyautogui.write(USUARIO, interval=0.00)
    time.sleep(0.0)
    pyautogui.press("tab")
    time.sleep(0.0)
    pyautogui.write(SENHA, interval=0.0)
    time.sleep(0.0)
    pyautogui.press("tab")
    time.sleep(0.00)
    pyautogui.press("enter")

if __name__ == "__main__":
    main()
