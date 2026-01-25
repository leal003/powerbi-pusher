import time
import logging
import psutil
import win32gui
import win32process
import win32con
from pathlib import Path
from pywinauto.application import Application
from .exceptions import LocalAutomationError

# Logger configurado para a marca Phaze
logger = logging.getLogger("phaze")

class Phaze:
    """
    Phaze Automation Library - Version 1.0
    -------------------------------------
    Especializada em automacao furtiva de Power BI Desktop.
    
    Recursos:
    - Limbo Mode: Move janelas para -30000px mantendo a arvore UIA ativa.
    - PID Isolation: Garante que apenas janelas do processo alvo sejam movidas.
    - Deep Search: Ignora janelas fantasmas conectando apenas na que possui conteudo.
    """
    
    VERSION = "1.0"

    def __init__(self, pbi_executable_path: str = None):
        self.pbi_path = pbi_executable_path or r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe"
        self.window = None
        self.pid = None
        self.hwnd = None
        self.limbo_coords = (-30000, -30000)
        logger.info(f"Phaze v{self.VERSION} inicializada.")

    # =========================================================================
    # GESTAO DE JANELAS E CONEXAO
    # =========================================================================

    def connect(self, file_path: str = None, window_name: str = None) -> bool:
        """Conecta ao BI e o isola no limbo imediatamente."""
        target_title = window_name.lower().strip() if window_name else \
                       Path(file_path).stem.lower().strip() if file_path else None

        if not target_title:
            raise LocalAutomationError("Phaze precisa de um titulo ou caminho de arquivo.")

        def enum_callback(hwnd, results):
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
                if target_title in win32gui.GetWindowText(hwnd).lower():
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    results.append((hwnd, pid))
            return True

        matches = []
        win32gui.EnumWindows(enum_callback, matches)
        if not matches: return False

        # Captura identidade do processo
        self.hwnd, self.pid = matches[0]
        
        # Teletransporta TUDO do processo (janela principal + popups futuros)
        self.teleport_process_to_limbo()
        
        # Vincula o motor UIA
        self.app = Application(backend="uia").connect(handle=self.hwnd)
        self.window = self.app.window(handle=self.hwnd)
        
        logger.info(f"Phaze conectada ao PID {self.pid} (Janela no Limbo).")
        return True

    def teleport_process_to_limbo(self):
        """Varre o sistema e joga qualquer janela do PID alvo para fora do monitor."""
        def enum_and_move(hwnd, ctx):
            if win32gui.IsWindow(hwnd):
                _, win_pid = win32process.GetWindowThreadProcessId(hwnd)
                if win_pid == self.pid:
                    rect = win32gui.GetWindowRect(hwnd)
                    # So move se a janela aparecer no monitor (x > -10000)
                    if rect[0] > -10000:
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        win32gui.SetWindowPos(
                            hwnd, win32con.HWND_TOP, 
                            self.limbo_coords[0], self.limbo_coords[1], 
                            1280, 720, win32con.SWP_SHOWWINDOW
                        )
            return True
        win32gui.EnumWindows(enum_and_move, None)

    # =========================================================================
    # ACOES EXECUTORAS (METODO TOOLKIT)
    # =========================================================================

    def go_to_home_tab(self):
        """Localiza e clica na aba 'Pagina Inicial'."""
        try:
            for item in self.window.descendants(control_type="TabItem"):
                name = item.window_text().lower()
                if "pagina inicial" in name or "home" in name or item.element_info.automation_id == "home":
                    if hasattr(item, "select"): item.select()
                    else: item.click_input()
                    logger.info("   Aba Home ativada.")
                    return True
        except: pass
        return False

    def click_refresh(self):
        """Localiza e clica no botao 'Atualizar'."""
        try:
            for btn in self.window.descendants(control_type="Button"):
                t = btn.window_text().lower()
                aid = btn.element_info.automation_id
                if t == "atualizar" or t == "refresh" or aid == "refreshQueries":
                    if hasattr(btn, "invoke"): btn.invoke()
                    else: btn.click()
                    logger.info("   Botao Atualizar acionado.")
                    return True
        except: pass
        return False

    def close_refresh_popup(self):
        """Localiza janelas de progresso e clica em 'Fechar'."""
        self.teleport_process_to_limbo() 
        try:
            desktop = Application(backend="uia").connect(process=self.pid)
            for win in desktop.windows():
                if any(x in win.window_text().lower() for x in ["atualizar", "refresh", "concluido"]):
                    for btn in win.descendants(control_type="Button"):
                        if btn.window_text() in ["Fechar", "Close"]:
                            if hasattr(btn, "invoke"): btn.invoke()
                            else: btn.click()
                            logger.info("   Popup de atualizacao fechado.")
                            return True
        except: pass
        return False

    def save(self):
        """Salva o arquivo via ID mapeado ou comando de teclado."""
        try:
            for btn in self.window.descendants(control_type="Button"):
                if btn.element_info.automation_id == "save" or btn.window_text() == "Salvar":
                    if hasattr(btn, "invoke"): btn.invoke()
                    else: btn.click()
                    logger.info("   Salvo via clique no botao.")
                    return True
        except: pass

        self.window.type_keys("^s")
        logger.info("   Salvo via sinal de teclado (Ctrl+S).")
        return True

    # =========================================================================
    # UTILITARIOS
    # =========================================================================

    def bring_back(self):
        """Traz o Power BI de volta para a coordenada visivel (100, 100)."""
        if self.hwnd:
            win32gui.SetWindowPos(self.hwnd, win32con.HWND_TOP, 100, 100, 1280, 720, win32con.SWP_SHOWWINDOW)
            logger.info("Janela retornada a area visivel.")

    def close(self):
        """Encerra o processo do Power BI."""
        if self.pid:
            logger.info(f"Phaze: Encerrando processo {self.pid}...")
            psutil.Process(self.pid).terminate()