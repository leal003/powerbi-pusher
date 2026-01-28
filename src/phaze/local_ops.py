import time
import logging
import psutil
import win32gui
import win32process
import win32con
import win32api
from pathlib import Path
from pywinauto.application import Application
from pywinauto import Desktop
from pyvda import AppView, VirtualDesktop, get_virtual_desktops
from .exceptions import LocalAutomationError, ConnectionError

logger = logging.getLogger("phaze")

class Phaze:
    """
    Phaze Automation Library - Version 1.2.0
    -------------------------------------
    1. REFRESH: Strobe Strategy (Smart Interval).
       - Verifica visualmente a cada 60s.
       - Se detectar estagnacao, confirma em 5s.
    2. SAVE: Kernel Force Mode (PostMessage).
       - Injeta Scan Codes diretamente na fila de mensagens (Hardware Simulation).
    3. VISUAL: Virtual Desktops + Teleport.
    """
    
    VERSION = "1.2.0"

    def __init__(self, pbi_executable_path: str = None):
        self.pbi_path = pbi_executable_path or r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe"
        self.window = None
        self.pid = None
        self.main_hwnd = None
        logger.info(f"Phaze v{self.VERSION} inicializada.")

    # =========================================================================
    # 1. CONEXAO
    # =========================================================================

    def connect(self, file_path: str = None, window_name: str = None) -> bool:
        target_title = window_name.lower().strip() if window_name else \
                       Path(file_path).stem.lower().strip() if file_path else None

        if not target_title:
            raise ConnectionError("Phaze precisa de um titulo.")

        def enum_callback(hwnd, results):
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
                if target_title in win32gui.GetWindowText(hwnd).lower():
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    results.append((hwnd, pid))
            return True

        matches = []
        win32gui.EnumWindows(enum_callback, matches)
        if not matches: return False

        self.main_hwnd, self.pid = matches[0]
        self.app = Application(backend="uia").connect(handle=self.main_hwnd)
        self.window = self.app.window(handle=self.main_hwnd)
        
        time.sleep(1)
        self._move_main_to_background()
        
        logger.info(f"Conectado ao PID {self.pid}. Janela em Background (Desktop 2).")
        return True

    def _move_main_to_background(self):
        if self.main_hwnd and win32gui.IsWindow(self.main_hwnd):
            try:
                if len(get_virtual_desktops()) < 2:
                    VirtualDesktop.create()
                target_desktop = VirtualDesktop(2)
                AppView(self.main_hwnd).move(target_desktop)
                if VirtualDesktop.current().number != 1:
                    VirtualDesktop(1).go()
            except: pass

    def _bring_main_to_foreground(self):
        if self.main_hwnd and win32gui.IsWindow(self.main_hwnd):
            try:
                current_desktop = VirtualDesktop.current()
                AppView(self.main_hwnd).move(current_desktop)
                win32gui.SetWindowPos(self.main_hwnd, win32con.HWND_TOP, 0, 0, 0, 0, 
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)
                try: self.window.set_focus()
                except: pass
            except: pass

    # =========================================================================
    # 2. FERRAMENTA: REFRESH (MANTIDA V1.2.2 - SUCESSO)
    # =========================================================================

    def refresh(self):
        if not self.window: raise ConnectionError("Nao conectado.")
        logger.info("Ferramenta REFRESH iniciada.")
        
        self._bring_main_to_foreground()
        time.sleep(1)
        if not self._click_refresh_button():
            self._move_main_to_background()
            return False
        
        logger.info(" > Aguardando popup iniciar...")
        if not self._wait_for_popup_visual(timeout=30):
            logger.error("   Popup nao abriu. Abortando.")
            self._move_main_to_background()
            return False

        logger.info(" > Popup detectado. Iniciando monitoramento em background...")
        self._move_main_to_background()
        return self._monitor_smart(timeout=3600)

    def _wait_for_popup_visual(self, timeout=30):
        start = time.time()
        while (time.time() - start) < timeout:
            try:
                dialog = self.window.child_window(auto_id="KoLoadToReportDialog")
                if dialog.exists(timeout=0.5):
                    return True
            except: pass
            time.sleep(0.5)
        return False

    def _monitor_smart(self, timeout=3600):
        start_time = time.time()
        last_content = "INICIO"
        stability_count = 0
        next_wait_time = 60 

        while (time.time() - start_time) < timeout:
            logger.info(f"   Aguardando {next_wait_time}s em background...")
            time.sleep(next_wait_time)

            self._bring_main_to_foreground()
            status, content = self._check_popup_state()
            self._move_main_to_background()

            if status == "GONE":
                logger.info(" > Popup sumiu (Concluido).")
                time.sleep(2)
                self._clean_webview_crash()
                return True

            logger.info(f"   Status Lido: '{content}'")

            if content == last_content:
                stability_count += 1
                if stability_count == 1:
                    logger.info("   . Texto igual. Reduzindo intervalo para 5s (Confirmacao).")
                    next_wait_time = 5
                    continue

                if stability_count >= 2:
                    logger.info(" > Confirmado: Atualizacao estagnada/concluida.")
                    self._bring_main_to_foreground()
                    self._try_close_popup()
                    self._move_main_to_background()
                    time.sleep(2)
                    self._clean_webview_crash()
                    return True
            else:
                logger.info("   . Progresso detectado. Mantendo intervalo de 60s.")
                stability_count = 0
                last_content = content
                next_wait_time = 60

        raise LocalAutomationError("Timeout na atualizacao.")

    def _check_popup_state(self):
        try:
            time.sleep(0.5)
            dialog = self.window.child_window(auto_id="KoLoadToReportDialog")
            if not dialog.exists(timeout=0.5):
                return "GONE", None
            try:
                summary = dialog.child_window(auto_id="modalDialog")
                texts = summary.descendants(control_type="Text")
                content = " | ".join([t.window_text() for t in texts])
                return "EXISTS", content
            except:
                return "EXISTS", "Unreadable"
        except:
            return "GONE", None

    def _try_close_popup(self):
        try:
            dialog = self.window.child_window(auto_id="KoLoadToReportDialog")
            btn = dialog.child_window(auto_id="view_7")
            if btn.exists():
                if hasattr(btn, "invoke"): btn.invoke()
                else: btn.click_input()
        except: pass

    def _ensure_home_tab(self):
        try:
            for item in self.window.descendants(control_type="TabItem"):
                if any(x in item.window_text().lower() for x in ["página inicial", "home"]):
                    if hasattr(item, "select"): item.select()
                    else: item.click_input()
                    break
        except: pass

    def _click_refresh_button(self):
        try:
            self._ensure_home_tab()
            for btn in self.window.descendants(control_type="Button"):
                if any(x in btn.window_text().lower() for x in ["atualizar", "refresh"]):
                    if hasattr(btn, "invoke"): btn.invoke()
                    else: btn.click()
                    return True
        except: pass
        return False

    def _clean_webview_crash(self):
        try:
            desktop = Desktop(backend="uia")
            self._bring_main_to_foreground()
            crash_window = desktop.window(title_re=".*WebView2.*")
            if crash_window.exists(timeout=2):
                logger.warning("   ERRO WEBVIEW2 DETECTADO! Limpando...")
                crash_window.close()
                self._move_main_to_background()
                return True
            self._move_main_to_background()
        except: pass
        return False

    # =========================================================================
    # 3. FERRAMENTA: SAVE (KERNEL FORCE MODE - RESTAURADO)
    # =========================================================================

    def _make_lparam(self, vk, down=True):
        """
        [KERNEL METHOD] Constroi o lParam exato de hardware.
        """
        scancode = win32api.MapVirtualKey(vk, 0)
        lparam = 1 
        lparam |= (scancode << 16)
        if not down:
            lparam |= (1 << 30) 
            lparam |= (1 << 31) 
        return lparam

    def save(self):
        logger.info("Ferramenta SAVE iniciada (Kernel Mode / PostMessage).")
        
        # Traz para frente para garantir que o Windows roteie a mensagem
        self._bring_main_to_foreground()
        time.sleep(1.0) 

        try:
            # Pega a janela principal e todas as filhas (foco disperso)
            hwnd_list = [self.main_hwnd]
            def child_callback(hwnd, ctx):
                if win32gui.IsWindowVisible(hwnd): ctx.append(hwnd)
            try: win32gui.EnumChildWindows(self.main_hwnd, child_callback, hwnd_list)
            except: pass

            logger.info(f" > Injetando Scan Codes em {len(hwnd_list)} alvos...")

            # Prepara os Pacotes de Kernel
            lp_s_down = self._make_lparam(0x53, True)    # 'S' down
            lp_s_up = self._make_lparam(0x53, False)     # 'S' up
            lp_ctrl_down = self._make_lparam(win32con.VK_CONTROL, True)
            lp_ctrl_up = self._make_lparam(win32con.VK_CONTROL, False)

            # Disparo Massivo (Para garantir que quem tem o foco receba)
            for target_hwnd in hwnd_list:
                # Sequencia Ctrl+S
                win32api.PostMessage(target_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, lp_ctrl_down)
                win32api.PostMessage(target_hwnd, win32con.WM_KEYDOWN, 0x53, lp_s_down)
                win32api.PostMessage(target_hwnd, win32con.WM_KEYUP, 0x53, lp_s_up)
                win32api.PostMessage(target_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, lp_ctrl_up)

            logger.info("   Scan Codes injetados com sucesso.")
            
            # Reforco (Opcional, mas util): Alt+1 (Atalho padrao de salvar na barra rapida)
            time.sleep(0.5)
            lp_1_down = self._make_lparam(0x31, True)
            lp_1_up = self._make_lparam(0x31, False)
            win32api.PostMessage(self.main_hwnd, win32con.WM_SYSKEYDOWN, 0x31, lp_1_down)
            win32api.PostMessage(self.main_hwnd, win32con.WM_SYSKEYUP, 0x31, lp_1_up)

            logger.info(" > Aguardando gravacao em disco (5s)...")
            time.sleep(5)
            
            self._move_main_to_background()
            return True

        except Exception as e:
            logger.error(f"   Erro fatal no save kernel: {e}")
            self._move_main_to_background()
            return False

    def bring_back(self):
        self._bring_main_to_foreground()

    def close(self):
        if self.pid:
            psutil.Process(self.pid).terminate()