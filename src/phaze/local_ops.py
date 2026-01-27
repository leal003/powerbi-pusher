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
from .exceptions import LocalAutomationError, ConnectionError

logger = logging.getLogger("phaze")

class Phaze:
    """
    Phaze Automation Library - Version 1.1.4
    -------------------------------------
    1. REFRESH: Monitoramento de estabilidade, deteccao de Auto-Close e Crash Recovery.
    2. SAVE: Implementacao de 'Hardware Simulation'.
       - Constroi 'lParams' complexos com Scan Codes reais.
       - Dispara para a janela principal e filhos para garantir recepcao.
    """
    
    VERSION = "1.1.4"

    def __init__(self, pbi_executable_path: str = None):
        self.pbi_path = pbi_executable_path or r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe"
        self.window = None
        self.pid = None
        self.main_hwnd = None
        self.limbo_coords = (-30000, -30000)
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
        
        self._move_main_to_limbo()
        
        logger.info(f"Conectado ao PID {self.pid}. Janela oculta no Limbo.")
        return True

    def _move_main_to_limbo(self):
        if self.main_hwnd and win32gui.IsWindow(self.main_hwnd):
            win32gui.SetWindowPos(
                self.main_hwnd, win32con.HWND_TOP, 
                self.limbo_coords[0], self.limbo_coords[1], 
                1280, 720, 
                win32con.SWP_NOACTIVATE | win32con.SWP_ASYNCWINDOWPOS
            )

    # =========================================================================
    # 2. FERRAMENTA: REFRESH
    # =========================================================================

    def refresh(self):
        if not self.window: raise ConnectionError("Nao conectado.")
        logger.info("Ferramenta REFRESH iniciada.")
        self._ensure_home_tab()
        if not self._click_refresh_button():
            return False
        return self._monitor_popup_logic()

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
            for btn in self.window.descendants(control_type="Button"):
                if any(x in btn.window_text().lower() for x in ["atualizar", "refresh"]):
                    if hasattr(btn, "invoke"): 
                        btn.invoke()
                    else: 
                        btn.click()
                    logger.info(" > Clique enviado para 'Atualizar'.")
                    return True
        except: pass
        return False

    def _clean_webview_crash(self):
        """FALLBACK DE ERRO: Fecha popup de erro WebView2 se aparecer."""
        logger.info("   Verificando se houve crash do WebView2...")
        try:
            desktop = Desktop(backend="uia")
            crash_window = desktop.window(title_re=".*WebView2.*")
            
            if crash_window.exists(timeout=3):
                logger.warning("   ERRO WEBVIEW2 DETECTADO! Executando limpeza...")
                crash_window.close()
                crash_window.wait_not('visible', timeout=5)
                logger.info("   Janela de erro fechada. Interface liberada.")
                return True
            else:
                logger.info("   Nenhum erro detectado.")
        except Exception as e:
            logger.debug(f"   (Check de erro finalizado sem acoes: {e})")
        return False

    def _monitor_popup_logic(self, timeout=3600):
        logger.info(" > Monitorando popup...")
        start_time = time.time()
        
        last_content = ""
        stability_counter = 0
        STABILITY_THRESHOLD = 3 
        dialog_seen = False 

        while (time.time() - start_time) < timeout:
            try:
                self._move_main_to_limbo()
                dialog = self.window.child_window(auto_id="KoLoadToReportDialog")
                
                if dialog.exists(timeout=0.5):
                    dialog_seen = True 
                    close_btn = dialog.child_window(auto_id="view_7", control_type="Button")
                    
                    if close_btn.exists(timeout=0.1):
                        btn_text = close_btn.window_text().lower()
                        if "cancel" in btn_text: 
                            stability_counter = 0
                            continue

                        summary = dialog.child_window(auto_id="modalDialog")
                        if summary.exists(timeout=0.1):
                            all_texts = [t.window_text() for t in summary.descendants(control_type="Text")]
                            current_content = " | ".join(all_texts)

                            forbidden = ["carregando", "avaliando", "aguardando", "loading", "waiting", "model"]
                            if any(bad in current_content.lower() for bad in forbidden):
                                stability_counter = 0
                                last_content = current_content
                                time.sleep(2)
                                continue

                            if current_content == last_content and current_content != "":
                                stability_counter += 1
                                logger.info(f"   . Estavel ({stability_counter}/{STABILITY_THRESHOLD})")
                            else:
                                stability_counter = 0
                                last_content = current_content

                            if stability_counter >= STABILITY_THRESHOLD:
                                logger.info(" > Estabilidade confirmada. Clicando Fechar.")
                                if hasattr(close_btn, "invoke"): 
                                    close_btn.invoke()
                                else: 
                                    close_btn.click_input()
                                
                                try: dialog.wait_not('visible', timeout=10)
                                except: pass
                                
                                time.sleep(2)
                                self._clean_webview_crash()
                                self._apply_cooldown()
                                return True
                else:
                    if dialog_seen:
                        logger.info(" > Popup desapareceu (Auto-Close).")
                        time.sleep(2)
                        self._clean_webview_crash()
                        self._apply_cooldown()
                        return True
                    else:
                        pass 

            except Exception:
                pass 
            
            time.sleep(2.0)
            
        raise LocalAutomationError("Timeout na atualizacao.")

    def _apply_cooldown(self):
        logger.info(" > Cool-down: 3s para estabilizacao...")
        time.sleep(3)

    # =========================================================================
    # 3. FERRAMENTA: SAVE (HARDWARE SIMULATION)
    # =========================================================================

    def _make_lparam(self, vk, down=True):
        """
        Cria o lParam real de um evento de teclado.
        Isso e crucial para que aplicacoes .NET/WPF aceitem o input injetado.
        """
        scancode = win32api.MapVirtualKey(vk, 0)
        lparam = 1 # Repeat count
        lparam |= (scancode << 16)
        if not down:
            lparam |= (1 << 30) # Previous key state
            lparam |= (1 << 31) # Transition state
        return lparam

    def save(self):
        """
        Envia Ctrl+S com simulacao completa de hardware (Scan Codes).
        Tenta enviar para a janela principal e para todos os filhos (ex: WebView).
        """
        logger.info("Ferramenta SAVE iniciada (Hardware Injection).")
        
        try:
            hwnd_list = [self.main_hwnd]
            
            # Tenta pegar janelas filhas (o foco real pode estar nelas)
            def child_callback(hwnd, ctx):
                if win32gui.IsWindowVisible(hwnd): 
                    ctx.append(hwnd)
            try: win32gui.EnumChildWindows(self.main_hwnd, child_callback, hwnd_list)
            except: pass

            logger.info(f" > Disparando Ctrl+S para {len(hwnd_list)} alvos...")

            # Prepara os lParams (Scan Codes Reais)
            lp_s_down = self._make_lparam(0x53, True) # 0x53 = 'S'
            lp_s_up = self._make_lparam(0x53, False)
            lp_ctrl_down = self._make_lparam(win32con.VK_CONTROL, True)
            lp_ctrl_up = self._make_lparam(win32con.VK_CONTROL, False)

            # Loop de envio (Tiro de dispersao para garantir que quem tem o foco receba)
            for target_hwnd in hwnd_list:
                # 1. Ctrl Down
                win32api.PostMessage(target_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, lp_ctrl_down)
                # 2. S Down
                win32api.PostMessage(target_hwnd, win32con.WM_KEYDOWN, 0x53, lp_s_down)
                # 3. S Up
                win32api.PostMessage(target_hwnd, win32con.WM_KEYUP, 0x53, lp_s_up)
                # 4. Ctrl Up
                win32api.PostMessage(target_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, lp_ctrl_up)

            logger.info("   Comandos injetados.")
            
            # REFORCO: Atalho Alt+1 (Quick Access Save) tambem com Scan Codes
            time.sleep(0.5)
            logger.info(" > Reforco: Enviando Alt+1 (Hardware)...")
            
            lp_1_down = self._make_lparam(0x31, True) # '1'
            lp_1_up = self._make_lparam(0x31, False)
            
            win32api.PostMessage(self.main_hwnd, win32con.WM_SYSKEYDOWN, 0x31, lp_1_down)
            win32api.PostMessage(self.main_hwnd, win32con.WM_SYSKEYUP, 0x31, lp_1_up)
            
            time.sleep(5) # Delay generoso para garantir gravacao em disco
            return True

        except Exception as e:
            logger.error(f"   Erro fatal no save: {e}")
            return False

    # =========================================================================
    # 4. UTILS
    # =========================================================================

    def bring_back(self):
        if self.main_hwnd and win32gui.IsWindow(self.main_hwnd):
            win32gui.SetWindowPos(self.main_hwnd, win32con.HWND_TOP, 0, 0, 1280, 720, win32con.SWP_SHOWWINDOW)

    def close(self):
        if self.pid:
            psutil.Process(self.pid).terminate()