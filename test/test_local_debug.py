import logging
import os
import sys
import time

# Hack para importar a lib 'src'
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from powerbi_pusher.local_ops import PowerBIDriver
from powerbi_pusher.exceptions import LocalAutomationError

# Configuração de Logs
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# --- CONFIGURAÇÃO ---
ARQUIVO_TESTE = r"C:\Users\U5512793\Downloads\Painel de Pendências - BackOffice.pbix"
NOME_JANELA_FORCADO = "Painel de Pendências - BackOffice" 
FECHAR_AO_FINAL = False 

# Tempo que você sabe que o BI leva para atualizar (ex: 5 minutos)
TEMPO_ESPERA_ATUALIZACAO = 60 
# --------------------

def teste_fluxo_cronometrado():
    print(f"--- 🚀 INICIANDO TESTE (Fluxo Cronometrado) ---")
    driver = PowerBIDriver()

    try:
        # 1. Conectar e enviar para o limbo
        if not driver.connect(file_path=ARQUIVO_TESTE, window_name=NOME_JANELA_FORCADO):
            print("❌ Falha na conexão inicial.")
            return

        # 2. Preparar aba
        driver.go_to_home_tab()

        # 3. Disparar Refresh
        if driver.click_refresh():
            print(f"🔄 Atualização iniciada. Aguardando {TEMPO_ESPERA_ATUALIZACAO}s...")
            
            # ESPERA MANUAL: Aqui o script aguarda o tempo definido
            time.sleep(TEMPO_ESPERA_ATUALIZACAO)
            
            # 4. Forçar o fechamento do popup (agora que o tempo passou)
            print("🎯 Chamando função para fechar popup...")
            driver.close_refresh_popup()

            # 5. Salvar
            driver.save()
            print("\n✅ PROCESSO CONCLUÍDO E SALVO.")
        else:
            print("❌ Não foi possível clicar no botão Atualizar.")

    except Exception as e:
        print(f"\n☠️ ERRO DURANTE O TESTE: {e}")
    finally:
        if FECHAR_AO_FINAL:
            driver.close()

if __name__ == "__main__":
    teste_fluxo_cronometrado()