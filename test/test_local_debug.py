import logging
import os
import sys
import time

# Setup para importar a biblioteca da pasta src
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))
from phaze.local_ops import Phaze

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s', datefmt='%H:%M:%S')

# --- CONFIGURACAO DO TESTE ---
NOME_JANELA = "NOME_JANELA"

def main():
    print("\n--- TESTE DE FLUXO MODULAR ---")
    
    # 1. Instancia a biblioteca
    bot = Phaze()

    # 2. Conecta (Isso move a janela para o Limbo -30000 para protecao)
    print("\n[1] Conectando...")
    if not bot.connect(window_name=NOME_JANELA):
        print("Erro: Janela nao encontrada.")
        return

    # 3. Chama a ferramenta REFRESH
    # O robo monitora o popup e fecha. A janela principal continua oculta.
    print("\n[2] Executando Atualizacao...")
    if bot.refresh():
        print("Atualizacao finalizada (Popup fechado).")
    else:
        print("Falha na atualizacao.")
        bot.bring_back() # Traz de volta para ver o erro
        return

    # 4. Chama a ferramenta SAVE
    # ATENCAO: Agora ele salva via Teclado (Alt+1 / Ctrl+S) SEM trazer a janela
    # de volta para a tela, evitando o crash do WebView2.
    print("\n[3] Salvando Arquivo (Modo Blind/Teclado)...")
    if bot.save():
        print("Comandos de salvar enviados.")
    else:
        print("Falha ao salvar.")

    # 5. Fim (So AGORA trazemos a janela de volta para conferencia visual)
    print("\n[4] Processo Concluido. Restaurando janela...")
    bot.bring_back() 
    
    # bot.close() # Descomente se quiser fechar o app no final

if __name__ == "__main__":
    main()
