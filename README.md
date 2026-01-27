# Phaze v1.1.4

**[PT]** Biblioteca para automação silenciosa, robusta e "headless" do Microsoft Power BI Desktop.
**[EN]** Library for silent, robust, and "headless" automation of Microsoft Power BI Desktop.

---

## Principais Recursos / Key Features

### Português

* **Limbo Mode:** Move a janela do Power BI para coordenadas negativas (-30000), permitindo que você continue usando o computador enquanto o robô trabalha.
* **Kernel Force Save:** Ignora falhas de clique visual. Utiliza injeção direta de mensagens no Kernel do Windows (`PostMessage` com Scan Codes reais) para forçar o salvamento (Ctrl+S) mesmo com a janela oculta.
* **Crash Recovery:** Detecta e trata automaticamente o erro "Há um problema com o WebView2" que ocorre ao renderizar gráficos fora da tela.
* **Smart Refresh:** Monitora a estabilidade visual do popup de atualização e detecta janelas que fecham sozinhas (Auto-Close).

### English

* **Limbo Mode:** Moves the Power BI window to negative coordinates (-30000), allowing you to use the computer while the bot works.
* **Kernel Force Save:** Bypasses visual click failures. Uses direct Kernel message injection (`PostMessage` with real Scan Codes) to force save (Ctrl+S) even when the window is hidden.
* **Crash Recovery:** Automatically detects and handles the "There is a problem with WebView2" error that occurs when rendering charts off-screen.
* **Smart Refresh:** Monitors the visual stability of the refresh popup and detects windows that close automatically (Auto-Close).

---

## Instalação / Installation

```bash
pip install phaze.

## Como Usar / How to Use

A Phaze v1.1.4 funciona como um "Toolkit Modular". Você decide a ordem das ações.Phaze v1.1.4 operates as a "Modular Toolkit". 
You decide the order of actions.

# Exemplo Completo (The Maestro Script)

import time from phaze.local_ops import Phaze

# Configuração
NOME_JANELA = "Nome do Seu Arquivo PBI"

def main():
    # 1. Inicializa
    bot = Phaze()

    # 2. Conecta (A janela será movida para o Limbo imediatamente)
    print("Conectando...")
    if not bot.connect(window_name=NOME_JANELA):
        print("Janela não encontrada.")
        return

    # 3. Atualiza (Refresh)
    # O robô clica, monitora o popup, fecha e aguarda estabilização (cool-down)
    print("Atualizando dados...")
    if bot.refresh():
        print("Atualização concluída com sucesso.")
    else:
        print("Falha na atualização.")
        bot.bring_back() # Traz de volta para ver o erro
        return

    # 4. Salva (Save)
    # Usa injeção de Kernel. Funciona mesmo com a janela invisível.
    print("Salvando...")
    if bot.save():
        print("Comando de salvar enviado.")
    else:
        print("Erro ao salvar.")

    # 5. Finaliza
    # Traz a janela de volta para a tela principal
    bot.bring_back() 
    print("Processo finalizado.")

if __name__ == "__main__":
    main()

## Índice de Drivers (API Reference)

A classe `Phaze` expõe os seguintes métodos para controle total:

| Método / Method | Descrição (PT) | Description (EN) |
| :--- | :--- | :--- |
| `connect(window_name)` | Localiza a janela, conecta ao processo e ativa o **Limbo Mode** (oculta a janela). | Locates the window, connects to the process, and activates **Limbo Mode** (hides the window). |
| `refresh()` | Executa o ciclo completo: Clica em Atualizar > Monitora Estabilidade > Fecha Popup > Trata Erros WebView2. | Runs full cycle: Click Refresh > Monitor Stability > Close Popup > Handle WebView2 Errors. |
| `save()` | Tenta salvar usando **Hardware Injection** (PostMessage) e Atalhos de Teclado. Não requer mouse. | Attempts to save using **Hardware Injection** (PostMessage) and Keyboard Shortcuts. Mouse-free. |
| `bring_back()` | Restaura a janela do Limbo para a posição (0,0) na tela principal. | Restores the window from Limbo to position (0,0) on the main screen. |
| `close()` | Encerra forçadamente o processo do Power BI (Kill process). | Forcefully terminates the Power BI process. |

## Contribuição / Contributing

- Faça um Fork do projeto.
- Crie uma branch para sua modificação: git checkout -b feature/melhoria-direta.
- Envie um Pull Request detalhando o que foi feito.