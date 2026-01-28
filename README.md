# Phaze v1.2.0

**[PT]** Biblioteca para automação silenciosa, robusta e "headless" do Microsoft Power BI Desktop.
**[EN]** Library for silent, robust, and "headless" automation of Microsoft Power BI Desktop.

---

## Principais Recursos / Key Features

### Português

* **Virtual Desktop Isolation:** Em vez de apenas esconder a janela, a Phaze move o Power BI para uma Área de Trabalho Virtual (Desktop 2). Isso garante que o Power BI continue renderizando gráficos (evitando erros do WebView2) sem atrapalhar o seu mouse ou teclado no Desktop Principal.
* **Hybrid Force Save:** O sistema de salvamento mais robusto até agora. Combina **Cliques Visuais** (botão disquete) com **Injeção de Hardware** (`SendInput` + `PostMessage`). É impossível para o Power BI ignorar o comando de salvar.
* **Strobe Smart Refresh:** Monitoramento intermitente inteligente. O robô mantém o PBI oculto e o traz para o topo por milissegundos a cada 60s para verificar o status visualmente (leitura 100% precisa), devolvendo-o imediatamente. Se detectar estabilidade, reduz o intervalo para 5s para confirmação final.
* **Crash Recovery:** Detecta e trata automaticamente o erro "Há um problema com o WebView2" e janelas de Auto-Close.

### English

* **Virtual Desktop Isolation:** Instead of just hiding the window, Phaze moves Power BI to a Virtual Desktop (Desktop 2). This ensures Power BI continues rendering charts (avoiding WebView2 errors) without disturbing your mouse or keyboard on the Main Desktop.
* **Hybrid Force Save:** The most robust saving system to date. Combines **Visual Clicks** (save button) with **Hardware Injection** (`SendInput` + `PostMessage`). It is impossible for Power BI to ignore the save command.
* **Strobe Smart Refresh:** Intelligent intermittent monitoring. The bot keeps PBI hidden and brings it to the front for milliseconds every 60s to check status visually (100% accurate reading), returning it immediately. If stability is detected, it reduces the interval to 5s for final confirmation.
* **Crash Recovery:** Automatically detects and handles "There is a problem with WebView2" errors and Auto-Close windows.

---

## Instalação / Installation

bash

```pip install phaze ```

## Como Usar / How to Use

A Phaze v1.2.0 funciona como um "Toolkit Modular". Você decide a ordem das ações. 
Phaze v1.2.0 operates as a "Modular Toolkit". You decide the order of actions.

# Exemplo
```
import time
from phaze.local_ops import Phaze

# Configuração
NOME_JANELA = "Nome do Seu Arquivo PBI"

def main():
    # 1. Inicializa / Initialize
    driver = Phaze()

    # 2. Conecta / Connect
    # A janela será movida para o Desktop 2 (Background) imediatamente.
    # The window will be moved to Desktop 2 (Background) immediately.
    print("Conectando...")
    if not driver.connect(window_name=NOME_JANELA):
        print("Janela não encontrada.")
        return

    # 3. Atualiza / Refresh (Strobe Strategy)
    print("Atualizando dados...")
    
    # O robô vai checar o status visualmente a cada 60s (Strobe)
    if driver.refresh():
        print("Atualização concluída com sucesso.")
    else:
        print("Falha na atualização.")
        driver.bring_back()
        return

    # 4. Salva / Save (Hybrid)
    # Tenta clicar no botão, se falhar, usa injeção de teclado físico.
    print("Salvando...")
    if driver.save():
        print("Comando de salvar enviado.")
    else:
        print("Erro ao salvar.")

    # 5. Finaliza / Finish
    driver.bring_back() 
    print("Processo finalizado.")

if __name__ == "__main__":
    main()
```
## Índice de Drivers (API Reference)

A classe `Phaze` expõe os seguintes métodos para controle total:

| Método / Method | Descrição (PT) | Description (EN) |
| :--- | :--- | :--- |
| `connect(window_name)` | Localiza a janela, conecta ao processo e a move para o **Virtual Desktop 2**. | Locates the window, connects to the process, and moves it to **Virtual Desktop 2**. |
| `refresh()` | Executa a **Strobe Strategy**: Traz para frente > Lê Status > Esconde. Intervalos inteligentes de 60s/5s. | Runs **Strobe Strategy**: Bring to front > Read Status > Hide. Smart intervals of 60s/5s. |
| `save()` | Salva usando método **Híbrido**: Clique Visual + Hardware Injection (Ctrl+S). | Saves using **Hybrid** method: Visual Click + Hardware Injection (Ctrl+S). |
| `bring_back()` | Traz a janela do Desktop Virtual de volta para o Desktop Principal. | Brings the window from the Virtual Desktop back to the Main Desktop. |
| `close()` | Encerra forçadamente o processo do Power BI (Kill process). | Forcefully terminates the Power BI process. |

## Contribuição / Contributing

- Faça um Fork do projeto.
- Crie uma branch para sua modificação: ``` git checkout -b feature/melhoria-direta ```.
- Envie um Pull Request detalhando o que foi feito.
