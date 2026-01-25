##Phaze v1.0

Biblioteca para automação silenciosa de atualização e salvamento de arquivos do Power BI Desktop.

## Principais Recursos

- **Limbo Mode**: Executa ações em coordenadas invisíveis (-30000px).

- **Process Hunter**: Identifica a janela correta analisando a densidade de elementos UI.

- **Stealth**: Atualiza e salva sem interferir no uso do mouse ou teclado do usuário.

## English Description

Library for silent automation of updating and saving Power BI Desktop files.

## Key Features

- **Limbo Mode**: Executes actions at invisible coordinates (-30000px).

- **Process Hunter**: Identifies the correct window by analyzing UI element density.

- **Stealth**: Updates and saves without interfering with the user's mouse or keyboard.

## Nota do Autor / Author's Note

Português: Sou novo no desenvolvimento e utilizei inteligencia artificial para materializar minhas ideias e colocar este projeto para rodar. O foco aqui e funcionalidade bruta para resolver o problema. Aceito colaboracoes de quem quiser chegar junto para otimizar e melhorar o codigo.

English: I am new to development and used AI to materialize my ideas and get this project running. The focus here is raw functionality to get the job done. I welcome collaborations from anyone who wants to help optimize and improve the code.

## Instalação / Installation
Bash
pip install .
Como Usar / How to Use
Python
from phaze.local_ops import Phaze
import time

driver = Phaze()

if driver.connect(window_name="NOME_DA_JANELA"):
    driver.go_to_home_tab()
    driver.click_refresh()
    time.sleep(300) 
    driver.close_refresh_popup()
    driver.save()

## Contribuição / Contributing

- Faca um Fork do projeto.

- Crie uma branch para sua modificacao: git checkout -b feature/melhoria-direta.

- Envie um Pull Request detalhando o que foi feito.