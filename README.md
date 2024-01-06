# Automatização de Atualização de Produtos - Script Python



## Este script em Python foi desenvolvido para automatizar a atualização de produtos em um site desktop a partir de uma planilha Excel. O processo envolve copiar dados específicos da planilha e colá-los nos campos correspondentes do site usando a biblioteca pyautogui.




### Pré-requisitos

 1. Python: Certifique-se de ter o Python instalado em seu sistema. Você pode baixá-lo em python.org.
 2. Bibliotecas Python: As bibliotecas necessárias são openpyxl, pyperclip e pyautogui. Você pode instalá-las utilizando o comando:

    
      ``` pip install openpyxl pyperclip pyautogui```
 3. Abra o seguinte site <https://cadastro-produtos-devaprender.netlify.app/> ao lado do editor de código, divindo a tela onde o navegador fica do lado esquerdo e o editor de código no lado direto.

### Como Usar

 - Planilha de Produtos: Certifique-se de ter uma planilha chamada produtos_ficticios.xlsx no mesmo diretório do script, com a estrutura esperada.
 - Execução do Script: Execute o script Python. Ele abrirá a planilha, lerá os dados e os inserirá nos campos correspondentes do site.

### Funcionamento do Script

- O script realiza as seguintes etapas:

    1. Leitura da Planilha: Abre a planilha produtos_ficticios.xlsx e seleciona a aba 'Produtos'.
    2. Atualização de Campos no Site: Itera sobre cada linha da planilha, copia os dados e cola nos campos correspondentes do site.
    3. Cliques e Atalhos: Utiliza cliques do mouse e atalhos do teclado (ctrl + v) para inserir as informações.
    4. Tratamento Específico: Trata especificamente o campo 'tamanho', clicando em posições diferentes com base no valor na planilha.
    5. Conclusão e Navegação: Clica em botões 'Próximo' e 'Concluir' no site para avançar nas etapas.
    6. Finalização do Processo: Realiza cliques adicionais para finalizar o processo de atualização.

### Observações

- Este script foi projetado para um contexto específico e pode necessitar de ajustes conforme as particularidades do site ou da planilha utilizada.
    Certifique-se de ajustar as coordenadas de clique (pyautogui.click()) de acordo com a interface do seu site.
    Recomenda-se testar o script em um ambiente de teste antes de aplicá-lo a dados reais.

 **Nota: Este script é apenas um projeto para portifólio baseado no video do Youtuber Dev Aprender | Jhonathan de Souza e seus dados são dados fictícios apenas para aprendizado.**
