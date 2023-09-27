# Automatização de Processo com Selenium e Python

Este é um script Python que automatiza um processo específico de preenchimento de dados e download de comprovantes em um site usando a biblioteca Selenium.

## Requisitos

Antes de usar este script, certifique-se de ter instalado os seguintes requisitos:

- Python
- Selenium
- pandas
- numpy
- pyautogui
- Chrome WebDriver

Você pode instalar as bibliotecas Python necessárias usando o `pip`.

```
pip install selenium pandas numpy pyautogui
```

## Como usar
Clone este repositório em seu computador.

Crie um arquivo `.env` na raiz do projeto com suas credenciais de login, por exemplo:

**LOGIN**=suaseuemail@example.com
**PASSWORD**=suasenha123

Execute o script Python e siga as instruções para fornecer a data inicial e final.

O script automatizará o processo de pesquisa e download de comprovantes para as POs especificadas.

## Detalhes do Script

- O script usa o Selenium para automatizar a interação com o site.
- Ele lê os dados de uma planilha Excel (comprovantes.xlsx) usando o Pandas.
- Ele preenche os campos no site com base nas informações da planilha.
- Ele verifica os valores e faz o download dos comprovantes quando apropriado.
- Os resultados são atualizados na planilha.

## Autor
Lucas Henrique Couto de Carvalho

### Nota: Este script é fornecido apenas como exemplo e pode ser necessário adaptá-lo para atender aos requisitos específicos do seu sistema e ambiente.