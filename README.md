# pdf-and-excel-scraping

## Resumo
Desenvolvi esse programa para automatizar uma tarefa repetitiva de buscar valores diariamente em 5 arquivos de PDF e 3 planilhas de Excel, para no fim formatar tudo de uma certa maneira e inserir numa grande planilha de Excel. 

## Disclaimer
Não é possível compartilhar os arquivos de pdf e excel que esse script estava buscando, mas posso dizer que os arquivos não eram limpos e tinham pouca padronização, por isso a tarefa acabou se mostrando mais complicada do que eu esperava.

## Estrutura
* `app.py` -> GUI básica para tornar rodar esse script mais acessível.
* `script_final.py` -> Faz o trabalho sujo
  * sap(cod): um bot feito com pyautogui para gerar planilhas dentro de um software usado para buscar informações, variável COD altera o polo que será buscado.
  * excel_write(dic_descarga: dict, action: str): parte dos dados extraidos precisavam ser inseridos pela primeira vez no excel, e a outra parte precisava ser inserida nas informações que já estavam lá, a fim de atualizar elas. dic_descarga é justamente uma variável com os valores antigos e a sua correspondência dos dados atualizados. action determina se a função vai somente inserir dados novos, ou vai verificar a planilha para atualizar o que está lá.
  * pdf_reader(operador: str): faz leitura e extração dos PDF's, de início tentei usar read_pdf da biblioteca tabula e passar para um DataFrame, mas alguns PDF'S não passavam todas as informações dessa forma, por isso tive que usar PdfReader e ir buscando os valores numa string gigante. Num escopo de trabalho com PDF's mais padronizados dá para enxugar bastante essa função.
  * excel_reader(operador: str): faz leitura e extração de alguns arquivos de Excel, usei openpyxl sem grandes problemas, pois as planilhas eram bem mais padronizadas e com os dados 'limpos' para serem extraidos.

## Abordagem
* função sap: clica na tela em lugares específicos e digita valores a fim de navegar em um software e obter planilhas do Excel que serão lidas adiante, e que devem ser salvas na mesma pasta do arquivo.
* função excel_write: usando a ação de digitar informações novas a função busca a última linha da planilha e vai colando os novos dados extraídos do Excel e PDF seguindo uma certa formatação, na ação de atualizar informações primeiro faz-se um check das informações a serem correspondidas pelas atualizações que são inseridas a seguir.
* função pdf_reader: lê os PDF's mais padronizados com biblioteca tabula, corrigindo alguns possíveis erros e limpando algumas coisas para entrarem na formatação necessária, e lê os PDF's menos padronizados com a biblioteca PdfReader, buscando os valores em uma string. Movimentações são computadas no script para exibir acumulados com base em uma série de condições (tipo do modal e polo onde ocorrem as movimentações).
* função excel_reader: lê os arquivos de Excel com uma maior simplicidade do que a função anterior.

## Resultado
![Final](https://i.imgur.com/KUrXufj.png)

