![background](https://github.com/kawanbez/relatorio_prefeitura_barueri_vba/blob/main/bg.jpg)


# Extração Relatório de Faturamento: Prefeitura Barueri  - VBA

Em Junho de 2019 a empresa em que atuo transferiu-se de São Paulo para Barueri (interior de São Paulo). Com isso, sua tributação começou a ser recolhida pela Prefeitura de Barueri.
Após a mudança, a área financeira teve que adaptar-se aos novos padrões de relatórios oriundos da Prefeitura de Barueri, pois anteriormente em São Paulo, o relatório era disponibilizado em Excel e em Barueri era somente disponibilizado em txt. Devido a essa mudança, todos os relatórios e integrações com sistemas que estavam adaptados para a Prefeitura de São Paulo, necessitávam ser atualizados para o novo modelo de exportação de dados.
Com isso, houve a necessidade de criação de um código em VBA que transformasse o txt da Prefeitura de Barueri para Excel de acordo com as antigas orientações do relatório de São Paulo.

### Analisando o Manual da Prefeitura de Barueri

O primeiro passo foi analisar o manual de layout para exportação de dados disponibilizado pela prefeitura de Barueri. Nele encontraremos as orientações de layout que o arquivo em txt segue, com informações como: Descrição do Campo, Tipo, Tamanho, Posição Inicial, Posição Inicial e Conteúdo:
[Documentação Prefeitura](https://github.com/kawanbez/relatorio_prefeitura_barueri_vba/blob/main/NFE_Layout.pdf)

### Criação do Código em VBA

Após a análise da documentação, desenvolvi o código em VBA com base nas orientações citadas no manual: [Código de Extração](https://github.com/kawanbez/relatorio_prefeitura_barueri_vba/blob/main/codigo_de_extracao)  
Em resumo, o código busca dentro do txt as orientações que defini em código, como por exemplo a posição e tamanho dos campos número da nota, data de emissão, tributação, razão social do tomador, etc. - [Arquivo txt de Exemplo](https://github.com/kawanbez/relatorio_prefeitura_barueri_vba/blob/main/48941143DE20201030%20-%20MODELO.txt) 

### Definição do Processo

Após a criação do código e obtenção das informações, definimos a seguinte rotina dentro do departamento de faturamento:

* Download dos arquivos txt da Prefeitura de Barueri (cada arquivo se refere a um dia de faturamento)
* Salvar os arquivos em uma pasta específica dentro do drive compartilhado (no arquivo de exemplo determinei a pasta: HostFolder = "C:\Users\kawan.bezerra\Downloads\NFs Barueri\", portanto, em caso de testes, altere o caminho do arquivo) - [Planilha de Exemplo](https://github.com/kawanbez/relatorio_prefeitura_barueri_vba)  
* Após o salvamento dos arquivos, é feita a abertura da planilha em Excel com o código em VBA inserido e é realizado o start no código
* Após o preenchimento da planilha com base na execução do código, é criada uma base mensal de faturamento composta pelos dados individuais das execuções diarias pelo time de faturametno

### Conclusão

Após a criação do código e criação do processo, o faturamento pôde novamente ser apurado e seguir para fechamento contábil e fiscal.



