![background](https://github.com/kawanbez/relatorio_prefeitura_barueri_vba/blob/main/bg.jpg)

# [PORTUGUÊS] PT-BR (English version bellow)
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

# [ENGLISH]
# Billing Report Exportation: Barueri City - VBA

In June 2019, the company I work for moved from São Paulo to Barueri. With that, its taxation began to be collected by the City Hall of Barueri.
After the change, the financial area had to adapt to the new reporting standards from the City of Barueri, as previously in São Paulo, the report was made available in Excel and in Barueri it was only made available in txt. Due to this change, all the adapted reports and integrations with systems that were being adapted for the São Paulo City Hall needed to be updated to the new data export model.
As a result, there was a need to create a code in VBA that would transform the txt from the City of Barueri to Excel according to the old guidelines of the São Paulo report.

### Analyzing the Barueri City Hall Manual

The first step was to analyze the layout manual for data export provided by the city of Barueri. In it we will find the layout guidelines that the txt file follows, with information such as: Field Description, Type, Size, Initial Position, Initial Position and Content:
[City Documentation](https://github.com/kawanbez/relatorio_prefeitura_barueri_vba/blob/main/NFE_Layout.pdf)

### Code Creation in VBA

After analyzing the documentation, I developed the code in VBA based on the guidelines cited in the manual: [Extraction Code](https://github.com/kawanbez/relatorio_prefeitura_barueri_vba/blob/main/codigo_de_extracao)
In summary, the code searches within the txt for the guidelines that I defined in code, such as the position and size of the invoice number, date of issue, taxation, corporate name of the borrower, etc. - [Example txt file](https://github.com/kawanbez/relatorio_prefeitura_barueri_vba/blob/main/48941143DE20201030%20-%20MODELO.txt)

### Process Definition

After creating the code and obtaining the information, we defined the following routine within the billing department:

* Download txt files from Barueri City Hall (each file refers to one billing day)
* Save the files in a specific folder inside the shared drive (in the example file I determined the folder: HostFolder = "C:\Users\kawan.bezerra\Downloads\NFs Barueri\", therefore, in case of tests, change the path from the file) - [Example Worksheet](https://github.com/kawanbez/relatorio_prefeitura_barueri_vba)
* After saving the files, the Excel spreadsheet is opened with the VBA code inserted and the start in the code is performed
* After filling in the worksheet based on the execution of the code, a monthly billing base is created consisting of the individual data of the daily executions by the billing team

### Conclusion

After creating the code and creating the process, the billing could be calculated again and proceed to accounting and tax closing.
