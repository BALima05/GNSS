# :satellite: Processamento automatizado GNSS

Projeto IC para a coleta e processamento de dados GNSS. Oferecido atráves de projeto PUB pelo professor orientador Edvaldo Simoes da Fonseca Junior (PTR).

## :pencil: Informações iniciais

Projeto para coleta e proccessamento automatizado de dados GNSS de estações presentes no país, com enfoque em dados providos pela estação POLI.

Os dados são coletados pelo serviço RBMC do IBGE, e processados através da ferramenta open-source RTKlib para posterior análise temporal da variação dos dados.

## :clipboard: Índice

- [IBGE-RBMC](#ibge-rbmc)
- [Processamento pelo RTKlib](#processamento-pelo-rtklib)

---
## IBGE-RBMC

### :satellite: Serviço RBMC
A primeira etapa para o processamento dos dados é a obtenção deles, que é feita diretamente através da ferramenta disponibilizada pelo IBGE, o [RBMC - Rede Brasileira de Monitoramento Contínuo dos Sistemas GNSS](https://www.ibge.gov.br/geociencias/informacoes-sobre-posicionamento-geodesico/rede-geodesica/16258-rede-brasileira-de-monitoramento-continuo-dos-sistemas-gnss-rbmc.html?=&t=dados-diarios-e-situacao-operacional), da qual é possível obter os dados de GNSS em qualquer intervalo de tempo para as estações brasileiras. 

A obtenção destes dados ... (em progresso - automatização)

### :file_folder: Conversão Hatanaka -> RINEX

Os arquivos obtidos pelo RBMC são disponibilizados em formato [Hatanaka](https://gnss.be/hatanaka.php), e precisam ser convertidos para [RINEX](https://igs.org/wg/rinex/) de forma a serem processados pelo RTKlib. Dessa forma, o primeiro código ```1IBGE-RBMC.py``` faz essa conversão através da ferramenta [CRX2RNX](https://terras.gsi.go.jp/ja/crx2rnx.html).

Após a conversão, os arquivos são separados em constelação (GPS; GLONASS; GPS e GLONASS) através da ferramenta [TEQC](https://www.unavco.org/software/data-processing/teqc/teqc.html), separando-as em diferentes pastas.

---
## Processamento pelo RTKlib
### :computer: Obtenção das efemérides
Para a realização do PPP - Processamento por Ponto Preciso, é necessária a obtenção das Efemérides Precisas, a qual o RTKlib necessita para realizar o processamento. 

Essas efemérides contém dados de órbita e relógio de satélites de alta qualidade, sendo cruciais na determinação daposição da estação com precisão centimétrica.

Portanto, no código ```2BAIXAR-PRODUTOS.py```, é realizada a obtenção automática das efemérides para os dias dos dados baixados e convertidos do RBMC, fazendo a obtenção de todos os dados necessários para o RTKlib processar os dados enviados. 

Foi optado por se obter os dados através do site do [IGN - Institut National de L'information Géographique et Forestière](https://webigs-rf.ign.fr/), o qual é uma das fontes oficiais de distribuição de informações e produtos do [IGS - International GNSS Service](https://igs.org/), sendo acessado forma fácil pelo código.