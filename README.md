# :satellite: Processamento automatizado GNSS

Projeto IC para a coleta e processamento de dados GNSS. Oferecido atráves de projeto PUB pelo professor orientador Edvaldo Simoes da Fonseca Junior (PTR).

## :pencil: Informações iniciais

Projeto para coleta e proccessamento automatizado de dados GNSS de estações presentes no país, com enfoque em dados providos pela estação POLI.

Os dados são coletados pelo serviço RBMC do IBGE, e processados através da ferramenta open-source RTKlib para posterior análise temporal da variação dos dados.

## :clipboard: Índice

- [IBGE-RBMC](#ibge-rbmc)
- [Obtenção das efemérides](#obtenção-das-efemérides)
- [Processamento pelo RTKlib](#processamento-pelo-rtklib)
- [Como utilizar](#como-utilizar)

---
## IBGE-RBMC

### :satellite: Serviço RBMC

A primeira etapa para o processamento dos dados é a obtenção deles, que é feita diretamente através da ferramenta disponibilizada pelo IBGE, o [RBMC - Rede Brasileira de Monitoramento Contínuo dos Sistemas GNSS](https://www.ibge.gov.br/geociencias/informacoes-sobre-posicionamento-geodesico/rede-geodesica/16258-rede-brasileira-de-monitoramento-continuo-dos-sistemas-gnss-rbmc.html?=&t=dados-diarios-e-situacao-operacional), da qual é possível obter os dados de GNSS em qualquer intervalo de tempo para as estações brasileiras. 

A obtenção destes dados ... (em progresso - automação)

Todos os dados são disponibilizados em arquivos .d (dados de observação - comprimidos em Hatanaka), .g (arquivos de navegação do GLONASS), .n (arquivos de navegação do GPS), .l (logs, dados brutos de observação).

Esses dados têm que ser separados conforme suas funções, sendo os principais os dados de observação, os quais contém as informações de geolocalização a serem processados para análise. Os dados de navegação, necessários para o processamento pelo RTKlib, são separados em outra pasta para fácil acesso posteriormente.

### :file_folder: Conversão Hatanaka -> RINEX

Os arquivos obtidos pelo RBMC são disponibilizados em formato [Hatanaka](https://gnss.be/hatanaka.php), e precisam ser convertidos para [RINEX](https://igs.org/wg/rinex/) de forma a serem processados pelo RTKlib. Dessa forma, o primeiro código ```1IBGE-RBMC.py``` faz essa conversão através da ferramenta [CRX2RNX](https://terras.gsi.go.jp/ja/crx2rnx.html).

Após a conversão, os arquivos são separados em constelação (GPS; GLONASS; GPS e GLONASS) através da ferramenta [GFZRNX](https://gnss.gfz.de/services/gfzrnx), separando-as em diferentes pastas. Estas pastas serão posteriormente enviadas para os próximos scripts, de forma a baixar as devidas efemérides e realizar o processamento dos dados pelo RTKlib.

---
## Obtenção das efemérides

### :satellite: Efemérides no PPP
Para a realização do PPP - Processamento por Ponto Preciso, é necessária a obtenção das Efemérides Precisas, a qual o RTKlib necessita para realizar o processamento. 

Essas efemérides contém dados de órbita e relógio de satélites de alta qualidade, sendo cruciais na determinação da posição da estação com precisão centimétrica.

Portanto, no código ```2BAIXAR-PRODUTOS.py```, é realizada a obtenção automática das efemérides para os dias dos dados baixados e convertidos do RBMC, fazendo a obtenção de todos os dados necessários para o RTKlib processar os dados enviados. 

### :computer: Método de obtenção

Foi optado por se obter os dados através do site do [IGN - Institut National de L'information Géographique et Forestière](https://webigs-rf.ign.fr/), o qual é uma das fontes oficiais de distribuição de informações e produtos do [IGS - International GNSS Service](https://igs.org/), sendo acessado forma fácil pelo código.

## Processamento pelo RTKlib

:construction: Em progresso :construction:

## Como utilizar

Primeiro, após clonar o repositório, deve ser criado um arquivo ```config.py``` na pasta principal, de forma a definir os caminhos dos programas a serem utilizados no código.

A maioria dos programas está presente na pasta ```Programas```, de onde os seus caminhos podem ser obtidos diretamente. 

*Obs.: Os programas presentes nessa pasta são executáveis, não sendo necessário sua instalação. Verificar sempre a instalação correta de todas as dependências necessárias. É possível que o binário não seja compatível com sua máquina. Nesse caso, instalar todos os programas necessários separadamente diretamente de suas fontes.*

O arquivo ```config.py``` deve conter a seguinte arquitetura:

```python
import os

# Definição global de caminhos e parâmetros

CRX2RNX_PATH = r"/exemplo_de_caminho_de_linux"
TEQC_PATH = r"C:\exemplo_de_caminho_de_windows"
RNX2RTKP_PATH = r"preencher"
GFZRNX_PATH = r"preencher"

IBGE_ZIP = r"preencher"
PASTA_BASE = r"Caminho para uma pasta base onde o código vai depositar os dados processados"

CONFIG_FILE = r"caminho para o arquivo ppp-static.conf do rtklib /GNSS/ppp-static.conf"
```

Também é necessário instalar o RTKLIB nativamente. Este projeto utiliza o RTKLIB 2.5.1 que pode ser instalado seguindo as instruções em seu [repositório oficial do Github](https://github.com/rtklibexplorer/RTKLIB/releases/tag/v2.5.1).

Após a criação adequada do arquivo de configuração e verificação das dependências, é necessário apenas rodar o ```run_all.py``` pelo seguinte comando:

```bash
python3 run_all.py
```

Também é possível rodar as etapas individualmente com o mesmo comando, trocando o nome do arquivo desejado.

Obs.: Este projeto utiliza **Python 3.14.4** Sua instalação é imperativa para a utilização deste projeto.

Obs².: O processamento de 1 mês inteiro de dados pode ocupar de 7 a 10 gigabytes (gb) antes da rotina de limpeza dependendo da quantidade de observações em cada dia. Tenha a certeza de que há armazenamento suficiente antes de executar a rotina (após a limpeza, cada mês ocupa pouco menos que 1gb).

 ### 1IBGE-RBMC.py

 :construction: Em progresso :construction:

### 2BAIXAR-PRODUTOS.py

:construction: Em progresso :construction:

### 3RTKlib-PPP.py

:construction: Em progresso :construction:

### 4ANALISE-DADOS.py

:construction: Em progresso :construction:

### 4-LIMPEZA.py

Após o término do processamento dos scripts 1 a 4 anteriores, é interessante remover os arquivos intermediários de processamento necessários pelo RTKLIB, uma vez que estes não serão mais necessários. 

Para isso, pode-se rodar o programa **"4-LIMPEZA"**, cuja função é remover todos os arquivos desnecessários após o processamento. O script naturalmente tenta limpar a pasta presente no caminho salvo do atual processamento, no arquivo ```.json```, caso não haja nenhuma flag quando for chamada.

O que é removido pelo programa:
- Pasta ```1 - Dados tipos .d``` e todos os seus conteúdos;
- Pasta ```1.1 - Navegacao Broadcast``` e todos os seus conteudos;
- Pasta ```2 - Dados separados por satélite (Prontos para RTKLIB)``` e todos os seus conteúdos;
- Pasta ```IGS_PRODUCTS``` e todos os seus conteúdos;
- Na pasta ```RESULTADOS_PPP```, são removidos:
  - Arquivos do tipo ```.stat```;
  - Arquivos do tipo ```.trace```.

Dessa forma, são mantidos os arquivos ```.pos``` de posicionamento do dia, gráficos da série mensal e tabelas estatísticas geradas. 

O arquivo também conta com flags com diversas funções para utilização, apresentando os seguintes comportamentos:

- ```--pasta /caminho_da_pasta_mensal``` limpa a pasta do mês especificado no caminho;
- ```--simular``` apenas lista o que seria removido e o espaço liberado, sem remover;
- ```--forcar``` ignora as travas de segurança descritas abaixo.
  
**Travas de segurança** estão presentes no arquivo, para evitar remover arquivos errados ou em momentos inadequados (caso algum script anterior falhe).
- Só atua dentro da pasta do mês atual, com nome no formato ```MMM_YY```;
- Só apaga nomes exatos da lista, sem varredura recursiva na pasta ano, então os outros meses não são afetados.
- Aborta se não houver ```.pos``` e ```Dados_Serie_*.csv``` em ```RESULTADOS_PPP```. Isso evita apagar os insumos quando o script 3 ou o 4 falhou.

*Obs.: Embora o programa naturalmente aborte se for executado antes do script 4, este pode ser executado a qualquer momento após a execução do script 3 sem interferir no funcionamento normal dos scripts 4 e 5.*

*Obs².: Caso deseje remover os arquivos de observação ```.pos``` após a execução das rotinas de análise, pode-se alterar a ordem de chamada da execução desse script e apenas adicionar ```".pos"``` entre os arquivos para remoção na constante ```SUFIXOS_RESULTADOS``` ao início do arquivo:*

```python
SUFIXOS_RESULTADOS = (".stat", ".trace", ".pos")
```

### 5ANALISE-ANUAL.py

:construction: Em progresso :construction: