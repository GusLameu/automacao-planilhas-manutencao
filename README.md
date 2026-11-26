# Excel Sync v11

**Sincronização automática de dados de um arquivo Excel para o Google Sheets utilizando Python.**

## 1. Sobre o projeto

O **Excel Sync** é uma aplicação desenvolvida em Python para automatizar a transferência de dados de uma planilha Excel armazenada em um diretório de rede para uma planilha do Google Sheets.

O sistema monitora periodicamente o arquivo de origem, identifica alterações, aguarda a estabilização do arquivo e executa a sincronização automaticamente.

Além da transferência dos dados, a aplicação realiza transformações específicas nas informações antes de enviá-las ao destino, incluindo a geração automática de cópias das demandas concluídas e o tratamento de campos de data.

O projeto foi desenvolvido para reduzir atividades manuais, padronizar os dados e manter atualizadas as informações utilizadas nos relatórios do Looker Studio.

## 2. Funcionalidades

- Menu inicial com dois fluxos: **Planilha de Manutenção** e **Planilha Medidores**.
- **Planilha de Manutenção:** exportação da planilha Excel local para o Google Sheets.
- **Planilha Medidores:** importação de 3 planilhas Google Sheets para uma única planilha Excel local.
- Monitoramento periódico de um arquivo Excel armazenado em rede.
- Identificação de alterações no arquivo de origem.
- Espera pela estabilização do arquivo antes da leitura.
- Leitura simultânea de múltiplas abas.
- Validação dos cabeçalhos e da estrutura das planilhas.
- Combinação dos dados de diferentes abas em uma única tabela.
- Ignorar automaticamente abas vazias e linhas sem dados.
- Geração automática de cópias das demandas concluídas.
- Transformação de campos específicos das cópias.
- Normalização de datas para interpretação correta pelo Google Sheets.
- Substituição integral dos dados existentes na aba de destino.
- Preservação da formatação visual da planilha de destino, quando não sobrescrita por outras operações.
- Registro das operações e ocorrências em arquivos de log.

### 2.1. Menu inicial

Ao executar `python main.py`, o programa abre uma janela com duas opções:

| Opção | Fluxo |
|---|---|
| **Planilha de Manutenção** | Exporta a planilha Excel local para o Google Sheets. |
| **Planilha Medidores** | Importa 3 planilhas Google Sheets diferentes para uma única planilha Excel local. |

Cada opção abre seu próprio painel, com botões de execução manual, monitoramento contínuo, parada, status e histórico de logs. Os dois fluxos podem ser monitorados ao mesmo tempo, pois usam serviços e buffers de log independentes.

### 2.2. Planilha de Manutenção

Mantém o comportamento original do projeto: lê as abas configuradas em `SOURCE_SHEETS`, combina os registros, gera as cópias da aba `Concluídos` e substitui o conteúdo da aba de destino no Google Sheets.

### 2.3. Planilha Medidores

Lê as 3 origens configuradas em `MEDIDORES_SOURCE_1_ID`, `MEDIDORES_SOURCE_2_ID` e `MEDIDORES_SOURCE_3_ID`, valida se todas possuem os mesmos cabeçalhos, empilha os registros em uma única tabela e substitui o conteúdo da aba de destino no arquivo Excel local indicado por `MEDIDORES_EXCEL_FILE`.

As origens podem ser informadas pelo ID ou pela URL completa da planilha. Números e datas são preservados como números e datas no Excel. Se o arquivo de destino não existir, ele é criado automaticamente. As demais abas do arquivo são preservadas.

## 3. Tecnologias utilizadas

| Tecnologia | Finalidade |
|---|---|
| Python | Linguagem principal |
| OpenPyXL | Leitura de arquivos Excel |
| gspread | Integração com Google Sheets |
| google-auth | Autenticação com a API do Google |
| python-dotenv | Gerenciamento das variáveis de ambiente |
| Google Sheets API | Atualização da planilha de destino |

### Bibliotecas Python

As dependências do projeto estão especificadas no arquivo `requirements.txt`.

```text
openpyxl>=3.1
gspread>=6.0
google-auth>=2.0
python-dotenv>=1.0
```

## 4. Estrutura do projeto

A organização dos arquivos segue a seguinte estrutura:

```text
excel_sync/
│
├── main.py
├── config.py
├── monitor.py
├── excel_reader.py
├── google_sheets.py
├── medidores.py
├── processor.py
├── sync_panel.py
├── logger.py
│
├── .env
├── .env.example
├── .gitignore
├── requirements.txt
├── README.md
│
├── credentials/
│   └── service_account.json
│
├── logs/
│   └── excel_sync.log
│
└── .venv/
```

### Responsabilidade dos módulos

| Arquivo | Responsabilidade |
|---|---|
| `main.py` | Ponto de entrada da aplicação. Abre o painel com o menu de fluxos ou executa em modo console. |
| `config.py` | Carrega as variáveis de ambiente e valida as configurações dos dois fluxos. |
| `monitor.py` | Monitora o arquivo Excel e identifica alterações para iniciar uma sincronização. |
| `excel_reader.py` | Realiza a leitura, validação, combinação e transformação dos dados das abas do Excel. |
| `google_sheets.py` | Gerencia a conexão com o Google Sheets e a substituição dos dados na planilha de destino. |
| `medidores.py` | Lê as 3 origens do Google Sheets, valida cabeçalhos, empilha os registros e grava a planilha Excel local. |
| `processor.py` | Coordena o processo de sincronização entre a leitura do Excel e o envio ao Google Sheets. |
| `sync_panel.py` | Painel local com o menu Planilha de Manutenção / Planilha Medidores, status e logs. |
| `logger.py` | Configura os registros de execução e os arquivos de log. |
| `.env` | Armazena as configurações específicas do ambiente. |
| `.env.example` | Apresenta um modelo das variáveis de configuração. |
| `requirements.txt` | Lista as dependências necessárias para executar o projeto. |
| `.gitignore` | Define arquivos e diretórios que não devem ser versionados. |

## 5. Requisitos

Para executar o projeto, são necessários:

- Windows com acesso ao arquivo Excel de origem.
- Python instalado.
- Acesso à internet para comunicação com os serviços do Google.
- Acesso de leitura ao arquivo Excel armazenado na rede.
- Uma planilha do Google Sheets.
- Um projeto no Google Cloud com a API do Google Sheets habilitada.
- Uma conta de serviço com permissão de edição na planilha de destino.

## 6. Configuração do ambiente

### 6.1. Clonar ou obter o projeto

Acesse o diretório em que deseja armazenar o projeto.

Caso esteja utilizando Git:

```bash
git clone URL_DO_REPOSITORIO
```

Entre na pasta:

```bash
cd excel_sync
```

### 6.2. Criar o ambiente virtual

Crie um ambiente virtual para isolar as dependências:

```bash
python -m venv .venv
```

Ative o ambiente no Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Caso esteja utilizando o Prompt de Comando:

```cmd
.venv\Scripts\activate.bat
```

### 6.3. Instalar as dependências

Com o ambiente virtual ativado:

```bash
python -m pip install -r requirements.txt
```

## 7. Configuração do Google Sheets

A aplicação utiliza uma conta de serviço para realizar a autenticação e atualizar a planilha de destino.

### 7.1. Criar um projeto no Google Cloud

1. Acesse o [Google Cloud Console](https://console.cloud.google.com/).
2. Crie um projeto ou selecione um projeto existente.
3. Acesse a biblioteca de APIs.
4. Habilite a **Google Sheets API**.

### 7.2. Criar uma conta de serviço

1. Acesse a área de contas de serviço do Google Cloud.
2. Crie uma conta de serviço.
3. Gere uma chave no formato JSON.
4. Baixe o arquivo de credenciais.

### 7.3. Armazenar as credenciais

Coloque o arquivo JSON na pasta `credentials`:

```text
credentials/
└── service_account.json
```

**Importante:** o arquivo de credenciais contém informações privadas. Ele não deve ser enviado para repositórios públicos ou compartilhado indevidamente.

### 7.4. Compartilhar a planilha com a conta de serviço

Abra o arquivo JSON e localize o campo:

```json
"client_email": "nome-da-conta@projeto.iam.gserviceaccount.com"
```

Copie o endereço de e-mail e compartilhe a planilha do Google Sheets com essa conta, concedendo permissão de **Editor**.

Sem essa permissão, a aplicação não conseguirá atualizar os dados.

## 8. Variáveis de ambiente

As configurações são armazenadas no arquivo `.env`, localizado na raiz do projeto.

Crie o arquivo a partir do modelo `.env.example` ou configure manualmente:

```env
# Planilha de Manutenção (Excel local -> Google Sheets)
EXCEL_FILE=localdaplanilha\nome_planilha.xlsx

SOURCE_SHEETS=Pendentes,Concluídos

GOOGLE_SPREADSHEET_ID=id_generico_gerado_no_drive

GOOGLE_WORKSHEET_NAME=Todas as demandas

GOOGLE_CREDENTIALS=credentials/service_account.json

POLL_INTERVAL_SECONDS=5

DEBOUNCE_SECONDS=3

STABILITY_TIMEOUT_SECONDS=120

STABILITY_INTERVAL_SECONDS=1

# Planilha Medidores (3 Google Sheets -> Excel local)
MEDIDORES_EXCEL_FILE=planilha_medidores.xlsx

MEDIDORES_WORKSHEET_NAME=Medidores

MEDIDORES_IMPORT_INTERVAL_SECONDS=60

MEDIDORES_SOURCE_1_LABEL=Medidores A
MEDIDORES_SOURCE_1_ID=id_ou_url_da_planilha_1
MEDIDORES_SOURCE_1_SHEET=

MEDIDORES_SOURCE_2_LABEL=Medidores B
MEDIDORES_SOURCE_2_ID=id_ou_url_da_planilha_2
MEDIDORES_SOURCE_2_SHEET=

MEDIDORES_SOURCE_3_LABEL=Medidores C
MEDIDORES_SOURCE_3_ID=id_ou_url_da_planilha_3
MEDIDORES_SOURCE_3_SHEET=
```

### Descrição das variáveis

| Variável | Descrição |
|---|---|
| `EXCEL_FILE` | Caminho completo do arquivo Excel de origem. |
| `SOURCE_SHEETS` | Abas que serão utilizadas na leitura. |
| `GOOGLE_SPREADSHEET_ID` | Identificador da planilha do Google Sheets. |
| `GOOGLE_WORKSHEET_NAME` | Nome da aba que receberá os dados. |
| `GOOGLE_CREDENTIALS` | Caminho do arquivo JSON de autenticação. |
| `POLL_INTERVAL_SECONDS` | Intervalo entre as verificações do arquivo. |
| `DEBOUNCE_SECONDS` | Tempo de espera para evitar múltiplas sincronizações decorrentes de alterações consecutivas. |
| `STABILITY_TIMEOUT_SECONDS` | Tempo máximo de espera pela estabilização do arquivo. |
| `STABILITY_INTERVAL_SECONDS` | Intervalo entre as verificações de estabilidade. |
| `MEDIDORES_EXCEL_FILE` | Arquivo Excel local que receberá os dados dos medidores. Se não existir, é criado. |
| `MEDIDORES_WORKSHEET_NAME` | Aba do Excel local que será substituída. Padrão: `Medidores`. |
| `MEDIDORES_IMPORT_INTERVAL_SECONDS` | Intervalo entre as importações automáticas das 3 origens. |
| `MEDIDORES_SOURCE_n_ID` | ID ou URL completa do Google Sheets da origem `n` (1 a 3). |
| `MEDIDORES_SOURCE_n_SHEET` | Aba da origem `n`. Se vazio, usa a primeira aba da planilha. |
| `MEDIDORES_SOURCE_n_LABEL` | Nome amigável da origem `n`, usado nos logs. |

Após alterar o `.env`, reinicie o programa para que as novas configurações sejam carregadas.

### Identificando o ID da planilha

O ID é encontrado na URL do Google Sheets:

```text
https://docs.google.com/spreadsheets/d/ID_DA_PLANILHA/edit
```

O valor entre `/d/` e `/edit` corresponde ao `GOOGLE_SPREADSHEET_ID`. Para as origens de medidores, o campo aceita tanto o ID puro quanto a URL completa.

**Atenção:** o caminho do arquivo Excel precisa estar acessível pelo usuário que executa a aplicação. Caso seja utilizado um diretório de rede mapeado, como `Z:`, a unidade deverá estar disponível durante a execução.

## 9. Funcionamento da sincronização

O fluxo de execução pode ser dividido em seis etapas.

### Etapa 1: Inicialização

A aplicação carrega as configurações do `.env`, valida os parâmetros necessários e inicializa o sistema de logs.

### Etapa 2: Monitoramento

O módulo `monitor.py` verifica periodicamente se o arquivo Excel sofreu alterações.

O intervalo é definido pela variável:

```env
POLL_INTERVAL_SECONDS=5
```

### Etapa 3: Estabilização do arquivo

Quando uma alteração é identificada, o sistema aguarda o arquivo permanecer estável antes de iniciar a leitura.

Esse procedimento evita que a aplicação tente acessar o Excel enquanto ele ainda está sendo salvo ou modificado.

As configurações são:

```env
DEBOUNCE_SECONDS=3
STABILITY_TIMEOUT_SECONDS=120
STABILITY_INTERVAL_SECONDS=1
```

### Etapa 4: Leitura e validação

O módulo `excel_reader.py` abre o arquivo utilizando OpenPyXL e realiza a leitura das abas configuradas.

Atualmente, são utilizadas:

- `Pendentes`
- `Concluídos`

A leitura considera as colunas de A até N.

O sistema verifica:

- Existência das abas configuradas.
- Presença dos cabeçalhos.
- Ausência de cabeçalhos duplicados.
- Compatibilidade dos cabeçalhos entre as abas.
- Existência de registros válidos.

As abas vazias e as linhas completamente vazias são ignoradas.

### Etapa 5: Transformação dos dados

Após a leitura, os dados são combinados em uma única lista.

A aplicação também executa regras específicas para os registros da aba `Concluídos`, descritas na próxima seção.

### Etapa 6: Atualização do Google Sheets

Depois que todas as abas são lidas e validadas, o módulo `google_sheets.py` realiza a conexão com o Google Sheets.

O sistema limpa os valores existentes na aba de destino e grava os novos dados a partir da célula `A1`.

A operação utiliza:

```python
worksheet.clear()
```

E:

```python
worksheet.update(
    range_name="A1",
    values=values,
    value_input_option="USER_ENTERED"
)
```

A opção `USER_ENTERED` permite que o Google Sheets interprete valores enviados, incluindo datas, conforme suas regras de entrada.

A aplicação também aplica a formatação de data às colunas B e J:

```python
worksheet.format(
    ["B:B", "J:J"],
    {
        "numberFormat": {
            "type": "DATE",
            "pattern": "dd/MM/yyyy"
        }
    }
)
```

A substituição é integral: os registros anteriores são removidos e substituídos pelos dados da execução atual.

### Fluxo da Planilha Medidores

O fluxo de importação dos medidores é independente do fluxo de manutenção:

1. O módulo `medidores.py` autentica a conta de serviço e abre as 3 planilhas configuradas.
2. Cada origem é lida com valores formatados e não formatados, permitindo preservar números e datas.
3. A primeira linha de cada origem é tratada como cabeçalho.
4. Os cabeçalhos das 3 origens precisam ser iguais, na mesma ordem, sem colunas vazias e sem duplicidades.
5. Os registros são empilhados em uma única tabela.
6. O conteúdo anterior da aba `MEDIDORES_WORKSHEET_NAME` é substituído no arquivo `MEDIDORES_EXCEL_FILE`.
7. As demais abas do arquivo Excel são preservadas.

Antes de gravar, a aplicação calcula uma assinatura dos dados. No modo automático, se nada mudou desde a última importação, o arquivo local não é reescrito. Isso evita conflitos com o Excel aberto e gravações desnecessárias.

Se o arquivo de destino estiver aberto no Excel no momento da gravação, a importação falha com uma mensagem solicitando o fechamento do arquivo. A gravação é feita primeiro em um arquivo temporário, reduzindo o risco de corromper o destino.

## 10. Regras de transformação dos dados

Esta seção descreve uma das principais particularidades do projeto.

### 10.1. Aba Pendentes

Os registros da aba `Pendentes` são enviados sem a geração de cópias e sem as transformações específicas aplicadas aos registros concluídos.

### 10.2. Aba Concluídos

Para cada registro encontrado na aba `Concluídos`, o sistema gera duas linhas no destino:

1. Registro original.
2. Registro duplicado, identificado como cópia.

A cópia é criada em memória durante a leitura, sem modificar o arquivo Excel original.

### 10.3. Transformações das cópias

As cópias recebem as seguintes alterações:

| Coluna | Ação |
|---|---|
| B | Recebe o valor originalmente presente na coluna J. |
| J | Tem seu conteúdo removido. |
| N | Recebe o identificador `_copia_`. |

As demais colunas permanecem iguais às do registro original.

### Exemplo ilustrativo

**Registro original:**

| B | J | N |
|---|---|---|
| 10/11/2026 | 23/11/2026 | CONCLUIDO |

**Registro gerado automaticamente:**

| B | J | N |
|---|---|---|
| 23/11/2026 | *(vazio)* | `_copia_` |

O registro original permanece inalterado.

### 10.4. Resultado esperado

Considerando:

- 30 registros na aba `Pendentes`.
- 122 registros na aba `Concluídos`.

O sistema produzirá:

| Categoria | Quantidade |
|---|---:|
| Pendentes | 30 |
| Concluídos originais | 122 |
| Cópias dos concluídos | 122 |
| **Total de registros** | **274** |
| Cabeçalho | 1 |
| **Total de linhas no destino** | **275** |

Esses valores são um exemplo baseado em uma execução do projeto. A quantidade real varia conforme os dados existentes no Excel.

## 11. Tratamento das datas

O tratamento das datas é importante para garantir a compatibilidade com o Google Sheets e o Looker Studio.

As colunas B e J devem ser reconhecidas como datas, e não como textos.

### 11.1. Normalização no Python

Para as colunas B e J, os valores são convertidos para o formato ISO:

```python
value.strftime("%Y-%m-%d")
```

Exemplo:

```text
23/11/2026 15:30:00
```

É convertido para:

```text
2026-11-23
```

### 11.2. Interpretação no Google Sheets

A atualização utiliza:

```python
value_input_option="USER_ENTERED"
```

Com isso, o Google Sheets interpreta a data enviada como um valor de data, em vez de armazená-la simplesmente como texto.

### 11.3. Formatação visual

As colunas B e J recebem o formato:

```text
dd/MM/yyyy
```

O resultado visual será:

```text
23/11/2026
```

Essa abordagem permite que os campos sejam utilizados em filtros, agrupamentos por mês e ano e gráficos temporais no Looker Studio.

### 11.4. Configuração no Looker Studio

Após a primeira sincronização:

1. Abra a fonte de dados do relatório.
2. Atualize os campos da fonte de dados.
3. Localize as colunas B e J.
4. Verifique se ambas estão identificadas com o tipo `Data`.
5. Configure os gráficos temporais utilizando o campo de data apropriado.

O reconhecimento correto das datas permite utilizar dimensões como ano, mês e ano/mês.

## 12. Execução da aplicação

Com o ambiente virtual ativado e as configurações preenchidas, execute:

```bash
python main.py
```

O painel abre no menu inicial. Escolha **Planilha de Manutenção** ou **Planilha Medidores**.

Em cada painel:

| Botão | Planilha de Manutenção | Planilha Medidores |
|---|---|---|
| Executar | Sincronizar agora | Importar agora |
| Monitorar | Manter sincronia | Importar automaticamente |
| Parar | Encerra o monitoramento | Encerra o monitoramento |

Para executar sem a janela:

```bash
# Planilha de Manutenção, monitorando alterações no Excel
python main.py --console

# Planilha Medidores, importando as 3 origens periodicamente
python main.py --console --medidores
```

Para interromper a execução em modo console, utilize:

```text
Ctrl + C
```

## 13. Logs e acompanhamento

O sistema registra informações sobre as operações realizadas.

Exemplo de log:

```text
2026-11-23 16:08:44 | INFO | Aba 'Pendentes': 30 registro(s) lido(s) em A:N.

2026-11-23 16:08:46 | INFO | Aba 'Concluídos': 122 registro(s) lido(s) em A:N.

2026-11-23 16:08:46 | INFO | Abas configuradas: ['Pendentes', 'Concluídos']

2026-11-23 16:08:46 | INFO | Abas efetivamente utilizadas: ['Pendentes', 'Concluídos']

2026-11-23 16:08:46 | INFO | Abas ignoradas: []

2026-11-23 16:08:50 | INFO | Destino 'Todas as demandas' atualizado com 275 linhas e 14 colunas.

2026-11-23 16:08:50 | INFO | Concluído: 274 registros enviados.
```

Os registros permitem acompanhar:

- Inicialização da aplicação.
- Arquivo monitorado.
- Leitura das abas.
- Quantidade de registros encontrados.
- Abas utilizadas e ignoradas.
- Atualização da planilha de destino.
- Erros ocorridos durante a execução.

O arquivo de log é armazenado no diretório `logs/`, conforme a configuração do módulo `logger.py`.

## 14. Tratamento de erros e limitações

### 14.1. Arquivo Excel indisponível

Se o arquivo de origem não estiver acessível, a aplicação poderá falhar durante a leitura.

Verifique:

- Conexão com a rede corporativa.
- Disponibilidade do servidor de arquivos.
- Permissão de leitura.
- Existência da unidade de rede mapeada.

### 14.2. Abas inexistentes

Se uma aba configurada em `SOURCE_SHEETS` não existir no arquivo Excel, a validação deverá interromper o processamento.

Confira os nomes no `.env`.

### 14.3. Cabeçalhos incompatíveis

As abas precisam possuir cabeçalhos preenchidos e iguais, tanto nos nomes quanto na ordem das colunas A:N.

### 14.4. Falha de autenticação

Se ocorrer um erro de autenticação, verifique:

- Existência do arquivo JSON.
- Caminho configurado em `GOOGLE_CREDENTIALS`.
- Validade das credenciais.
- Permissão de Editor da conta de serviço na planilha.

### 14.5. Problemas com datas

Se o Looker Studio apresentar valores `null`:

1. Verifique se as colunas B e J estão sendo enviadas no formato ISO.
2. Confira se a atualização utiliza `USER_ENTERED`.
3. Verifique a formatação das colunas no Google Sheets.
4. Atualize os campos da fonte de dados no Looker Studio.
5. Confirme se os campos estão configurados como `Data`.

### 14.6. Aviso do OpenPyXL

Durante a leitura de arquivos `.xlsm`, pode aparecer o aviso:

```text
UserWarning: Data Validation extension is not supported and will be removed
```

Esse aviso está relacionado às extensões de validação de dados do Excel não suportadas pelo OpenPyXL.

A aplicação realiza apenas a leitura do arquivo. Portanto, esse aviso não significa que a planilha original será modificada.

### 14.7. Fórmulas no Excel

A leitura utiliza `data_only=True`.

Isso significa que o OpenPyXL utiliza os últimos resultados de fórmulas salvos pelo Excel, em vez de recalcular as fórmulas.

Se os valores calculados não estiverem atualizados no arquivo, os dados enviados também poderão estar desatualizados.

### 14.8. Disponibilidade do computador

A aplicação precisa permanecer em execução para detectar alterações e realizar novas sincronizações.

Se o computador for desligado, perder acesso à rede ou encerrar o processo Python, o monitoramento será interrompido.

### 14.9. Segurança e substituição dos dados

A aplicação utiliza a operação `clear()` antes de gravar os novos valores.

Por isso, uma falha de conexão ocorrida entre a limpeza e a gravação poderá deixar a aba de destino sem os dados anteriores.

É recomendável manter uma cópia de segurança da planilha de destino e considerar mecanismos adicionais de recuperação caso o projeto passe a ser utilizado em produção.

### 14.10. Planilha de destino dos medidores aberta no Excel

Se o arquivo `MEDIDORES_EXCEL_FILE` estiver aberto no Excel no momento da importação, o Windows pode impedir a substituição do arquivo. A aplicação exibe uma mensagem solicitando o fechamento do arquivo. Feche-o e tente novamente.

### 14.11. Cabeçalhos diferentes nas origens de medidores

As 3 origens precisam ter os mesmos cabeçalhos, na mesma ordem, sem colunas vazias e sem nomes duplicados. Quando houver diferença, a importação é interrompida antes de alterar o arquivo local, e o log mostra os cabeçalhos esperados e os encontrados.

### 14.12. Origens de medidores não configuradas

Se `MEDIDORES_SOURCE_1_ID`, `MEDIDORES_SOURCE_2_ID` ou `MEDIDORES_SOURCE_3_ID` estiverem vazios, o painel mostra **Configuração pendente** e não permite iniciar a importação. Preencha os IDs ou URLs no `.env` e reinicie o programa.

## 15. Segurança

Algumas práticas importantes:

- Nunca versionar o arquivo `service_account.json`.
- Nunca publicar o arquivo `.env` com informações privadas.
- Não compartilhar credenciais da conta de serviço.
- Utilizar permissões mínimas necessárias.
- Manter o acesso ao arquivo Excel restrito aos usuários autorizados.

O `.gitignore` deve conter, no mínimo:

```gitignore
.env
.venv/
__pycache__/
*.py[cod]

credentials/
logs/
```

O arquivo `.env.example` pode ser versionado desde que contenha apenas valores fictícios ou campos sem informações privadas.

## 16. Possíveis melhorias futuras

Algumas melhorias que podem ser consideradas para evoluir o projeto:

- Implementar tentativas automáticas de reconexão em caso de falha.
- Criar um mecanismo de backup antes da substituição dos dados.
- Implementar atualização transacional ou estratégia de recuperação para evitar que o destino fique vazio em caso de falha.
- Adicionar notificações de erro e sucesso.
- Implementar execução automática do programa junto à inicialização do Windows.
- Adicionar uma tela de configuração das origens dentro do painel, sem editar o `.env`.
- Adicionar testes automatizados para as regras de transformação.
- Implementar validações adicionais de integridade dos dados.
- Avaliar a utilização de sincronização incremental para arquivos maiores.

## 17. Considerações finais

O Excel Sync automatiza o processo de consolidação e disponibilização de informações de manutenção, integrando uma fonte de dados Excel a uma estrutura centralizada no Google Sheets. O mesmo painel também concentra o fluxo de medidores, importando 3 planilhas Google Sheets para uma única planilha Excel local.

A aplicação elimina a necessidade de transferências manuais recorrentes, padroniza os registros e prepara os dados para utilização em relatórios e indicadores desenvolvidos no Looker Studio.

A arquitetura modular permite que novas regras de transformação, validações e integrações sejam adicionadas sem concentrar toda a lógica em um único arquivo.

---

**Excel Sync**  
*Automatização, integração e padronização de dados.*