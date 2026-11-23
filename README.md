# Excel Sync

**Sincronização automática de dados de um arquivo Excel para o Google Sheets utilizando Python.**

## 1. Sobre o projeto

O **Excel Sync** é uma aplicação desenvolvida em Python para automatizar a transferência de dados de uma planilha Excel armazenada em um diretório de rede para uma planilha do Google Sheets.

O sistema monitora periodicamente o arquivo de origem, identifica alterações, aguarda a estabilização do arquivo e executa a sincronização automaticamente.

Além da transferência dos dados, a aplicação realiza transformações específicas nas informações antes de enviá-las ao destino, incluindo a geração automática de cópias das demandas concluídas e o tratamento de campos de data.

O projeto foi desenvolvido para reduzir atividades manuais, padronizar os dados e manter atualizadas as informações utilizadas nos relatórios do Looker Studio.

## 2. Funcionalidades

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
├── processor.py
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
| `main.py` | Ponto de entrada da aplicação. Inicializa as configurações, o logger e o monitoramento. |
| `config.py` | Carrega as variáveis de ambiente e valida as configurações necessárias. |
| `monitor.py` | Monitora o arquivo Excel e identifica alterações para iniciar uma sincronização. |
| `excel_reader.py` | Realiza a leitura, validação, combinação e transformação dos dados das abas do Excel. |
| `google_sheets.py` | Gerencia a conexão com o Google Sheets e a substituição dos dados na planilha de destino. |
| `processor.py` | Coordena o processo de sincronização entre a leitura do Excel e o envio ao Google Sheets. |
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
EXCEL_FILE=localdaplanilha\nome_planilha.xlsx

SOURCE_SHEETS=Pendentes,Concluídos

GOOGLE_SPREADSHEET_ID=id_generico_gerado_no_drive

GOOGLE_WORKSHEET_NAME=Todas as demandas

GOOGLE_CREDENTIALS=credentials/service_account.json

POLL_INTERVAL_SECONDS=5

DEBOUNCE_SECONDS=3

STABILITY_TIMEOUT_SECONDS=120

STABILITY_INTERVAL_SECONDS=1
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

### Identificando o ID da planilha

O ID é encontrado na URL do Google Sheets:

```text
https://docs.google.com/spreadsheets/d/ID_DA_PLANILHA/edit
```

O valor entre `/d/` e `/edit` corresponde ao `GOOGLE_SPREADSHEET_ID`.

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

A aplicação iniciará o monitoramento do arquivo Excel.

A primeira sincronização é realizada durante a inicialização do monitoramento, conforme o fluxo implementado.

Depois disso, novas sincronizações ocorrerão quando alterações forem identificadas.

Para interromper a aplicação, utilize:

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
- Criar uma interface para acompanhamento do status da sincronização.
- Adicionar testes automatizados para as regras de transformação.
- Implementar validações adicionais de integridade dos dados.
- Avaliar a utilização de sincronização incremental para arquivos maiores.

## 17. Considerações finais

O Excel Sync automatiza o processo de consolidação e disponibilização de informações de manutenção, integrando uma fonte de dados Excel a uma estrutura centralizada no Google Sheets.

A aplicação elimina a necessidade de transferências manuais recorrentes, padroniza os registros e prepara os dados para utilização em relatórios e indicadores desenvolvidos no Looker Studio.

A arquitetura modular permite que novas regras de transformação, validações e integrações sejam adicionadas sem concentrar toda a lógica em um único arquivo.

---

**Excel Sync**  
*Automatização, integração e padronização de dados.*