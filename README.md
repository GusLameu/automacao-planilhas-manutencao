# Excel Sync - PlanilhaSuporteBI

## Configuração
- Origem: `\\172.31.136.227\05. Indicadores\Indicadores 2026\Indicador Setorial\PlanilhaSuporteBI.xlsm`
- Abas: `Pendentes` e `Concluídos`
- Colunas lidas: A:N
- Destino: planilha Google informada, aba `Todas as demandas`
- Cabeçalhos devem ser idênticos e estar na mesma ordem nas duas abas.
- Abas sem registros e linhas totalmente vazias são ignoradas.
- Não é adicionada coluna de origem.
- Para cada linha importada de `Concluídos`, a coluna B recebe o valor original da coluna J, a coluna J é limpa e a coluna N recebe `_copia_`. Essa transformação ocorre apenas na cópia enviada ao Google Sheets; o arquivo Excel original não é alterado.
- Os registros de `Pendentes` são enviados sem essa transformação.

O monitoramento usa consulta periódica a cada 5 segundos, pois eventos de arquivos nem sempre são confiáveis em compartilhamentos SMB/UNC. O programa também faz uma sincronização inicial ao iniciar.

## Instalação (Windows)
```powershell
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Copie `.env.example` para `.env`. Os caminhos e identificadores já estão preenchidos.

## Google Sheets API
1. No Google Cloud Console, crie um projeto.
2. Ative Google Sheets API e Google Drive API.
3. Crie uma conta de serviço e baixe a chave JSON.
4. Salve-a em `credentials/service_account.json`.
5. Compartilhe a planilha Google com o e-mail da conta de serviço como Editor.
6. Confirme que a aba `Todas as demandas` existe.

## Execução
```powershell
python main.py
```

O computador deve permanecer ligado, conectado à rede e com o programa executando. O usuário do Windows precisa ter permissão de leitura no compartilhamento.

## Segurança e limitações
- O programa lê e valida todas as abas antes de limpar o destino. Se houver erro ou nenhum registro, não limpa a aba Google.
- A substituição apaga os valores anteriores da aba de destino e grava os dados atuais. Formatação da aba é preservada; dados e fórmulas manuais no destino serão substituídos.
- Teste inicialmente em uma cópia da planilha Google.
- O arquivo `.xlsm` é aberto somente para leitura; macros não são executadas. Fórmulas são lidas pelo último resultado calculado e salvo no Excel.
- Consulte `logs/sync.log` para verificar execuções.
