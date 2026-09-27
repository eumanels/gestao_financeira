# 💰 Controle Financeiro Pessoal (Streamlit + Excel no OneDrive)

App para controlar salário, dívidas (recorrentes e parceladas) e cofrinho, com **vários
usuários** (uma aba por usuário) no mesmo arquivo `.xlsx`. Toda ação salva
automaticamente na planilha.

## Estrutura do projeto

```
├── app.py                      # ponto de entrada (Streamlit Cloud usa este arquivo)
├── requirements.txt            # dependências de execução
├── requirements-dev.txt        # + pytest
├── get_token.py                # gera o refresh token do OneDrive pessoal (roda 1x)
├── finance/
│   ├── config.py               # lê secrets.toml / variáveis de ambiente
│   ├── storage.py              # leitura/escrita do arquivo: local ou Graph API (msal)
│   ├── repository.py           # Excel <-> DataFrame, abas por usuário, cabeçalhos
│   ├── operations.py           # regras que alteram dados (salário, dívida, pagamento, cofrinho)
│   ├── calculations.py         # saldo do mês, totais, históricos
│   ├── state.py                # cache em st.session_state + salvamento automático
│   └── ui.py                   # componentes de interface
├── tests/                      # pytest (regras, planilha e backend Graph simulado)
└── .streamlit/
    ├── config.toml
    └── secrets.toml.example    # modelo de configuração (copie para secrets.toml)
```

## Estrutura da planilha

Cada **aba = um usuário**. Todas as abas têm os mesmos cabeçalhos (linha 1):

`Registro | ID | Mes_Referencia | Data | Descricao | Categoria | Tipo | Valor | Parcelas_Total | Parcelas_Pagas | Status_Mes | Ativa | ID_Divida`

A coluna **Registro** indica o que é cada linha:

| Registro    | O que representa | Campos principais |
|-------------|------------------|-------------------|
| `SALARIO`   | salário de um mês (1 linha por mês) | Mes_Referencia, Valor |
| `DIVIDA`    | cadastro da dívida | Descricao (nome), Categoria, Tipo (Recorrente/Parcelada), Valor (parcela), Parcelas_Total (vazio = recorrente), Parcelas_Pagas, Status_Mes, Ativa |
| `PAGAMENTO` | uma parcela paga | ID_Divida, Mes_Referencia, Valor |
| `COFRINHO`  | depósito/retirada | Tipo (Depósito/Retirada), Valor, Descricao |

* **Saldo do mês = salário do mês − soma dos `PAGAMENTO` do mês.**
* Planilha vazia ou só com cabeçalhos: o app cria/completa a estrutura sozinho. Se o arquivo
  não existir, ele é criado (com uma aba `_modelo`). Abas que começam com `_` e abas padrão
  vazias (`Planilha1`, `Sheet1`) não aparecem como usuário.
* Dá para filtrar a coluna `Registro` no Excel para ver só dívidas, só pagamentos etc.
* `Status_Mes` é informativo (Pago/Pendente/Quitada/Encerrada no mês corrente, atualizado a cada
  salvamento). O que vale para os cálculos são as linhas `PAGAMENTO`.

## 1. Testar localmente (pasta sincronizada pelo OneDrive)

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows  (Linux/Mac: source .venv/bin/activate)
pip install -r requirements-dev.txt
copy .streamlit\secrets.toml.example .streamlit\secrets.toml   # Linux/Mac: cp
```

Em `.streamlit/secrets.toml` deixe `mode = "local"` e aponte `local_path` para o arquivo
dentro da sua pasta do OneDrive (use `/` no caminho). Depois:

```bash
pytest -q            # opcional: roda os testes
streamlit run app.py
```

> ⚠️ O Excel desktop **bloqueia** o arquivo enquanto ele está aberto. Feche a planilha antes
> de usar o app (o app mostra uma mensagem clara se isso acontecer). Se editar a planilha na
> mão, clique em **🔄 Recarregar planilha** na barra lateral.

## 2. Acessar o OneDrive pela Microsoft Graph API (necessário para o Streamlit Cloud)

No Streamlit Community Cloud o app roda num servidor, que não enxerga o seu computador —
por isso lá é preciso usar `mode = "graph"`.

### OneDrive pessoal (conta Microsoft/Outlook) — `auth = "refresh_token"`

1. Acesse <https://entra.microsoft.com> → **Registros de aplicativo** → **Novo registro**
   (com conta pessoal pode ser preciso criar antes uma conta Azure gratuita).
   * Tipos de conta: **"Contas em qualquer diretório organizacional e contas Microsoft pessoais"**.
2. No app criado → **Autenticação** → ative **"Permitir fluxos de clientes públicos"**.
3. **Permissões de API** → Microsoft Graph → *Permissões delegadas* → `Files.ReadWrite`.
4. Copie o **ID do aplicativo (cliente)** e rode no seu PC:
   ```bash
   python get_token.py SEU_CLIENT_ID
   ```
   Abra o link, digite o código, entre com a conta dona do OneDrive e copie o token impresso.
5. No `secrets.toml`: `mode = "graph"` e preencha `[graph]` (`client_id`, `refresh_token`,
   `file_path` relativo à raiz do OneDrive, ex.: `Financas/controle_financeiro.xlsx`).

> O refresh token pode expirar (por inatividade ou troca de senha). Se o app mostrar
> "Falha na autenticação", rode o `get_token.py` de novo e atualize o Secret.

### OneDrive corporativo (Microsoft 365) — `auth = "client_credentials"`

Registre o app no tenant, crie um **client secret**, adicione a permissão de **aplicativo**
`Files.ReadWrite.All` e peça **consentimento do administrador**. Preencha `tenant_id`,
`client_id`, `client_secret` e `user_id` (e-mail do dono do OneDrive).

## 3. Deploy no Streamlit Community Cloud

1. Suba o projeto para um repositório no GitHub (o `.gitignore` já impede subir
   `secrets.toml` e arquivos `.xlsx`). **Prefira repositório privado.**
2. Em <https://share.streamlit.io> → **Create app** → escolha o repositório, branch e
   `app.py` como *Main file path*.
3. Em **Advanced settings → Secrets**, cole o conteúdo do seu `secrets.toml` com
   `mode = "graph"`.
4. **Segurança:** a URL do app pode ser acessada por qualquer pessoa que a tenha. Defina
   `app_password` nos Secrets (o app pede senha) e/ou restrinja quem pode ver o app nas
   configurações de compartilhamento do Streamlit Cloud.

## Como funciona o salvamento

* A planilha é lida **uma vez por sessão** e fica em `st.session_state`.
* Cada ação altera os dados em memória e grava **só a aba do usuário ativo**, preservando as
  demais abas.
* Antes de gravar, o app compara a versão do arquivo (data de modificação local ou eTag do
  OneDrive). Se outra pessoa alterou o arquivo, ele baixa a versão nova antes de salvar, para
  não apagar as outras abas.
* Se a gravação falhar, a alteração é desfeita na tela e aparece a mensagem de erro — tela e
  planilha nunca ficam diferentes.

Limitações: o openpyxl não preserva gráficos/imagens que você inserir manualmente no
arquivo; o upload simples do Graph aceita arquivos de até 4 MB (bem acima do necessário
para uso pessoal).
