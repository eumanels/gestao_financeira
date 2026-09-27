"""Pacote do Controle Financeiro Pessoal.

Módulos:
    config        -> leitura das configurações (st.secrets / variáveis de ambiente)
    storage       -> onde o arquivo .xlsx mora (pasta local sincronizada ou OneDrive via Graph API)
    repository    -> conversão planilha <-> DataFrame (estrutura/abas por usuário)
    operations    -> regras de negócio que ALTERAM os dados (funções puras, testáveis)
    calculations  -> cálculos de saldo/resumos (funções puras, testáveis)
    state         -> cache em st.session_state + salvamento automático
    ui            -> componentes de interface Streamlit
"""
