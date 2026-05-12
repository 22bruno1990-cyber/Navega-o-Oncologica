# Navegacao Oncologica MVP

Aplicativo em `Streamlit` para apoiar a navegacao de pacientes oncologicos, com foco em:

- carteira por medico
- previsao do proximo ciclo
- status de prescricao
- status de autorizacao do convenio
- status de agendamento da quimioterapia
- medicacoes de suporte e pendencias

## Como rodar

1. Entre na pasta do projeto:

```bash
cd navegacao_oncologica_mvp
```

2. Instale as dependencias:

```bash
pip install -r requirements.txt
```

3. Rode o app:

```bash
streamlit run app.py
```

## Publicar no Streamlit Community Cloud

1. Suba a pasta do projeto para um repositorio no GitHub.
2. No Streamlit Community Cloud, escolha esse repositorio e selecione o arquivo `app.py`.
3. Em `App settings > Secrets`, configure o login do app:

```toml
[auth]
username = "juliana"
password = "troque-esta-senha"
```

4. Depois que o app abrir na web, use a aba `Planilha principal` para enviar um novo arquivo `.xlsx` e definir a fonte principal do painel.

Observacao:

- no ambiente web, o app nao enxerga o OneDrive do seu Mac
- por isso a fonte da planilha na nuvem deve ser enviada pela interface do app
- essa versao funciona bem para uso pessoal e validacao do fluxo

## O que esta incluido nesta versao

- leitura das abas dos medicos a partir da planilha principal `.xlsx`
- calendario por dia com acesso ao resumo do paciente
- controle do fluxo por ciclo/data: prescricao, autorizacao e agendamento
- campo de antecedencia para avisar quando cobrar novo ciclo ao medico
- aba de alertas para identificar quem corre risco de ficar fora da agenda
- sincronizacao manual da planilha principal
- troca da planilha principal pela interface do app

## Banco de dados

O app cria automaticamente um banco SQLite local em `data/oncologia_cuidado.db`.

## Proximos passos sugeridos

- permitir editar pacientes ja cadastrados
- criar filtros por convenio e por status de autorizacao
- exportar a fila prioritaria para Excel
- disparar alertas diarios por data
- registrar exames, laboratorio e toxicidades pre-ciclo
