# OncoNavega

Produto em `Streamlit` para apoiar profissionais de navegação oncológica na operação dos pacientes em tratamento sistêmico, com possibilidade de evolução posterior para clínicas.

**Promessa:** navegação oncológica para proteger ciclos, agenda e receita.

O foco é antecipar o próximo ciclo, acompanhar prescrição, autorização do convênio e agendamento da quimioterapia antes que o paciente fique fora da agenda.

## Proposta de valor

O produto transforma planilhas e controles manuais em uma fila operacional clara:

- carteira por médico
- calendário de infusões por paciente
- previsão do próximo ciclo
- status de prescrição
- status de autorização do convênio
- status de agendamento da quimioterapia
- alertas de pacientes em risco
- medicamentos de suporte e pendências
- painel comercial com precificação, ROI, apresentação externa e proposta comercial

## Como ganhar dinheiro com o projeto

O caminho mais forte é vender produtividade profissional, redução de retrabalho e previsibilidade da jornada. Em oncologia, um ciclo atrasado pode gerar cobrança manual, ruído entre áreas e perda de controle operacional. O app deve começar como uma ferramenta premium para profissionais de alto nível que precisam organizar carteira, ciclos, prescrição, autorização e agenda.

### Estratégia comercial recomendada

Use um modelo em duas etapas:

```text
profissional premium adota -> produto prova valor -> clínica pode virar etapa futura
```

O profissional de navegação, enfermeira navegadora, médica assistente ou profissional de coordenação sente a dor no dia a dia e pode começar com uma carteira individual. A clínica só vira etapa posterior quando houver necessidade real de multiusuário, dados centralizados, governança, backup, relatórios e padronização do processo.

### Cliente ideal inicial

- profissional que acompanha pacientes oncológicos em uma ou mais clínicas
- enfermeiras navegadoras, médicas assistentes e profissionais de acompanhamento oncológico
- operações que usam Excel, WhatsApp e memória da equipe
- serviços com dor real em autorização de convênio e agenda de infusão
- clínicas que ainda não têm um módulo de navegação bem resolvido no sistema principal, apenas como expansão futura

### Oferta recomendada

O OncoNavega não deve ser vendido como agenda barata. A tese é produtividade profissional premium em navegação oncológica: proteger ciclo, prescrição, autorização, agenda e clareza operacional.

1. Piloto profissional 90 dias: `R$ 850`, com instância separada, implantação leve e limite de uso
2. Plano profissional: `R$ 297` por mês, com instância separada e até 30 casos ativos
3. Plano profissional Pro: `R$ 497` por mês, com até 80 casos ativos, templates, exportações e prioridade de evolução
4. Plano profissional Premium: `R$ 697` por mês, com até 150 casos ativos, acompanhamento mensal e prioridade ampliada
5. Mensalidade clínica futura: a partir de `R$ 1.500` por unidade/mês
6. Clínica com customização e suporte: `R$ 2.500` a `R$ 6.000` por unidade/mês
7. Implantação enterprise: `R$ 15.000+` para múltiplas unidades, integrações e governança

### Modelo de licença MVP

Para novos profissionais, use instância separada:

```text
1 licença profissional = 1 instância + 1 usuário principal + até 30 casos ativos
```

Cada profissional deve ter link próprio, banco próprio, planilha/base própria e usuário/senha próprios no Streamlit Secrets. O limite de uso deve ser por casos ativos, exportações, templates e suporte, não por tempo conectado ao app. Isso evita mistura de pacientes, reduz risco de acesso indevido, simplifica LGPD no MVP e deixa a cobrança por licença mais clara.

Para clínicas, o modelo evolui para plano institucional com multiusuário, governança, suporte, backup e relatórios.

### Como o profissional ajuda a vender para a clínica

O app deve gerar uma conversa simples:

```text
Minha carteira tem X pacientes ativos, Y infusões nos próximos 30 dias e Z pacientes com pendência de prescrição, autorização, agenda ou protocolo. Se a clínica contratar o plano institucional, a equipe passa a acompanhar isso com acesso compartilhado, processo padronizado e relatório gerencial.
```

O profissional não precisa convencer a clínica com discurso abstrato. Ele mostra a fila real, os riscos e o que está sendo perdido por falta de visibilidade.

### Tese de ROI

Uma boa conversa comercial pode partir desta conta:

```text
pacientes ativos x % de ciclos protegidos x receita média por ciclo = receita operacional preservada
```

Exemplo:

```text
120 pacientes ativos x 5% de ciclos protegidos x R$ 4.500 por ciclo = R$ 27.000/mês protegidos
```

Nesse cenário, uma mensalidade de `R$ 3.500` fica ancorada em uma tese de retorno de aproximadamente `7,7x`. O número exato deve ser ajustado com dados reais da clínica.

## Roteiro de venda

1. Profissional testa: organizar uma carteira individual com dados mínimos, autorizados ou anonimizados.
2. Relatório de valor: mostrar pendências, ciclos próximos e riscos que a clínica deveria acompanhar.
3. Convite interno: profissional apresenta o resumo para gestor, médico líder ou faturamento.
4. Piloto profissional 90 dias: profissional usa com instância própria, limite de casos e meta objetiva.
5. Contrato mensal: converter em assinatura individual de R$ 297/mês, Pro de R$ 497/mês ou Premium de R$ 697/mês.
6. Expansão: só depois adicionar clínica, multiusuário, integrações, indicadores financeiros e treinamento recorrente.

## Materiais comerciais gerados pelo app

A aba `Modelo comercial` inclui duas entregas para apoiar venda:

- `Apresentação Externa`: resumo rápido para o profissional apresentar à clínica, com carteira, infusões dos próximos 30 dias e pendências.
- `Gerar Proposta Comercial`: proposta editável com nome da clínica, contato, plano, duração do piloto, implantação, mensalidade, ROI estimado e escopo.
- `Baixar one-page do produto`: página HTML com nome, promessa, oferta e tese de valor do OncoNavega.
- mensagens prontas para WhatsApp e e-mail de abordagem.
- playbook de venda por modo: profissional navegador ou clínica/gestor.
- pacote profissional de 90 dias por R$ 850, com etapas, entregáveis e meta.

Os materiais podem ser baixados em Markdown. A proposta também pode ser baixada como uma página HTML simples para envio ou abertura no navegador.

## Atenção a dados sensíveis

Para uso profissional individual, o ideal é trabalhar com autorização da clínica ou com dados mínimos: iniciais, datas, status e códigos internos. Dados identificáveis, acesso multiusuário, relatórios gerenciais e backup devem migrar para o plano clínica, com responsável institucional e controle de acesso.

## Como rodar

1. Entre na pasta do projeto:

```bash
cd navegacao_oncologica_mvp
```

2. Instale as dependências:

```bash
pip install -r requirements.txt
```

3. Rode o app:

```bash
streamlit run app.py
```

O login deve ser configurado por variável de ambiente ou por `.streamlit/secrets.toml`. Se nada for configurado, o app bloqueia o acesso para evitar credenciais padrão em produção.

## Publicar no Streamlit Community Cloud

1. Suba a pasta do projeto para um repositório no GitHub.
2. No Streamlit Community Cloud, escolha esse repositório e selecione o arquivo `app.py`.
3. Em `App settings > Secrets`, configure o login do app:

```toml
[auth]
username = "juliana"
password = "troque-esta-senha"
```

4. Depois que o app abrir na web, use a aba `Planilha principal` para enviar um novo arquivo `.xlsx` e definir a fonte principal do painel.

Observações:

- no ambiente web, o app não enxerga o OneDrive do Mac
- por isso a fonte da planilha na nuvem deve ser enviada pela interface do app
- evite subir dados reais identificáveis de pacientes para repositórios públicos

## O que está incluído nesta versão

- leitura das abas dos médicos a partir da planilha principal `.xlsx`
- calendário por dia com acesso ao resumo do paciente
- controle do fluxo por ciclo/data: prescrição, autorização e agendamento
- campo de antecedência para avisar quando cobrar novo ciclo ao médico
- aba de alertas para identificar quem corre risco de ficar fora da agenda
- sincronização manual da planilha principal
- troca da planilha principal pela interface do app
- vínculo com planilha online Microsoft por link compartilhado do Excel Online, OneDrive ou SharePoint
- aba `Modelo comercial` com funil profissional primeiro, pacotes premium, ROI, modelo de licença por instância, apresentação externa, proposta comercial e roteiro de venda
- mensagens de abordagem comercial e pacote profissional de 90 dias por R$ 850

## Vincular planilha Microsoft online

Na aba `Planilha principal`, use a seção `Vincular planilha online Microsoft`.

Fluxo recomendado:

1. Abra a planilha no Excel Online, OneDrive ou SharePoint.
2. Clique em `Compartilhar`.
3. Gere um link em que qualquer pessoa com o link possa visualizar, ou um link institucional com acesso liberado para o ambiente onde o app roda.
4. Cole o link no app.
5. Clique em `Vincular e sincronizar planilha online`.

O app baixa uma cópia `.xlsx` para `data/cloud_primary_workbook.xlsx` e usa esse arquivo como fonte principal. A cada sincronização, ele tenta baixar novamente a versão online antes de atualizar o banco local.

Observações:

- o link precisa permitir download do arquivo `.xlsx`
- se o link abrir uma tela de login ou uma página web, o app não conseguirá baixar sem integração Microsoft Graph
- nesse caso, sincronize o arquivo pelo OneDrive no Finder e cole o caminho local em `Usar arquivo sincronizado no Mac`
- para produção com dados reais, o caminho mais robusto é usar Microsoft Graph com autenticação institucional e permissões controladas

## Banco de dados

O app cria automaticamente um banco SQLite local em `data/oncologia_cuidado.db`.

## Próximos passos sugeridos

- exportar a fila prioritária para Excel ou PDF
- disparar alertas diários por e-mail ou WhatsApp
- criar perfis de usuário por função: médica, navegação, faturamento e gestão
- registrar exames, laboratório e toxicidades pré-ciclo
- adicionar indicadores financeiros: ciclos protegidos, atrasos evitados e agenda recuperada
- preparar termo de implantação e contrato mensal simples
- gerar resumo comercial automático para o profissional apresentar à clínica
