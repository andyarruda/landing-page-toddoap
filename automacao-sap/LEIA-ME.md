# Automação diária SAP — MM x EWM e Perfis de acesso x GPA

## Arquivos
| Arquivo | Para que serve |
|---|---|
| `coletor_sap.py` | Robô SAP GUI Scripting: entra nas transações, executa e exporta |
| `analise_mm_ewm.py` | Cruza MM x EWM x filas e gera o relatório |
| `roteiros_sap.json` | Passos de cada transação (criado na 1ª execução) |
| `coleta_automatica.bat` | Dois cliques: coleta + análise + abre o relatório |
| `rodar_analise.bat` | Modo manual (arquivos exportados por você em `entrada\`) |
| `analise_perfis.py` | Cruza perfis SAP x catálogo GPA e gera o relatório estilo Fiori |
| `modelo_relatorio_perfis.html` | Layout do relatório de perfis (não apague) |
| `coleta_perfis.bat` | Dois cliques: coleta de perfis no SAP + análise x GPA |
| `rodar_analise_perfis.bat` | Modo manual da análise de perfis |
| `config_perfis.exemplo.json` | Modelo de ajustes (colunas do GPA, perfis críticos, prazos) |

## Instalação (uma vez)
1. `pip install pywin32 pandas openpyxl`
2. **Scripting habilitado** (confirme com o Basis/Segurança antes):
   - Servidor: parâmetro `sapgui/user_scripting = TRUE` (RZ11).
   - Seu PC: SAP Logon → Opções → Acessibilidade e scripts → Scripts →
     marque *Ativar scripts* e desmarque *Notificar quando um script se conectar*.
3. Teste a conexão: `python coletor_sap.py listar`

## Ensinar as transações ao robô (gravar uma vez cada)
1. No SAP: **Alt+F12 → Gravação e reprodução de scripts → Gravar**.
2. Faça o caminho normal, de preferência usando **variante** e **layout** salvos:
   - **mm**: MB52 → variante → executar → exportar *Arquivo local → Texto com tabulações* → salvar.
   - **ewm**: /SCWM/MON (estoque disponível) ou o relatório que vocês usam → exportar da mesma forma.
   - **filas**: SMQ2 (no EWM e/ou SMQ1 no ERP) → exportar a lista.
3. Pare a gravação e importe:
   ```
   python coletor_sap.py importar gravacao_mb52.vbs --nome mm
   python coletor_sap.py importar gravacao_ewm.vbs  --nome ewm
   python coletor_sap.py importar gravacao_smq2.vbs --nome filas
   ```
   O caminho e o nome do arquivo gravados viram `{PASTA_ENTRADA}` e `{ARQUIVO}` automaticamente,
   e a data de hoje vira `{DATA_SAP}`.
4. Em `roteiros_sap.json`, ajuste `"sistema"` e `"mandante"` (ex.: PRD / 300).
   Valores fixos podem virar variáveis: troque `"1000"` por `"{CENTRO}"` e defina em `"variaveis"`.

## Uso diário
Dois cliques em `coleta_automatica.bat` (com o SAP logado). O robô abre uma sessão nova
(não atrapalha a sua tela), roda tudo, fecha a sessão e abre o relatório.

Opções: `python coletor_sap.py coletar --somente mm,ewm` · `--sem-analise` · `--nao-abrir`

## Agendar (opcional)
Agendador de Tarefas do Windows → ação `coleta_automatica.bat`, horário ex. 07:30,
opção **"Executar somente quando o usuário estiver conectado"** (o SAP GUI precisa de tela).

## Quando quebrar
- Mensagem "Passo N falhou em wnd[...]": a tela mudou (patch, layout, popup novo). Grave e reimporte.
- Popup que só aparece às vezes: no JSON, adicione `"opcional": true` ao passo.
- Detalhes de cada execução ficam em `logs\`.

## Etapa 2 — Perfis de acesso x GPA
Responde, para cada usuário x perfil no SAP: **o perfil é solicitável via GPA?** E mais:
existe solicitação GPA aprovada? a atribuição está vencendo/vencida? o perfil é crítico?
o usuário está bloqueado ou sem logon?

Teste sem SAP: `python analise_perfis.py --demo` → `saida\<data>\perfis\relatorio_perfis.html`

### Entradas (pasta `entrada\`)
| Arquivo | De onde vem | Obrigatório |
|---|---|---|
| `perfis.txt/.xlsx` | SE16N tabela **AGR_USERS** (ou SUIM → *Usuários por função*). Robô: roteiro `perfis` | Sim |
| `catalogo_gpa.xlsx` | **Exportado do GPA**: lista de perfis com *Solicitável*, *Status*, *Descrição*, *Aprovador* | Sim |
| `usuarios.txt/.xlsx` | SE16N tabela **USR02** (BNAME, USTYP, UFLAG, TRDAT, GLTGB). Robô: roteiro `usuarios` | Não |
| `solicitacoes_gpa.xlsx` | **Exportado do GPA**: usuário x perfil x nº x status (x validade) | Não |

O GPA é sistema web, então o robô do SAP GUI **não** entra nele: exporte o catálogo (e as solicitações,
se possível) e salve em `entrada\` com esses nomes. Os nomes de coluna são reconhecidos automaticamente
(sem acento/maiúsculas); se o GPA usar outro nome, ajuste em `config_perfis.json`
(copie de `config_perfis.exemplo.json`).

### Gravar as transações (uma vez)
```
python coletor_sap.py importar gravacao_agr_users.vbs --nome perfis
python coletor_sap.py importar gravacao_usr02.vbs    --nome usuarios
```
Na SE16N, deixe o *Número máx. de ocorrências* em branco antes de executar (senão corta em 500 linhas)
e exporte por *Exportar → Planilha / Arquivo local → Texto com tabulações*.
Se você não tiver acesso à SE16N nessas tabelas, use a SUIM (relatórios RSUSR002 / RSUSR200) e grave igual.

### Uso diário
Dois cliques em `coleta_perfis.bat` (SAP logado + catálogo GPA em `entrada\`).
Ou: `python coletor_sap.py coletar --fluxo perfis`. O fluxo de estoque continua sendo o padrão.

### Como cada atribuição é classificada
| Situação no GPA | Significado |
|---|---|
| Solicitável | Está no catálogo, ativo e marcado como solicitável |
| Não solicitável | Está no catálogo mas não pode ser pedido via GPA → exige exceção ou remoção |
| Inativo no GPA | Perfil desativado no catálogo |
| Fora do catálogo | Não existe no GPA → atribuído fora do processo |
| Via função composta | Herdado de uma composta (COL_FLAG = X); quem é avaliada é a composta |

Risco = soma dos motivos (fora do catálogo 3, não solicitável 3, crítico 3, sem solicitação 2,
usuário inativo 2, inativo no GPA 2, bloqueado/vencendo/vencida 1): **Alto ≥ 5, Médio ≥ 3, Baixo ≥ 1**.
Atribuições já vencidas não pontuam nos demais motivos (não dão mais acesso).

### Relatório (estilo SAP Fiori)
Blocos de KPI clicáveis, abas por tipo de achado, busca, ordenação por coluna, painel lateral com todos
os perfis do usuário e ações sugeridas, visão *Por perfil* e *Catálogo sem uso*, exportação CSV,
checklist e tendência. Também gera `achados_perfis.csv` e `rascunhos_servicenow_perfis.json`.

## Regras
- Só transações de **consulta**. Nada de correção automática de estoque.
- Nenhum acesso é atribuído ou removido automaticamente. Regularizações seguem o GPA.
- Os arquivos de perfis/usuários são dados pessoais: `entrada\`, `saida\` e `*.db` ficam fora do Git
  (veja `.gitignore`). Combine com Segurança da Informação onde guardar e por quanto tempo.
- Ler AGR_USERS/USR02 via SE16N exige autorização (S_TABU_NAM/S_TABU_DIS). Peça pelo GPA, com justificativa.
- Nenhuma senha é armazenada; usa a sessão já logada.
