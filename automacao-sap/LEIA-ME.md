# Análise diária MM x EWM — coleta automática no SAP

## Arquivos
| Arquivo | Para que serve |
|---|---|
| `coletor_sap.py` | Robô SAP GUI Scripting: entra nas transações, executa e exporta |
| `analise_mm_ewm.py` | Cruza MM x EWM x filas e gera o relatório |
| `roteiros_sap.json` | Passos de cada transação (criado na 1ª execução) |
| `coleta_automatica.bat` | Dois cliques: coleta + análise + abre o relatório |
| `rodar_analise.bat` | Modo manual (arquivos exportados por você em `entrada\`) |

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

## Regras
- Só transações de **consulta**. Nada de correção automática de estoque.
- Nenhuma senha é armazenada; usa a sessão já logada.
