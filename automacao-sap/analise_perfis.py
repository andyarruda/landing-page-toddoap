#!/usr/bin/env python3
"""
Análise diária de perfis de acesso SAP x catálogo GPA
-----------------------------------------------------
Cruza as atribuições de perfis (funções/roles) dos usuários no SAP com o
catálogo de perfis do GPA e responde, para cada atribuição:

  * o perfil é SOLICITÁVEL via GPA?  (está no catálogo, ativo e marcado como solicitável)
  * existe SOLICITAÇÃO GPA aprovada para esse usuário + perfil?  (se o arquivo for informado)
  * a atribuição está vigente, vencendo ou vencida?
  * o perfil é crítico (SAP_ALL, *ADMIN*, ...)?
  * o usuário está bloqueado ou sem logon há muito tempo?

Entradas (xlsx, csv ou txt exportado do SAP):
  --perfis        AGR_USERS (SE16N) ou SUIM RSUSR002 -> usuário x perfil x validade   [obrigatório]
  --catalogo      Exportação do catálogo de perfis do GPA                              [obrigatório]
  --solicitacoes  Exportação das solicitações do GPA (usuário x perfil x status)       [opcional]
  --usuarios      USR02 (SE16N) ou SUIM RSUSR200 -> tipo, bloqueio, último logon       [opcional]

Saída em saida/AAAA-MM-DD/perfis/:
  relatorio_perfis.html            -> relatório estilo SAP Fiori (filtros, busca, painel por usuário)
  achados_perfis.csv               -> todas as atribuições classificadas, para o Excel
  rascunhos_servicenow_perfis.json -> rascunhos de chamado (NÃO envia nada)
historico_perfis.db                 -> marca achados novos x recorrentes

Uso:
  python analise_perfis.py --demo
  python analise_perfis.py --perfis entrada\\perfis.txt --catalogo entrada\\catalogo_gpa.xlsx
        [--solicitacoes entrada\\solicitacoes_gpa.xlsx] [--usuarios entrada\\usuarios.txt] [--config config_perfis.json]

Somente leitura: o script NÃO conecta, NÃO atribui e NÃO remove nada no SAP nem no GPA.
"""
import argparse
import datetime as dt
import fnmatch
import json
import re
import sqlite3
import sys
from pathlib import Path

import pandas as pd

from analise_mm_ewm import achar_coluna, ler_arquivo, normalizar_nome, texto

BASE = Path(__file__).resolve().parent

CONFIG_PADRAO = {
    "dias_alerta_vencimento": 30,      # atribuição que vence em até N dias -> alerta
    "dias_inatividade": 90,            # usuário de diálogo sem logon há N dias -> inativo
    "dias_historico": 30,
    "sistema_sap": "",                 # ex.: "PRD" -> filtra o catálogo GPA pela coluna Sistema, se existir
    "grupo_atribuicao_servicenow": "SEGURANCA-SAP",
    # Padrões (curingas * e ?) de perfis críticos
    "perfis_criticos": ["SAP_ALL", "SAP_NEW", "*ADMIN*", "*BASIS*", "*SUPERUSER*", "*_FULL*",
                        "Z*SE16*", "Z*DEBUG*"],
    # Usuários/perfis fora da análise (técnicos, padrão SAP etc.)
    "ignorar_usuarios": ["DDIC", "SAP*", "TMSADM", "WF-BATCH", "EARLYWATCH"],
    "ignorar_perfis": [],
    # Valores que significam "sim, solicitável" na coluna do catálogo GPA
    "valores_solicitavel_sim": ["S", "SIM", "X", "Y", "YES", "TRUE", "1", "SOLICITAVEL", "VERDADEIRO"],
    # Status do perfil no catálogo considerados ativos (vazio = sem coluna de status)
    "status_catalogo_ativo": ["ATIVO", "ACTIVE", "LIBERADO", "DISPONIVEL", "PUBLICADO", ""],
    # Status de solicitação GPA que comprovam a concessão
    "status_solicitacao_ok": ["APROVADA", "APROVADO", "CONCLUIDA", "CONCLUIDO", "ATENDIDA",
                              "ATENDIDO", "FINALIZADA", "APPROVED", "CLOSED", "COMPLETED"],
    "colunas_perfis": {
        "usuario": ["Usuário", "Usuario", "Nome do usuário", "User", "User Name", "UNAME", "BNAME"],
        "perfil": ["Função", "Funcao", "Nome da função", "Perfil", "Role", "AGR_NAME"],
        "inicio": ["Data de início", "Data início", "Início validade", "Válido desde", "From", "FROM_DAT"],
        "fim": ["Data de fim", "Data fim", "Data final", "Fim validade", "Válido até", "To", "TO_DAT"],
        "nome": ["Nome completo", "Nome completo do usuário", "Full Name", "NAME_TEXT"],
        "composta": ["COL_FLAG", "Função composta", "Indicador função composta", "Atribuição indireta"],
    },
    "colunas_catalogo": {
        "perfil": ["Perfil SAP", "Perfil", "Função", "Funcao", "Role", "AGR_NAME", "Nome técnico"],
        "solicitavel": ["Solicitável", "Solicitavel", "Solicitável via GPA", "Permite solicitação", "Requestable"],
        "status": ["Status", "Situação", "Situacao"],
        "descricao": ["Descrição", "Descricao", "Descrição do perfil", "Description", "Texto"],
        "aprovador": ["Aprovador", "Responsável", "Responsavel", "Dono do perfil", "Owner"],
        "sistema": ["Sistema", "System", "Sistema SAP"],
    },
    "colunas_solicitacoes": {
        "usuario": ["Usuário", "Usuario", "Chave", "Usuário SAP", "User"],
        "perfil": ["Perfil SAP", "Perfil", "Função", "Funcao", "Role"],
        "numero": ["Solicitação", "Solicitacao", "Nº solicitação", "Número", "Numero", "Request", "Chamado"],
        "status": ["Status", "Situação", "Situacao"],
        "fim": ["Válido até", "Data fim", "Data de fim", "Validade", "Fim validade"],
    },
    "colunas_usuarios": {
        "usuario": ["Usuário", "Usuario", "Nome do usuário", "User", "BNAME"],
        "tipo": ["Tipo de usuário", "Tipo usuário", "Tipo", "User Type", "USTYP"],
        "bloqueio": ["Status bloqueio usuário", "Bloqueio de usuário", "Bloqueio", "Status de bloqueio", "Lock", "UFLAG"],
        "ultimo_logon": ["Data do último logon", "Último logon", "Ultimo logon", "Última data de logon",
                         "Last Logon", "TRDAT"],
        "valido_ate": ["Válido até", "Valido ate", "Valid To", "GLTGB"],
    },
}

# Pontuação de risco por motivo (soma -> Alto >= 5, Médio >= 3, Baixo >= 1)
PESOS = {
    "fora_catalogo": 3, "nao_solicitavel": 3, "inativo_gpa": 2, "sem_solicitacao": 2,
    "critico": 3, "usuario_inativo": 2, "usuario_bloqueado": 1, "vencendo": 1, "vencido": 1,
}

ACOES = {
    "fora_catalogo": "Perfil não existe no catálogo GPA: verificar como foi atribuído e regularizar via GPA "
                     "ou pedir a remoção ao time de Segurança.",
    "nao_solicitavel": "Perfil marcado como NÃO solicitável no GPA: confirmar exceção aprovada (evidência) "
                       "ou pedir a remoção.",
    "inativo_gpa": "Perfil inativo no catálogo GPA: confirmar com o dono do perfil se ainda deve existir.",
    "sem_solicitacao": "Atribuição sem solicitação GPA aprovada: levantar a evidência de concessão.",
    "critico": "Perfil crítico: validar a necessidade com o gestor e com o dono do perfil.",
    "usuario_inativo": "Usuário de diálogo sem logon recente: avaliar revogação dos acessos.",
    "usuario_bloqueado": "Usuário bloqueado com perfis atribuídos: avaliar limpeza na próxima revisão.",
    "vencendo": "Atribuição vence em breve: avisar o usuário para renovar no GPA, se ainda precisar.",
    "vencido": "Atribuição já vencida (não dá mais acesso): pode ser limpa na revisão periódica.",
}

CHECKLIST = [
    "Achados de risco Alto analisados",
    "Perfis fora do catálogo GPA levados ao time de Segurança",
    "Perfis não solicitáveis com exceção confirmada ou remoção pedida",
    "Usuários avisados sobre acessos vencendo",
    "Usuários inativos/bloqueados avaliados para revogação",
    "Relatório arquivado / compartilhado com o time",
]

TIPOS_USUARIO = {"A": "Diálogo", "B": "Sistema", "C": "Comunicação", "S": "Serviço", "L": "Referência"}


# --------------------------------------------------------------------------
# Utilitários
# --------------------------------------------------------------------------
def carregar_config(caminho):
    cfg = json.loads(json.dumps(CONFIG_PADRAO))
    if caminho:
        with open(caminho, encoding="utf-8") as f:
            for k, v in json.load(f).items():
                if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                    cfg[k].update(v)
                else:
                    cfg[k] = v
    return cfg


def data_sap(v):
    """'31.12.9999', '31/12/9999', '9999-12-31', '20261231', '2026-12-31 00:00:00' -> date (ou None)."""
    s = texto(v).split(" ")[0].split("T")[0]
    if not s or set(s) <= set("0./-"):
        return None
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y%m%d", "%d-%m-%Y", "%d.%m.%y", "%d/%m/%y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def bate(valor, padroes):
    return any(fnmatch.fnmatchcase(valor.upper(), p.upper()) for p in padroes)


def fmt_data(d):
    if not d:
        return ""
    return "ilimitado" if d.year >= 9999 else d.strftime("%d/%m/%Y")


def colunas(df, aliases, obrigatorias, nome):
    return {k: achar_coluna(df, v, k in obrigatorias, nome) for k, v in aliases.items()}


def celula(r, col):
    return texto(r[col]) if col else ""


# --------------------------------------------------------------------------
# Leitura
# --------------------------------------------------------------------------
def ler_perfis(caminho, cfg):
    df = ler_arquivo(caminho, cfg["colunas_perfis"]["usuario"] + cfg["colunas_perfis"]["perfil"])
    col = colunas(df, cfg["colunas_perfis"], ("usuario", "perfil"), caminho)
    linhas = []
    for _, r in df.iterrows():
        usuario, perfil = celula(r, col["usuario"]).upper(), celula(r, col["perfil"]).upper()
        if not usuario or not perfil:
            continue
        if bate(usuario, cfg["ignorar_usuarios"]) or bate(perfil, cfg["ignorar_perfis"]):
            continue
        linhas.append({
            "usuario": usuario, "perfil": perfil, "nome": celula(r, col["nome"]),
            "inicio": data_sap(r[col["inicio"]]) if col["inicio"] else None,
            "fim": data_sap(r[col["fim"]]) if col["fim"] else None,
            "via_composta": celula(r, col["composta"]).upper() in ("X", "S", "SIM", "1", "TRUE"),
        })
    # mesma função atribuída em vários intervalos -> fica a de maior validade
    linhas.sort(key=lambda x: (x["usuario"], x["perfil"], x["fim"] or dt.date.max))
    unicas = {}
    for x in linhas:
        unicas[(x["usuario"], x["perfil"])] = x
    return list(unicas.values())


def ler_catalogo(caminho, cfg, avisos):
    df = ler_arquivo(caminho, cfg["colunas_catalogo"]["perfil"])
    col = colunas(df, cfg["colunas_catalogo"], ("perfil",), caminho)
    if not col["solicitavel"]:
        avisos.append("Catálogo GPA sem coluna 'Solicitável': todo perfil presente no catálogo foi considerado "
                      "solicitável.")
    sim = {normalizar_nome(x) for x in cfg["valores_solicitavel_sim"]}
    ativos = {normalizar_nome(x) for x in cfg["status_catalogo_ativo"]}
    sistema = cfg.get("sistema_sap", "").upper()
    cat = {}
    for _, r in df.iterrows():
        perfil = celula(r, col["perfil"]).upper()
        if not perfil:
            continue
        if sistema and col["sistema"] and celula(r, col["sistema"]).upper() not in ("", sistema):
            continue
        cat[perfil] = {
            "solicitavel": (normalizar_nome(celula(r, col["solicitavel"])) in sim) if col["solicitavel"] else True,
            "ativo": (normalizar_nome(celula(r, col["status"])) in ativos) if col["status"] else True,
            "status": celula(r, col["status"]),
            "descricao": celula(r, col["descricao"]),
            "aprovador": celula(r, col["aprovador"]),
        }
    return cat


def ler_solicitacoes(caminho, cfg, hoje):
    df = ler_arquivo(caminho, cfg["colunas_solicitacoes"]["usuario"])
    col = colunas(df, cfg["colunas_solicitacoes"], ("usuario", "perfil", "status"), caminho)
    ok = {normalizar_nome(x) for x in cfg["status_solicitacao_ok"]}
    sol = {}
    for _, r in df.iterrows():
        k = (celula(r, col["usuario"]).upper(), celula(r, col["perfil"]).upper())
        if not all(k):
            continue
        fim = data_sap(r[col["fim"]]) if col["fim"] else None
        valida = normalizar_nome(celula(r, col["status"])) in ok and (fim is None or fim >= hoje)
        atual = sol.get(k)
        if atual and atual["valida"] and not valida:
            continue
        sol[k] = {"numero": celula(r, col["numero"]) or "s/nº", "valida": valida,
                  "status": celula(r, col["status"]), "fim": fim}
    return sol


def ler_usuarios(caminho, cfg):
    df = ler_arquivo(caminho, cfg["colunas_usuarios"]["usuario"])
    col = colunas(df, cfg["colunas_usuarios"], ("usuario",), caminho)
    us = {}
    for _, r in df.iterrows():
        u = celula(r, col["usuario"]).upper()
        if not u:
            continue
        bloq = celula(r, col["bloqueio"]).upper()
        bloqueado = (bloq.isdigit() and int(bloq) != 0) or bool(re.search(r"BLOQ|LOCK|^X$", bloq))
        tipo = celula(r, col["tipo"]).upper()[:1] if col["tipo"] else "A"
        us[u] = {"tipo": tipo or "A", "bloqueado": bloqueado,
                 "ultimo_logon": data_sap(r[col["ultimo_logon"]]) if col["ultimo_logon"] else None,
                 "tem_logon": bool(col["ultimo_logon"]),
                 "valido_ate": data_sap(r[col["valido_ate"]]) if col["valido_ate"] else None}
    return us


# --------------------------------------------------------------------------
# Classificação
# --------------------------------------------------------------------------
def classificar(atribs, cat, sol, usuarios, cfg, hoje):
    dias_venc, dias_inat = cfg["dias_alerta_vencimento"], cfg["dias_inatividade"]
    out = []
    for a in atribs:
        motivos = []
        c = cat.get(a["perfil"])
        vencido = bool(a["fim"] and a["fim"] < hoje)
        restante = (a["fim"] - hoje).days if a["fim"] and not vencido else None

        # 1) Situação no GPA (perfil herdado de função composta segue a composta)
        if a["via_composta"]:
            gpa = "Via função composta"
        elif c is None:
            gpa = "Fora do catálogo"
        elif not c["ativo"]:
            gpa = "Inativo no GPA"
        elif not c["solicitavel"]:
            gpa = "Não solicitável"
        else:
            gpa = "Solicitável"
        if not vencido:
            motivos += {"Fora do catálogo": ["fora_catalogo"], "Inativo no GPA": ["inativo_gpa"],
                        "Não solicitável": ["nao_solicitavel"]}.get(gpa, [])

        # 2) Solicitação GPA
        if sol is None:
            solicitacao = ""
        elif a["via_composta"]:
            solicitacao = "n/a"
        else:
            s = sol.get((a["usuario"], a["perfil"]))
            if s and s["valida"]:
                solicitacao = s["numero"]
            else:
                solicitacao = f"Sem aprovação ({s['numero']}: {s['status']})" if s else "Sem solicitação"
                if not vencido and gpa != "Fora do catálogo":
                    motivos.append("sem_solicitacao")

        # 3) Validade da atribuição
        if vencido:
            validade = f"Vencida em {fmt_data(a['fim'])}"
            motivos.append("vencido")
        elif restante is not None and restante <= dias_venc:
            validade = f"Vence em {restante} dia(s)"
            motivos.append("vencendo")
        else:
            validade = "Vigente"

        # 4) Perfil crítico
        critico = bate(a["perfil"], cfg["perfis_criticos"])
        if critico and not vencido:
            motivos.append("critico")

        # 5) Usuário
        u = usuarios.get(a["usuario"]) if usuarios is not None else None
        if usuarios is None:
            status_u = ""
        elif u is None:
            status_u = "Não encontrado"
        elif u["bloqueado"]:
            status_u = "Bloqueado"
            if not vencido:
                motivos.append("usuario_bloqueado")
        elif u["valido_ate"] and u["valido_ate"] < hoje:
            status_u = "Validade expirada"
            if not vencido:
                motivos.append("usuario_bloqueado")
        elif u["tipo"] != "A":
            status_u = f"Técnico ({TIPOS_USUARIO.get(u['tipo'], u['tipo'])})"
        elif u["tem_logon"] and (u["ultimo_logon"] is None or (hoje - u["ultimo_logon"]).days > dias_inat):
            status_u = ("Nunca logou" if u["ultimo_logon"] is None
                        else f"Sem logon há {(hoje - u['ultimo_logon']).days} dias")
            if not vencido:
                motivos.append("usuario_inativo")
        else:
            status_u = "Ativo"

        pontos = sum(PESOS[m] for m in motivos)
        risco = "Alto" if pontos >= 5 else "Médio" if pontos >= 3 else "Baixo" if pontos >= 1 else "OK"
        out.append({
            "risco": risco, "pontos": pontos, "usuario": a["usuario"], "nome": a["nome"],
            "perfil": a["perfil"], "descricao": (c or {}).get("descricao", ""),
            "aprovador": (c or {}).get("aprovador", ""), "situacao_gpa": gpa,
            "solicitacao": solicitacao, "inicio": fmt_data(a["inicio"]), "fim": fmt_data(a["fim"]),
            "validade": validade, "critico": critico, "status_usuario": status_u,
            "motivos": motivos, "acoes": [ACOES[m] for m in motivos],
        })
    return out


def perfis_sem_uso(cat, achados):
    usados = {x["perfil"] for x in achados}
    return [{"perfil": p, **v} for p, v in sorted(cat.items()) if p not in usados]


def registrar_historico(achados, data, cfg):
    db = sqlite3.connect(BASE / "historico_perfis.db")
    db.execute("""CREATE TABLE IF NOT EXISTS achado (data TEXT, usuario TEXT, perfil TEXT, risco TEXT,
                  motivos TEXT, PRIMARY KEY (data, usuario, perfil))""")
    inicio = (dt.date.fromisoformat(data) - dt.timedelta(days=cfg["dias_historico"])).isoformat()
    anteriores = {(u, p): n for u, p, n in db.execute(
        "SELECT usuario, perfil, COUNT(DISTINCT data) FROM achado WHERE data >= ? AND data < ? GROUP BY 1, 2",
        (inicio, data))}
    db.execute("DELETE FROM achado WHERE data = ?", (data,))
    com_risco = [x for x in achados if x["risco"] != "OK"]
    db.executemany("INSERT INTO achado VALUES (?,?,?,?,?)",
                   [(data, x["usuario"], x["perfil"], x["risco"], ",".join(x["motivos"])) for x in com_risco])
    tendencia = db.execute("""SELECT data, SUM(risco = 'Alto'), COUNT(*) FROM achado
                              GROUP BY data ORDER BY data DESC LIMIT 10""").fetchall()
    db.commit()
    db.close()
    for x in achados:
        n = anteriores.get((x["usuario"], x["perfil"]), 0)
        x["situacao"] = "" if x["risco"] == "OK" else ("Novo" if n == 0 else f"Recorrente ({n + 1} dias)")
    return [{"data": d, "alto": int(a or 0), "total": int(t)} for d, a, t in reversed(tendencia)]


def rascunhos_servicenow(achados, cfg, data):
    grupos = [
        ("fora_catalogo", "perfil(is) atribuído(s) fora do catálogo GPA"),
        ("nao_solicitavel", "atribuição(ões) de perfil não solicitável no GPA"),
        ("sem_solicitacao", "atribuição(ões) sem solicitação GPA aprovada"),
        ("critico", "atribuição(ões) de perfil crítico para revisão"),
        ("usuario_inativo", "usuário(s) inativo(s) com perfis atribuídos"),
    ]
    inc = []
    for motivo, texto_curto in grupos:
        g = [x for x in achados if motivo in x["motivos"]]
        if not g:
            continue
        linhas = [f"- {x['usuario']} | {x['perfil']} | {x['situacao_gpa']} | {x['validade']} | risco {x['risco']}"
                  for x in sorted(g, key=lambda x: -x["pontos"])]
        inc.append({
            "short_description": f"[Perfis SAP x GPA] {len(g)} {texto_curto} - {data}",
            "description": ACOES[motivo] + "\n\n" + "\n".join(linhas[:50])
                           + ("\n(... lista completa em achados_perfis.csv)" if len(linhas) > 50 else ""),
            "urgency": "1" if any(x["risco"] == "Alto" for x in g) else "2", "impact": "2",
            "assignment_group": cfg["grupo_atribuicao_servicenow"],
        })
    return inc


# --------------------------------------------------------------------------
# Relatório HTML (padrão visual SAP Fiori / Horizon)
# --------------------------------------------------------------------------
def gerar_html(achados, sem_uso, tendencia, avisos, data, origem, cfg):
    dados = {"data": data, "achados": achados, "sem_uso": sem_uso, "tendencia": tendencia,
             "avisos": avisos, "origem": origem, "checklist": CHECKLIST,
             "tem_solicitacoes": bool(origem.get("solicitacoes")), "tem_usuarios": bool(origem.get("usuarios")),
             "dias_venc": cfg["dias_alerta_vencimento"], "dias_inat": cfg["dias_inatividade"]}
    js = json.dumps(dados, ensure_ascii=False, default=str).replace("</", "<\\/")
    modelo = (BASE / "modelo_relatorio_perfis.html").read_text(encoding="utf-8")
    return modelo.replace("/*__DADOS__*/null", js).replace("__DATA__", data)


# --------------------------------------------------------------------------
# Dados de demonstração (fictícios)
# --------------------------------------------------------------------------
def gerar_demo(pasta, hoje):
    pasta.mkdir(parents=True, exist_ok=True)
    d = lambda n: (hoje + dt.timedelta(days=n)).strftime("%d.%m.%Y")  # noqa: E731
    perfis = pd.DataFrame([
        ["USR0001", "Ana Lima", "Z_MM_CONSULTA", d(-400), "31.12.9999", ""],
        ["USR0001", "Ana Lima", "Z_EWM_OPERADOR", d(-200), d(12), ""],
        ["USR0002", "Bruno Reis", "Z_MM_CONSULTA", d(-300), "31.12.9999", ""],
        ["USR0002", "Bruno Reis", "Z_FI_LANCAMENTO", d(-90), "31.12.9999", ""],
        ["USR0003", "Carla Souza", "Z_EWM_SUPERVISOR", d(-60), "31.12.9999", ""],
        ["USR0003", "Carla Souza", "Z_EWM_OPERADOR", d(-60), "31.12.9999", "X"],
        ["USR0004", "Diego Alves", "SAP_ALL", d(-10), d(20), ""],
        ["USR0005", "Elisa Prado", "Z_MM_CONSULTA", d(-500), d(-30), ""],
        ["USR0005", "Elisa Prado", "Z_MM_COMPRAS", d(-500), "31.12.9999", ""],
        ["USR0006", "Fábio Nunes", "Z_BASIS_ADMIN", d(-700), "31.12.9999", ""],
        ["USR0007", "Gabi Torres", "Z_SD_VENDAS", d(-100), "31.12.9999", ""],
        ["USR0007", "Gabi Torres", "Z_TESTE_PROJETO_X", d(-100), "31.12.9999", ""],
        ["RFC_EWM", "Usuário RFC EWM", "Z_RFC_EWM_INTEGRACAO", d(-900), "31.12.9999", ""],
    ], columns=["Usuário", "Nome completo", "Função", "Data de início", "Data de fim", "COL_FLAG"])
    catalogo = pd.DataFrame([
        ["Z_MM_CONSULTA", "MM - Consulta de estoque", "Sim", "Ativo", "Key user MM"],
        ["Z_MM_COMPRAS", "MM - Compras", "Sim", "Ativo", "Key user MM"],
        ["Z_EWM_OPERADOR", "EWM - Operador de armazém", "Sim", "Ativo", "Key user EWM"],
        ["Z_EWM_SUPERVISOR", "EWM - Supervisor (composta)", "Sim", "Ativo", "Key user EWM"],
        ["Z_FI_LANCAMENTO", "FI - Lançamentos contábeis", "Sim", "Ativo", "Key user FI"],
        ["Z_SD_VENDAS", "SD - Vendas", "Não", "Ativo", "Key user SD"],
        ["Z_BASIS_ADMIN", "Basis - Administração", "Não", "Ativo", "Segurança SAP"],
        ["Z_PP_PLANEJAMENTO", "PP - Planejamento", "Sim", "Inativo", "Key user PP"],
        ["Z_QM_INSPECAO", "QM - Inspeção", "Sim", "Ativo", "Key user QM"],
    ], columns=["Perfil SAP", "Descrição", "Solicitável", "Status", "Aprovador"])
    solicitacoes = pd.DataFrame([
        ["USR0001", "Z_MM_CONSULTA", "GPA-2025-0101", "Aprovada", ""],
        ["USR0001", "Z_EWM_OPERADOR", "GPA-2026-0450", "Aprovada", d(12)],
        ["USR0002", "Z_MM_CONSULTA", "GPA-2025-0233", "Concluída", ""],
        ["USR0002", "Z_FI_LANCAMENTO", "GPA-2026-0611", "Em aprovação", ""],
        ["USR0003", "Z_EWM_SUPERVISOR", "GPA-2026-0700", "Aprovada", ""],
        ["USR0005", "Z_MM_COMPRAS", "GPA-2024-0012", "Aprovada", ""],
    ], columns=["Usuário", "Perfil SAP", "Solicitação", "Status", "Válido até"])
    usuarios = pd.DataFrame([
        ["USR0001", "A", "0", d(-1)], ["USR0002", "A", "0", d(-3)], ["USR0003", "A", "0", d(0)],
        ["USR0004", "A", "0", d(-2)], ["USR0005", "A", "0", d(-180)], ["USR0006", "A", "64", d(-400)],
        ["USR0007", "A", "0", d(-5)], ["RFC_EWM", "B", "0", ""],
    ], columns=["Usuário", "Tipo de usuário", "Status bloqueio usuário", "Data do último logon"])
    arquivos = {"perfis": perfis, "catalogo": catalogo, "solicitacoes": solicitacoes, "usuarios": usuarios}
    for nome, df in arquivos.items():
        df.to_csv(pasta / f"{nome}.csv", sep=";", index=False, encoding="utf-8-sig")
    return {k: str(pasta / f"{k}.csv") for k in arquivos}


# --------------------------------------------------------------------------
def main():
    p = argparse.ArgumentParser(description="Perfis de acesso SAP x catálogo GPA (somente leitura)")
    p.add_argument("--perfis", help="AGR_USERS / SUIM: usuário x perfil x validade")
    p.add_argument("--catalogo", help="Catálogo de perfis exportado do GPA")
    p.add_argument("--solicitacoes", help="Solicitações exportadas do GPA (opcional)")
    p.add_argument("--usuarios", help="USR02 / SUIM: tipo, bloqueio, último logon (opcional)")
    p.add_argument("--config", help="JSON com ajustes (colunas, perfis críticos, prazos)")
    p.add_argument("--data", default=dt.date.today().isoformat(), help="Data da análise (AAAA-MM-DD)")
    p.add_argument("--saida", default=str(BASE / "saida"))
    p.add_argument("--demo", action="store_true", help="Roda com dados fictícios")
    a = p.parse_args()

    hoje = dt.date.fromisoformat(a.data)
    if a.demo:
        for k, v in gerar_demo(BASE / "demo" / "perfis", hoje).items():
            setattr(a, k, v)
    if not a.perfis or not a.catalogo:
        p.error("informe --perfis e --catalogo (ou use --demo)")

    cfg = carregar_config(a.config)
    avisos = []
    atribs = ler_perfis(a.perfis, cfg)
    cat = ler_catalogo(a.catalogo, cfg, avisos)
    sol = ler_solicitacoes(a.solicitacoes, cfg, hoje) if a.solicitacoes else None
    usuarios = ler_usuarios(a.usuarios, cfg) if a.usuarios else None
    if not cat:
        avisos.append("Catálogo GPA vazio ou sem perfis reconhecidos: todos os perfis aparecerão como fora do catálogo.")
    if sol is None:
        avisos.append("Solicitações GPA não informadas: não foi verificado se cada atribuição tem pedido aprovado.")
    if usuarios is None:
        avisos.append("Dados de usuário (USR02) não informados: bloqueio e inatividade não foram avaliados.")

    achados = classificar(atribs, cat, sol, usuarios, cfg, hoje)
    tendencia = registrar_historico(achados, a.data, cfg)
    sem_uso = perfis_sem_uso(cat, achados)

    pasta = Path(a.saida) / a.data / "perfis"
    pasta.mkdir(parents=True, exist_ok=True)
    cols = ["risco", "situacao", "usuario", "nome", "perfil", "descricao", "situacao_gpa", "solicitacao",
            "inicio", "fim", "validade", "critico", "status_usuario", "aprovador", "motivos"]
    df = pd.DataFrame(achados).reindex(columns=cols)
    if not df.empty:
        df["motivos"] = df["motivos"].map(lambda m: ", ".join(m))
        df["critico"] = df["critico"].map({True: "Sim", False: ""})
    df.to_csv(pasta / "achados_perfis.csv", sep=";", index=False, encoding="utf-8-sig")
    (pasta / "rascunhos_servicenow_perfis.json").write_text(
        json.dumps(rascunhos_servicenow(achados, cfg, a.data), indent=2, ensure_ascii=False), encoding="utf-8")
    origem = {k: Path(getattr(a, k)).name if getattr(a, k) else ""
              for k in ("perfis", "catalogo", "solicitacoes", "usuarios")}
    (pasta / "relatorio_perfis.html").write_text(
        gerar_html(achados, sem_uso, tendencia, avisos, a.data, origem, cfg), encoding="utf-8")

    cont = {r: sum(x["risco"] == r for x in achados) for r in ("Alto", "Médio", "Baixo", "OK")}
    print(f"Atribuições: {len(achados)} | Usuários: {len({x['usuario'] for x in achados})} | "
          f"Risco Alto {cont['Alto']} · Médio {cont['Médio']} · Baixo {cont['Baixo']} · OK {cont['OK']}")
    for av in avisos:
        print("[AVISO]", av)
    print(f"Relatório: {pasta / 'relatorio_perfis.html'}")


if __name__ == "__main__":
    main()
