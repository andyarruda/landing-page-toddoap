#!/usr/bin/env python3
"""
Análise diária de divergências MM x EWM
---------------------------------------
Lê as exportações de estoque do MM (MB52/MMBE) e do EWM (/SCWM/MON - estoque
disponível ou /SCWM/ERP_STOCKCHECK), opcionalmente a lista de filas qRFC
(SMQ1/SMQ2), cruza tudo e gera:

  saida/AAAA-MM-DD/relatorio.html         -> relatório com checklist do analista
  saida/AAAA-MM-DD/divergencias.csv       -> lista para filtrar no Excel
  saida/AAAA-MM-DD/rascunhos_servicenow.json -> rascunhos de incidente (não envia)
  historico.db                             -> histórico para marcar recorrência

Uso:
  python analise_mm_ewm.py --demo                       (gera dados fictícios e roda)
  python analise_mm_ewm.py --mm mb52.xlsx --ewm ewm.xlsx [--filas smq2.csv]
  python analise_mm_ewm.py --mm mb52.csv --ewm ewm.csv --config config.json

Somente leitura: o script NÃO conecta nem altera nada no SAP.
Requisitos: Python 3.9+, pandas (e openpyxl para ler .xlsx).
"""
import argparse
import datetime as dt
import html
import json
import re
import sqlite3
import sys
import unicodedata
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("Instale o pandas:  pip install pandas openpyxl")

BASE = Path(__file__).resolve().parent

# --------------------------------------------------------------------------
# Configuração padrão (pode ser sobrescrita por um config.json)
# --------------------------------------------------------------------------
CONFIG_PADRAO = {
    "tolerancia_absoluta": 0.001,     # diferenças menores que isso são ignoradas
    "tolerancia_percentual": 0.0,     # ex.: 0.5 = ignora diferenças até 0,5%
    "severidade_alta_qtd": 100,       # |diferença| >= isso -> Alta
    "severidade_media_qtd": 10,       # |diferença| >= isso -> Média
    "dias_historico": 30,
    "grupo_atribuicao_servicenow": "SAP-MM-EWM",
    # Tipo de estoque EWM -> categoria MM
    "mapa_tipo_estoque_ewm": {
        "F1": "livre", "F2": "livre",
        "Q1": "qualidade", "Q2": "qualidade",
        "B1": "bloqueado", "B2": "bloqueado",
    },
    # Se a exportação do EWM vier só com o nº do depósito EWM (ex.: "BR01"),
    # informe aqui para qual centro/depósito MM ele corresponde.
    "mapa_deposito_ewm": {
        # "BR01": {"centro": "1000", "deposito": "0001"}
    },
    "status_fila_erro": ["SYSFAIL", "CPICERR", "STOP", "WAITSTOP", "RETRY"],
    # Nomes de coluna aceitos (comparação ignora acento, maiúsculas e pontuação)
    "colunas_mm": {
        "material": ["Material", "Material No"],
        "centro": ["Centro", "Plant"],
        "deposito": ["Depósito", "Deposito", "Stor. Location", "Storage Location", "SLoc"],
        "lote": ["Lote", "Batch"],
        "livre": ["Utilização livre", "Utiliz.livre", "Utiliz. livre", "Livre", "Unrestricted"],
        "qualidade": ["Em controle qualidade", "Controle qualidade", "Qualidade",
                      "Ctrl.qualidade", "Quality Inspection", "In Quality Insp."],
        "bloqueado": ["Bloqueado", "Estoque bloqueado", "Blocked"],
    },
    "colunas_ewm": {
        "material": ["Produto", "Material", "Product"],
        "centro": ["Centro", "Plant"],
        "deposito": ["Depósito", "Deposito", "Storage Location", "SLoc"],
        "deposito_ewm": ["Nº depósito", "Depósito EWM", "Warehouse No.", "Warehouse Number", "Warehouse"],
        "lote": ["Lote", "Batch"],
        "tipo": ["Tipo de estoque", "Tipo estoque", "Stock Type", "Cat. estoque"],
        "qtd": ["Quantidade", "Qtd", "Qtd.", "Quantity"],
    },
    "colunas_filas": {
        "fila": ["Fila", "Nome da fila", "Queue Name", "Queue"],
        "status": ["Status"],
        "entradas": ["Entradas", "Entries", "Qtd entradas"],
        "destino": ["Destino", "Destination"],
    },
}

CATEGORIAS = ["livre", "qualidade", "bloqueado"]

ACOES = {
    "Só no MM": [
        "Confira em /SCWM/MON se o estoque existe com outro lote, tipo de estoque ou depósito.",
        "Verifique filas qRFC com erro (SMQ2 no EWM / SMQ1 no ERP) para o material.",
        "Procure movimento recente no MM (MB51) não replicado: remessa, transferência ou ajuste.",
        "Se confirmado, trate pela /SCWM/DIFF_ANALYZER, com aprovação do key user.",
    ],
    "Só no EWM": [
        "Confira em MMBE se o estoque está em outro depósito, lote ou categoria.",
        "Verifique se há lançamento EWM aguardando o ERP (filas de saída, SLG1, PPF).",
        "Cheque remessas com saída/entrada confirmada no EWM e não lançada no ERP (VL06O / /SCWM/PRDO / /SCWM/PRDI).",
        "Se confirmado, trate pela /SCWM/DIFF_ANALYZER, com aprovação do key user.",
    ],
    "Diferença de quantidade": [
        "Rode /SCWM/ERP_STOCKCHECK para o material e confirme a diferença.",
        "Compare MB51 (MM) com o histórico de tarefas e documentos do EWM no período.",
        "Verifique filas qRFC com erro e logs (SLG1) do dia.",
        "Divergência recorrente? Investigue a causa raiz antes de corrigir de novo.",
    ],
}

CHECKLIST = [
    "Filas qRFC com erro tratadas ou escaladas (SMQ1/SMQ2)",
    "Divergências de severidade Alta analisadas",
    "Divergências recorrentes avaliadas para causa raiz",
    "Novas divergências registradas no ServiceNow",
    "Key user informado sobre correções que precisam de aprovação",
    "Relatório arquivado / compartilhado com o time",
]


# --------------------------------------------------------------------------
# Utilitários
# --------------------------------------------------------------------------
def normalizar_nome(txt):
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", txt.lower())


def achar_coluna(df, aliases, obrigatoria=True, arquivo=""):
    mapa = {normalizar_nome(c): c for c in df.columns}
    for a in aliases:
        if normalizar_nome(a) in mapa:
            return mapa[normalizar_nome(a)]
    if obrigatoria:
        sys.exit(f"[ERRO] {arquivo}: coluna não encontrada. Aceitos: {aliases}. "
                 f"Colunas no arquivo: {list(df.columns)}. Ajuste no config.json.")
    return None


def numero(valor):
    """Converte '1.234,56', '1234.56', '1,234.56', '12-' (sinal SAP) em float."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace(" ", "")
    if not s:
        return 0.0
    negativo = s.endswith("-") or s.startswith("-")
    s = s.strip("-")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        v = float(s)
    except ValueError:
        return 0.0
    return -v if negativo else v


def texto(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") and s[:-2].isdigit() else s


def material(v):
    s = texto(v)
    return s.lstrip("0") or s if s.isdigit() else s


def ler_texto_sap(caminho, dicas):
    """Lê exportação de lista SAP (.txt): ignora título/linhas tracejadas e acha o cabeçalho."""
    bruto = None
    for enc in ("utf-8-sig", "utf-16", "latin-1"):
        try:
            bruto = Path(caminho).read_text(encoding=enc)
            if "\x00" not in bruto:
                break
        except Exception:
            continue
    linhas = [l.rstrip("\r\n") for l in bruto.splitlines()]
    dicas_n = {normalizar_nome(d) for d in dicas}
    for sep in ("\t", "|", ";"):
        for i, l in enumerate(linhas):
            celulas = [c.strip() for c in l.split(sep)]
            if len(celulas) > 1 and dicas_n & {normalizar_nome(c) for c in celulas}:
                def limpar(cs):
                    cs = [c.strip() for c in cs]
                    if cs and cs[0] == "":
                        cs = cs[1:]
                    if cs and cs[-1] == "":
                        cs = cs[:-1]
                    return cs
                cab = limpar(l.split(sep))
                dados = []
                for d in linhas[i + 1:]:
                    if not d.strip() or set(d.strip()) <= set("-|=_ \t"):
                        continue
                    cs = limpar(d.split(sep))
                    if [normalizar_nome(x) for x in cs] == [normalizar_nome(x) for x in cab]:
                        continue  # cabeçalho repetido a cada página
                    dados.append((cs + [""] * len(cab))[:len(cab)])
                return pd.DataFrame(dados, columns=cab, dtype=str)
    return None


def ler_arquivo(caminho, dicas=()):
    caminho = Path(caminho)
    if not caminho.exists():
        sys.exit(f"[ERRO] Arquivo não encontrado: {caminho}")
    if caminho.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(caminho, dtype=str)
    if dicas:
        df = ler_texto_sap(caminho, dicas)
        if df is not None:
            return df
    for sep in (";", ",", "\t"):
        for enc in ("utf-8-sig", "latin-1"):
            try:
                df = pd.read_csv(caminho, sep=sep, dtype=str, encoding=enc)
                if df.shape[1] > 1:
                    return df
            except Exception:
                continue
    sys.exit(f"[ERRO] Não consegui ler {caminho}. Use .xlsx ou .csv separado por ; ou ,")


def carregar_config(caminho):
    cfg = json.loads(json.dumps(CONFIG_PADRAO))
    if caminho:
        with open(caminho, encoding="utf-8") as f:
            usuario = json.load(f)
        for k, v in usuario.items():
            if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                cfg[k].update(v)
            else:
                cfg[k] = v
    return cfg


# --------------------------------------------------------------------------
# Leitura e normalização
# --------------------------------------------------------------------------
def preparar_mm(df, cfg, nome):
    c = cfg["colunas_mm"]
    col = {k: achar_coluna(df, c[k], k not in ("lote", "qualidade", "bloqueado"), nome) for k in c}
    linhas = []
    for _, r in df.iterrows():
        base = {
            "material": material(r[col["material"]]),
            "centro": texto(r[col["centro"]]),
            "deposito": texto(r[col["deposito"]]),
            "lote": texto(r[col["lote"]]) if col["lote"] else "",
        }
        if not base["material"]:
            continue
        for cat in CATEGORIAS:
            if col[cat]:
                linhas.append({**base, "categoria": cat, "qtd": numero(r[col[cat]])})
    out = pd.DataFrame(linhas)
    if out.empty:
        return pd.DataFrame(columns=["material", "centro", "deposito", "lote", "categoria", "qtd_mm"])
    return (out.groupby(["material", "centro", "deposito", "lote", "categoria"], as_index=False)["qtd"]
               .sum().rename(columns={"qtd": "qtd_mm"}))


def preparar_ewm(df, cfg, nome, avisos):
    c = cfg["colunas_ewm"]
    col = {
        "material": achar_coluna(df, c["material"], True, nome),
        "centro": achar_coluna(df, c["centro"], False, nome),
        "deposito": achar_coluna(df, c["deposito"], False, nome),
        "deposito_ewm": achar_coluna(df, c["deposito_ewm"], False, nome),
        "lote": achar_coluna(df, c["lote"], False, nome),
        "tipo": achar_coluna(df, c["tipo"], True, nome),
        "qtd": achar_coluna(df, c["qtd"], True, nome),
    }
    mapa_tipo = {k.upper(): v for k, v in cfg["mapa_tipo_estoque_ewm"].items()}
    mapa_dep = cfg["mapa_deposito_ewm"]
    tipos_desconhecidos, sem_mapa = set(), set()
    linhas = []
    for _, r in df.iterrows():
        mat = material(r[col["material"]])
        if not mat:
            continue
        centro = texto(r[col["centro"]]) if col["centro"] else ""
        dep = texto(r[col["deposito"]]) if col["deposito"] else ""
        if (not centro or not dep) and col["deposito_ewm"]:
            nr = texto(r[col["deposito_ewm"]])
            if nr in mapa_dep:
                centro = centro or mapa_dep[nr].get("centro", "")
                dep = dep or mapa_dep[nr].get("deposito", "")
            else:
                sem_mapa.add(nr)
        tipo = texto(r[col["tipo"]]).upper()
        cat = mapa_tipo.get(tipo)
        if not cat:
            tipos_desconhecidos.add(tipo)
            continue
        linhas.append({"material": mat, "centro": centro, "deposito": dep,
                       "lote": texto(r[col["lote"]]) if col["lote"] else "",
                       "categoria": cat, "qtd": numero(r[col["qtd"]])})
    if tipos_desconhecidos:
        avisos.append(f"Tipos de estoque EWM ignorados (sem mapeamento): {sorted(tipos_desconhecidos)}. "
                      "Inclua em 'mapa_tipo_estoque_ewm' no config.json se precisarem entrar na comparação.")
    if sem_mapa:
        avisos.append(f"Depósitos EWM sem correspondência centro/depósito: {sorted(sem_mapa)}. "
                      "Preencha 'mapa_deposito_ewm' no config.json.")
    out = pd.DataFrame(linhas)
    if out.empty:
        return pd.DataFrame(columns=["material", "centro", "deposito", "lote", "categoria", "qtd_ewm"])
    return (out.groupby(["material", "centro", "deposito", "lote", "categoria"], as_index=False)["qtd"]
               .sum().rename(columns={"qtd": "qtd_ewm"}))


def preparar_filas(df, cfg, nome):
    c = cfg["colunas_filas"]
    col = {k: achar_coluna(df, c[k], k in ("fila", "status"), nome) for k in c}
    erros = set(s.upper() for s in cfg["status_fila_erro"])
    linhas = []
    for _, r in df.iterrows():
        status = texto(r[col["status"]]).upper()
        if status in erros:
            linhas.append({
                "fila": texto(r[col["fila"]]),
                "status": status,
                "entradas": texto(r[col["entradas"]]) if col["entradas"] else "",
                "destino": texto(r[col["destino"]]) if col["destino"] else "",
            })
    return pd.DataFrame(linhas, columns=["fila", "status", "entradas", "destino"])


# --------------------------------------------------------------------------
# Comparação, classificação e histórico
# --------------------------------------------------------------------------
def comparar(mm, ewm, cfg):
    chave = ["material", "centro", "deposito", "lote", "categoria"]
    df = mm.merge(ewm, on=chave, how="outer")
    df["qtd_mm"] = df["qtd_mm"].fillna(0.0)
    df["qtd_ewm"] = df["qtd_ewm"].fillna(0.0)
    df["diferenca"] = (df["qtd_mm"] - df["qtd_ewm"]).round(3)
    base = df[["qtd_mm", "qtd_ewm"]].abs().max(axis=1).replace(0, 1)
    df["dif_pct"] = (df["diferenca"].abs() / base * 100).round(2)
    df = df[(df["diferenca"].abs() > cfg["tolerancia_absoluta"]) &
            (df["dif_pct"] > cfg["tolerancia_percentual"])].copy()

    def tipo(r):
        if r.qtd_ewm == 0:
            return "Só no MM"
        if r.qtd_mm == 0:
            return "Só no EWM"
        return "Diferença de quantidade"

    def sev(r):
        d = abs(r.diferenca)
        if d >= cfg["severidade_alta_qtd"]:
            return "Alta"
        if d >= cfg["severidade_media_qtd"]:
            return "Média"
        return "Baixa"

    if df.empty:
        df["tipo"] = []
        df["severidade"] = []
        return df
    df["tipo"] = df.apply(tipo, axis=1)
    df["severidade"] = df.apply(sev, axis=1)
    return df


def registrar_historico(df, data, cfg):
    db = sqlite3.connect(BASE / "historico.db")
    db.execute("""CREATE TABLE IF NOT EXISTS divergencia (
                    data TEXT, material TEXT, centro TEXT, deposito TEXT, lote TEXT,
                    categoria TEXT, diferenca REAL,
                    PRIMARY KEY (data, material, centro, deposito, lote, categoria))""")
    inicio = (dt.date.fromisoformat(data) - dt.timedelta(days=cfg["dias_historico"])).isoformat()
    anteriores = {}
    for row in db.execute("""SELECT material, centro, deposito, lote, categoria, COUNT(DISTINCT data)
                             FROM divergencia WHERE data >= ? AND data < ?
                             GROUP BY 1,2,3,4,5""", (inicio, data)):
        anteriores[tuple(row[:5])] = row[5]
    db.execute("DELETE FROM divergencia WHERE data = ?", (data,))
    db.executemany("INSERT INTO divergencia VALUES (?,?,?,?,?,?,?)",
                   [(data, r.material, r.centro, r.deposito, r.lote, r.categoria, r.diferenca)
                    for r in df.itertuples()])
    tendencia = db.execute("""SELECT data, COUNT(*) FROM divergencia
                              GROUP BY data ORDER BY data DESC LIMIT 10""").fetchall()
    db.commit()
    db.close()
    if not df.empty:
        dias = [anteriores.get((r.material, r.centro, r.deposito, r.lote, r.categoria), 0)
                for r in df.itertuples()]
        df["dias_anteriores"] = dias
        df["situacao"] = ["Nova" if d == 0 else f"Recorrente ({d + 1} dias)" for d in dias]
    else:
        df["dias_anteriores"] = []
        df["situacao"] = []
    return df, list(reversed(tendencia))


def rascunhos_servicenow(df, filas, cfg, data):
    incidentes = []
    if not filas.empty:
        incidentes.append({
            "short_description": f"[MM x EWM] {len(filas)} fila(s) qRFC com erro - {data}",
            "description": "Filas com erro:\n" + "\n".join(
                f"- {f.fila} | {f.status} | entradas: {f.entradas} | destino: {f.destino}"
                for f in filas.itertuples()),
            "urgency": "1", "impact": "2",
            "assignment_group": cfg["grupo_atribuicao_servicenow"],
        })
    if df.empty:
        return incidentes
    for (centro, dep), g in df.groupby(["centro", "deposito"]):
        alta = (g["severidade"] == "Alta").sum()
        linhas = [f"- {r.material} | lote {r.lote or '-'} | {r.categoria} | MM {r.qtd_mm:g} | "
                  f"EWM {r.qtd_ewm:g} | dif {r.diferenca:g} | {r.tipo} | {r.situacao}"
                  for r in g.sort_values("diferenca", key=abs, ascending=False).itertuples()]
        incidentes.append({
            "short_description": f"[MM x EWM] {len(g)} divergência(s) centro {centro} dep. {dep} - {data}",
            "description": "Divergências encontradas na análise diária:\n" + "\n".join(linhas[:50])
                           + ("\n(... lista completa no anexo divergencias.csv)" if len(linhas) > 50 else ""),
            "urgency": "1" if alta else "2", "impact": "2",
            "assignment_group": cfg["grupo_atribuicao_servicenow"],
        })
    return incidentes


# --------------------------------------------------------------------------
# Relatório HTML
# --------------------------------------------------------------------------
def gerar_html(df, filas, avisos, tendencia, data, origem):
    e = html.escape
    tot = len(df)
    altas = int((df["severidade"] == "Alta").sum()) if tot else 0
    novas = int((df["situacao"] == "Nova").sum()) if tot else 0
    recorrentes = tot - novas

    def cards():
        itens = [("Divergências", tot, ""), ("Severidade alta", altas, "alta" if altas else ""),
                 ("Novas hoje", novas, ""), ("Recorrentes", recorrentes, "media" if recorrentes else ""),
                 ("Filas com erro", len(filas), "alta" if len(filas) else "")]
        return "".join(f'<div class="card {c}"><div class="n">{n}</div><div>{e(t)}</div></div>'
                       for t, n, c in itens)

    ordem = {"Alta": 0, "Média": 1, "Baixa": 2}
    linhas = ""
    if tot:
        df = df.assign(_o=df["severidade"].map(ordem)).sort_values(["_o", "diferenca"],
                                                                     key=lambda s: s if s.name == "_o" else -s.abs())
        for r in df.itertuples():
            linhas += (f'<tr class="{r.severidade.lower().replace("é", "e")}"><td>{e(r.severidade)}</td>'
                       f"<td>{e(r.situacao)}</td><td>{e(r.tipo)}</td><td>{e(r.material)}</td>"
                       f"<td>{e(r.centro)}</td><td>{e(r.deposito)}</td><td>{e(r.lote or '-')}</td>"
                       f"<td>{e(r.categoria)}</td><td class='num'>{r.qtd_mm:,.3f}</td>"
                       f"<td class='num'>{r.qtd_ewm:,.3f}</td><td class='num'><b>{r.diferenca:,.3f}</b></td></tr>")
    tabela = (f"""<input id="filtro" placeholder="Filtrar por material, centro, tipo..." oninput="filtrar()">
      <table id="div"><thead><tr><th>Sev.</th><th>Situação</th><th>Tipo</th><th>Material</th><th>Centro</th>
      <th>Depósito</th><th>Lote</th><th>Categoria</th><th>Qtd MM</th><th>Qtd EWM</th><th>Diferença</th></tr></thead>
      <tbody>{linhas}</tbody></table>""" if tot else "<p class='ok'>Nenhuma divergência acima da tolerância. ✔</p>")

    tab_filas = ("<p class='ok'>Nenhuma fila com erro no arquivo informado. ✔</p>" if filas.empty and origem.get("filas")
                 else "<p class='muted'>Arquivo de filas não informado — verifique SMQ1/SMQ2 manualmente.</p>"
                 if filas.empty else
                 "<p class='alerta'>Trate as filas antes de corrigir estoque: muitas divergências se resolvem "
                 "sozinhas quando a fila é reprocessada.</p><table><thead><tr><th>Fila</th><th>Status</th>"
                 "<th>Entradas</th><th>Destino</th></tr></thead><tbody>" +
                 "".join(f"<tr class='alta'><td>{e(f.fila)}</td><td>{e(f.status)}</td><td>{e(f.entradas)}</td>"
                         f"<td>{e(f.destino)}</td></tr>" for f in filas.itertuples()) + "</tbody></table>")

    tipos_presentes = sorted(set(df["tipo"])) if tot else []
    acoes = "".join(f"<h3>{e(t)}</h3><ol>" + "".join(f"<li>{e(a)}</li>" for a in ACOES[t]) + "</ol>"
                    for t in tipos_presentes) or "<p class='muted'>Sem ações pendentes.</p>"

    chk = "".join(f'<label><input type="checkbox" data-k="{i}"> {e(t)}</label>' for i, t in enumerate(CHECKLIST))
    av = "".join(f"<li>{e(a)}</li>" for a in avisos)
    maxv = max([n for _, n in tendencia] + [1])
    barras = "".join(f'<div class="bar" title="{d}: {n}"><span style="height:{n / maxv * 100:.0f}%"></span>'
                     f"<small>{d[5:]}</small><b>{n}</b></div>" for d, n in tendencia)

    return f"""<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Divergências MM x EWM — {data}</title>
<style>
 body{{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f4f6f8;color:#1f2933}}
 header{{background:#123b5c;color:#fff;padding:18px 28px}} header h1{{margin:0;font-size:22px}}
 header p{{margin:4px 0 0;opacity:.8;font-size:13px}}
 main{{padding:20px 28px;max-width:1300px;margin:auto}}
 section{{background:#fff;border-radius:8px;padding:16px 20px;margin-bottom:18px;box-shadow:0 1px 3px #0001}}
 h2{{margin-top:0;font-size:17px;color:#123b5c}} h3{{font-size:14px;margin:12px 0 4px}}
 .cards{{display:flex;gap:12px;flex-wrap:wrap}}
 .card{{flex:1;min-width:140px;background:#fff;border-radius:8px;padding:14px;border-left:5px solid #5b8db8;box-shadow:0 1px 3px #0001}}
 .card .n{{font-size:28px;font-weight:700}} .card.alta{{border-color:#c0392b}} .card.media{{border-color:#e6a23c}}
 table{{border-collapse:collapse;width:100%;font-size:13px}} th,td{{padding:6px 8px;border-bottom:1px solid #e3e7eb;text-align:left}}
 th{{background:#eef2f6;position:sticky;top:0}} td.num{{text-align:right;font-variant-numeric:tabular-nums}}
 tr.alta td:first-child{{color:#c0392b;font-weight:700}} tr.media td:first-child{{color:#b7791f;font-weight:700}}
 #filtro{{width:100%;padding:8px;margin-bottom:10px;border:1px solid #cbd2d9;border-radius:6px;box-sizing:border-box}}
 .ok{{color:#2f855a}} .muted{{color:#7b8794}} .alerta{{background:#fdecea;padding:8px 12px;border-radius:6px}}
 label{{display:block;padding:5px 0}} .avisos li{{color:#b7791f}}
 .trend{{display:flex;gap:8px;align-items:flex-end;height:120px}}
 .bar{{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;font-size:11px}}
 .bar span{{width:70%;background:#5b8db8;border-radius:3px 3px 0 0;min-height:2px}}
 .grid{{display:grid;grid-template-columns:1fr 1fr;gap:18px}} @media(max-width:800px){{.grid{{grid-template-columns:1fr}}}}
 @media print{{#filtro{{display:none}}}}
</style></head><body>
<header><h1>Análise diária de divergências MM x EWM</h1>
<p>Data: {data} · MM: {e(origem.get('mm',''))} · EWM: {e(origem.get('ewm',''))} · Filas: {e(origem.get('filas') or 'não informado')}</p></header>
<main>
<div class="cards">{cards()}</div><br>
{f'<section><h2>Avisos de leitura</h2><ul class="avisos">{av}</ul></section>' if avisos else ''}
<section><h2>1. Filas qRFC com erro</h2>{tab_filas}</section>
<section><h2>2. Divergências de estoque</h2>{tabela}</section>
<div class="grid">
<section><h2>3. Ações sugeridas</h2>{acoes}</section>
<section><h2>4. Checklist do analista</h2>{chk}
<h2 style="margin-top:18px">Tendência (divergências por dia)</h2><div class="trend">{barras}</div></section>
</div>
<p class="muted">Relatório gerado automaticamente. Somente leitura — nenhuma alteração foi feita no SAP.
Correções de estoque exigem análise e aprovação do key user.</p>
</main>
<script>
function filtrar(){{const t=document.getElementById('filtro').value.toLowerCase();
 document.querySelectorAll('#div tbody tr').forEach(r=>r.style.display=r.innerText.toLowerCase().includes(t)?'':'none');}}
const K='chk-mmewm-{data}';let s={{}};try{{s=JSON.parse(localStorage.getItem(K)||'{{}}')}}catch(e){{}}
document.querySelectorAll('[data-k]').forEach(c=>{{c.checked=!!s[c.dataset.k];
 c.onchange=()=>{{s[c.dataset.k]=c.checked;try{{localStorage.setItem(K,JSON.stringify(s))}}catch(e){{}}}};}});
</script></body></html>"""


# --------------------------------------------------------------------------
# Dados de demonstração
# --------------------------------------------------------------------------
def gerar_demo(pasta):
    pasta.mkdir(parents=True, exist_ok=True)
    mm = pd.DataFrame([
        ["000000000010001234", "1000", "0001", "L001", "1.500,000", "0", "0"],
        ["000000000010001235", "1000", "0001", "", "320,000", "10,000", "0"],
        ["000000000010001236", "1000", "0001", "L777", "75,000", "0", "5,000"],
        ["000000000010001237", "1000", "0002", "", "980,000", "0", "0"],
        ["000000000010001238", "1000", "0001", "L010", "12,000", "0", "0"],
    ], columns=["Material", "Centro", "Depósito", "Lote", "Utilização livre",
                "Em controle qualidade", "Bloqueado"])
    ewm = pd.DataFrame([
        ["10001234", "BR01", "L001", "F2", "1.500,000"],
        ["10001235", "BR01", "", "F2", "300,000"],
        ["10001235", "BR01", "", "Q2", "10,000"],
        ["10001236", "BR01", "L777", "F2", "75,000"],
        ["10001237", "BR02", "", "F2", "860,000"],
        ["10001239", "BR01", "", "F2", "40,000"],
    ], columns=["Produto", "Nº depósito", "Lote", "Tipo de estoque", "Quantidade"])
    filas = pd.DataFrame([
        ["DLWS0000123456", "SYSFAIL", "3", "ERPCLNT100"],
        ["DLWS0000123460", "READY", "1", "ERPCLNT100"],
    ], columns=["Fila", "Status", "Entradas", "Destino"])
    mm.to_csv(pasta / "mm_mb52.csv", sep=";", index=False)
    ewm.to_csv(pasta / "ewm_estoque.csv", sep=";", index=False)
    filas.to_csv(pasta / "filas_smq2.csv", sep=";", index=False)
    cfg = {"mapa_deposito_ewm": {"BR01": {"centro": "1000", "deposito": "0001"},
                                 "BR02": {"centro": "1000", "deposito": "0002"}}}
    (pasta / "config_demo.json").write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")
    return pasta / "mm_mb52.csv", pasta / "ewm_estoque.csv", pasta / "filas_smq2.csv", pasta / "config_demo.json"


# --------------------------------------------------------------------------
def main():
    p = argparse.ArgumentParser(description="Análise diária de divergências MM x EWM (somente leitura)")
    p.add_argument("--mm", help="Exportação MB52/MMBE (.xlsx ou .csv)")
    p.add_argument("--ewm", help="Exportação de estoque EWM (.xlsx ou .csv)")
    p.add_argument("--filas", help="Exportação SMQ1/SMQ2 (opcional)")
    p.add_argument("--config", help="config.json com mapeamentos e tolerâncias")
    p.add_argument("--data", default=dt.date.today().isoformat(), help="Data da análise (AAAA-MM-DD)")
    p.add_argument("--saida", default=str(BASE / "saida"))
    p.add_argument("--demo", action="store_true", help="Roda com dados fictícios")
    a = p.parse_args()

    if a.demo:
        a.mm, a.ewm, a.filas, a.config = map(str, gerar_demo(BASE / "demo"))
    if not a.mm or not a.ewm:
        p.error("informe --mm e --ewm (ou use --demo)")

    cfg = carregar_config(a.config)
    avisos = []
    mm = preparar_mm(ler_arquivo(a.mm, cfg["colunas_mm"]["material"]), cfg, a.mm)
    ewm = preparar_ewm(ler_arquivo(a.ewm, cfg["colunas_ewm"]["material"]), cfg, a.ewm, avisos)
    filas = preparar_filas(ler_arquivo(a.filas, cfg["colunas_filas"]["fila"]), cfg, a.filas) if a.filas else pd.DataFrame(
        columns=["fila", "status", "entradas", "destino"])

    df = comparar(mm, ewm, cfg)
    df, tendencia = registrar_historico(df, a.data, cfg)

    pasta = Path(a.saida) / a.data
    pasta.mkdir(parents=True, exist_ok=True)
    colunas = ["severidade", "situacao", "tipo", "material", "centro", "deposito", "lote",
               "categoria", "qtd_mm", "qtd_ewm", "diferenca", "dif_pct"]
    df.reindex(columns=colunas).to_csv(pasta / "divergencias.csv", sep=";", index=False,
                                       decimal=",", encoding="utf-8-sig")
    (pasta / "rascunhos_servicenow.json").write_text(
        json.dumps(rascunhos_servicenow(df, filas, cfg, a.data), indent=2, ensure_ascii=False), encoding="utf-8")
    origem = {"mm": Path(a.mm).name, "ewm": Path(a.ewm).name, "filas": Path(a.filas).name if a.filas else ""}
    (pasta / "relatorio.html").write_text(gerar_html(df, filas, avisos, tendencia, a.data, origem), encoding="utf-8")

    print(f"MM: {len(mm)} linhas | EWM: {len(ewm)} linhas | Divergências: {len(df)} | Filas com erro: {len(filas)}")
    for av in avisos:
        print("[AVISO]", av)
    print(f"Relatório: {pasta / 'relatorio.html'}")


if __name__ == "__main__":
    main()
