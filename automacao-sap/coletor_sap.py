#!/usr/bin/env python3
"""
Coletor SAP GUI Scripting -> Análise MM x EWM
---------------------------------------------
Usa a sessão do SAP Logon JÁ ABERTA e logada (SSO/senha digitada por você),
entra nas transações, executa e exporta os resultados para a pasta "entrada",
depois chama analise_mm_ewm.py e abre o relatório.

Nenhuma senha é guardada. Use apenas transações de CONSULTA.

Comandos:
  python coletor_sap.py importar gravacao_mb52.vbs --nome mm
        -> converte uma gravação do SAP GUI (.vbs) em roteiro no roteiros_sap.json
  python coletor_sap.py listar
        -> mostra as conexões/sessões SAP abertas (para conferir o "sistema")
  python coletor_sap.py coletar [--somente mm,ewm] [--sem-analise]
        -> executa os roteiros e roda a análise

Requisitos: Windows, SAP GUI com scripting habilitado, Python 3.9+, pywin32
  pip install pywin32 pandas openpyxl
"""
import argparse
import datetime as dt
import json
import re
import subprocess
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent
ARQ_ROTEIROS = BASE / "roteiros_sap.json"
PASTA_ENTRADA = BASE / "entrada"
PASTA_LOGS = BASE / "logs"


# ==========================================================================
# 1) Importar gravação .vbs  ->  passos JSON
# ==========================================================================
RE_PROP = re.compile(r'session\.findById\("([^"]+)"\)\.(\w+)\s*=\s*(.+)$', re.I)
RE_METODO = re.compile(r'session\.findById\("([^"]+)"\)\.(\w+)\s*(.*)$', re.I)
IGNORAR = {"setfocus", "caretposition", "maximize", "resizeworkingpane", "restore"}


def _valor_vbs(txt):
    txt = txt.strip()
    if txt.startswith('"') and txt.endswith('"'):
        return txt[1:-1].replace('""', '"')
    if txt.lower() in ("true", "false"):
        return txt.lower() == "true"
    try:
        return int(txt)
    except ValueError:
        try:
            return float(txt)
        except ValueError:
            return txt


def _args_vbs(txt):
    txt = txt.strip()
    if txt.startswith("(") and txt.endswith(")"):
        txt = txt[1:-1]
    if not txt:
        return []
    partes = re.findall(r'"(?:[^"]|"")*"|[^,]+', txt)
    return [_valor_vbs(p) for p in partes if p.strip()]


def importar_vbs(caminho_vbs):
    passos = []
    for linha in Path(caminho_vbs).read_text(encoding="latin-1").splitlines():
        linha = linha.strip()
        if not linha.lower().startswith("session.findbyid"):
            continue
        m = RE_PROP.match(linha)
        if m:
            id_, prop, val = m.groups()
            if prop.lower() in IGNORAR:
                continue
            passos.append({"id": id_, "prop": prop, "valor": _valor_vbs(val)})
            continue
        m = RE_METODO.match(linha)
        if m:
            id_, metodo, args = m.groups()
            if metodo.lower() in IGNORAR:
                continue
            passos.append({"id": id_, "metodo": metodo, "args": _args_vbs(args)})
    return passos


def parametrizar(passos):
    """Troca caminhos/datas gravados por marcadores reutilizáveis."""
    hoje = dt.date.today()
    datas = {hoje.strftime("%d.%m.%Y"): "{DATA_SAP}", hoje.strftime("%d/%m/%Y"): "{DATA_SAP_BARRA}"}
    for p in passos:
        v = p.get("valor")
        if isinstance(v, str):
            if p["id"].endswith("DY_PATH"):
                p["valor"] = "{PASTA_ENTRADA}"
            elif p["id"].endswith("DY_FILENAME"):
                p["valor"] = "{ARQUIVO}"
            elif v in datas:
                p["valor"] = datas[v]
    return passos


def cmd_importar(a):
    roteiros = carregar_roteiros()
    passos = parametrizar(importar_vbs(a.vbs))
    if not passos:
        sys.exit("Nenhum comando session.findById encontrado na gravação.")
    roteiros["roteiros"][a.nome] = {
        "descricao": f"Importado de {Path(a.vbs).name} em {dt.datetime.now():%d/%m/%Y %H:%M}",
        "arquivo_saida": a.arquivo or f"{a.nome}.txt",
        "passos": passos,
    }
    salvar_roteiros(roteiros)
    print(f"Roteiro '{a.nome}' salvo com {len(passos)} passos em {ARQ_ROTEIROS.name}.")
    print("Revise: o caminho/nome do arquivo exportado viraram {PASTA_ENTRADA}/{ARQUIVO}.")


# ==========================================================================
# 2) Conexão com o SAP GUI
# ==========================================================================
def scripting_engine():
    try:
        import win32com.client  # noqa
    except ImportError:
        sys.exit("Instale o pywin32:  pip install pywin32   (somente Windows)")
    try:
        sapgui = win32com.client.GetObject("SAPGUI")
        return sapgui.GetScriptingEngine
    except Exception:
        sys.exit("SAP Logon não encontrado. Abra o SAP Logon e faça login no sistema antes de rodar.")


def listar_sessoes(app):
    out = []
    for i in range(app.Children.Count):
        con = app.Children(i)
        for j in range(con.Children.Count):
            s = con.Children(j)
            info = s.Info
            out.append((i, j, info.SystemName, info.Client, info.User, info.Transaction, s))
    return out


def obter_sessao(app, cfg):
    sessoes = listar_sessoes(app)
    if not sessoes:
        sys.exit("Nenhuma sessão SAP logada. Faça login no SAP Logon e rode de novo.")
    sistema = (cfg.get("sistema") or "").upper()
    mandante = str(cfg.get("mandante") or "")
    candidatas = [s for s in sessoes
                  if (not sistema or s[2].upper() == sistema) and (not mandante or s[3] == mandante)]
    if not candidatas:
        disp = ", ".join(f"{s[2]}/{s[3]}" for s in sessoes)
        sys.exit(f"Sessão do sistema {sistema}/{mandante} não encontrada. Abertas: {disp}")
    base = candidatas[0][-1]
    if not cfg.get("abrir_nova_sessao", True):
        return base
    # Abre uma sessão nova para não atrapalhar o que você está fazendo
    con = base.Parent
    antes = con.Children.Count
    if antes >= 6:
        print("[AVISO] Limite de 6 sessões atingido; usando a sessão existente.")
        return base
    base.CreateSession()
    for _ in range(40):
        time.sleep(0.5)
        if con.Children.Count > antes:
            return con.Children(con.Children.Count - 1)
    sys.exit("Não consegui abrir uma nova sessão SAP.")


# ==========================================================================
# 3) Execução dos roteiros
# ==========================================================================
class ErroSAP(Exception):
    pass


def substituir(valor, variaveis):
    if isinstance(valor, str):
        for k, v in variaveis.items():
            valor = valor.replace("{" + k + "}", str(v))
    return valor


def esperar_ocioso(sessao, timeout=300):
    fim = time.time() + timeout
    while getattr(sessao, "Busy", False):
        if time.time() > fim:
            raise ErroSAP("Tempo esgotado esperando o SAP responder.")
        time.sleep(0.3)


def checar_barra_status(sessao, passo_n):
    try:
        sbar = sessao.findById("wnd[0]/sbar")
        tipo, texto = sbar.MessageType, sbar.Text
    except Exception:
        return
    if tipo in ("E", "A"):
        raise ErroSAP(f"Passo {passo_n}: SAP retornou erro: {texto}")
    if tipo == "W":
        print(f"   [aviso SAP] {texto}")


def executar_passos(sessao, passos, variaveis, log):
    for n, p in enumerate(passos, 1):
        if p.get("acao") == "esperar":
            time.sleep(float(p.get("segundos", 1)))
            continue
        id_ = substituir(p["id"], variaveis)
        try:
            obj = sessao.findById(id_)
            if "prop" in p:
                setattr(obj, p["prop"], substituir(p["valor"], variaveis))
                desc = f'{p["prop"]} = {p["valor"]!r}'
            else:
                args = [substituir(x, variaveis) for x in p.get("args", [])]
                getattr(obj, p["metodo"])(*args)
                desc = f'{p["metodo"]}({", ".join(map(repr, args))})'
            log.write(f"  ok  {n:>3} {id_} .{desc}\n")
        except ErroSAP:
            raise
        except Exception as e:
            if p.get("opcional"):
                log.write(f"  --  {n:>3} {id_} (opcional, ignorado: {e})\n")
                continue
            raise ErroSAP(f"Passo {n} falhou em {id_}: {e}\n"
                          "Provável mudança de tela/ID. Grave de novo a transação e reimporte.")
        esperar_ocioso(sessao)
        if p.get("checar_status", True):
            checar_barra_status(sessao, n)


def voltar_menu(sessao):
    try:
        sessao.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        sessao.findById("wnd[0]").sendVKey(0)
    except Exception:
        pass


def cmd_coletar(a):
    roteiros = carregar_roteiros()
    cfg = roteiros["conexao"]
    selecionados = [r for r in roteiros["ordem"] if r in roteiros["roteiros"]]
    if a.somente:
        selecionados = [r for r in a.somente.split(",") if r in roteiros["roteiros"]]
    pendentes = [r for r in selecionados if roteiros["roteiros"][r].get("modelo")]
    if pendentes:
        sys.exit(f"Os roteiros {pendentes} ainda são MODELO. Grave as transações no SAP GUI e importe:\n"
                 "  python coletor_sap.py importar minha_gravacao.vbs --nome mm")

    PASTA_ENTRADA.mkdir(exist_ok=True)
    PASTA_LOGS.mkdir(exist_ok=True)
    hoje = dt.date.today()
    app = scripting_engine()
    sessao = obter_sessao(app, cfg)
    print(f"Conectado: {sessao.Info.SystemName}/{sessao.Info.Client} usuário {sessao.Info.User}")

    gerados, falhas = {}, []
    log_path = PASTA_LOGS / f"coleta_{dt.datetime.now():%Y%m%d_%H%M%S}.log"
    with open(log_path, "w", encoding="utf-8") as log:
        for nome in selecionados:
            r = roteiros["roteiros"][nome]
            arquivo = r.get("arquivo_saida", f"{nome}.txt")
            destino = PASTA_ENTRADA / arquivo
            if destino.exists():
                destino.unlink()
            variaveis = {
                "PASTA_ENTRADA": str(PASTA_ENTRADA) + "\\", "ARQUIVO": arquivo,
                "DATA_SAP": hoje.strftime("%d.%m.%Y"), "DATA_SAP_BARRA": hoje.strftime("%d/%m/%Y"),
                **roteiros.get("variaveis", {}), **r.get("variaveis", {}),
            }
            print(f"-> {nome}: {r.get('descricao', '')}")
            log.write(f"[{nome}] {dt.datetime.now():%H:%M:%S}\n")
            inicio = time.time()
            try:
                executar_passos(sessao, r["passos"], variaveis, log)
                for _ in range(20):  # espera o arquivo aparecer
                    if destino.exists() and destino.stat().st_size > 0:
                        break
                    time.sleep(0.5)
                if not destino.exists():
                    raise ErroSAP(f"Arquivo {arquivo} não foi gerado. Confira os passos de exportação.")
                gerados[nome] = destino
                print(f"   ok ({time.time() - inicio:.0f}s) -> {destino.name}")
            except ErroSAP as e:
                falhas.append(nome)
                print(f"   FALHOU: {e}")
                log.write(f"  ERRO: {e}\n")
            finally:
                voltar_menu(sessao)

    if cfg.get("abrir_nova_sessao", True) and cfg.get("fechar_sessao_ao_final", True):
        try:
            sessao.findById("wnd[0]/tbar[0]/okcd").text = "/i"
            sessao.findById("wnd[0]").sendVKey(0)
        except Exception:
            pass

    print(f"Log: {log_path}")
    if a.sem_analise:
        return
    if "mm" not in gerados or "ewm" not in gerados:
        sys.exit("Análise não executada: faltou a extração de 'mm' ou 'ewm'.")
    cmd = [sys.executable, str(BASE / "analise_mm_ewm.py"), "--mm", str(gerados["mm"]),
           "--ewm", str(gerados["ewm"])]
    if "filas" in gerados:
        cmd += ["--filas", str(gerados["filas"])]
    if (BASE / "config.json").exists():
        cmd += ["--config", str(BASE / "config.json")]
    subprocess.run(cmd, check=True)
    relatorio = BASE / "saida" / hoje.isoformat() / "relatorio.html"
    if relatorio.exists() and not a.nao_abrir:
        import os
        os.startfile(relatorio)  # noqa  (Windows)
    if falhas:
        print(f"[ATENÇÃO] Roteiros com falha: {falhas}")


def cmd_listar(_):
    app = scripting_engine()
    for i, j, sistema, mandante, usuario, tcode, _ in listar_sessoes(app):
        print(f"con[{i}]/ses[{j}]  sistema={sistema} mandante={mandante} usuário={usuario} transação={tcode}")


# ==========================================================================
# Roteiros (arquivo JSON)
# ==========================================================================
ROTEIROS_PADRAO = {
    "conexao": {
        "sistema": "",            # ex.: "PRD" (vazio = primeira sessão logada)
        "mandante": "",           # ex.: "300"
        "abrir_nova_sessao": True,
        "fechar_sessao_ao_final": True,
    },
    "variaveis": {"CENTRO": "1000"},
    "ordem": ["filas", "mm", "ewm"],
    "roteiros": {
        "mm": {
            "modelo": True,
            "descricao": "MB52 - estoque MM (MODELO: grave e importe o seu)",
            "arquivo_saida": "mm.txt",
            "passos": [
                {"id": "wnd[0]/tbar[0]/okcd", "prop": "text", "valor": "/nMB52"},
                {"id": "wnd[0]", "metodo": "sendVKey", "args": [0]},
                {"id": "wnd[0]/usr/ctxtWERKS-LOW", "prop": "text", "valor": "{CENTRO}"},
                {"id": "wnd[0]/tbar[1]/btn[8]", "metodo": "press", "args": []},
                {"id": "wnd[0]/mbar/menu[0]/menu[1]/menu[2]", "metodo": "select", "args": []},
                {"id": "wnd[1]/tbar[0]/btn[0]", "metodo": "press", "args": []},
                {"id": "wnd[1]/usr/ctxtDY_PATH", "prop": "text", "valor": "{PASTA_ENTRADA}"},
                {"id": "wnd[1]/usr/ctxtDY_FILENAME", "prop": "text", "valor": "{ARQUIVO}"},
                {"id": "wnd[1]/tbar[0]/btn[11]", "metodo": "press", "args": []},
            ],
        },
        "ewm": {
            "modelo": True,
            "descricao": "/SCWM/MON - estoque EWM (MODELO: grave e importe o seu)",
            "arquivo_saida": "ewm.txt",
            "passos": [],
        },
        "filas": {
            "modelo": True,
            "descricao": "SMQ2 - filas qRFC de entrada (MODELO: grave e importe o seu)",
            "arquivo_saida": "filas.txt",
            "passos": [],
        },
    },
}


def carregar_roteiros():
    if not ARQ_ROTEIROS.exists():
        salvar_roteiros(ROTEIROS_PADRAO)
    return json.loads(ARQ_ROTEIROS.read_text(encoding="utf-8"))


def salvar_roteiros(r):
    ARQ_ROTEIROS.write_text(json.dumps(r, indent=2, ensure_ascii=False), encoding="utf-8")


def main():
    p = argparse.ArgumentParser(description="Coletor SAP GUI Scripting para análise MM x EWM")
    sub = p.add_subparsers(dest="cmd", required=True)
    i = sub.add_parser("importar", help="Converte gravação .vbs do SAP GUI em roteiro")
    i.add_argument("vbs")
    i.add_argument("--nome", required=True, help="mm, ewm ou filas")
    i.add_argument("--arquivo", help="Nome do arquivo exportado (padrão: <nome>.txt)")
    i.set_defaults(func=cmd_importar)
    c = sub.add_parser("coletar", help="Executa os roteiros no SAP e roda a análise")
    c.add_argument("--somente", help="Ex.: mm,ewm")
    c.add_argument("--sem-analise", action="store_true")
    c.add_argument("--nao-abrir", action="store_true", help="Não abre o relatório no final")
    c.set_defaults(func=cmd_coletar)
    sub.add_parser("listar", help="Lista sessões SAP abertas").set_defaults(func=cmd_listar)
    a = p.parse_args()
    a.func(a)


if __name__ == "__main__":
    main()
