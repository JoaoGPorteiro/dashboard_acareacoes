"""Ferramentas de linha de comando do dashboard de acareações.

  python main.py importar "ACAREAÇÕES HRAO.xlsx"            gera nomes.csv para revisão (não grava nada)
  python main.py importar "ACAREAÇÕES HRAO.xlsx" --aplicar  grava no banco usando nomes.csv
  python main.py criar-admin nome.do.admin                   cria um administrador (pede a senha)
"""

import argparse
import csv
import getpass
import os
import re
import unicodedata
from collections import Counter
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path

import openpyxl
from fastapi import HTTPException
from pydantic import ValidationError

from api.index import INSERIR, Acareacao, NovoUsuario, chave_nome, conectar, criar_usuario

ANO = 2026
NOMES = Path("nomes.csv")
MESES = {"janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4, "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
         "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12}
PROCEDENCIAS = {"procedente": "Procedente", "improcedente": "Improcedente", "analizando": "Analisando",
                "analisando": "Analisando", "aguardando": "Aguardando", "sem retorno": "Sem retorno"}
SEMELHANCA = 0.9  # grafias 90%+ parecidas (sem acento/caixa/espaços) viram uma sugestão de grupo
GRAVAR_APELIDOS = (
    "insert into apelidos (campo, chave, nome) values (%(campo)s, %(chave)s, %(nome)s) "
    "on conflict (campo, chave) do update set nome = excluded.nome"
)


def limpar(v) -> str:
    return " ".join(str(v).split()) if v is not None else ""


def chave(v) -> str:
    return unicodedata.normalize("NFKD", limpar(v)).encode("ascii", "ignore").decode().casefold()


def ler_planilha(caminho) -> list[dict]:
    """Lê as abas de mês. Jan–Fev têm 12 colunas (com datas e valor); de março em diante, 9."""
    wb = openpyxl.load_workbook(caminho, data_only=True)
    registros = []
    for ws in wb.worksheets:
        mes = MESES.get(chave(ws.title))
        if mes is None:
            continue  # MODELO, RESULTADO, Sheet26...
        linhas = ws.iter_rows(values_only=True)
        cabecalho = next(linhas, ())
        com_datas = len(cabecalho) > 6 and chave(cabecalho[6]) == "data da reclamacao"
        dia = None
        for n, linha in enumerate(linhas, start=2):
            c = list(linha) + [None] * 12
            if not isinstance(c[0], (int, float)):
                if m := re.fullmatch(r"dia (\d{1,2})", chave(c[0])):
                    dia = int(m.group(1))
                continue  # linha vazia, separador "Dia NN" ou linha pré-preenchida sem ticket
            if com_datas:
                data, motorista, embarcador, resolucao, valor, detalhe = c[6], c[7], c[8], c[9], c[10], c[11]
            else:
                data, motorista, embarcador, resolucao, valor, detalhe = None, c[6], c[7], None, None, c[8]
            partes = [str(detalhe).strip() if detalhe else ""]
            if valor not in (None, ""):
                partes.append(f"Valor: {limpar(valor)}")
            if isinstance(resolucao, datetime):
                partes.append(f"Resolução: {resolucao:%d/%m/%Y}")
            registros.append({
                "aba": ws.title,
                "linha": n,
                "data_reclamacao": data.date() if isinstance(data, datetime) else date(ANO, mes, dia or 1),
                "ticket": int(c[0]),
                "entrega": int(c[1]) if isinstance(c[1], (int, float)) else c[1],
                "cliente": limpar(c[2]),
                "acareacao": limpar(c[3]) or "Acareação Motorista",
                "status": limpar(c[4]).capitalize() or "Aguardando",
                "procedencia": PROCEDENCIAS.get(chave(c[5]), limpar(c[5])) or "Aguardando",
                "motorista": limpar(motorista),
                "embarcador": limpar(embarcador),
                "detalhe": " · ".join(p for p in partes if p),
            })
    return registros


def validar(registros: list[dict]) -> tuple[list[dict], list[str]]:
    validos, erros = [], []
    for r in registros:
        try:
            validos.append(Acareacao(**r).model_dump())
        except ValidationError as e:
            campos = ", ".join(sorted({str(err["loc"][0]) for err in e.errors()}))
            erros.append(f"{r['aba']} linha {r['linha']}: {campos}")
    return validos, erros


def agrupar_nomes(registros: list[dict]) -> list[dict]:
    linhas = []
    for campo in ("motorista", "embarcador"):
        contagem = Counter(r[campo] for r in registros)
        grupos: list[list[str]] = []
        for nome, _ in contagem.most_common():  # mais frequente primeiro: vira a sugestão do grupo
            for grupo in grupos:
                if SequenceMatcher(None, chave_nome(nome), chave_nome(grupo[0])).ratio() >= SEMELHANCA:
                    grupo.append(nome)
                    break
            else:
                grupos.append([nome])
        for grupo in grupos:
            linhas += [{"campo": campo, "original": n, "sugerido": grupo[0], "ocorrencias": contagem[n]} for n in grupo]
    return linhas


def escrever_csv(linhas: list[dict], caminho=NOMES) -> None:
    # ; e utf-8-sig: abre direto no Excel em português
    with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["campo", "original", "sugerido", "ocorrencias"], delimiter=";")
        w.writeheader()
        w.writerows(linhas)


def ler_csv(caminho=NOMES) -> dict[tuple[str, str], str]:
    with open(caminho, newline="", encoding="utf-8-sig") as f:
        return {(l["campo"], l["original"]): limpar(l["sugerido"])
                for l in csv.DictReader(f, delimiter=";") if limpar(l["sugerido"])}


def aplicar_nomes(registros: list[dict], mapa: dict[tuple[str, str], str]) -> list[dict]:
    return [{**r, **{c: mapa.get((c, r[c]), r[c]) for c in ("motorista", "embarcador")}} for r in registros]


def montar_de_para(registros: list[dict], mapa: dict[tuple[str, str], str]) -> tuple[list[dict], list[str]]:
    """De/para a gravar: chave de cada grafia original e do nome padrão → nome padrão."""
    de_para: dict[tuple[str, str], str] = {}
    conflitos = set()
    for r in registros:
        for campo in ("motorista", "embarcador"):
            padrao = mapa.get((campo, r[campo]), r[campo])
            for grafia in (r[campo], padrao):
                k = (campo, chave_nome(grafia))
                if de_para.setdefault(k, padrao) != padrao:
                    conflitos.add(f"{campo} '{k[1]}': '{de_para[k]}' ou '{padrao}'?")
    return [{"campo": c, "chave": k, "nome": n} for (c, k), n in de_para.items()], sorted(conflitos)


def importar(caminho, aplicar: bool) -> None:
    registros = ler_planilha(caminho)
    if not aplicar:
        _, erros = validar(registros)
        print(f"{len(registros)} registros lidos, {len(erros)} com erro.", *erros, sep="\n")
        if NOMES.exists():  # já revisado: regerar apagaria o trabalho de revisão
            print(f"{NOMES} já existe e foi mantido. Apague-o para gerar as sugestões de novo.")
            return
        escrever_csv(agrupar_nomes(registros))
        print(f"Revise {NOMES} (coluna 'sugerido') e rode de novo com --aplicar.")
        return
    if not NOMES.exists():
        raise SystemExit(f"{NOMES} não encontrado. Rode primeiro sem --aplicar.")
    mapa = ler_csv()
    validos, erros = validar(aplicar_nomes(registros, mapa))
    de_para, conflitos = montar_de_para(registros, mapa)
    erros += [f"nomes.csv contraditório — {c}" for c in conflitos]
    if erros:
        raise SystemExit("\n".join(erros) + f"\n{len(erros)} linha(s) com erro. Nada foi gravado.")
    with conectar() as conn:
        if conn.execute("select count(*) from acareacoes").fetchone()["count"]:
            raise SystemExit("A tabela já tem registros. Importação cancelada para não duplicar.")
        cur = conn.cursor()
        cur.executemany(INSERIR, [{**v, "criado_por": "importacao"} for v in validos])
        cur.executemany(GRAVAR_APELIDOS, de_para)
    print(f"{len(validos)} registros importados, {len(de_para)} grafias no de/para.")


def criar_admin(usuario: str) -> None:
    senha = getpass.getpass("Senha (mínimo 8 caracteres): ")
    try:
        criar_usuario(NovoUsuario(usuario=usuario, senha=senha, admin=True))
    except ValidationError:
        raise SystemExit("Usuário: 3 a 40 letras minúsculas, números, ponto, _ ou -. Senha: 8+ caracteres.")
    except HTTPException as e:
        raise SystemExit(e.detail)
    print(f"Administrador {usuario} criado.")


def exigir(*nomes: str) -> None:
    faltando = [n for n in nomes if not os.environ.get(n)]
    if faltando:
        raise SystemExit(
            f"Faltam no .env (ou estão vazias): {', '.join(faltando)}.\n"
            "Confira se cada linha está como NOME=valor, com o nome exatamente igual ao do README "
            "(sem espaços nem caracteres invisíveis copiados junto)."
        )
    url = os.environ.get("SUPABASE_URL", "")
    if "SUPABASE_URL" in nomes and not url.startswith("https://"):
        raise SystemExit(f"SUPABASE_URL deve ser a URL completa, ex.: https://{url or '<projeto>'}.supabase.co")


def main(argv=None) -> None:
    from dotenv import load_dotenv  # só existe no ambiente local

    load_dotenv()
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="comando", required=True)
    imp = sub.add_parser("importar", help="importa o histórico da planilha")
    imp.add_argument("planilha")
    imp.add_argument("--aplicar", action="store_true", help="grava no banco usando nomes.csv")
    adm = sub.add_parser("criar-admin", help="cria um usuário administrador")
    adm.add_argument("usuario")
    args = p.parse_args(argv)
    if args.comando == "importar":
        if args.aplicar:
            exigir("DATABASE_URL")
        importar(args.planilha, args.aplicar)
    else:
        exigir("SUPABASE_URL", "SUPABASE_SERVICE_ROLE_KEY")
        criar_admin(args.usuario)


if __name__ == "__main__":
    main()
