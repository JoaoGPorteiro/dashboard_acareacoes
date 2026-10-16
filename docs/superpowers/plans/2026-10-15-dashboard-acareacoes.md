# Dashboard de Acareações — Plano de Implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Substituir a planilha `ACAREAÇÕES HRAO.xlsx` por um dashboard web (cadastro, painel com o RESULTADO automático, registros filtráveis, exportação `.xlsx`, login com admin).

**Architecture:** Uma função FastAPI (`api/index.py`) na Vercel valida o token do Supabase Auth e fala com o Postgres do Supabase via `psycopg`. O frontend React (Vite) baixa todos os registros uma vez e calcula filtros, KPIs e ranking no navegador. `main.py` é a CLI local que importa o histórico e cria a admin.

**Tech Stack:** Python 3.12+ (FastAPI 0.141, psycopg 3.3, openpyxl 3.1, httpx 0.28, pytest), React 18.3 + TypeScript 5.6 + Vite 6 + Vitest 3, `@supabase/supabase-js` 2, `@phosphor-icons/react` 2, Supabase (Auth + Postgres), Vercel.

**Spec:** `docs/superpowers/specs/2026-10-15-dashboard-acareacoes-design.md`

## Global Constraints

- Poucos arquivos, legíveis sem agente: nada de camadas, wrappers ou abstrações além do que está aqui.
- Repositório **público**: nunca commitar `*.xlsx`, `nomes.csv`, `.env*` (já no `.gitignore`).
- Listas fechadas (grafia exata): Acareação = `Whatsapp, Whats Motorista, Acareação Motorista, Recebido` (padrão `Acareação Motorista`); Status = `Aguardando, Recebido, Indenizado, Devolvido, Parado`; Procedência = `Procedente, Improcedente, Analisando, Aguardando, Sem retorno`.
- Login: usuário `fulano` ↔ e-mail `fulano@acareacoes.local`; perfil em `app_metadata.role` ∈ {`admin`, `usuario`}.
- Situação por média mensal no período: ≤1 Ótimo, ≤3 Bom, ≤5 Regular, >5 Atenção.
- Mensagens ao usuário em português.
- Cores só em tokens de `frontend/src/styles.css`. Azul Magalu `#0086FF`; botão `#006FD6`; Improcedente `#0086FF`/escuro `#1A8CFF`; Procedente `#E6007E`/escuro `#E8338C`.
- Fonte do sistema (`-apple-system, "Segoe UI Variable Text", "Segoe UI", system-ui`); tamanhos e espaçamentos em `rem`.
- Python roda local em 3.14 e na Vercel em 3.12: não usar sintaxe exclusiva de 3.13+.
- Rodar testes Python sempre da raiz com `python -m pytest` (põe a raiz no `sys.path`).

## Review Focus

1. Texto começando com `=` (cliente/detalhe) na exportação → precisa sair como texto, não fórmula do Excel. *Teste: Task 1 `test_gerar_xlsx_nao_vira_formula`.*
2. Cadastro público esquecido ligado no Supabase → conta sem `app_metadata.role` não pode usar a API (403). *Teste: Task 2 `test_token_sem_perfil_da_403`.*
3. Rodar `importar --aplicar` duas vezes → segunda vez recusa, sem duplicar. *Teste: Task 3 `test_aplicar_recusa_tabela_com_registros`.*
4. Trimestre/ano em andamento → média mensal não pode ser diluída por meses que ainda não aconteceram. *Teste: Task 4 `calcula intervalos e meses decorridos`.*
5. Ticket/entrega absurdos (0, negativos, 16+ dígitos) → recusados na API. *Teste: Task 1 `test_acareacao_rejeita_valor_invalido`.*

---

## Estrutura de arquivos

```
api/index.py            API inteira (modelo, auth, rotas, xlsx, usuários)
main.py                 CLI: importar planilha, criar admin
schema.sql              tabela, CHECKs, RLS
requirements.txt        dependências da função na Vercel
tests/test_api.py       pytest da API
tests/test_main.py      pytest da importação
frontend/package.json  frontend/tsconfig.json  frontend/vite.config.ts  frontend/index.html
frontend/src/main.tsx        entrada
frontend/src/App.tsx         login, menu, tema, trocar senha, roteamento por hash
frontend/src/Painel.tsx      página inicial
frontend/src/Registros.tsx   tabela + filtros
frontend/src/Usuarios.tsx    admin
frontend/src/Formulario.tsx  painel lateral criar/editar
frontend/src/dados.ts        chamadas à API + cálculos puros
frontend/src/dados.test.ts   vitest
frontend/src/styles.css      tokens e estilos
vercel.json  README.md
```

(A spec previa um só arquivo de teste Python; ficam dois, um por arquivo de código, para cada teste morar ao lado do que testa.)

---

### Task 1: Modelo, banco e exportação `.xlsx`

**Files:**
- Create: `requirements.txt`, `schema.sql`, `api/index.py`
- Test: `tests/test_api.py`

**Interfaces:**
- Produces: `api.index.Acareacao` (Pydantic), `COLUNAS: tuple[str, ...]`, `CABECALHO: tuple[str, ...]`, `INSERIR: str`, `ATUALIZAR: str`, `gerar_xlsx(linhas: list[dict]) -> bytes`, `DOMINIO = "acareacoes.local"`, `PERFIS = {"admin", "usuario"}`.

- [ ] **Step 1: Criar `requirements.txt`**

```
fastapi==0.141.1
httpx==0.28.1
openpyxl==3.1.5
psycopg[binary]==3.3.6
```

- [ ] **Step 2: Criar `schema.sql`**

```sql
-- Rodar uma vez no SQL Editor do Supabase.
create table if not exists acareacoes (
  id bigint generated always as identity primary key,
  data_reclamacao date not null,
  ticket bigint not null check (ticket > 0),
  entrega bigint not null check (entrega > 0),
  cliente text not null check (cliente <> ''),
  acareacao text not null default 'Acareação Motorista'
    check (acareacao in ('Whatsapp', 'Whats Motorista', 'Acareação Motorista', 'Recebido')),
  status text not null check (status in ('Aguardando', 'Recebido', 'Indenizado', 'Devolvido', 'Parado')),
  procedencia text not null
    check (procedencia in ('Procedente', 'Improcedente', 'Analisando', 'Aguardando', 'Sem retorno')),
  motorista text not null check (motorista <> ''),
  embarcador text not null check (embarcador <> ''),
  detalhe text not null default '',
  criado_por text not null,
  criado_em timestamptz not null default now()
);

-- RLS ligado e sem policies: a chave pública (anon) do navegador não lê nem grava esta tabela.
-- Só a API em Python acessa, pela DATABASE_URL.
alter table acareacoes enable row level security;
```

- [ ] **Step 3: Escrever os testes que falham** — `tests/test_api.py`

```python
import io
from datetime import date

import pytest
from openpyxl import load_workbook
from pydantic import ValidationError

from api import index
from api.index import Acareacao, gerar_xlsx

BASE = dict(
    data_reclamacao="2026-05-31",
    ticket=1000007,
    entrega=2000007,
    cliente="  FULANA   DE TAL ",
    status="Recebido",
    procedencia="Improcedente",
    motorista=" Motorista Teste ",
    embarcador="MAGAZINE LUIZA",
)


def linha(**kw):
    return {"id": 1, **Acareacao(**BASE).model_dump(), "criado_por": "teste", "criado_em": None, **kw}


def test_acareacao_limpa_espacos_e_aplica_padrao():
    a = Acareacao(**BASE)
    assert a.cliente == "FULANA DE TAL" and a.motorista == "Motorista Teste"
    assert a.acareacao == "Acareação Motorista" and a.detalhe == ""


@pytest.mark.parametrize(
    "campo,valor",
    [
        ("status", "Pendente"),
        ("procedencia", "procedente"),
        ("acareacao", "Email"),
        ("cliente", "   "),
        ("motorista", ""),
        ("ticket", 0),
        ("entrega", -5),
        ("ticket", 10**15),
    ],
)
def test_acareacao_rejeita_valor_invalido(campo, valor):
    with pytest.raises(ValidationError):
        Acareacao(**{**BASE, campo: valor})


def test_gerar_xlsx_cabecalho_linhas_e_tipos():
    ws = load_workbook(io.BytesIO(gerar_xlsx([linha(), linha(id=2, cliente="CICLANA")])))["Acareações"]
    assert tuple(c.value for c in ws[1]) == index.CABECALHO
    assert ws["A2"].value.date() == date(2026, 5, 31) and ws["A2"].number_format == "DD/MM/YYYY"
    assert ws["B2"].value == 1000007 and ws["D2"].value == "FULANA DE TAL" and ws["D3"].value == "CICLANA"
    assert ws.freeze_panes == "A2" and ws.auto_filter.ref == "A1:J3"


def test_gerar_xlsx_nao_vira_formula():
    ws = load_workbook(io.BytesIO(gerar_xlsx([linha(cliente='=HYPERLINK("http://x")')]))).active
    assert ws["D2"].data_type == "s" and ws["D2"].value == '=HYPERLINK("http://x")'
```

- [ ] **Step 4: Rodar e ver falhar**

Run: `python -m pytest tests/test_api.py -v`
Expected: ERROR `ModuleNotFoundError: No module named 'api'`

- [ ] **Step 5: Implementar** — `api/index.py`

```python
"""API do dashboard de acareações. Roda como função Python na Vercel (variável `app`)."""

import io
from datetime import date
from typing import Literal

from openpyxl import Workbook
from openpyxl.styles import Font
from pydantic import BaseModel, Field, field_validator

DOMINIO = "acareacoes.local"  # usuário "fulano" existe no Supabase Auth como fulano@acareacoes.local
PERFIS = {"admin", "usuario"}

COLUNAS = (
    "data_reclamacao", "ticket", "entrega", "cliente", "acareacao",
    "status", "procedencia", "motorista", "embarcador", "detalhe",
)
CABECALHO = (
    "DATA DA RECLAMAÇÃO", "TICKET", "Nº DA ENTREGA", "CLIENTE", "ACAREAÇÃO MOTORISTA",
    "STATUS", "PROCEDÊNCIA", "MOTORISTA", "EMBARCADOR", "DETALHE",
)
LARGURAS = (14, 12, 14, 28, 22, 13, 14, 22, 22, 60)

# Nomes de coluna vêm só de COLUNAS (constante), nunca do usuário.
INSERIR = (
    f"insert into acareacoes ({', '.join(COLUNAS)}, criado_por) "
    f"values ({', '.join(f'%({c})s' for c in COLUNAS)}, %(criado_por)s) returning *"
)
ATUALIZAR = f"update acareacoes set {', '.join(f'{c} = %({c})s' for c in COLUNAS)} where id = %(id)s returning *"


class Acareacao(BaseModel):
    data_reclamacao: date
    ticket: int = Field(gt=0, lt=10**15)
    entrega: int = Field(gt=0, lt=10**15)
    cliente: str = Field(min_length=1, max_length=200)
    acareacao: Literal["Whatsapp", "Whats Motorista", "Acareação Motorista", "Recebido"] = "Acareação Motorista"
    status: Literal["Aguardando", "Recebido", "Indenizado", "Devolvido", "Parado"]
    procedencia: Literal["Procedente", "Improcedente", "Analisando", "Aguardando", "Sem retorno"]
    motorista: str = Field(min_length=1, max_length=200)
    embarcador: str = Field(min_length=1, max_length=200)
    detalhe: str = Field(default="", max_length=5000)

    @field_validator("cliente", "motorista", "embarcador", mode="before")
    @classmethod
    def _espacos(cls, v):
        return " ".join(v.split()) if isinstance(v, str) else v

    @field_validator("detalhe", mode="before")
    @classmethod
    def _detalhe(cls, v):
        return "" if v is None else v.strip() if isinstance(v, str) else v


def gerar_xlsx(linhas: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Acareações"
    ws.append(CABECALHO)
    for cel in ws[1]:
        cel.font = Font(bold=True)
    for linha in linhas:
        ws.append([linha[c] for c in COLUNAS])
        for cel in ws[ws.max_row]:
            if isinstance(cel.value, str) and cel.value.startswith("="):
                cel.data_type = "s"  # texto digitado não pode virar fórmula no Excel
    for cel in ws["A"][1:]:
        cel.number_format = "DD/MM/YYYY"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, largura in enumerate(LARGURAS):
        ws.column_dimensions[chr(65 + i)].width = largura
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
```

- [ ] **Step 6: Rodar e ver passar**

Run: `python -m pytest tests/test_api.py -v`
Expected: 11 passed. Se `test_gerar_xlsx_nao_vira_formula` falhar porque o openpyxl salvou como fórmula mesmo com `data_type = "s"`, trocar a linha por `cel.value, cel.data_type = cel.value, "s"` e, se ainda falhar, investigar com superpowers:systematic-debugging antes de mudar a abordagem.

- [ ] **Step 7: Commit**

```bash
git add requirements.txt schema.sql api/index.py tests/test_api.py
git commit -m "feat(api): modelo de acareação, schema e exportação xlsx"
```

---

### Task 2: Rotas da API, login e usuários

**Files:**
- Modify: `api/index.py`
- Test: `tests/test_api.py`

**Interfaces:**
- Consumes: Task 1 (`Acareacao`, `INSERIR`, `ATUALIZAR`, `gerar_xlsx`, `DOMINIO`, `PERFIS`).
- Produces: `app` (FastAPI); `supabase(metodo, caminho, token=None, **kw) -> httpx.Response`; `usuario_atual(authorization) -> dict{id, usuario, admin, criado_em}`; `so_admin`; `conectar() -> psycopg.Connection` (row_factory `dict_row`); `NovoUsuario(usuario, senha, admin=False)`; `criar_usuario(novo: NovoUsuario) -> dict`. Rotas: `GET/POST /api/acareacoes`, `PUT/DELETE /api/acareacoes/{rid}`, `POST /api/exportar {ids}`, `GET/POST /api/usuarios`, `DELETE /api/usuarios/{uid}`. Erros como `{"detail": "<mensagem em português>"}`.

- [ ] **Step 1: Escrever os testes que falham** — acrescentar ao fim de `tests/test_api.py` (e somar `import os` e `from fastapi.testclient import TestClient` aos imports, além de `app, usuario_atual` no `from api.index import ...`)

```python
cliente = TestClient(app)
EU = "00000000-0000-0000-0000-000000000001"


class Resp:
    def __init__(self, status_code, corpo=None):
        self.status_code, self._corpo = status_code, corpo

    def json(self):
        return self._corpo


@pytest.fixture
def como():
    def _como(admin):
        app.dependency_overrides[usuario_atual] = lambda: {"id": EU, "usuario": "teste", "admin": admin, "criado_em": None}

    yield _como
    app.dependency_overrides.clear()


def test_sem_token_da_401():
    r = cliente.get("/api/acareacoes")
    assert r.status_code == 401 and r.json()["detail"] == "Faça login para continuar."


def test_token_invalido_da_401(monkeypatch):
    monkeypatch.setattr(index, "supabase", lambda *a, **k: Resp(401))
    assert cliente.get("/api/acareacoes", headers={"Authorization": "Bearer x"}).status_code == 401


def test_token_sem_perfil_da_403(monkeypatch):
    conta = {"id": EU, "email": "intruso@acareacoes.local", "app_metadata": {}}
    monkeypatch.setattr(index, "supabase", lambda *a, **k: Resp(200, conta))
    assert cliente.get("/api/acareacoes", headers={"Authorization": "Bearer x"}).status_code == 403


def test_usuario_comum_nao_gerencia_usuarios(como):
    como(admin=False)
    assert cliente.get("/api/usuarios").status_code == 403
    assert cliente.post("/api/usuarios", json={"usuario": "novo", "senha": "12345678"}).status_code == 403
    assert cliente.delete(f"/api/usuarios/{EU}").status_code == 403


def test_admin_nao_remove_a_si_mesmo(como):
    como(admin=True)
    r = cliente.delete(f"/api/usuarios/{EU}")
    assert r.status_code == 400 and "si mesmo" in r.json()["detail"]


def test_novo_usuario_valida_nome_e_senha(como):
    como(admin=True)
    assert cliente.post("/api/usuarios", json={"usuario": "A B", "senha": "12345678"}).status_code == 422
    assert cliente.post("/api/usuarios", json={"usuario": "novo", "senha": "123"}).status_code == 422


def test_registro_invalido_da_422_em_portugues(como):
    como(admin=False)
    r = cliente.post("/api/acareacoes", json={**BASE, "status": "X"})
    assert r.status_code == 422 and r.json()["detail"] == "Confira os campos: status"


def test_exportar_devolve_xlsx_dos_ids(como, monkeypatch):
    como(admin=False)

    class Conn:
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

        def execute(self, sql, params):
            self.params = params
            return self

        def fetchall(self):
            return [linha()]

    conn = Conn()
    monkeypatch.setattr(index, "conectar", lambda: conn)
    r = cliente.post("/api/exportar", json={"ids": [1]})
    assert r.status_code == 200 and conn.params == ([1],)
    assert r.headers["content-disposition"].startswith('attachment; filename="acareacoes_')
    assert load_workbook(io.BytesIO(r.content)).active["D2"].value == "FULANA DE TAL"


@pytest.mark.skipif(not os.environ.get("DATABASE_URL_TESTE"), reason="sem banco de teste (DATABASE_URL_TESTE)")
def test_crud_no_banco(como, monkeypatch):
    monkeypatch.setenv("DATABASE_URL", os.environ["DATABASE_URL_TESTE"])
    como(admin=False)
    criado = cliente.post("/api/acareacoes", json=BASE)
    assert criado.status_code == 201
    rid = criado.json()["id"]
    try:
        assert criado.json()["criado_por"] == "teste" and criado.json()["cliente"] == "FULANA DE TAL"
        editado = cliente.put(f"/api/acareacoes/{rid}", json={**BASE, "status": "Indenizado"})
        assert editado.json()["status"] == "Indenizado"
        assert any(r["id"] == rid for r in cliente.get("/api/acareacoes").json())
    finally:
        assert cliente.delete(f"/api/acareacoes/{rid}").status_code == 204
    assert cliente.put(f"/api/acareacoes/{rid}", json=BASE).status_code == 404
    assert cliente.delete(f"/api/acareacoes/{rid}").status_code == 404
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_api.py -v`
Expected: ERROR `ImportError: cannot import name 'app' from 'api.index'`

- [ ] **Step 3: Implementar** — em `api/index.py`, trocar o bloco de imports por:

```python
import io
import os
from datetime import date
from typing import Literal
from uuid import UUID

import httpx
import psycopg
from fastapi import Depends, FastAPI, Header, HTTPException, Request, Response
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse
from openpyxl import Workbook
from openpyxl.styles import Font
from psycopg.rows import dict_row
from pydantic import BaseModel, Field, field_validator
```

e acrescentar ao fim do arquivo:

```python
class NovoUsuario(BaseModel):
    usuario: str = Field(pattern=r"^[a-z0-9._-]{3,40}$")
    senha: str = Field(min_length=8, max_length=72)
    admin: bool = False


class Exportacao(BaseModel):
    ids: list[int] = Field(max_length=50_000)


app = FastAPI(title="Acareações", docs_url=None, redoc_url=None, openapi_url=None)


@app.exception_handler(RequestValidationError)
async def _erro_validacao(_: Request, exc: RequestValidationError) -> JSONResponse:
    campos = sorted({str(e["loc"][-1]) for e in exc.errors()})
    return JSONResponse({"detail": "Confira os campos: " + ", ".join(campos)}, status_code=422)


def supabase(metodo: str, caminho: str, token: str | None = None, **kw) -> httpx.Response:
    chave = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    cabecalhos = {"apikey": chave, "Authorization": f"Bearer {token or chave}"}
    return httpx.request(metodo, os.environ["SUPABASE_URL"].rstrip("/") + caminho, headers=cabecalhos, timeout=10, **kw)


def _resumo(u: dict) -> dict:
    return {
        "id": u["id"],
        "usuario": u["email"].split("@")[0],
        "admin": u.get("app_metadata", {}).get("role") == "admin",
        "criado_em": u.get("created_at"),
    }


def usuario_atual(authorization: str = Header(default="")) -> dict:
    token = authorization.removeprefix("Bearer ").strip()
    if not token:
        raise HTTPException(401, "Faça login para continuar.")
    r = supabase("GET", "/auth/v1/user", token)
    if r.status_code != 200:
        raise HTTPException(401, "Sessão expirada. Entre novamente.")
    u = r.json()
    # Conta sem perfil (ex.: cadastro público ligado por engano no Supabase) não entra.
    if u.get("app_metadata", {}).get("role") not in PERFIS:
        raise HTTPException(403, "Usuário sem acesso ao dashboard.")
    return _resumo(u)


def so_admin(u: dict = Depends(usuario_atual)) -> dict:
    if not u["admin"]:
        raise HTTPException(403, "Apenas administradores podem fazer isso.")
    return u


def conectar() -> psycopg.Connection:
    # prepare_threshold=None: o pooler do Supabase (modo transação) não aceita prepared statements.
    return psycopg.connect(os.environ["DATABASE_URL"], prepare_threshold=None, row_factory=dict_row)


@app.get("/api/acareacoes")
def listar(_: dict = Depends(usuario_atual)) -> list[dict]:
    # shortcut: devolve tudo (~1.200/ano); paginar se passar de ~20 mil registros.
    with conectar() as c:
        return c.execute("select * from acareacoes order by data_reclamacao desc, id desc").fetchall()


@app.post("/api/acareacoes", status_code=201)
def criar(a: Acareacao, u: dict = Depends(usuario_atual)) -> dict:
    with conectar() as c:
        return c.execute(INSERIR, {**a.model_dump(), "criado_por": u["usuario"]}).fetchone()


@app.put("/api/acareacoes/{rid}")
def editar(rid: int, a: Acareacao, _: dict = Depends(usuario_atual)) -> dict:
    with conectar() as c:
        linha = c.execute(ATUALIZAR, {**a.model_dump(), "id": rid}).fetchone()
    if not linha:
        raise HTTPException(404, "Registro não encontrado.")
    return linha


@app.delete("/api/acareacoes/{rid}", status_code=204)
def excluir(rid: int, _: dict = Depends(usuario_atual)) -> Response:
    with conectar() as c:
        if c.execute("delete from acareacoes where id = %s", (rid,)).rowcount == 0:
            raise HTTPException(404, "Registro não encontrado.")
    return Response(status_code=204)


@app.post("/api/exportar")
def exportar(e: Exportacao, _: dict = Depends(usuario_atual)) -> Response:
    with conectar() as c:
        linhas = c.execute(
            "select * from acareacoes where id = any(%s) order by data_reclamacao desc, id desc", (e.ids,)
        ).fetchall()
    return Response(
        gerar_xlsx(linhas),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="acareacoes_{date.today().isoformat()}.xlsx"'},
    )


def criar_usuario(novo: NovoUsuario) -> dict:
    r = supabase("POST", "/auth/v1/admin/users", json={
        "email": f"{novo.usuario}@{DOMINIO}",
        "password": novo.senha,
        "email_confirm": True,
        "app_metadata": {"role": "admin" if novo.admin else "usuario"},
    })
    if r.status_code >= 400:
        if "email_exists" in r.text or "already" in r.text:
            raise HTTPException(409, "Esse usuário já existe.")
        raise HTTPException(400, f"Não foi possível criar o usuário (Supabase {r.status_code}).")
    return _resumo(r.json())


@app.get("/api/usuarios")
def listar_usuarios(_: dict = Depends(so_admin)) -> list[dict]:
    r = supabase("GET", "/auth/v1/admin/users", params={"per_page": 1000})
    r.raise_for_status()
    return sorted((_resumo(u) for u in r.json()["users"]), key=lambda u: u["usuario"])


@app.post("/api/usuarios", status_code=201)
def novo_usuario(novo: NovoUsuario, _: dict = Depends(so_admin)) -> dict:
    return criar_usuario(novo)


@app.delete("/api/usuarios/{uid}", status_code=204)
def remover_usuario(uid: UUID, u: dict = Depends(so_admin)) -> Response:
    if str(uid) == u["id"]:
        raise HTTPException(400, "Você não pode remover a si mesmo.")
    r = supabase("DELETE", f"/auth/v1/admin/users/{uid}")
    if r.status_code == 404:
        raise HTTPException(404, "Usuário não encontrado.")
    r.raise_for_status()
    return Response(status_code=204)
```

- [ ] **Step 4: Rodar e ver passar**

Run: `python -m pytest tests/test_api.py -v`
Expected: 19 passed, 1 skipped (`test_crud_no_banco` sem `DATABASE_URL_TESTE`).

- [ ] **Step 5: Commit**

```bash
git add api/index.py tests/test_api.py
git commit -m "feat(api): rotas de acareações, exportação, login e usuários"
```

---

### Task 3: Importação do histórico e criação da admin (`main.py`)

**Files:**
- Modify: `main.py` (hoje vazio)
- Test: `tests/test_main.py`

**Interfaces:**
- Consumes: `api.index.Acareacao`, `INSERIR`, `NovoUsuario`, `conectar`, `criar_usuario`.
- Produces: `ler_planilha(caminho) -> list[dict]` (cada dict tem `aba`, `linha` + campos de `Acareacao`); `validar(registros) -> tuple[list[dict], list[str]]`; `agrupar_nomes(registros) -> list[dict{campo, original, sugerido, ocorrencias}]`; `escrever_csv(linhas, caminho=NOMES)`; `ler_csv(caminho=NOMES) -> dict[(campo, original), sugerido]`; `aplicar_nomes(registros, mapa) -> list[dict]`; `importar(caminho, aplicar: bool)`; `criar_admin(usuario)`; `main(argv=None)`.

Regras da planilha real (conferidas em 2026-10-15): jan–fev com 12 colunas (data na G); mar–out com 9 colunas e cabeçalhos às vezes errados ("TICKET, TICKET, TICKET") — vale a posição; ago–out com linhas "Dia NN"/"DIA 06"; linhas pré-preenchidas só com "Acareação Motorista" e sem ticket; 4 linhas de janeiro sem data; 8 linhas sem status/procedência e 3 sem acareação. Decisões: sem data → dia 1º do mês da aba (ou dia do último "Dia NN"); status/procedência vazios → `Aguardando`; acareação vazia → `Acareação Motorista`.

- [ ] **Step 1: Escrever os testes que falham** — `tests/test_main.py`

```python
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook

import main

CAB_JAN = ["TICKET", "ENTREGA", "CLIENTE", "ACAREAÇÃO MOTORISTA", "STATUS", "PROCEDÊNCIA", "DATA DA RECLAMAÇÃO",
           "MOTORISTA", "EMBARCADOR", "DATA DA RESOLUÇÃO", "VALOR", "DETALHE "]
CAB_NOVO = ["TICKET", "TICKET", "TICKET", "ACAREAÇÃO MOTORISTA", "STATUS", "PROCEDÊNCIA", "MOTORISTA", "EMBARCADOR",
            "DETALHE "]
REAL = Path(__file__).parent.parent / "ACAREAÇÕES HRAO.xlsx"


@pytest.fixture
def planilha(tmp_path):
    wb = Workbook()
    wb.remove(wb.active)
    modelo = wb.create_sheet("MODELO")
    modelo.append(CAB_JAN)
    modelo.append([None, 1, "X", "Whats Motorista", "Aguardando", "Analizando"])
    jan = wb.create_sheet("JANEIRO")
    jan.append(CAB_JAN)
    jan.append([1000001.0, 2000001.0, "CLIENTE  UM ", "Acareação Motorista", "Recebido", "Improcedente",
                datetime(2026, 1, 2), "Motorista Alfa", "KABUM", datetime(2026, 3, 1), "R$ 100,00", None])
    jan.append([1000002.0, 2000002.0, "CLIENTE DOIS", "Acareação Motorista", "Aguardando", "Analizando", None,
                "Motorista Graca", "ELECTROLUX", None, None, "cliente sumiu"])
    mar = wb.create_sheet("MARÇO")
    mar.append(CAB_NOVO)
    mar.append([1000003.0, 2000003.0, "CLIENTE TRES", "Acareação Motorista", "Recebido", "procedente", "Motorista Graça",
                "MAGAZINE LUIZA", None])
    mar.append([1000004.0, 2000004.0, "CLIENTE QUATRO", None, None, None, "Motorista  Graça", "DAFITTI", None])
    ago = wb.create_sheet("AGOSTO")
    ago.append(CAB_NOVO)
    ago.append(["DIA 06"])
    ago.append([1000005.0, 2000005.0, "CLIENTE CINCO", "Acareação Motorista", "Recebido", "Improcedente", "Motorista Beta",
                "DAFITI", "falta interna"])
    ago.append(["Dia 07 "])
    ago.append([None, None, None, "Acareação Motorista"])
    ago.append([1000006.0, 2000006.0, "CLIENTE SEIS", "Acareação Motorista", "Devolvido", "Improcedente", "Motorista Beta",
                "DAFITI", None])
    res = wb.create_sheet("RESULTADO")
    res.append(["MOTORISTA", "IMPROCEDENTE"])
    res.append(["Alex", 5])
    caminho = tmp_path / "planilha.xlsx"
    wb.save(caminho)
    return caminho


class Conn:
    """Conexão falsa: guarda o que seria gravado."""

    def __init__(self, ja_tem=0):
        self.ja_tem, self.gravados = ja_tem, []

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def execute(self, sql, *a):
        return self

    def fetchone(self):
        return {"count": self.ja_tem}

    def cursor(self):
        return self

    def executemany(self, sql, linhas):
        self.gravados.extend(linhas)


def test_le_so_abas_de_mes_e_ignora_linhas_sem_ticket(planilha):
    assert [r["ticket"] for r in main.ler_planilha(planilha)] == [1000001, 1000002, 1000003, 1000004, 1000005, 1000006]


def test_layout_janeiro_preserva_valor_e_resolucao_no_detalhe(planilha):
    r = main.ler_planilha(planilha)[0]
    assert (r["data_reclamacao"], r["motorista"], r["embarcador"]) == (date(2026, 1, 2), "Motorista Alfa", "KABUM")
    assert r["cliente"] == "CLIENTE UM"
    assert r["detalhe"] == "Valor: R$ 100,00 · Resolução: 01/03/2026"


def test_datas_sem_dia_usam_dia_1_e_dia_nn_vale_para_as_linhas_abaixo(planilha):
    datas = [r["data_reclamacao"] for r in main.ler_planilha(planilha)]
    assert datas[1:] == [date(2026, 1, 1), date(2026, 3, 1), date(2026, 3, 1), date(2026, 8, 6), date(2026, 8, 7)]


def test_normaliza_listas_e_preenche_vazios(planilha):
    regs = main.ler_planilha(planilha)
    assert regs[1]["procedencia"] == "Analisando" and regs[1]["detalhe"] == "cliente sumiu"
    assert regs[2]["procedencia"] == "Procedente"
    assert (regs[3]["acareacao"], regs[3]["status"], regs[3]["procedencia"]) == (
        "Acareação Motorista", "Aguardando", "Aguardando")


def test_tudo_valido(planilha):
    validos, erros = main.validar(main.ler_planilha(planilha))
    assert erros == [] and len(validos) == 6 and "aba" not in validos[0]


def test_validar_aponta_aba_linha_e_campo():
    r = {"aba": "MAIO", "linha": 7, "data_reclamacao": date(2026, 5, 1), "ticket": 1, "entrega": 2, "cliente": "A",
         "acareacao": "Acareação Motorista", "status": "Sumido", "procedencia": "Aguardando", "motorista": "M",
         "embarcador": "E", "detalhe": ""}
    assert main.validar([r]) == ([], ["MAIO linha 7: status"])


def test_agrupa_grafias_parecidas_na_mais_frequente(planilha):
    g = {(x["campo"], x["original"]): x["sugerido"] for x in main.agrupar_nomes(main.ler_planilha(planilha))}
    assert g[("motorista", "Motorista Graca")] == g[("motorista", "Motorista Graça")] == "Motorista Graça"
    assert g[("embarcador", "DAFITTI")] == "DAFITI"
    assert g[("motorista", "Motorista Beta")] == "Motorista Beta"


def test_csv_ida_e_volta_ignora_sugestao_vazia(tmp_path):
    linhas = [
        {"campo": "motorista", "original": "Motorista Graca", "sugerido": "Motorista Graça", "ocorrencias": 1},
        {"campo": "motorista", "original": "Motorista Beta", "sugerido": "", "ocorrencias": 2},
    ]
    main.escrever_csv(linhas, tmp_path / "n.csv")
    assert main.ler_csv(tmp_path / "n.csv") == {("motorista", "Motorista Graca"): "Motorista Graça"}


def test_aplicar_sem_csv_para(planilha, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    with pytest.raises(SystemExit, match="Rode primeiro sem --aplicar"):
        main.importar(planilha, aplicar=True)


def test_aplicar_com_erro_nao_abre_o_banco(planilha, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    main.importar(planilha, aplicar=False)
    ruim = {**main.ler_planilha(planilha)[0], "status": "Sumido"}
    monkeypatch.setattr(main, "ler_planilha", lambda c: [ruim])
    monkeypatch.setattr(main, "conectar", lambda: pytest.fail("não deveria abrir o banco"))
    with pytest.raises(SystemExit, match="Nada foi gravado"):
        main.importar(planilha, aplicar=True)


def test_aplicar_recusa_tabela_com_registros(planilha, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    main.importar(planilha, aplicar=False)
    conn = Conn(ja_tem=3)
    monkeypatch.setattr(main, "conectar", lambda: conn)
    with pytest.raises(SystemExit, match="já tem registros"):
        main.importar(planilha, aplicar=True)
    assert conn.gravados == []


def test_aplicar_grava_tudo_com_nomes_unificados(planilha, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    main.importar(planilha, aplicar=False)
    assert (tmp_path / "nomes.csv").exists()
    conn = Conn()
    monkeypatch.setattr(main, "conectar", lambda: conn)
    main.importar(planilha, aplicar=True)
    assert len(conn.gravados) == 6
    assert {g["motorista"] for g in conn.gravados} == {"Motorista Alfa", "Motorista Graça", "Motorista Beta"}
    assert {g["embarcador"] for g in conn.gravados} >= {"DAFITI"} and "DAFITTI" not in {g["embarcador"] for g in conn.gravados}
    assert all(g["criado_por"] == "importacao" for g in conn.gravados)


@pytest.mark.skipif(not REAL.exists(), reason="planilha real fora da pasta (dados pessoais, fora do git)")
def test_planilha_real_importa_os_956_registros():
    validos, erros = main.validar(main.ler_planilha(REAL))
    assert erros == [] and len(validos) == 956
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_main.py -v`
Expected: FAIL — `AttributeError: module 'main' has no attribute 'ler_planilha'`

- [ ] **Step 3: Implementar** — `main.py`

```python
"""Ferramentas de linha de comando do dashboard de acareações.

  python main.py importar "ACAREAÇÕES HRAO.xlsx"            gera nomes.csv para revisão (não grava nada)
  python main.py importar "ACAREAÇÕES HRAO.xlsx" --aplicar  grava no banco usando nomes.csv
  python main.py criar-admin nome.do.admin                   cria um administrador (pede a senha)
"""

import argparse
import csv
import getpass
import re
import unicodedata
from collections import Counter
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path

import openpyxl
from fastapi import HTTPException
from pydantic import ValidationError

from api.index import INSERIR, Acareacao, NovoUsuario, conectar, criar_usuario

ANO = 2026
NOMES = Path("nomes.csv")
MESES = {"janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4, "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
         "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12}
PROCEDENCIAS = {"procedente": "Procedente", "improcedente": "Improcedente", "analizando": "Analisando",
                "analisando": "Analisando", "aguardando": "Aguardando", "sem retorno": "Sem retorno"}
SEMELHANCA = 0.9  # grafias 90%+ parecidas (sem acento/caixa/espaços) viram uma sugestão de grupo


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
                if SequenceMatcher(None, chave(nome), chave(grupo[0])).ratio() >= SEMELHANCA:
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


def importar(caminho, aplicar: bool) -> None:
    registros = ler_planilha(caminho)
    if not aplicar:
        escrever_csv(agrupar_nomes(registros))
        _, erros = validar(registros)
        print(f"{len(registros)} registros lidos, {len(erros)} com erro.", *erros, sep="\n")
        print(f"Revise {NOMES} (coluna 'sugerido') e rode de novo com --aplicar.")
        return
    if not NOMES.exists():
        raise SystemExit(f"{NOMES} não encontrado. Rode primeiro sem --aplicar.")
    validos, erros = validar(aplicar_nomes(registros, ler_csv()))
    if erros:
        raise SystemExit("\n".join(erros) + f"\n{len(erros)} linha(s) com erro. Nada foi gravado.")
    with conectar() as conn:
        if conn.execute("select count(*) from acareacoes").fetchone()["count"]:
            raise SystemExit("A tabela já tem registros. Importação cancelada para não duplicar.")
        conn.cursor().executemany(INSERIR, [{**v, "criado_por": "importacao"} for v in validos])
    print(f"{len(validos)} registros importados.")


def criar_admin(usuario: str) -> None:
    senha = getpass.getpass("Senha (mínimo 8 caracteres): ")
    try:
        criar_usuario(NovoUsuario(usuario=usuario, senha=senha, admin=True))
    except ValidationError:
        raise SystemExit("Usuário: 3 a 40 letras minúsculas, números, ponto, _ ou -. Senha: 8+ caracteres.")
    except HTTPException as e:
        raise SystemExit(e.detail)
    print(f"Administrador {usuario} criado.")


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
        importar(args.planilha, args.aplicar)
    else:
        criar_admin(args.usuario)


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Rodar e ver passar**

Run: `python -m pytest tests/ -v`
Expected: todos passam (o teste da planilha real roda porque o arquivo está na pasta; 956 registros, nenhum erro); `test_crud_no_banco` pulado.

- [ ] **Step 5: Conferir os agrupamentos na planilha real (sem gravar)**

Run: `python -c "import main; [print(l) for l in main.agrupar_nomes(main.ler_planilha('ACAREAÇÕES HRAO.xlsx')) if l['original'] != l['sugerido']]"`
Expected: só pares claramente iguais (ex.: `Motorista Graca → Motorista Graça`, `DAFITTI → DAFITI`, `Fulano Mecchia ↔ Fulano Mechia`). Se aparecer junção errada de pessoas diferentes, subir `SEMELHANCA` (ex.: 0.92) e rerodar Steps 4–5. A revisão final é humana, no `nomes.csv`.

- [ ] **Step 6: Commit**

```bash
git add main.py tests/test_main.py
git commit -m "feat: importação do histórico com unificação de nomes e criação de admin"
```

---

### Task 4: Frontend — projeto e cálculos (`dados.ts`)

**Files:**
- Create: `frontend/package.json`, `frontend/tsconfig.json`, `frontend/vite.config.ts`, `frontend/src/dados.ts`
- Test: `frontend/src/dados.test.ts`

**Interfaces:**
- Consumes: rotas da Task 2.
- Produces (todas exportadas de `dados.ts`): `DOMINIO`, `ACAREACOES`, `STATUS`, `PROCEDENCIAS`; tipos `Registro`, `Campos`, `Usuario`, `Escala`, `Periodo{escala, ref:"AAAA-MM"}`, `Intervalo{de, ate, meses}`, `Situacao`, `LinhaRanking`, `Kpis`, `Item{rotulo, valor}`, `Filtros`; `supabase()`, `listar()`, `salvar(campos, id?)`, `excluir(id)`, `exportar(ids)`, `listarUsuarios()`, `criarUsuario({usuario, senha, admin})`, `removerUsuario(id)`; `fmtNum`, `fmtPct`, `fmtData`, `fmtMes(ym, longo?)`, `normalizar`, `mesDe`, `somaMes`, `hoje`, `intervalo(p, atual, primeiro)`, `anterior(p)`, `navegar(p, dir)`, `rotuloPeriodo(p)`, `noIntervalo(rs, i)`, `situacao(total, meses)`, `ranking(rs, meses)`, `kpis(rs, meses)`, `contar(rs, campo, limite?)`, `porMes(rs, ate, n?)`, `SEM_FILTRO`, `filtrar(rs, f)`, `unicos(rs, campo)`, `duplicado(rs, ticket, id?)`.

- [ ] **Step 1: Criar `frontend/package.json`**

```json
{
  "name": "dashboard-acareacoes",
  "private": true,
  "version": "1.0.0",
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "tsc -b && vite build",
    "test": "vitest run"
  },
  "dependencies": {
    "@phosphor-icons/react": "^2.1.10",
    "@supabase/supabase-js": "^2.100.0",
    "react": "18.3.1",
    "react-dom": "18.3.1"
  },
  "devDependencies": {
    "@types/react": "^18.3.12",
    "@types/react-dom": "^18.3.1",
    "@vitejs/plugin-react": "^4.3.4",
    "typescript": "~5.6.3",
    "vite": "^6.4.4",
    "vitest": "^3.2.7"
  }
}
```

- [ ] **Step 2: Criar `frontend/tsconfig.json`**

```json
{
  "compilerOptions": {
    "target": "ES2020",
    "lib": ["ES2020", "DOM", "DOM.Iterable"],
    "module": "ESNext",
    "moduleResolution": "bundler",
    "jsx": "react-jsx",
    "strict": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "skipLibCheck": true,
    "noEmit": true,
    "isolatedModules": true,
    "types": ["vite/client"]
  },
  "include": ["src", "vite.config.ts"]
}
```

- [ ] **Step 3: Criar `frontend/vite.config.ts`**

```ts
/// <reference types="vitest/config" />
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  envDir: "..", // um .env só, na raiz; o Vite expõe ao navegador apenas as variáveis VITE_*
  server: { proxy: { "/api": "http://127.0.0.1:8000" } },
  test: { environment: "node" },
});
```

- [ ] **Step 4: Instalar dependências**

Run: `npm install --prefix frontend`
Expected: cria `frontend/node_modules` e `frontend/package-lock.json` sem erros.

- [ ] **Step 5: Escrever os testes que falham** — `frontend/src/dados.test.ts`

```ts
import { describe, expect, it } from "vitest";
import {
  anterior, contar, duplicado, filtrar, intervalo, kpis, navegar, porMes, ranking, rotuloPeriodo,
  SEM_FILTRO, situacao, somaMes, type Registro,
} from "./dados";

let seq = 0;
const reg = (p: Partial<Registro> = {}): Registro => ({
  id: ++seq, data_reclamacao: "2026-05-10", ticket: 1000 + seq, entrega: 2000 + seq, cliente: "CLIENTE",
  acareacao: "Acareação Motorista", status: "Recebido", procedencia: "Improcedente", motorista: "Motorista Padrão",
  embarcador: "MAGAZINE LUIZA", detalhe: "", criado_por: "teste", criado_em: "2026-05-10T00:00:00Z", ...p,
});

describe("situação", () => {
  it("segue a tabela da planilha num mês", () => {
    expect([0, 1, 2, 3, 4, 5, 6, 16].map((t) => situacao(t, 1))).toEqual(
      ["Ótimo", "Ótimo", "Bom", "Bom", "Regular", "Regular", "Atenção", "Atenção"]);
  });
  it("usa a média mensal em períodos maiores", () => {
    expect(situacao(12, 3)).toBe("Regular");
    expect(situacao(6, 3)).toBe("Bom");
  });
});

describe("período", () => {
  it("soma meses atravessando o ano", () => {
    expect(somaMes("2026-01", -1)).toBe("2025-12");
    expect(somaMes("2025-12", 1)).toBe("2026-01");
  });
  it("calcula intervalos e meses decorridos", () => {
    const i = (escala: "mes" | "trimestre" | "ano" | "tudo", ref: string) => intervalo({ escala, ref }, "2026-10", "2026-01");
    expect(i("mes", "2026-05")).toEqual({ de: "2026-05", ate: "2026-05", meses: 1 });
    expect(i("trimestre", "2026-05")).toEqual({ de: "2026-04", ate: "2026-06", meses: 3 });
    expect(i("ano", "2026-05")).toEqual({ de: "2026-01", ate: "2026-12", meses: 10 });
    expect(i("tudo", "2026-05")).toEqual({ de: "2026-01", ate: "2026-10", meses: 10 });
    expect(i("trimestre", "2026-10").meses).toBe(1);
  });
  it("anda e acha o período anterior", () => {
    expect(anterior({ escala: "trimestre", ref: "2026-05" })).toEqual({ escala: "trimestre", ref: "2026-02" });
    expect(anterior({ escala: "tudo", ref: "2026-05" })).toBeNull();
    expect(navegar({ escala: "ano", ref: "2026-05" }, 1)).toEqual({ escala: "ano", ref: "2027-05" });
  });
  it("rotula", () => {
    expect(rotuloPeriodo({ escala: "mes", ref: "2026-05" })).toBe("Maio de 2026");
    expect(rotuloPeriodo({ escala: "trimestre", ref: "2026-05" })).toBe("2º trimestre de 2026");
    expect(rotuloPeriodo({ escala: "ano", ref: "2026-05" })).toBe("2026");
  });
});

describe("indicadores", () => {
  const rs = [
    ...Array.from({ length: 6 }, () => reg({ motorista: "Motorista A" })),
    reg({ motorista: "Motorista B", procedencia: "Procedente", status: "Indenizado" }),
    reg({ motorista: "Motorista B", procedencia: "Aguardando" }),
    reg({ motorista: "Motorista C" }),
  ];
  it("ranking soma por motorista e ordena por total", () => {
    expect(ranking(rs, 1).map((l) => [l.motorista, l.improcedente, l.procedente, l.outras, l.total, l.situacao])).toEqual([
      ["Motorista A", 6, 0, 0, 6, "Atenção"],
      ["Motorista B", 0, 1, 1, 2, "Bom"],
      ["Motorista C", 1, 0, 0, 1, "Ótimo"],
    ]);
  });
  it("kpis", () => {
    expect(kpis(rs, 1)).toEqual({ total: 9, procedentes: 1 / 9, indenizados: 1, atencao: 1 });
    expect(kpis([], 1).procedentes).toBe(0);
  });
  it("contar junta o resto em Outros", () => {
    const e = ["A", "A", "A", "B", "B", "C", "D", "E"].map((embarcador) => reg({ embarcador }));
    expect(contar(e, "embarcador", 2)).toEqual([
      { rotulo: "A", valor: 3 }, { rotulo: "B", valor: 2 }, { rotulo: "Outros (3)", valor: 3 },
    ]);
  });
  it("porMes corta meses vazios do início", () => {
    const m = porMes([reg({ data_reclamacao: "2026-03-02" }), reg({ data_reclamacao: "2026-05-01" })], "2026-05");
    expect(m).toEqual([{ rotulo: "2026-03", valor: 1 }, { rotulo: "2026-04", valor: 0 }, { rotulo: "2026-05", valor: 1 }]);
  });
});

describe("registros", () => {
  const rs = [reg({ cliente: "JOSÉ", ticket: 1000007 }), reg({ cliente: "ANA", status: "Parado", data_reclamacao: "2026-04-01" })];
  it("busca ignora acento e caixa e combina filtros", () => {
    expect(filtrar(rs, { ...SEM_FILTRO, busca: "jose" })).toHaveLength(1);
    expect(filtrar(rs, { ...SEM_FILTRO, busca: "000007" })[0].cliente).toBe("JOSÉ");
    expect(filtrar(rs, { ...SEM_FILTRO, mes: "2026-04", status: "Parado" })[0].cliente).toBe("ANA");
    expect(filtrar(rs, { ...SEM_FILTRO, mes: "2026-04", status: "Recebido" })).toHaveLength(0);
  });
  it("acha ticket duplicado ignorando o próprio registro", () => {
    expect(duplicado(rs, 1000007)?.cliente).toBe("JOSÉ");
    expect(duplicado(rs, 1000007, rs[0].id)).toBeUndefined();
  });
});
```

- [ ] **Step 6: Rodar e ver falhar**

Run: `npm test --prefix frontend`
Expected: FAIL — `Failed to resolve import "./dados"`

- [ ] **Step 7: Implementar** — `frontend/src/dados.ts`

```ts
import { createClient, type SupabaseClient } from "@supabase/supabase-js";

export const DOMINIO = "acareacoes.local"; // usuário "fulano" entra como fulano@acareacoes.local
export const ACAREACOES = ["Acareação Motorista", "Whatsapp", "Whats Motorista", "Recebido"] as const;
export const STATUS = ["Aguardando", "Recebido", "Indenizado", "Devolvido", "Parado"] as const;
export const PROCEDENCIAS = ["Improcedente", "Procedente", "Analisando", "Aguardando", "Sem retorno"] as const;

export type Registro = {
  id: number; data_reclamacao: string; ticket: number; entrega: number; cliente: string; acareacao: string;
  status: string; procedencia: string; motorista: string; embarcador: string; detalhe: string;
  criado_por: string; criado_em: string;
};
export type Campos = Omit<Registro, "id" | "criado_por" | "criado_em">;
export type Usuario = { id: string; usuario: string; admin: boolean; criado_em: string | null };

// ---- Servidor ----

let cliente: SupabaseClient | undefined;
export const supabase = () =>
  (cliente ??= createClient(import.meta.env.VITE_SUPABASE_URL, import.meta.env.VITE_SUPABASE_ANON_KEY));

async function api(caminho: string, init: RequestInit = {}): Promise<Response> {
  const { data } = await supabase().auth.getSession();
  const r = await fetch(`/api${caminho}`, {
    ...init,
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${data.session?.access_token ?? ""}` },
  });
  if (!r.ok) throw new Error((await r.json().catch(() => null))?.detail ?? `Erro ${r.status} no servidor.`);
  return r;
}

export const listar = async (): Promise<Registro[]> => (await api("/acareacoes")).json();
export const salvar = (c: Campos, id?: number) =>
  api(id ? `/acareacoes/${id}` : "/acareacoes", { method: id ? "PUT" : "POST", body: JSON.stringify(c) });
export const excluir = (id: number) => api(`/acareacoes/${id}`, { method: "DELETE" });
export const listarUsuarios = async (): Promise<Usuario[]> => (await api("/usuarios")).json();
export const criarUsuario = (u: { usuario: string; senha: string; admin: boolean }) =>
  api("/usuarios", { method: "POST", body: JSON.stringify(u) });
export const removerUsuario = (id: string) => api(`/usuarios/${id}`, { method: "DELETE" });

export async function exportar(ids: number[]): Promise<void> {
  const r = await api("/exportar", { method: "POST", body: JSON.stringify({ ids }) });
  const nome = /filename="(.+)"/.exec(r.headers.get("Content-Disposition") ?? "")?.[1] ?? "acareacoes.xlsx";
  const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(await r.blob()), download: nome });
  a.click();
  URL.revokeObjectURL(a.href);
}

// ---- Formatação ----

const MESES = ["jan.", "fev.", "mar.", "abr.", "maio", "jun.", "jul.", "ago.", "set.", "out.", "nov.", "dez."];
const MESES_LONGOS = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro",
  "Outubro", "Novembro", "Dezembro"];

export const fmtNum = (n: number) => n.toLocaleString("pt-BR", { maximumFractionDigits: 1 });
export const fmtPct = (fracao: number) => `${fmtNum(fracao * 100)}%`;
export const fmtData = (iso: string) => iso.slice(0, 10).split("-").reverse().join("/");
export function fmtMes(ym: string, longo = false): string {
  const [a, m] = ym.split("-").map(Number);
  return longo ? `${MESES_LONGOS[m - 1]} de ${a}` : `${MESES[m - 1]} ${a}`;
}
export const normalizar = (s: string) => s.normalize("NFD").replace(/\p{Diacritic}/gu, "").toLowerCase();
export const hoje = () => new Date().toLocaleDateString("sv-SE"); // AAAA-MM-DD no fuso local

// ---- Período ----

export type Escala = "mes" | "trimestre" | "ano" | "tudo";
export type Periodo = { escala: Escala; ref: string }; // ref = mês âncora "AAAA-MM"
export type Intervalo = { de: string; ate: string; meses: number };

const PASSO: Record<Escala, number> = { mes: 1, trimestre: 3, ano: 12, tudo: 0 };

export const mesDe = (iso: string) => iso.slice(0, 7);
export function somaMes(ym: string, n: number): string {
  const [a, m] = ym.split("-").map(Number);
  const t = a * 12 + (m - 1) + n;
  return `${Math.floor(t / 12)}-${String((t % 12) + 1).padStart(2, "0")}`;
}
function distancia(de: string, ate: string): number {
  const [a1, m1] = de.split("-").map(Number);
  const [a2, m2] = ate.split("-").map(Number);
  return (a2 - a1) * 12 + (m2 - m1);
}

export function intervalo(p: Periodo, atual: string, primeiro: string): Intervalo {
  const [a, m] = p.ref.split("-").map(Number);
  const deslocamento = (m - 1) % 3;
  const [de, ate] =
    p.escala === "mes" ? [p.ref, p.ref]
    : p.escala === "trimestre" ? [somaMes(p.ref, -deslocamento), somaMes(p.ref, 2 - deslocamento)]
    : p.escala === "ano" ? [`${a}-01`, `${a}-12`]
    : [primeiro, atual];
  // Trimestre/ano em andamento: a média mensal não conta meses que ainda não aconteceram.
  const fim = ate < atual ? ate : atual;
  return { de, ate, meses: Math.max(1, distancia(de, fim) + 1) };
}

export const anterior = (p: Periodo): Periodo | null =>
  PASSO[p.escala] ? { escala: p.escala, ref: somaMes(p.ref, -PASSO[p.escala]) } : null;
export const navegar = (p: Periodo, dir: 1 | -1): Periodo => ({ ...p, ref: somaMes(p.ref, dir * PASSO[p.escala]) });

export function rotuloPeriodo(p: Periodo): string {
  const [a, m] = p.ref.split("-").map(Number);
  if (p.escala === "mes") return fmtMes(p.ref, true);
  if (p.escala === "trimestre") return `${Math.floor((m - 1) / 3) + 1}º trimestre de ${a}`;
  if (p.escala === "ano") return String(a);
  return "Todo o histórico";
}

export const noIntervalo = (rs: Registro[], i: Intervalo) =>
  rs.filter((r) => mesDe(r.data_reclamacao) >= i.de && mesDe(r.data_reclamacao) <= i.ate);

// ---- Indicadores ----

export type Situacao = "Ótimo" | "Bom" | "Regular" | "Atenção";
export function situacao(total: number, meses: number): Situacao {
  const media = total / meses; // regra da aba RESULTADO, por mês
  return media <= 1 ? "Ótimo" : media <= 3 ? "Bom" : media <= 5 ? "Regular" : "Atenção";
}

export type LinhaRanking = {
  motorista: string; improcedente: number; procedente: number; outras: number; total: number; situacao: Situacao;
};
export function ranking(rs: Registro[], meses: number): LinhaRanking[] {
  const por = new Map<string, LinhaRanking>();
  for (const r of rs) {
    const l = por.get(r.motorista) ?? { motorista: r.motorista, improcedente: 0, procedente: 0, outras: 0, total: 0, situacao: "Ótimo" };
    if (r.procedencia === "Improcedente") l.improcedente++;
    else if (r.procedencia === "Procedente") l.procedente++;
    else l.outras++;
    l.total++;
    por.set(r.motorista, l);
  }
  return [...por.values()]
    .map((l) => ({ ...l, situacao: situacao(l.total, meses) }))
    .sort((a, b) => b.total - a.total || a.motorista.localeCompare(b.motorista, "pt-BR"));
}

export type Kpis = { total: number; procedentes: number; indenizados: number; atencao: number };
export function kpis(rs: Registro[], meses: number): Kpis {
  const total = rs.length;
  return {
    total,
    procedentes: total ? rs.filter((r) => r.procedencia === "Procedente").length / total : 0,
    indenizados: rs.filter((r) => r.status === "Indenizado").length,
    atencao: ranking(rs, meses).filter((l) => l.situacao === "Atenção").length,
  };
}

export type Item = { rotulo: string; valor: number };
export function contar(rs: Registro[], campo: "status" | "embarcador" | "procedencia", limite = Infinity): Item[] {
  const c = new Map<string, number>();
  for (const r of rs) c.set(r[campo], (c.get(r[campo]) ?? 0) + 1);
  const itens = [...c]
    .map(([rotulo, valor]) => ({ rotulo, valor }))
    .sort((a, b) => b.valor - a.valor || a.rotulo.localeCompare(b.rotulo, "pt-BR"));
  if (itens.length <= limite) return itens;
  const resto = itens.slice(limite);
  return [...itens.slice(0, limite), { rotulo: `Outros (${resto.length})`, valor: resto.reduce((s, i) => s + i.valor, 0) }];
}

export function porMes(rs: Registro[], ate: string, n = 12): Item[] {
  const c = new Map<string, number>();
  for (const r of rs) c.set(mesDe(r.data_reclamacao), (c.get(mesDe(r.data_reclamacao)) ?? 0) + 1);
  const meses = Array.from({ length: n }, (_, i) => somaMes(ate, i - n + 1)).map((ym) => ({ rotulo: ym, valor: c.get(ym) ?? 0 }));
  const primeiro = meses.findIndex((m) => m.valor > 0);
  return primeiro < 0 ? meses.slice(-1) : meses.slice(primeiro);
}

// ---- Registros ----

export type Filtros = { busca: string; mes: string; motorista: string; embarcador: string; status: string; procedencia: string };
export const SEM_FILTRO: Filtros = { busca: "", mes: "", motorista: "", embarcador: "", status: "", procedencia: "" };

export function filtrar(rs: Registro[], f: Filtros): Registro[] {
  const b = normalizar(f.busca.trim());
  return rs.filter((r) =>
    (!b || String(r.ticket).includes(b) || String(r.entrega).includes(b) || normalizar(r.cliente).includes(b)) &&
    (!f.mes || mesDe(r.data_reclamacao) === f.mes) &&
    (!f.motorista || r.motorista === f.motorista) &&
    (!f.embarcador || r.embarcador === f.embarcador) &&
    (!f.status || r.status === f.status) &&
    (!f.procedencia || r.procedencia === f.procedencia));
}

export const unicos = (rs: Registro[], campo: "motorista" | "embarcador") =>
  [...new Set(rs.map((r) => r[campo]))].sort((a, b) => a.localeCompare(b, "pt-BR"));
export const duplicado = (rs: Registro[], ticket: number, id?: number) => rs.find((r) => r.ticket === ticket && r.id !== id);
```

- [ ] **Step 8: Rodar e ver passar**

Run: `npm test --prefix frontend`
Expected: 12 passed.

- [ ] **Step 9: Commit**

```bash
git add frontend/package.json frontend/package-lock.json frontend/tsconfig.json frontend/vite.config.ts frontend/src/dados.ts frontend/src/dados.test.ts
git commit -m "feat(front): projeto Vite e cálculos do painel

```

---

### Task 5: Visual base, login e menu (`styles.css`, `App.tsx`)

**Files:**
- Create: `frontend/index.html`, `frontend/src/main.tsx`, `frontend/src/styles.css`, `frontend/src/App.tsx`
- Create (provisórios, substituídos nas Tasks 6–9): `frontend/src/Painel.tsx`, `frontend/src/Registros.tsx`, `frontend/src/Usuarios.tsx`, `frontend/src/Formulario.tsx`

**Interfaces:**
- Consumes: `dados.ts` (Task 4).
- Produces: `export type Acoes = { registros: Registro[]; abrirForm: (r?: Registro) => void; recarregar: () => Promise<void> }` em `App.tsx`. Páginas: `Painel(props: Acoes)`, `Registros(props: Acoes)`, `Usuarios({ eu }: { eu: string })`, `Formulario({ aberto, registro, registros, onFechar, onSalvo })`. Classes CSS usadas pelas próximas tasks (todas definidas aqui): `cabecalho acoes botao primario botao-icone link perigo periodo seg passo grade kpis span-2 card kpi kpi-valor kpi-var melhor pior legenda rank rank-pos rank-nome trilho s1 s2 s3 sit ot bo re at colunas barra on bh bh-trilho tabela-rolagem tabela detalhe pill acoes-linha barra-tabela busca filtro ativo campo inteiro aviso-dup sheet sheet-topo sheet-corpo sheet-rodape caixa caixa-acoes estado erro ok faixa-erro form-usuario marcar sr-only num depois`.

- [ ] **Step 1: Criar `frontend/index.html`**"
```html
<!doctype html>
<html lang="pt-BR">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Acareações</title>
    <script>
      (function () {
        var t = null;
        try { t = localStorage.getItem("tema"); } catch (e) {}
        if (t !== "claro" && t !== "escuro") t = matchMedia("(prefers-color-scheme: dark)").matches ? "escuro" : "claro";
        document.documentElement.dataset.theme = t;
      })();
    </script>
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/src/main.tsx"></script>
  </body>
</html>
```

- [ ] **Step 2: Criar `frontend/src/main.tsx`**

```tsx
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import "./styles.css";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
```

- [ ] **Step 3: Criar `frontend/src/styles.css`**

```css
/* Tokens — único lugar com cor. Marca Magalu; tipografia, espaço e movimento pela skill apple-design. */
:root,
[data-theme="claro"] {
  --fundo: #f5f5f7;
  --superficie: #ffffff;
  --superficie-2: #f2f2f5; /* campos e trilhos */
  --borda: rgba(0, 0, 0, 0.08);
  --texto: #1d1d1f;
  --secundario: #6e6e73;
  --material: rgba(236, 236, 240, 0.8); /* menu lateral */
  --acento: #0086ff; /* azul Magalu: barras e destaques */
  --acento-texto: #006fd6; /* links e item ativo */
  --acento-suave: rgba(0, 134, 255, 0.1);
  --botao: #006fd6; /* branco sobre ele: 4,95:1 (o #0086ff dá 3,59:1) */
  --serie-1: #0086ff; /* Improcedente */
  --serie-2: #e6007e; /* Procedente — rosa da faixa Magalu */
  --serie-3: #c7c7cc; /* Em análise */
  --coluna: #d2d2d7;
  --ok: #1e7f37;
  --ok-bg: #e6f4ea;
  --neutro: #48484a;
  --neutro-bg: #efeff2;
  --aviso: #9a5200;
  --aviso-bg: #fff2de;
  --critico: #c5162e;
  --critico-bg: #fde7ea;
  --sombra: 0 1px 2px rgba(0, 0, 0, 0.04), 0 4px 16px rgba(0, 0, 0, 0.04);
  --sombra-alta: 0 24px 80px rgba(0, 0, 0, 0.25);
  color-scheme: light;
}

[data-theme="escuro"] {
  --fundo: #000000;
  --superficie: #1c1c1e;
  --superficie-2: #2c2c2e;
  --borda: rgba(255, 255, 255, 0.1);
  --texto: #f5f5f7;
  --secundario: #98989d;
  --material: rgba(28, 28, 30, 0.72);
  --acento: #1a8cff;
  --acento-texto: #4da3ff;
  --acento-suave: rgba(26, 140, 255, 0.16);
  --botao: #006fd6;
  --serie-1: #1a8cff;
  --serie-2: #e8338c;
  --serie-3: #48484a;
  --coluna: #3a3a3c;
  --ok: #4cd368;
  --ok-bg: rgba(48, 209, 88, 0.16);
  --neutro: #c7c7cc;
  --neutro-bg: rgba(255, 255, 255, 0.1);
  --aviso: #ffb340;
  --aviso-bg: rgba(255, 159, 10, 0.16);
  --critico: #ff6961;
  --critico-bg: rgba(255, 69, 58, 0.16);
  --sombra: 0 0 0 1px rgba(255, 255, 255, 0.06);
  --sombra-alta: 0 24px 80px rgba(0, 0, 0, 0.6);
  color-scheme: dark;
}

*,
*::before,
*::after {
  box-sizing: border-box;
}

body {
  margin: 0;
  background: var(--fundo);
  color: var(--texto);
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif;
  font-size: 0.875rem;
  line-height: 1.5;
  -webkit-font-smoothing: antialiased;
}

h1,
h2,
p {
  margin: 0;
}

a {
  color: var(--acento-texto);
}

button,
input,
select,
textarea {
  font: inherit;
  color: inherit;
}

:focus-visible {
  outline: 2px solid var(--acento);
  outline-offset: 2px;
}

.num {
  font-variant-numeric: tabular-nums;
}

.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  overflow: hidden;
  clip: rect(0 0 0 0);
  white-space: nowrap;
}

/* Layout */
.app {
  display: grid;
  grid-template-columns: 14.5rem minmax(0, 1fr);
  min-height: 100vh;
}

.lateral {
  position: sticky;
  top: 0;
  height: 100vh;
  display: flex;
  flex-direction: column;
  gap: 1.75rem;
  padding: 1.5rem 0.875rem;
  background: var(--material);
  -webkit-backdrop-filter: blur(30px) saturate(180%);
  backdrop-filter: blur(30px) saturate(180%);
  border-right: 1px solid var(--borda);
}

.marca {
  display: grid;
  padding: 0 0.625rem;
}

.marca strong {
  font-size: 1.1875rem;
  font-weight: 700;
  letter-spacing: -0.025em;
  line-height: 1.15;
}

.marca span {
  color: var(--secundario);
  font-size: 0.75rem;
}

/* Assinatura Magalu: as 7 cores da faixa, iguais nos dois temas. */
.faixa {
  height: 3px;
  margin-top: 0.625rem;
  border-radius: 3px;
  background: linear-gradient(90deg, #ffc800 0 14.3%, #ff8a00 0 28.6%, #f0263c 0 42.9%, #e6007e 0 57.1%, #8e36c4 0 71.4%, #0086ff 0 85.7%, #2fb34a 0);
}

.lateral ul {
  list-style: none;
  margin: 0;
  padding: 0;
  display: grid;
  gap: 2px;
}

.lateral nav a {
  display: flex;
  align-items: center;
  gap: 0.625rem;
  padding: 0.5rem 0.625rem;
  border-radius: 8px;
  color: var(--texto);
  text-decoration: none;
  font-size: 0.84375rem;
  transition: background-color 150ms;
}

.lateral nav a svg {
  color: var(--secundario);
}

.lateral nav a:hover,
.lateral nav a[aria-current="page"] {
  background: var(--acento-suave);
}

.lateral nav a[aria-current="page"],
.lateral nav a[aria-current="page"] svg {
  color: var(--acento-texto);
  font-weight: 600;
}

.rodape {
  margin-top: auto;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 0.5rem 0.625rem;
  padding: 0.75rem 0.625rem 0;
  border-top: 1px solid var(--borda);
  font-size: 0.78125rem;
}

.avatar {
  display: grid;
  place-items: center;
  width: 1.75rem;
  height: 1.75rem;
  border-radius: 50%;
  background: var(--botao);
  color: #fff;
  font-size: 0.75rem;
  font-weight: 600;
}

.rodape-nome {
  flex: 1;
  min-width: 0;
  display: grid;
  line-height: 1.25;
  overflow-wrap: anywhere;
}

.rodape-nome small {
  color: var(--secundario);
}

.rodape-acoes {
  display: flex;
  gap: 0.375rem;
  width: 100%;
}

.principal {
  min-width: 0;
  max-width: 90rem;
  padding: 1.75rem 2rem 2.5rem;
}

.cabecalho {
  display: flex;
  flex-wrap: wrap;
  align-items: flex-end;
  justify-content: space-between;
  gap: 1rem;
  margin-bottom: 1.25rem;
}

.cabecalho h1 {
  font-size: 2.125rem;
  font-weight: 700;
  letter-spacing: -0.022em;
  line-height: 1.1;
}

.cabecalho p {
  margin-top: 0.25rem;
  color: var(--secundario);
}

.acoes {
  display: flex;
  flex-wrap: wrap;
  gap: 0.5rem;
}

/* Botões */
.botao,
.botao-icone {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 0.4375rem;
  height: 2.25rem;
  border: 1px solid var(--borda);
  border-radius: 9px;
  background: var(--superficie);
  font-size: 0.8125rem;
  font-weight: 600;
  cursor: pointer;
  transition: transform 100ms ease-out, background-color 150ms;
}

.botao {
  padding: 0 0.875rem;
}

.botao-icone {
  flex: none;
  width: 2rem;
  height: 2rem;
  color: var(--secundario);
}

.botao:hover,
.botao-icone:hover {
  background: var(--superficie-2);
}

.botao.primario {
  border-color: transparent;
  background: var(--botao);
  color: #fff;
}

.botao:active:not(:disabled),
.botao-icone:active:not(:disabled) {
  transform: scale(0.97);
}

.botao:disabled,
.botao-icone:disabled {
  opacity: 0.5;
  cursor: default;
}

.link {
  padding: 0;
  border: 0;
  background: none;
  color: var(--acento-texto);
  font-size: 0.8125rem;
  font-weight: 500;
  cursor: pointer;
}

.link.perigo {
  color: var(--critico);
}

/* Período */
.periodo {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 0.75rem;
  margin-bottom: 1.25rem;
}

.seg {
  display: inline-flex;
  padding: 3px;
  border-radius: 9px;
  background: var(--neutro-bg);
}

.seg button {
  padding: 0.3125rem 0.875rem;
  border: 0;
  border-radius: 7px;
  background: none;
  font-size: 0.8125rem;
  cursor: pointer;
  transition: background-color 150ms;
}

.seg button[aria-pressed="true"] {
  background: var(--superficie);
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.12);
  font-weight: 600;
}

.passo {
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
  font-weight: 600;
}

.passo span {
  min-width: 10rem;
  text-align: center;
}

/* Grade e cards */
.grade {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 1rem;
}

.kpis {
  grid-column: 1 / -1;
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 1rem;
}

.span-2 {
  grid-column: span 2;
}

.card {
  min-width: 0;
  padding: 1.125rem 1.25rem;
  border-radius: 14px;
  background: var(--superficie);
  box-shadow: var(--sombra);
}

.card h2 {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  margin-bottom: 0.875rem;
  font-size: 0.9375rem;
  font-weight: 600;
  letter-spacing: -0.01em;
}

.card h2 a {
  margin-left: auto;
  font-size: 0.78125rem;
  font-weight: 500;
  text-decoration: none;
}

.card h2.depois {
  margin-top: 1.5rem;
}

.estado {
  padding: 1.5rem 0;
  color: var(--secundario);
  font-size: 0.8125rem;
}

.erro,
.ok {
  font-size: 0.8125rem;
  font-weight: 500;
}

.erro {
  color: var(--critico);
}

.ok {
  color: var(--ok);
}

.faixa-erro {
  margin-bottom: 1rem;
  padding: 0.625rem 0.875rem;
  border-radius: 10px;
  background: var(--critico-bg);
}

/* KPI */
.kpi h2 {
  margin: 0;
  color: var(--secundario);
  font-size: 0.78125rem;
  font-weight: 500;
  letter-spacing: 0;
}

.kpi-valor {
  margin-top: 0.5rem;
  font-size: 2rem;
  font-weight: 700;
  letter-spacing: -0.025em;
  line-height: 1.15;
  font-variant-numeric: tabular-nums;
}

.kpi-var {
  margin-top: 0.375rem;
  color: var(--secundario);
  font-size: 0.75rem;
}

.kpi-var b {
  color: var(--texto);
  font-weight: 600;
}

.kpi-var b.melhor {
  color: var(--ok);
}

.kpi-var b.pior {
  color: var(--critico);
}

/* Legenda e ranking */
.legenda {
  display: flex;
  flex-wrap: wrap;
  gap: 1rem;
  margin: -0.25rem 0 0.75rem;
  padding: 0;
  list-style: none;
  color: var(--secundario);
  font-size: 0.75rem;
}

.legenda li {
  display: flex;
  align-items: center;
  gap: 0.375rem;
}

.legenda i {
  width: 10px;
  height: 10px;
  border-radius: 3px;
}

.rank {
  display: grid;
  gap: 0.5625rem;
  max-height: 27.5rem;
  margin: 0;
  padding: 0 0.25rem 0 0;
  overflow-y: auto;
  list-style: none;
}

.rank li {
  display: grid;
  grid-template-columns: 1.375rem minmax(6rem, 10rem) minmax(0, 1fr) 2rem 5.75rem;
  align-items: center;
  gap: 0.625rem;
  font-size: 0.8125rem;
}

.rank-pos {
  color: var(--secundario);
  font-size: 0.75rem;
  font-variant-numeric: tabular-nums;
}

.rank-nome {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.rank b {
  text-align: right;
}

.trilho {
  display: flex;
  gap: 2px;
  height: 10px;
}

.trilho span {
  min-width: 3px;
  border-radius: 2px;
}

.trilho span:last-child {
  border-radius: 2px 4px 4px 2px;
}

.s1 {
  background: var(--serie-1);
}

.s2 {
  background: var(--serie-2);
}

.s3 {
  background: var(--serie-3);
}

.sit {
  justify-self: start;
  display: inline-flex;
  align-items: center;
  gap: 0.3125rem;
  padding: 0.125rem 0.5rem;
  border-radius: 999px;
  font-size: 0.6875rem;
  font-weight: 600;
}

.sit::before {
  content: "";
  width: 6px;
  height: 6px;
  border-radius: 50%;
  background: currentColor;
}

.sit.ot {
  background: var(--ok-bg);
  color: var(--ok);
}

.sit.bo {
  background: var(--neutro-bg);
  color: var(--neutro);
}

.sit.re {
  background: var(--aviso-bg);
  color: var(--aviso);
}

.sit.at {
  background: var(--critico-bg);
  color: var(--critico);
}

/* Colunas por mês */
.colunas {
  display: flex;
  gap: 0.375rem;
  height: 11rem;
}

.colunas > div {
  flex: 1;
  min-width: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 0.25rem;
  color: var(--secundario);
  font-size: 0.65625rem;
}

.colunas small {
  min-height: 1em;
  font-size: 0.6875rem;
}

.colunas .barra {
  flex: 1;
  width: 100%;
  display: flex;
  align-items: flex-end;
  justify-content: center;
  border-bottom: 1px solid var(--borda);
}

.colunas .barra span {
  width: 100%;
  max-width: 1.25rem;
  min-height: 1px;
  border-radius: 4px 4px 0 0;
  background: var(--coluna);
}

.colunas .on {
  color: var(--texto);
  font-weight: 600;
}

.colunas .on .barra span {
  background: var(--acento);
}

/* Barras horizontais */
.bh {
  display: grid;
  gap: 0.6875rem;
  margin: 0;
  padding: 0;
  list-style: none;
}

.bh li {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  row-gap: 0.25rem;
  font-size: 0.78125rem;
}

.bh li > span:first-child {
  overflow: hidden;
  color: var(--secundario);
  text-overflow: ellipsis;
  white-space: nowrap;
}

.bh small {
  margin-left: 0.375rem;
  color: var(--secundario);
  font-weight: 400;
}

.bh-trilho {
  grid-column: 1 / -1;
  height: 6px;
  overflow: hidden;
  border-radius: 4px;
  background: var(--superficie-2);
}

.bh-trilho span {
  display: block;
  height: 100%;
  border-radius: 0 4px 4px 0;
  background: var(--acento);
}

/* Tabela */
.tabela-rolagem {
  overflow-x: auto;
}

.tabela {
  width: 100%;
  border-collapse: collapse;
  font-size: 0.78125rem;
}

.tabela th {
  padding: 0 0.75rem 0.5rem 0;
  color: var(--secundario);
  font-size: 0.71875rem;
  font-weight: 500;
  text-align: left;
  white-space: nowrap;
}

.tabela td {
  padding: 0.625rem 0.75rem 0.625rem 0;
  border-top: 1px solid var(--borda);
  white-space: nowrap;
}

.tabela td.detalhe {
  max-width: 16rem;
  overflow: hidden;
  color: var(--secundario);
  text-overflow: ellipsis;
}

.pill {
  padding: 0.125rem 0.5rem;
  border-radius: 999px;
  background: var(--neutro-bg);
  font-size: 0.71875rem;
}

.acoes-linha {
  text-align: right;
}

.acoes-linha .link + .link {
  margin-left: 0.75rem;
}

/* Busca e filtros da tabela */
.barra-tabela {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 0.5rem;
  margin-bottom: 0.875rem;
}

.busca {
  flex: 1 1 14rem;
  display: flex;
  align-items: center;
  gap: 0.5rem;
  height: 2.125rem;
  padding: 0 0.75rem;
  border-radius: 9px;
  background: var(--superficie-2);
  color: var(--secundario);
}

.busca:focus-within {
  outline: 2px solid var(--acento);
  outline-offset: 0;
}

.busca input {
  flex: 1;
  min-width: 0;
  border: 0;
  outline: none;
  background: none;
  color: var(--texto);
  font-size: 0.8125rem;
}

.filtro select {
  max-width: 12rem;
  height: 2.125rem;
  padding: 0 0.625rem;
  border: 1px solid var(--borda);
  border-radius: 9px;
  background: var(--superficie);
  font-size: 0.78125rem;
}

.filtro.ativo select {
  border-color: var(--acento);
  background: var(--acento-suave);
  color: var(--acento-texto);
  font-weight: 600;
}

/* Campos de formulário */
.campo {
  display: grid;
  gap: 0.3125rem;
  color: var(--secundario);
  font-size: 0.78125rem;
  font-weight: 500;
}

.campo input,
.campo select,
.campo textarea {
  width: 100%;
  min-height: 2.375rem;
  padding: 0.5rem 0.75rem;
  border: 1px solid transparent;
  border-radius: 9px;
  background: var(--superficie-2);
  color: var(--texto);
  font-size: 0.84375rem;
  font-weight: 400;
}

.campo textarea {
  resize: vertical;
}

.campo :focus {
  outline: none;
  border-color: var(--acento);
  background: var(--superficie);
  box-shadow: 0 0 0 3px var(--acento-suave);
}

.campo :user-invalid {
  border-color: var(--critico);
}

.inteiro,
.aviso-dup {
  grid-column: 1 / -1;
}

.aviso-dup {
  padding: 0.5rem 0.625rem;
  border-radius: 8px;
  background: var(--aviso-bg);
  color: var(--aviso);
  font-size: 0.75rem;
  font-weight: 500;
}

/* Painel lateral (cadastro/edição) — entra e sai pela direita, mesma curva nos dois sentidos */
.sheet {
  position: fixed;
  inset: 0.75rem 0.75rem 0.75rem auto;
  width: min(27.5rem, calc(100vw - 1.5rem));
  max-width: none;
  height: auto;
  max-height: none;
  margin: 0;
  padding: 0;
  border: 0;
  border-radius: 16px;
  background: var(--superficie);
  color: var(--texto);
  box-shadow: var(--sombra-alta);
  flex-direction: column;
  transform: translateX(calc(100% + 1.5rem));
  transition: transform 380ms cubic-bezier(0.32, 0.72, 0, 1), overlay 380ms allow-discrete, display 380ms allow-discrete;
}

.sheet[open] {
  display: flex;
  transform: none;
}

@starting-style {
  .sheet[open] {
    transform: translateX(calc(100% + 1.5rem));
  }
}

.sheet::backdrop,
.caixa::backdrop {
  background: rgba(0, 0, 0, 0.32);
  opacity: 0;
  transition: opacity 300ms, overlay 300ms allow-discrete, display 300ms allow-discrete;
}

.sheet[open]::backdrop,
.caixa[open]::backdrop {
  opacity: 1;
}

@starting-style {
  .sheet[open]::backdrop,
  .caixa[open]::backdrop {
    opacity: 0;
  }
}

.sheet-topo {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 1.25rem 1.375rem 0.5rem;
}

.sheet-topo h2 {
  font-size: 1.375rem;
  letter-spacing: -0.02em;
}

.sheet-corpo {
  flex: 1;
  display: grid;
  grid-template-columns: 1fr 1fr;
  align-content: start;
  gap: 0.875rem;
  padding: 0.5rem 1.375rem 1rem;
  overflow-y: auto;
}

.sheet-rodape {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: flex-end;
  gap: 0.5rem;
  padding: 1rem 1.375rem;
  border-top: 1px solid var(--borda);
}

.sheet-rodape .erro {
  flex-basis: 100%;
}

/* Caixa de diálogo pequena (trocar senha) */
.caixa {
  width: min(22rem, calc(100vw - 2rem));
  padding: 1.5rem;
  border: 0;
  border-radius: 16px;
  background: var(--superficie);
  color: var(--texto);
  box-shadow: var(--sombra-alta);
}

.caixa form {
  display: grid;
  gap: 1rem;
}

.caixa h2 {
  font-size: 1.25rem;
  letter-spacing: -0.015em;
}

.caixa-acoes {
  display: flex;
  justify-content: flex-end;
  gap: 0.5rem;
}

/* Login */
.login {
  display: grid;
  place-items: center;
  min-height: 100vh;
  padding: 1rem;
}

.login-cartao {
  display: grid;
  gap: 1rem;
  width: min(22rem, 100%);
  padding: 2rem 1.75rem;
}

.login-cartao .marca {
  padding: 0 0 0.5rem;
}

.login-cartao .botao {
  height: 2.5rem;
}

/* Usuários */
.form-usuario {
  display: grid;
  gap: 0.875rem;
}

.marcar {
  display: flex;
  align-items: flex-start;
  gap: 0.5rem;
  font-size: 0.8125rem;
}

/* Movimento */
@media (prefers-reduced-motion: no-preference) {
  .grade > *,
  .kpis > * {
    animation: entrar 500ms cubic-bezier(0.16, 1, 0.3, 1) both;
  }
}

@keyframes entrar {
  from {
    opacity: 0;
    transform: translateY(0.5rem);
  }
}

@media (prefers-reduced-motion: reduce) {
  .sheet {
    opacity: 0;
    transform: none;
    transition: opacity 200ms, overlay 200ms allow-discrete, display 200ms allow-discrete;
  }

  .sheet[open] {
    opacity: 1;
  }

  @starting-style {
    .sheet[open] {
      opacity: 0;
      transform: none;
    }
  }

  .botao:active:not(:disabled),
  .botao-icone:active:not(:disabled) {
    transform: none;
  }
}

@media (prefers-reduced-transparency: reduce) {
  .lateral {
    background: var(--superficie);
    -webkit-backdrop-filter: none;
    backdrop-filter: none;
  }
}

/* Telas menores */
@media (max-width: 75rem) {
  .kpis {
    grid-template-columns: repeat(2, minmax(0, 1fr));
  }
}

@media (max-width: 60rem) {
  .app {
    grid-template-columns: minmax(0, 1fr);
  }

  .lateral {
    position: static;
    height: auto;
    flex-direction: row;
    flex-wrap: wrap;
    align-items: center;
    gap: 0.75rem 1rem;
    padding: 0.75rem 1rem;
    border-right: 0;
    border-bottom: 1px solid var(--borda);
  }

  .lateral ul {
    display: flex;
  }

  .rodape {
    flex-wrap: nowrap;
    margin: 0 0 0 auto;
    padding: 0;
    border: 0;
  }

  .rodape-acoes {
    width: auto;
  }

  .principal {
    padding: 1.5rem 1rem;
  }

  .grade {
    grid-template-columns: minmax(0, 1fr);
  }

  .span-2 {
    grid-column: auto;
  }

  .rank li {
    grid-template-columns: 1.25rem minmax(5rem, 8rem) minmax(0, 1fr) 1.75rem 5rem;
  }
}

@media (max-width: 30rem) {
  .kpis,
  .sheet-corpo {
    grid-template-columns: minmax(0, 1fr);
  }

  .rodape-nome {
    display: none;
  }
}
```

- [ ] **Step 4: Conferir contraste dos textos coloridos (WCAG AA ≥ 4,5:1)**

Run:
```bash
node -e "
const L=h=>{const c=[1,3,5].map(i=>parseInt(h.substr(i,2),16)/255).map(v=>v<=.03928?v/12.92:((v+.055)/1.055)**2.4);return .2126*c[0]+.7152*c[1]+.0722*c[2]};
const cr=(a,b)=>{const x=L(a),y=L(b);return ((Math.max(x,y)+.05)/(Math.min(x,y)+.05)).toFixed(2)};
for (const [t,f] of [['#1e7f37','#e6f4ea'],['#48484a','#efeff2'],['#9a5200','#fff2de'],['#c5162e','#fde7ea'],['#6e6e73','#ffffff'],['#006fd6','#ffffff'],['#ffffff','#006fd6'],['#98989d','#1c1c1e'],['#4da3ff','#1c1c1e'],['#ff6961','#1c1c1e'],['#4cd368','#1c1c1e'],['#ffb340','#1c1c1e']]) console.log(t,'sobre',f,cr(t,f));"
```
Expected: todos ≥ 4.5. Se algum ficar abaixo, escurecer (claro) ou clarear (escuro) o token de texto até passar e atualizar o `styles.css`.

- [ ] **Step 5: Criar `frontend/src/App.tsx`**

```tsx
import { useCallback, useEffect, useRef, useState, type FormEvent, type RefObject } from "react";
import { Key, ListBullets, Moon, SignOut, SquaresFour, Sun, UsersThree, type Icon } from "@phosphor-icons/react";
import type { Session } from "@supabase/supabase-js";
import { DOMINIO, listar, supabase, type Registro } from "./dados";
import Formulario from "./Formulario";
import Painel from "./Painel";
import Registros from "./Registros";
import Usuarios from "./Usuarios";

export type Acoes = { registros: Registro[]; abrirForm: (r?: Registro) => void; recarregar: () => Promise<void> };

const PAGINAS = { painel: "Painel", registros: "Registros", usuarios: "Usuários" } as const;
type Pagina = keyof typeof PAGINAS;
const paginaAtual = (): Pagina => {
  const p = location.hash.replace(/^#\/?/, "");
  return p in PAGINAS ? (p as Pagina) : "painel";
};

export default function App() {
  const [sessao, setSessao] = useState<Session | null | undefined>(undefined);
  useEffect(() => {
    const { data } = supabase().auth.onAuthStateChange((_, s) => setSessao(s));
    return () => data.subscription.unsubscribe();
  }, []);
  if (sessao === undefined) return null;
  return sessao ? <Dashboard sessao={sessao} /> : <Login />;
}

function Marca() {
  return (
    <div className="marca">
      <strong>Acareações</strong>
      <span>HRAO · Motoristas</span>
      <div className="faixa" aria-hidden />
    </div>
  );
}

function Login() {
  const [erro, setErro] = useState("");
  const [entrando, setEntrando] = useState(false);

  async function entrar(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    const f = new FormData(e.currentTarget);
    setEntrando(true);
    setErro("");
    const { error } = await supabase().auth.signInWithPassword({
      email: `${String(f.get("usuario")).trim().toLowerCase()}@${DOMINIO}`,
      password: String(f.get("senha")),
    });
    setEntrando(false);
    if (error) setErro(error.status === 400 ? "Usuário ou senha incorretos." : "Não foi possível entrar. Tente de novo.");
  }

  return (
    <main className="login">
      <form className="card login-cartao" onSubmit={entrar}>
        <Marca />
        <label className="campo">
          Usuário
          <input name="usuario" autoComplete="username" required autoFocus />
        </label>
        <label className="campo">
          Senha
          <input name="senha" type="password" autoComplete="current-password" required />
        </label>
        {erro && <p className="erro" role="alert">{erro}</p>}
        <button className="botao primario" disabled={entrando}>{entrando ? "Entrando…" : "Entrar"}</button>
      </form>
    </main>
  );
}

function Item({ p, atual, Icone }: { p: Pagina; atual: Pagina; Icone: Icon }) {
  return (
    <li>
      <a href={`#/${p}`} aria-current={p === atual ? "page" : undefined}>
        <Icone size={18} weight={p === atual ? "fill" : "regular"} aria-hidden />
        {PAGINAS[p]}
      </a>
    </li>
  );
}

function Dashboard({ sessao }: { sessao: Session }) {
  const usuario = sessao.user.email?.split("@")[0] ?? "";
  const admin = sessao.user.app_metadata?.role === "admin";
  const [pagina, setPagina] = useState(paginaAtual);
  const [tema, setTema] = useState(() => (document.documentElement.dataset.theme === "escuro" ? "escuro" : "claro"));
  const [registros, setRegistros] = useState<Registro[]>([]);
  const [carregando, setCarregando] = useState(true);
  const [erro, setErro] = useState("");
  const [form, setForm] = useState<{ aberto: boolean; registro: Registro | null }>({ aberto: false, registro: null });
  const senha = useRef<HTMLDialogElement>(null);

  const recarregar = useCallback(async () => {
    try {
      setRegistros(await listar());
      setErro("");
    } catch (e) {
      setErro((e as Error).message);
    } finally {
      setCarregando(false);
    }
  }, []);

  useEffect(() => {
    recarregar();
  }, [recarregar]);

  useEffect(() => {
    const ao = () => setPagina(paginaAtual());
    addEventListener("hashchange", ao);
    return () => removeEventListener("hashchange", ao);
  }, []);

  useEffect(() => {
    document.documentElement.dataset.theme = tema;
    try {
      localStorage.setItem("tema", tema);
    } catch {
      /* storage bloqueado: vale só nesta visita */
    }
  }, [tema]);

  const acoes: Acoes = { registros, abrirForm: (r) => setForm({ aberto: true, registro: r ?? null }), recarregar };
  const visivel: Pagina = pagina === "usuarios" && !admin ? "painel" : pagina;
  const fechar = () => setForm((f) => ({ ...f, aberto: false }));

  return (
    <div className="app">
      <aside className="lateral">
        <Marca />
        <nav aria-label="Páginas">
          <ul>
            <Item p="painel" atual={visivel} Icone={SquaresFour} />
            <Item p="registros" atual={visivel} Icone={ListBullets} />
            {admin && <Item p="usuarios" atual={visivel} Icone={UsersThree} />}
          </ul>
        </nav>
        <div className="rodape">
          <span className="avatar" aria-hidden>{usuario[0]?.toUpperCase()}</span>
          <span className="rodape-nome">
            {usuario}
            <small>{admin ? "Administrador(a)" : "Usuário"}</small>
          </span>
          <div className="rodape-acoes">
            <button type="button" className="botao-icone" aria-label="Trocar senha" title="Trocar senha" onClick={() => senha.current?.showModal()}>
              <Key size={16} weight="bold" aria-hidden />
            </button>
            <button
              type="button"
              className="botao-icone"
              aria-label={tema === "escuro" ? "Usar tema claro" : "Usar tema escuro"}
              title={tema === "escuro" ? "Usar tema claro" : "Usar tema escuro"}
              onClick={() => setTema(tema === "escuro" ? "claro" : "escuro")}
            >
              {tema === "escuro" ? <Sun size={16} weight="bold" aria-hidden /> : <Moon size={16} weight="bold" aria-hidden />}
            </button>
            <button type="button" className="botao-icone" aria-label="Sair" title="Sair" onClick={() => supabase().auth.signOut()}>
              <SignOut size={16} weight="bold" aria-hidden />
            </button>
          </div>
        </div>
      </aside>

      <main className="principal">
        {erro && (
          <p className="erro faixa-erro" role="alert">
            {erro}{" "}
            <button type="button" className="link" onClick={recarregar}>Tentar de novo</button>
          </p>
        )}
        {carregando ? (
          <p className="estado">Carregando acareações…</p>
        ) : visivel === "painel" ? (
          <Painel {...acoes} />
        ) : visivel === "registros" ? (
          <Registros {...acoes} />
        ) : (
          <Usuarios eu={sessao.user.id} />
        )}
      </main>

      <Formulario
        aberto={form.aberto}
        registro={form.registro}
        registros={registros}
        onFechar={fechar}
        onSalvo={() => {
          fechar();
          recarregar();
        }}
      />
      <TrocarSenha dialogo={senha} />
    </div>
  );
}

function TrocarSenha({ dialogo }: { dialogo: RefObject<HTMLDialogElement> }) {
  const [erro, setErro] = useState("");

  async function enviar(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    const form = e.currentTarget;
    const { error } = await supabase().auth.updateUser({ password: String(new FormData(form).get("senha")) });
    if (error) return setErro("Não foi possível trocar a senha. Use outra, com 8 ou mais caracteres.");
    form.reset();
    dialogo.current?.close();
  }

  return (
    <dialog ref={dialogo} className="caixa" aria-labelledby="titulo-senha" onClose={() => setErro("")}>
      <form onSubmit={enviar}>
        <h2 id="titulo-senha">Trocar senha</h2>
        <label className="campo">
          Nova senha
          <input name="senha" type="password" minLength={8} maxLength={72} required autoComplete="new-password" />
        </label>
        {erro && <p className="erro" role="alert">{erro}</p>}
        <div className="caixa-acoes">
          <button type="button" className="botao" onClick={() => dialogo.current?.close()}>Cancelar</button>
          <button className="botao primario">Salvar</button>
        </div>
      </form>
    </dialog>
  );
}
```

- [ ] **Step 6: Criar páginas provisórias (para compilar; substituídas nas próximas tasks)**

`frontend/src/Painel.tsx`:
```tsx
import type { Acoes } from "./App";

export default function Painel({ registros }: Acoes) {
  return <p className="estado">{registros.length} acareações</p>;
}
```
`frontend/src/Registros.tsx`:
```tsx
import type { Acoes } from "./App";

export default function Registros({ registros }: Acoes) {
  return <p className="estado">{registros.length} acareações</p>;
}
```
`frontend/src/Usuarios.tsx`:
```tsx
export default function Usuarios({ eu }: { eu: string }) {
  return <p className="estado">Usuários ({eu})</p>;
}
```
`frontend/src/Formulario.tsx`:
```tsx
import type { Registro } from "./dados";

type Props = { aberto: boolean; registro: Registro | null; registros: Registro[]; onFechar: () => void; onSalvo: () => void };

export default function Formulario({ aberto }: Props) {
  return aberto ? null : null;
}
```

- [ ] **Step 7: Compilar**

Run: `npm run build --prefix frontend`
Expected: build sem erros de TypeScript; gera `frontend/dist/`.

- [ ] **Step 8: Ver o login no navegador**

Run (em segundo plano): `npm run dev --prefix frontend -- --port 5173`
Depois: `python -c "from playwright.sync_api import sync_playwright as s; p=s().start(); b=p.chromium.launch(channel='msedge'); g=b.new_page(viewport={'width':1280,'height':800}); g.goto('http://localhost:5173'); g.wait_for_timeout(1500); g.screenshot(path=r'C:/Users/hacin/AppData/Local/Temp/claude/C--Users-hacin-Desktop-Acareacoes/c9591135-e316-4890-bf6b-77c6a097109b/scratchpad/login.png'); b.close(); p.stop()"` e abrir a imagem.
Expected: cartão de login centralizado com marca + faixa colorida. Sem `.env` o Supabase ainda não responde — só o visual importa aqui. (Se a tela vier em branco por falta de `VITE_SUPABASE_URL`, criar `.env` na raiz com `VITE_SUPABASE_URL=http://localhost:54321` e `VITE_SUPABASE_ANON_KEY=x` só para esta checagem e apagar depois.)

- [ ] **Step 9: Commit**

```bash
git add frontend/index.html frontend/src/main.tsx frontend/src/styles.css frontend/src/App.tsx frontend/src/Painel.tsx frontend/src/Registros.tsx frontend/src/Usuarios.tsx frontend/src/Formulario.tsx
git commit -m "feat(front): visual base, login, menu lateral, tema e troca de senha

```

---

### Task 6: Painel

**Files:**
- Modify (substituir inteiro): `frontend/src/Painel.tsx`

**Interfaces:**
- Consumes: `Acoes` (Task 5); de `dados.ts`: `anterior, contar, exportar, fmtData, fmtMes, fmtNum, fmtPct, hoje, intervalo, kpis, mesDe, navegar, noIntervalo, porMes, ranking, rotuloPeriodo, type Escala, type Item, type Periodo, type Situacao`.

- [ ] **Step 1: Implementar** — `frontend/src/Painel.tsx`"
```tsx
import { useState } from "react";
import { CaretLeft, CaretRight, DownloadSimple, Plus } from "@phosphor-icons/react";
import type { Acoes } from "./App";
import {
  anterior, contar, exportar, fmtData, fmtMes, fmtNum, fmtPct, hoje, intervalo, kpis, mesDe, navegar, noIntervalo,
  porMes, ranking, rotuloPeriodo, type Escala, type Item, type Periodo, type Situacao,
} from "./dados";

const ESCALAS: { valor: Escala; rotulo: string }[] = [
  { valor: "mes", rotulo: "Mês" },
  { valor: "trimestre", rotulo: "Trimestre" },
  { valor: "ano", rotulo: "Ano" },
  { valor: "tudo", rotulo: "Tudo" },
];
const CLASSE: Record<Situacao, string> = { "Ótimo": "ot", Bom: "bo", Regular: "re", "Atenção": "at" };
const pp = (d: number) => `${fmtNum(d * 100)} p.p.`;

export default function Painel({ registros, abrirForm }: Acoes) {
  const atual = mesDe(hoje());
  const primeiro = registros.length ? mesDe(registros[registros.length - 1].data_reclamacao) : atual; // API ordena por data desc
  const [periodo, setPeriodo] = useState<Periodo>({ escala: "mes", ref: atual });
  const [exportando, setExportando] = useState(false);
  const [erro, setErro] = useState("");

  const int = intervalo(periodo, atual, primeiro);
  const doPeriodo = noIntervalo(registros, int);
  const k = kpis(doPeriodo, int.meses);
  const ant = anterior(periodo);
  const intAnt = ant && intervalo(ant, atual, primeiro);
  const kAnt = intAnt ? kpis(noIntervalo(registros, intAnt), intAnt.meses) : null;
  const rotuloAnt = ant ? rotuloPeriodo(ant) : "";
  const rank = ranking(doPeriodo, int.meses);
  const maxRank = rank[0]?.total ?? 0;
  const meses = porMes(registros, int.ate < atual ? int.ate : atual);
  const maxMes = Math.max(0, ...meses.map((m) => m.valor));
  const mostraValor = periodo.escala === "mes" || periodo.escala === "trimestre";

  async function baixar() {
    setExportando(true);
    setErro("");
    try {
      await exportar(doPeriodo.map((r) => r.id));
    } catch (e) {
      setErro((e as Error).message);
    } finally {
      setExportando(false);
    }
  }

  return (
    <>
      <header className="cabecalho">
        <div>
          <h1>Painel</h1>
          <p>Acareações de motoristas e procedência das reclamações</p>
        </div>
        <div className="acoes">
          <button type="button" className="botao" disabled={exportando || !doPeriodo.length} onClick={baixar}>
            <DownloadSimple size={16} weight="bold" aria-hidden />
            {exportando ? "Exportando…" : "Exportar .xlsx"}
          </button>
          <button type="button" className="botao primario" onClick={() => abrirForm()}>
            <Plus size={15} weight="bold" aria-hidden />
            Nova acareação
          </button>
        </div>
      </header>

      <div className="periodo">
        <div className="seg" role="group" aria-label="Tamanho do período">
          {ESCALAS.map((e) => (
            <button key={e.valor} type="button" aria-pressed={periodo.escala === e.valor} onClick={() => setPeriodo({ ...periodo, escala: e.valor })}>
              {e.rotulo}
            </button>
          ))}
        </div>
        {periodo.escala === "tudo" ? (
          <span className="passo">Desde {fmtMes(primeiro, true).toLowerCase()}</span>
        ) : (
          <div className="passo">
            <button type="button" className="botao-icone" aria-label="Período anterior" onClick={() => setPeriodo(navegar(periodo, -1))}>
              <CaretLeft size={14} weight="bold" aria-hidden />
            </button>
            <span aria-live="polite">{rotuloPeriodo(periodo)}</span>
            <button type="button" className="botao-icone" aria-label="Próximo período" disabled={int.ate >= atual} onClick={() => setPeriodo(navegar(periodo, 1))}>
              <CaretRight size={14} weight="bold" aria-hidden />
            </button>
          </div>
        )}
      </div>

      {erro && <p className="erro faixa-erro" role="alert">{erro}</p>}

      <div className="grade">
        <div className="kpis">
          <Kpi rotulo="Acareações" atual={k.total} anterior={kAnt?.total} fmt={fmtNum} rotuloAnt={rotuloAnt} />
          <Kpi rotulo="Procedentes" atual={k.procedentes} anterior={kAnt?.procedentes} fmt={fmtPct} fmtDif={pp} rotuloAnt={rotuloAnt} />
          <Kpi rotulo="Indenizados" atual={k.indenizados} anterior={kAnt?.indenizados} fmt={fmtNum} rotuloAnt={rotuloAnt} />
          <Kpi rotulo="Motoristas em atenção" atual={k.atencao} anterior={kAnt?.atencao} fmt={fmtNum} rotuloAnt={rotuloAnt} />
        </div>

        <section className="card span-2">
          <h2>Reclamações por motorista</h2>
          <ul className="legenda">
            <li><i className="s1" />Improcedente</li>
            <li><i className="s2" />Procedente</li>
            <li><i className="s3" />Em análise</li>
          </ul>
          {rank.length ? (
            <ol className="rank">
              {rank.map((l, i) => (
                <li key={l.motorista} title={`${l.motorista}: ${l.improcedente} improcedente(s), ${l.procedente} procedente(s), ${l.outras} em análise`}>
                  <span className="rank-pos">{i + 1}</span>
                  <span className="rank-nome">{l.motorista}</span>
                  <span className="trilho" aria-hidden>
                    {([["s1", l.improcedente], ["s2", l.procedente], ["s3", l.outras]] as const).map(([classe, n]) =>
                      n > 0 ? <span key={classe} className={classe} style={{ width: `${(n / maxRank) * 100}%` }} /> : null,
                    )}
                  </span>
                  <b className="num">{l.total}</b>
                  <span className={`sit ${CLASSE[l.situacao]}`}>{l.situacao}</span>
                </li>
              ))}
            </ol>
          ) : (
            <p className="estado">Nenhuma acareação neste período.</p>
          )}
        </section>

        <section className="card">
          <h2>Acareações por mês</h2>
          <div className="colunas" role="img" aria-label={`Acareações por mês: ${meses.map((m) => `${fmtMes(m.rotulo)} ${m.valor}`).join(", ")}`}>
            {meses.map((m) => {
              const dentro = m.rotulo >= int.de && m.rotulo <= int.ate;
              return (
                <div key={m.rotulo} className={dentro ? "on" : ""} title={`${fmtMes(m.rotulo)}: ${m.valor}`}>
                  <small>{dentro && mostraValor ? m.valor : ""}</small>
                  <span className="barra">
                    <span style={{ height: `${maxMes ? (m.valor / maxMes) * 100 : 0}%` }} />
                  </span>
                  {fmtMes(m.rotulo)[0].toUpperCase()}
                </div>
              );
            })}
          </div>
          <h2 className="depois">Status</h2>
          <Barras itens={contar(doPeriodo, "status")} total={doPeriodo.length} />
        </section>

        <section className="card">
          <h2>Por embarcador</h2>
          <Barras itens={contar(doPeriodo, "embarcador", 4)} total={doPeriodo.length} />
        </section>

        <section className="card span-2">
          <h2>
            Últimos registros
            <a href="#/registros">Ver todos os registros →</a>
          </h2>
          {doPeriodo.length ? (
            <div className="tabela-rolagem">
              <table className="tabela">
                <caption className="sr-only">Últimas acareações do período</caption>
                <thead>
                  <tr>
                    <th scope="col">Data</th>
                    <th scope="col">Ticket</th>
                    <th scope="col">Cliente</th>
                    <th scope="col">Motorista</th>
                    <th scope="col">Embarcador</th>
                    <th scope="col">Status</th>
                    <th scope="col">Procedência</th>
                  </tr>
                </thead>
                <tbody>
                  {doPeriodo.slice(0, 5).map((r) => (
                    <tr key={r.id}>
                      <td>{fmtData(r.data_reclamacao)}</td>
                      <td className="num">{r.ticket}</td>
                      <td>{r.cliente}</td>
                      <td>{r.motorista}</td>
                      <td>{r.embarcador}</td>
                      <td><span className="pill">{r.status}</span></td>
                      <td>{r.procedencia}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="estado">Nenhuma acareação neste período.</p>
          )}
        </section>
      </div>
    </>
  );
}

function Kpi(props: { rotulo: string; atual: number; anterior?: number; fmt: (n: number) => string; fmtDif?: (n: number) => string; rotuloAnt: string }) {
  const { rotulo, atual, anterior, fmt, fmtDif = fmt, rotuloAnt } = props;
  return (
    <section className="card kpi">
      <h2>{rotulo}</h2>
      <p className="kpi-valor">{fmt(atual)}</p>
      <p className="kpi-var">
        {anterior === undefined ? (
          "Todo o histórico"
        ) : atual === anterior ? (
          `Igual a ${rotuloAnt.toLowerCase()} (${fmt(anterior)})`
        ) : (
          <>
            {/* nos quatro indicadores, menos é melhor */}
            <b className={atual < anterior ? "melhor" : "pior"}>{atual < anterior ? "↓" : "↑"} {fmtDif(Math.abs(atual - anterior))}</b>{" "}
            vs. {rotuloAnt.toLowerCase()} ({fmt(anterior)})
          </>
        )}
      </p>
    </section>
  );
}

function Barras({ itens, total }: { itens: Item[]; total: number }) {
  const max = Math.max(0, ...itens.map((i) => i.valor));
  if (!itens.length) return <p className="estado">Nenhuma acareação neste período.</p>;
  return (
    <ul className="bh">
      {itens.map((i) => (
        <li key={i.rotulo}>
          <span>{i.rotulo}</span>
          <b className="num">
            {fmtNum(i.valor)}
            <small>{fmtPct(i.valor / total)}</small>
          </b>
          <span className="bh-trilho" aria-hidden>
            <span style={{ width: `${(i.valor / max) * 100}%` }} />
          </span>
        </li>
      ))}
    </ul>
  );
}
```

- [ ] **Step 2: Compilar e testar**

Run: `npm run build --prefix frontend && npm test --prefix frontend`
Expected: build ok; 12 testes passam.

- [ ] **Step 3: Commit**

```bash
git add frontend/src/Painel.tsx
git commit -m "feat(front): painel com KPIs, ranking de motoristas e gráficos

```

---

### Task 7: Painel lateral de cadastro e edição

**Files:**
- Modify (substituir inteiro): `frontend/src/Formulario.tsx`

**Interfaces:**
- Consumes: de `dados.ts`: `ACAREACOES, PROCEDENCIAS, STATUS, duplicado, fmtData, hoje, salvar, unicos, type Campos, type Registro`.
- Produces: `Formulario({ aberto, registro, registros, onFechar, onSalvo })` — mesmo contrato do provisório da Task 5.

- [ ] **Step 1: Implementar** — `frontend/src/Formulario.tsx`"
```tsx
import { useEffect, useRef, useState, type FormEvent } from "react";
import { X } from "@phosphor-icons/react";
import { ACAREACOES, PROCEDENCIAS, STATUS, duplicado, fmtData, hoje, salvar, unicos, type Campos, type Registro } from "./dados";

type Props = { aberto: boolean; registro: Registro | null; registros: Registro[]; onFechar: () => void; onSalvo: () => void };

export default function Formulario({ aberto, registro: r, registros, onFechar, onSalvo }: Props) {
  const dialogo = useRef<HTMLDialogElement>(null);
  const [ticket, setTicket] = useState("");
  const [erro, setErro] = useState("");
  const [salvando, setSalvando] = useState(false);
  const [versao, setVersao] = useState(0); // muda a cada abertura: o form remonta com os valores do registro

  useEffect(() => {
    const d = dialogo.current!;
    if (aberto && !d.open) {
      setTicket(r ? String(r.ticket) : "");
      setErro("");
      setVersao((v) => v + 1);
      d.showModal();
    }
    if (!aberto && d.open) d.close();
  }, [aberto, r]);

  const dup = /^\d+$/.test(ticket) ? duplicado(registros, Number(ticket), r?.id) : undefined;

  async function enviar(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    const f = Object.fromEntries(new FormData(e.currentTarget)) as Record<string, string>;
    setSalvando(true);
    setErro("");
    try {
      await salvar({ ...f, ticket: Number(f.ticket), entrega: Number(f.entrega) } as Campos, r?.id);
      onSalvo();
    } catch (err) {
      setErro((err as Error).message);
    } finally {
      setSalvando(false);
    }
  }

  return (
    <dialog ref={dialogo} className="sheet" aria-labelledby="titulo-form" onClose={onFechar}>
      <header className="sheet-topo">
        <h2 id="titulo-form">{r ? "Editar acareação" : "Nova acareação"}</h2>
        <button type="button" className="botao-icone" aria-label="Fechar" onClick={onFechar}>
          <X size={16} weight="bold" aria-hidden />
        </button>
      </header>
      <form key={versao} id="form-acareacao" className="sheet-corpo" onSubmit={enviar}>
        <label className="campo">
          Data da reclamação
          <input type="date" name="data_reclamacao" required defaultValue={r?.data_reclamacao ?? hoje()} />
        </label>
        <label className="campo">
          Ticket
          <input name="ticket" inputMode="numeric" pattern="\d{1,15}" required defaultValue={r?.ticket} onChange={(e) => setTicket(e.target.value)} />
        </label>
        {dup && (
          <p className="aviso-dup" role="status">
            Ticket já cadastrado em {fmtData(dup.data_reclamacao)} ({dup.cliente}). Pode salvar mesmo assim.
          </p>
        )}
        <label className="campo">
          Nº da entrega
          <input name="entrega" inputMode="numeric" pattern="\d{1,15}" required defaultValue={r?.entrega} />
        </label>
        <label className="campo">
          Cliente
          <input name="cliente" required maxLength={200} defaultValue={r?.cliente} placeholder="Nome do cliente" />
        </label>
        <label className="campo inteiro">
          Motorista
          <input name="motorista" list="lista-motoristas" required maxLength={200} autoComplete="off" defaultValue={r?.motorista} />
        </label>
        <label className="campo inteiro">
          Embarcador
          <input name="embarcador" list="lista-embarcadores" required maxLength={200} autoComplete="off" defaultValue={r?.embarcador} placeholder="Ex.: MAGAZINE LUIZA" />
        </label>
        <Opcoes nome="acareacao" rotulo="Acareação" opcoes={ACAREACOES} valor={r?.acareacao ?? "Acareação Motorista"} />
        <Opcoes nome="status" rotulo="Status" opcoes={STATUS} valor={r?.status ?? "Aguardando"} />
        <Opcoes nome="procedencia" rotulo="Procedência" opcoes={PROCEDENCIAS} valor={r?.procedencia ?? "Aguardando"} inteiro />
        <label className="campo inteiro">
          Detalhe
          <textarea name="detalhe" rows={3} maxLength={5000} defaultValue={r?.detalhe} placeholder="O que aconteceu (opcional)" />
        </label>
        <datalist id="lista-motoristas">
          {unicos(registros, "motorista").map((m) => <option key={m} value={m} />)}
        </datalist>
        <datalist id="lista-embarcadores">
          {unicos(registros, "embarcador").map((m) => <option key={m} value={m} />)}
        </datalist>
      </form>
      <footer className="sheet-rodape">
        {erro && <p className="erro" role="alert">{erro}</p>}
        <button type="button" className="botao" onClick={onFechar}>Cancelar</button>
        <button type="submit" form="form-acareacao" className="botao primario" disabled={salvando}>
          {salvando ? "Salvando…" : "Salvar acareação"}
        </button>
      </footer>
    </dialog>
  );
}

function Opcoes(props: { nome: string; rotulo: string; opcoes: readonly string[]; valor: string; inteiro?: boolean }) {
  return (
    <label className={`campo${props.inteiro ? " inteiro" : ""}`}>
      {props.rotulo}
      <select name={props.nome} defaultValue={props.valor}>
        {props.opcoes.map((o) => <option key={o}>{o}</option>)}
      </select>
    </label>
  );
}
```

- [ ] **Step 2: Compilar**

Run: `npm run build --prefix frontend`
Expected: build ok.

- [ ] **Step 3: Commit**

```bash
git add frontend/src/Formulario.tsx
git commit -m "feat(front): painel lateral de cadastro e edição com aviso de ticket repetido

```

---

### Task 8: Registros

**Files:**
- Modify (substituir inteiro): `frontend/src/Registros.tsx`

**Interfaces:**
- Consumes: `Acoes`; de `dados.ts`: `PROCEDENCIAS, SEM_FILTRO, STATUS, excluir, exportar, filtrar, fmtData, fmtMes, fmtNum, mesDe, unicos, type Filtros, type Registro`.

- [ ] **Step 1: Implementar** — `frontend/src/Registros.tsx`"
```tsx
import { useState, type ChangeEvent } from "react";
import { DownloadSimple, MagnifyingGlass, Plus } from "@phosphor-icons/react";
import type { Acoes } from "./App";
import {
  PROCEDENCIAS, SEM_FILTRO, STATUS, excluir, exportar, filtrar, fmtData, fmtMes, fmtNum, mesDe, unicos,
  type Filtros, type Registro,
} from "./dados";

const pares = (valores: readonly string[]): [string, string][] => valores.map((v) => [v, v]);

export default function Registros({ registros, abrirForm, recarregar }: Acoes) {
  const [f, setF] = useState<Filtros>(SEM_FILTRO);
  const [erro, setErro] = useState("");
  const [exportando, setExportando] = useState(false);
  // shortcut: renderiza todas as linhas filtradas (~1.200/ano); paginar se a tabela ficar lenta.
  const lista = filtrar(registros, f);
  const meses = [...new Set(registros.map((r) => mesDe(r.data_reclamacao)))]; // já em ordem desc
  const muda = (campo: keyof Filtros) => (e: ChangeEvent<HTMLInputElement | HTMLSelectElement>) => setF({ ...f, [campo]: e.target.value });
  const desde = registros.length ? ` desde ${fmtMes(meses[meses.length - 1], true).toLowerCase()}` : "";

  async function baixar() {
    setExportando(true);
    setErro("");
    try {
      await exportar(lista.map((r) => r.id));
    } catch (e) {
      setErro((e as Error).message);
    } finally {
      setExportando(false);
    }
  }

  async function apagar(r: Registro) {
    if (!confirm(`Excluir a acareação do ticket ${r.ticket} (${r.cliente})? Isso não pode ser desfeito.`)) return;
    try {
      await excluir(r.id);
      await recarregar();
    } catch (e) {
      setErro((e as Error).message);
    }
  }

  return (
    <>
      <header className="cabecalho">
        <div>
          <h1>Registros</h1>
          <p>{fmtNum(registros.length)} acareações{desde}</p>
        </div>
        <div className="acoes">
          <button type="button" className="botao" disabled={exportando || !lista.length} onClick={baixar}>
            <DownloadSimple size={16} weight="bold" aria-hidden />
            {exportando ? "Exportando…" : `Exportar ${fmtNum(lista.length)} .xlsx`}
          </button>
          <button type="button" className="botao primario" onClick={() => abrirForm()}>
            <Plus size={15} weight="bold" aria-hidden />
            Nova acareação
          </button>
        </div>
      </header>

      {erro && <p className="erro faixa-erro" role="alert">{erro}</p>}

      <section className="card">
        <div className="barra-tabela">
          <label className="busca">
            <MagnifyingGlass size={15} aria-hidden />
            <span className="sr-only">Buscar</span>
            <input type="search" placeholder="Buscar ticket, entrega ou cliente" value={f.busca} onChange={muda("busca")} />
          </label>
          <Filtro rotulo="Período" valor={f.mes} onChange={muda("mes")} opcoes={meses.map((m): [string, string] => [m, fmtMes(m)])} />
          <Filtro rotulo="Motorista" valor={f.motorista} onChange={muda("motorista")} opcoes={pares(unicos(registros, "motorista"))} />
          <Filtro rotulo="Embarcador" valor={f.embarcador} onChange={muda("embarcador")} opcoes={pares(unicos(registros, "embarcador"))} />
          <Filtro rotulo="Status" valor={f.status} onChange={muda("status")} opcoes={pares(STATUS)} />
          <Filtro rotulo="Procedência" valor={f.procedencia} onChange={muda("procedencia")} opcoes={pares(PROCEDENCIAS)} />
          {Object.values(f).some(Boolean) && (
            <button type="button" className="link" onClick={() => setF(SEM_FILTRO)}>Limpar filtros</button>
          )}
        </div>

        <div className="tabela-rolagem">
          <table className="tabela">
            <caption className="sr-only">Acareações filtradas</caption>
            <thead>
              <tr>
                <th scope="col">Data</th>
                <th scope="col">Ticket</th>
                <th scope="col">Entrega</th>
                <th scope="col">Cliente</th>
                <th scope="col">Acareação</th>
                <th scope="col">Motorista</th>
                <th scope="col">Embarcador</th>
                <th scope="col">Status</th>
                <th scope="col">Procedência</th>
                <th scope="col">Detalhe</th>
                <th scope="col"><span className="sr-only">Ações</span></th>
              </tr>
            </thead>
            <tbody>
              {lista.map((r) => (
                <tr key={r.id}>
                  <td>{fmtData(r.data_reclamacao)}</td>
                  <td className="num">{r.ticket}</td>
                  <td className="num">{r.entrega}</td>
                  <td>{r.cliente}</td>
                  <td>{r.acareacao}</td>
                  <td>{r.motorista}</td>
                  <td>{r.embarcador}</td>
                  <td><span className="pill">{r.status}</span></td>
                  <td>{r.procedencia}</td>
                  <td className="detalhe" title={r.detalhe}>{r.detalhe}</td>
                  <td className="acoes-linha">
                    <button type="button" className="link" onClick={() => abrirForm(r)}>Editar</button>
                    <button type="button" className="link perigo" onClick={() => apagar(r)}>Excluir</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {!lista.length && <p className="estado">Nenhuma acareação com esses filtros.</p>}
      </section>
    </>
  );
}

function Filtro(props: { rotulo: string; valor: string; opcoes: [string, string][]; onChange: (e: ChangeEvent<HTMLSelectElement>) => void }) {
  return (
    <label className={`filtro${props.valor ? " ativo" : ""}`}>
      <span className="sr-only">{props.rotulo}</span>
      <select value={props.valor} onChange={props.onChange}>
        <option value="">{props.rotulo}: todos</option>
        {props.opcoes.map(([v, t]) => <option key={v} value={v}>{t}</option>)}
      </select>
    </label>
  );
}
```

- [ ] **Step 2: Compilar**

Run: `npm run build --prefix frontend`
Expected: build ok.

- [ ] **Step 3: Commit**

```bash
git add frontend/src/Registros.tsx
git commit -m "feat(front): página de registros com busca, filtros, edição, exclusão e exportação

```

---

### Task 9: Usuários (admin)

**Files:**
- Modify (substituir inteiro): `frontend/src/Usuarios.tsx`

**Interfaces:**
- Consumes: de `dados.ts`: `criarUsuario, fmtData, listarUsuarios, removerUsuario, type Usuario`.

- [ ] **Step 1: Implementar** — `frontend/src/Usuarios.tsx`"
```tsx
import { useCallback, useEffect, useState, type FormEvent } from "react";
import { criarUsuario, fmtData, listarUsuarios, removerUsuario, type Usuario } from "./dados";

export default function Usuarios({ eu }: { eu: string }) {
  const [lista, setLista] = useState<Usuario[] | null>(null);
  const [erro, setErro] = useState("");
  const [ok, setOk] = useState("");

  const carregar = useCallback(() => listarUsuarios().then(setLista, (e: Error) => setErro(e.message)), []);
  useEffect(() => {
    carregar();
  }, [carregar]);

  async function criar(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    const form = e.currentTarget;
    const f = new FormData(form);
    setErro("");
    setOk("");
    try {
      await criarUsuario({ usuario: String(f.get("usuario")).trim().toLowerCase(), senha: String(f.get("senha")), admin: f.get("admin") === "on" });
      form.reset();
      setOk("Usuário criado. Passe o usuário e a senha inicial para a pessoa.");
      carregar();
    } catch (err) {
      setErro((err as Error).message);
    }
  }

  async function remover(u: Usuario) {
    if (!confirm(`Remover o acesso de ${u.usuario}?`)) return;
    setErro("");
    setOk("");
    try {
      await removerUsuario(u.id);
      carregar();
    } catch (err) {
      setErro((err as Error).message);
    }
  }

  return (
    <>
      <header className="cabecalho">
        <div>
          <h1>Usuários</h1>
          <p>Quem pode entrar no dashboard</p>
        </div>
      </header>
      <div className="grade">
        <section className="card span-2">
          <h2>Com acesso</h2>
          {lista ? (
            <div className="tabela-rolagem">
              <table className="tabela">
                <caption className="sr-only">Usuários com acesso</caption>
                <thead>
                  <tr>
                    <th scope="col">Usuário</th>
                    <th scope="col">Perfil</th>
                    <th scope="col">Desde</th>
                    <th scope="col"><span className="sr-only">Ações</span></th>
                  </tr>
                </thead>
                <tbody>
                  {lista.map((u) => (
                    <tr key={u.id}>
                      <td>{u.usuario}</td>
                      <td>{u.admin ? "Administrador(a)" : "Usuário"}</td>
                      <td>{u.criado_em ? fmtData(u.criado_em) : "—"}</td>
                      <td className="acoes-linha">
                        {u.id !== eu && <button type="button" className="link perigo" onClick={() => remover(u)}>Remover</button>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="estado">{erro ? "Não foi possível carregar." : "Carregando…"}</p>
          )}
        </section>
        <section className="card">
          <h2>Novo usuário</h2>
          <form className="form-usuario" onSubmit={criar}>
            <label className="campo">
              Usuário
              <input name="usuario" required pattern="[A-Za-z0-9._\-]{3,40}" title="3 a 40 letras, números, ponto, _ ou -" autoComplete="off" />
            </label>
            <label className="campo">
              Senha inicial
              <input name="senha" type="password" required minLength={8} maxLength={72} autoComplete="new-password" />
            </label>
            <label className="marcar">
              <input type="checkbox" name="admin" />
              Administrador(a): pode criar e remover usuários
            </label>
            {erro && <p className="erro" role="alert">{erro}</p>}
            {ok && <p className="ok" role="status">{ok}</p>}
            <button className="botao primario">Criar usuário</button>
          </form>
        </section>
      </div>
    </>
  );
}
```

- [ ] **Step 2: Compilar e testar**

Run: `npm run build --prefix frontend && npm test --prefix frontend && python -m pytest tests/ -q`
Expected: build ok; todos os testes passam (1 pulado sem banco de teste).

- [ ] **Step 3: Commit**

```bash
git add frontend/src/Usuarios.tsx
git commit -m "feat(front): gestão de usuários para administradores

```

---

### Task 10: Publicação, documentação e verificação ponta a ponta

**Files:**
- Create: `vercel.json`, `README.md`

**Interfaces:**
- Consumes: tudo.

- [ ] **Step 1: Criar `vercel.json`**"
```json
{
  "installCommand": "npm ci --prefix frontend",
  "buildCommand": "npm run build --prefix frontend",
  "outputDirectory": "frontend/dist",
  "regions": ["gru1"],
  "rewrites": [{ "source": "/api/(.*)", "destination": "/api/index" }]
}
```

- [ ] **Step 2: Criar `README.md`**

````markdown
# Dashboard de Acareações

Substitui a planilha de acareações: cadastro, painel com o resultado por motorista, registros filtráveis e exportação `.xlsx`.

- `api/index.py` — API (FastAPI) que roda na Vercel
- `main.py` — comandos locais (importar planilha, criar admin)
- `schema.sql` — tabela do banco
- `frontend/` — telas (React)

## Variáveis de ambiente

Local: arquivo `.env` na raiz (nunca commitar). Na Vercel: Settings → Environment Variables.

| Variável | Onde achar no Supabase | Pública? |
|---|---|---|
| `VITE_SUPABASE_URL` | Project Settings → API → Project URL | sim |
| `VITE_SUPABASE_ANON_KEY` | Project Settings → API → anon public | sim |
| `SUPABASE_URL` | igual a `VITE_SUPABASE_URL` | não |
| `SUPABASE_SERVICE_ROLE_KEY` | Project Settings → API → service_role | **secreta** |
| `DATABASE_URL` | Connect → Transaction pooler (porta 6543) | **secreta** |

## Configuração (uma vez)

1. **Supabase** (região São Paulo):
   - SQL Editor → colar e rodar `schema.sql`.
   - Authentication → Sign In / Providers → **desligar "Allow new users to sign up"** (só a admin cria contas).
2. **Local**: `pip install -r requirements.txt uvicorn python-dotenv pytest` e criar o `.env`.
3. **Importar o histórico** (a planilha fica só no seu computador):
   ```
   python main.py importar "ACAREAÇÕES HRAO.xlsx"
   ```
   Abra `nomes.csv` no Excel, confira a coluna `sugerido` (apague a sugestão para manter o nome original), salve e rode:
   ```
   python main.py importar "ACAREAÇÕES HRAO.xlsx" --aplicar
   ```
4. **Criar a admin**: `python main.py criar-admin nome.do.admin` (pede a senha).
5. **Vercel**: Add New → Project → importar o repositório do GitHub, cadastrar as 5 variáveis e fazer o deploy. Cada push na `main` publica sozinho.

## Rodar localmente

```
uvicorn api.index:app --env-file .env --port 8000 --reload
npm install --prefix frontend
npm run dev --prefix frontend
```
Abrir http://localhost:5173.

## Testes

```
python -m pytest
npm test --prefix frontend
```
Para testar gravação no banco, defina `DATABASE_URL_TESTE` (cria e apaga um registro).

## Regras

- **Situação do motorista**: média mensal de reclamações no período — até 1 Ótimo, até 3 Bom, até 5 Regular, acima de 5 Atenção.
- **Backup**: o plano gratuito do Supabase não tem backup para baixar; exporte o `.xlsx` em Registros periodicamente.
````

- [ ] **Step 3: Pedir ao usuário as credenciais e o banco**

Parar e pedir ao usuário: criar o projeto Supabase (São Paulo), rodar `schema.sql`, desligar cadastro público e criar o `.env` na raiz com as 5 variáveis. Não seguir sem isso.

- [ ] **Step 4: Importar o histórico (com revisão humana)**

Run: `python main.py importar "ACAREAÇÕES HRAO.xlsx"`
Expected: `956 registros lidos, 0 com erro.` e `nomes.csv` criado. Mostrar ao usuário os pares em que `original != sugerido` e esperar a aprovação/edição do `nomes.csv`. Depois:
Run: `python main.py importar "ACAREAÇÕES HRAO.xlsx" --aplicar`
Expected: `956 registros importados.`

- [ ] **Step 5: Criar a admin**

Pedir ao usuário que rode no terminal dele (a senha não deve passar pelo agente): `! python main.py criar-admin nome.do.admin`
Expected: `Administrador nome.do.admin criado.`

- [ ] **Step 6: Teste de gravação no banco**

Run: `DATABASE_URL_TESTE="$(python -c "from dotenv import dotenv_values; print(dotenv_values()['DATABASE_URL'])")" python -m pytest tests/test_api.py::test_crud_no_banco -v`
Expected: PASS (cria, edita, lista, apaga um registro de teste).

- [ ] **Step 7: Verificação no navegador**

Subir API e frontend (comandos do README, em segundo plano). Com o usuário logado no navegador dele (ou com uma conta de teste criada via `python -c "..."` que o usuário aprovar), conferir e registrar com screenshots:
1. Painel em **maio/2026**: o total de acareações e o ranking por motorista (total, procedentes, situação) são iguais aos da aba RESULTADO da planilha local, salvo junções de nomes aprovadas no `nomes.csv`.
2. Trimestre/Ano/Tudo trocam os números sem recarregar a página; ‹ › navegam; › desabilita no mês atual.
3. Nova acareação: abre pela direita, Esc fecha, ticket existente mostra o aviso, salvar fecha e o total sobe 1.
4. Registros: busca "jose" acha "JOSÉ"; filtros combinam; Editar abre com os dados; Excluir pede confirmação e remove (excluir o registro criado no item 3).
5. Exportar baixa `.xlsx` que abre no Excel com as colunas da planilha e as linhas filtradas.
6. Tema escuro legível; largura 390px sem rolagem horizontal da página.
7. Usuários: criar usuário comum, entrar com ele (sem item Usuários no menu), remover.

- [ ] **Step 8: Commit e push (com autorização do usuário)**

```bash
git add vercel.json README.md
git commit -m "docs: publicação na Vercel e guia de configuração"
```
Perguntar ao usuário antes de `git push -u origin feat/dashboard` e do merge na `main` (que dispara o deploy). Depois do deploy, repetir os itens 1, 3 e 5 do Step 7 na URL da Vercel.
