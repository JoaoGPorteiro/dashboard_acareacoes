# Dashboard de Acareações — Design

Data: 2026-10-15 · Status: aprovado em conversa, aguardando revisão desta spec

## Objetivo

Substituir a planilha `ACAREAÇÕES HRAO.xlsx` (uma aba por mês + aba RESULTADO preenchida à mão) por um dashboard web onde:

- a página inicial permite cadastrar uma nova acareação sem abrir o Excel nem criar abas;
- o RESULTADO (reclamações por motorista + SITUAÇÃO) é calculado sozinho, para qualquer período;
- os registros podem ser filtrados, editados e exportados para `.xlsx`.

Sucesso = a equipe para de usar a planilha: cadastra, consulta e exporta tudo pelo dashboard, com o histórico de jan–out/2026 já dentro.

## Decisões tomadas

| Tema | Decisão |
|---|---|
| Hospedagem | Vercel (frontend + função Python, região `gru1`) + Supabase Postgres (região `sa-east-1`) |
| Repositório | `github.com/JoaoGPorteiro/dashboard_acareacoes` (**público**) — deploy automático da Vercel a partir da `main` |
| Backend | Um arquivo FastAPI (`api/index.py`), SQL direto com `psycopg`, sem ORM |
| Frontend | React + TSX (Vite), Recharts para gráficos, `@supabase/supabase-js` só para login |
| Login | Supabase Auth com usuário + senha; o usuário vira e-mail interno `<usuario>@acareacoes.local` |
| Perfis | `admin` (cria/remove usuários) e usuário comum (todo o resto) |
| Usuário inicial | `nome.do.admin`, admin |
| Histórico | Importar jan–out/2026 (956 registros) unificando grafias de nomes, com aprovação humana |
| Visual | Estrutura do dashboard RH do usuário (`inspiracao/dashboard_rh-main`), sem barra de filtros no topo; visual decidido pela skill apple-design dentro da marca Magalu |
| Código | Poucos arquivos, legível sem agente (preferência explícita do usuário) |

## 1. Dados

### Tabela `acareacoes`

| Coluna | Tipo | Regra |
|---|---|---|
| `id` | `bigint identity` | PK |
| `data_reclamacao` | `date` | obrigatório; formulário sugere hoje |
| `ticket` | `bigint` | obrigatório; **não único** (há 25 repetidos legítimos no histórico) |
| `entrega` | `bigint` | obrigatório |
| `cliente` | `text` | obrigatório |
| `acareacao` | `text` | um de: Whatsapp, Whats Motorista, Acareação Motorista, Recebido; padrão Acareação Motorista |
| `status` | `text` | um de: Aguardando, Recebido, Indenizado, Devolvido, Parado |
| `procedencia` | `text` | um de: Procedente, Improcedente, Analisando, Aguardando, Sem retorno |
| `motorista` | `text` | obrigatório |
| `embarcador` | `text` | obrigatório |
| `detalhe` | `text` | opcional |
| `criado_por` | `text` | usuário logado, preenchido pela API |
| `criado_em` | `timestamptz` | `now()` |

Listas fechadas validadas na API (Pydantic `Literal`) e por `CHECK` no banco. Motorista e embarcador são texto livre com autocompletar dos valores já existentes — sem tabela de cadastro; o de/para de grafias fica em `apelidos` (abaixo).

RLS **ativado sem nenhuma policy**: a chave anônima (que vai para o navegador) não lê nem escreve a tabela via API REST do Supabase. Só a função Python acessa, pela `DATABASE_URL`.

### De/para de nomes — tabela `apelidos` (adicionado em 2026-10-15, a pedido do usuário)

`(campo, chave, nome)`, chave primária `(campo, chave)`. `chave` = nome sem acento, sem caixa e **sem espaços** (`"Fulano  Graça"` → `fulanograca`), a forma justa de comparar. Fica no banco (nomes de motoristas são dados pessoais; o repositório é público), com RLS como `acareacoes`.

- A importação grava, na mesma transação dos registros, a chave de cada grafia original e do nome padrão aprovado no `nomes.csv`. Duas grafias com a mesma chave apontando para nomes diferentes = `nomes.csv` contraditório → nada é gravado.
- Ao criar/editar pela API, motorista e embarcador passam pelo de/para: chave conhecida → grava o nome padrão (ex.: um apelido aprovado vira o nome oficial); chave nova → grava como digitado e registra a grafia como padrão dela mesma.
- Ajustes futuros do de/para: Table Editor do Supabase (sem tela própria).
- A prévia da importação não sobrescreve um `nomes.csv` existente (revisão é trabalho humano).

### Importação do histórico (`python main.py importar <arquivo.xlsx>`)

Lê as abas de mês (ignora MODELO, RESULTADO, Sheet26). Os cabeçalhos mudam entre abas, então o mapeamento é por posição conforme o layout detectado:

- **Jan–Fev** (12 colunas): `TICKET, ENTREGA, CLIENTE, ACAREAÇÃO, STATUS, PROCEDÊNCIA, DATA DA RECLAMAÇÃO, MOTORISTA, EMBARCADOR, DATA DA RESOLUÇÃO, VALOR, DETALHE`.
- **Mar–Out** (9 colunas): `TICKET, ENTREGA, CLIENTE, ACAREAÇÃO, STATUS, PROCEDÊNCIA, MOTORISTA, EMBARCADOR, DETALHE` (cabeçalhos repetidos/errados como "TICKET, TICKET, TICKET" são ignorados — vale a posição).

Regras:

- Linha válida = coluna A numérica (ticket). Linhas vazias e linhas pré-preenchidas só com "Acareação Motorista" são ignoradas.
- Linha separadora "Dia NN" (Ago–Out) define o dia dos registros abaixo dela.
- Data: Jan–Fev usa a data da linha; Ago–Out usa o "Dia NN" corrente; Mar–Jul (sem dia) usa dia 1º do mês da aba. Ano: 2026.
- Jan–Fev: `VALOR` e `DATA DA RESOLUÇÃO` são anexados ao `detalhe` (`Valor: R$ 100,00 · Resolução: 01/03/2026`) — nada é perdido.
- Texto: `strip`, espaços internos colapsados; procedência/status normalizados para a grafia das listas (`procedente`→`Procedente`, `Analizando`→`Analisando`). Valor fora das listas → linha reportada como erro, import não grava nada.
- **Unificação de nomes em duas etapas:**
  1. `importar` sem `--aplicar` gera `nomes.csv` (`campo, original, sugerido, ocorrencias`), agrupando grafias iguais ignorando acento, caixa e espaços (ex.: `Motorista Graca`/`Motorista  Graca` → `Motorista Graça`; `DAFITTI` → `DAFITI`). A sugestão é a grafia mais frequente do grupo. Não grava nada no banco.
  2. O usuário revisa/edita `nomes.csv`; `importar --aplicar` lê o CSV, aplica os mapeamentos e grava tudo numa única transação (falhou → nada gravado).
- `criado_por` = `importacao`.

`nomes.csv` e a planilha contêm dados pessoais → ficam no `.gitignore`.

## 2. API (`api/index.py`)

Toda rota exige `Authorization: Bearer <token do Supabase>`. A API valida o token chamando `GET {SUPABASE_URL}/auth/v1/user` (mesma região, poucos ms) e lê `app_metadata.role`.

| Rota | Quem | Faz |
|---|---|---|
| `GET /api/acareacoes` | logado | todos os registros (ordenados por data desc). *shortcut: devolve tudo; ~1.200/ano cabe folgado — paginar se passar de ~20 mil* |
| `POST /api/acareacoes` | logado | cria; `criado_por` = usuário do token |
| `PUT /api/acareacoes/{id}` | logado | edita (404 se não existe) |
| `DELETE /api/acareacoes/{id}` | logado | exclui (frontend pede confirmação) |
| `POST /api/exportar` | logado | recebe lista de `id`s (o que está filtrado na tela) e devolve `.xlsx` |
| `GET /api/usuarios` | admin | lista usuários (nome, perfil, criado em) |
| `POST /api/usuarios` | admin | cria usuário (`usuario`, `senha`, `admin: bool`) via Supabase Admin API com `SUPABASE_SERVICE_ROLE_KEY` |
| `DELETE /api/usuarios/{id}` | admin | remove usuário (não pode remover a si mesmo) |

Erros: 401 sem/ token inválido, 403 não-admin em rota de admin, 422 validação — mensagens em português.

**Exportação `.xlsx`** (openpyxl): uma aba "Acareações" com colunas `DATA DA RECLAMAÇÃO, TICKET, Nº DA ENTREGA, CLIENTE, ACAREAÇÃO MOTORISTA, STATUS, PROCEDÊNCIA, MOTORISTA, EMBARCADOR, DETALHE`; cabeçalho em negrito, primeira linha congelada, autofiltro, data no formato `dd/mm/aaaa`, ticket/entrega como número inteiro. Nome do arquivo: `acareacoes_AAAA-MM-DD.xlsx`.

**Troca da própria senha:** feita no frontend com `supabase.auth.updateUser` — sem rota na API.

**Admin inicial:** `python main.py criar-admin nome.do.admin` (pede a senha no terminal, cria o usuário com `app_metadata.role = "admin"`). `app_metadata` só é gravável com a service role, então ninguém se promove sozinho.

## 3. Frontend

Estrutura herdada de `inspiracao/dashboard_rh-main/frontend` (menu lateral, cabeçalho com título + ações, grade de cards, KPIs com comparação ao período anterior, barras horizontais em HTML, tabelas, tema claro/escuro), **sem a barra de filtros no topo**. Mockup aprovado: `.superpowers/brainstorm/81-1792112496/content/layout-v2.html`.

### Telas

**Login** — usuário + senha; erro inline ("Usuário ou senha incorretos").

**Menu lateral** — marca "Acareações / HRAO · Motoristas" com a faixa multicolorida Magalu; itens Painel, Registros, Usuários (só admin); rodapé com usuário, perfil, trocar senha, tema claro/escuro e sair.

**Painel (página inicial)**
- Cabeçalho: título "Painel" + ações "Exportar .xlsx" (exporta os registros do período selecionado) e **"Nova acareação"** (primário).
- Abaixo do título: controle segmentado **Mês · Trimestre · Ano · Tudo** + navegação `‹ Maio de 2026 ›`. Padrão: mês atual.
- KPIs (cada um com comparação ao período anterior de mesmo tamanho; "Tudo" não compara):
  - Acareações (contagem)
  - Procedentes (% procedentes sobre o total; detalhe em p.p.)
  - Indenizados (status = Indenizado)
  - Motoristas em atenção (situação = Atenção)
- Cards: **Reclamações por motorista** (ranking em barras horizontais empilhadas Improcedente/Procedente, total e selo de situação — substitui a pizza 3D; mostra todos com rolagem), **Acareações por mês** (colunas, mês selecionado em destaque), **Status**, **Por embarcador** (top 4 + "Outros (n)", com "Ver todos (n)" que lista todos com nome e quantidade), **Últimos registros** (5 linhas + "Ver todos").

**Situação do motorista** — sobre a **média mensal** de reclamações do motorista no período (`total ÷ meses do período`; em "Mês" é o próprio total, idêntico à planilha):

| Média mensal | Situação |
|---|---|
| ≤ 1 | Ótimo |
| ≤ 3 | Bom |
| ≤ 5 | Regular |
| > 5 | Atenção |

Em "Tudo", meses = de janeiro/2026 até o mês atual. Selo sempre com texto + cor (nunca só cor).

**Nova / editar acareação** — painel que desliza da direita sobre véu escurecido; fecha com Esc, botão ✕ ou Cancelar; foco preso no painel. Campos: Data da reclamação (hoje), Ticket, Nº da entrega, Cliente, Motorista e Embarcador (autocompletar com nomes existentes via `<datalist>`), Acareação, Status, Procedência, Detalhe. Ticket já existente → aviso inline não bloqueante ("Ticket já cadastrado em 31/05 (Cliente). Pode salvar mesmo assim."). Validação inline nos obrigatórios. Salvar → fecha e os dados recarregam.

**Registros** — título com total; barra dentro do card: busca (ticket, entrega, cliente) + selects Período, Motorista, Embarcador, Status, Procedência; tabela com todas as colunas, ordenada por data desc, ações Editar (abre o painel) e Excluir (diálogo de confirmação). "Exportar N .xlsx" exporta exatamente as linhas filtradas.

**Usuários (admin)** — lista + formulário "Novo usuário" (usuário, senha, admin sim/não) + remover com confirmação.

### Dados no navegador

`dados.ts` busca todos os registros uma vez após o login e calcula no navegador: filtros, período anterior, KPIs, ranking, situação, agregações por mês/status/embarcador. Trocar filtro não chama o servidor. Após criar/editar/excluir, recarrega a lista.

### Visual (apple-design dentro da marca Magalu)

- **Cor** (tokens em `styles.css`, únicos lugares com cor):
  - Ação/acento: azul Magalu `#0086FF`; botões com texto branco usam `#006FD6` (contraste 4,95:1 — o `#0086FF` dá 3,59:1 e reprova AA).
  - Séries e etiquetas: Improcedente azul Magalu `#0086FF`, Procedente vermelho da faixa Magalu `#F0263C` (escuro `#FF453A`) — é o resultado ruim (pedido do usuário em 2026-10-16). A coluna Procedência das tabelas usa etiquetas nas mesmas cores; outras procedências em cinza. Validado com `dataviz/validate_palette.js` no modo claro (todas as checagens passam); o modo escuro tem passos próprios e será validado na implementação.
  - Situação: Ótimo verde, Bom neutro, Regular âmbar, Atenção vermelho — fundos suaves + texto escuro do mesmo tom.
  - Faixa multicolorida Magalu (amarelo, laranja, vermelho, rosa, roxo, azul, verde) só como assinatura sob a marca.
  - Neutros: fundo `#F5F5F7`, superfície `#FFFFFF`, texto `#1D1D1F`, secundário `#6E6E73`; tema escuro com passos próprios.
- **Tipografia**: fonte do sistema (`-apple-system, "Segoe UI Variable", "Segoe UI", system-ui`) — a fonte Magalu é proprietária. Títulos com tracking negativo (−0,022em em 34px), corpo em 0; hierarquia por peso; números tabulares.
- **Materiais**: menu lateral translúcido (`backdrop-filter`) mais pesado que o conteúdo; cards brancos com sombra leve sobre fundo cinza.
- **Movimento**: botões com feedback ao pressionar (`scale(.97)`); painel lateral entra e sai pela direita (mesmo caminho) com mola de pouco quique (amortecimento ~0,8, resposta ~0,3s) via CSS; véu com fade. `prefers-reduced-motion` → só fade. `prefers-reduced-transparency` → materiais sólidos.
- **Acessibilidade**: foco visível, rótulos em todos os campos, gráficos com `aria-label` resumindo os valores, identidade nunca só por cor, layout responsivo (menu vira faixa horizontal abaixo de 960px).

## 4. Arquivos

```
api/index.py            API inteira
main.py                 CLI: importar planilha, criar admin
schema.sql              tabela, CHECKs, RLS
tests/test_api.py       pytest
frontend/index.html
frontend/src/main.tsx
frontend/src/App.tsx         login, menu, tema, roteamento por hash
frontend/src/Painel.tsx
frontend/src/Registros.tsx
frontend/src/Usuarios.tsx
frontend/src/Formulario.tsx  painel lateral criar/editar
frontend/src/dados.ts        API + cálculos puros
frontend/src/dados.test.ts   vitest
frontend/src/styles.css
requirements.txt  vercel.json  README.md  .gitignore
```

`.gitignore`: `*.xlsx`, `nomes.csv`, `.env*`, `node_modules/`, `dist/`, `.superpowers/`, `inspiracao/`, `__pycache__/`.

## 5. Testes

- **Python (pytest)**:
  - importação: mapeamento por layout, "Dia NN", dia 1º em Mar–Jul, valor/resolução no detalhe, normalização, agrupamento de nomes, valor fora da lista falha sem gravar;
  - validação de campos (Pydantic);
  - autorização: sem token → 401; usuário comum em `/api/usuarios` → 403 (validação de token substituída por override de dependência);
  - exportação: abre o `.xlsx` gerado e confere cabeçalho, linhas, tipos.
  - CRUD contra banco: roda só se `DATABASE_URL_TESTE` estiver definida (senão pulado).
- **Frontend (vitest)**: situação por média mensal, ranking, período anterior, KPIs, filtros.
- **Verificação manual**: rodar local (Vite + uvicorn) apontando para o Supabase, abrir no navegador, cadastrar/editar/excluir/exportar, conferir números de maio contra a aba RESULTADO.

## 6. Publicação

1. Supabase (região São Paulo): rodar `schema.sql` no SQL Editor; copiar URL, anon key, service role key e connection string do **pooler (transaction, porta 6543)**.
2. Local: `.env` com as 4 variáveis → `python main.py importar ...` (revisar `nomes.csv`) → `--aplicar` → `python main.py criar-admin nome.do.admin`.
3. Vercel: importar o repositório GitHub; `vercel.json` define build do frontend, função `api/index.py`, região `gru1` e rewrite de `/api/*`. Variáveis: `VITE_SUPABASE_URL`, `VITE_SUPABASE_ANON_KEY` (públicas), `SUPABASE_URL`, `SUPABASE_SERVICE_ROLE_KEY`, `DATABASE_URL` (secretas).
4. Push na `main` → deploy automático.

## Fora do escopo (não pedido)

Histórico de alterações/auditoria por registro, anexos, notificações, tabelas de cadastro de motoristas/embarcadores, recuperação de senha por e-mail (admin redefine removendo e recriando, ou o próprio usuário troca logado), paginação no servidor, backups automáticos (plano gratuito do Supabase não oferece; exportação `.xlsx` serve de backup manual).

## Riscos conhecidos

- Mar–Jul do histórico ficam no dia 1º do mês (a planilha não tem o dia).
- Repositório público: nenhum dado real pode ser commitado; segredos só nas variáveis da Vercel/`.env`.
- Primeira chamada após inatividade pode levar 1–2 s (cold start da função Python).
