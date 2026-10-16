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
   Abra `nomes.csv` no Excel, confira a coluna `sugerido` (apague a sugestão para manter o nome original; escreva outro nome para juntar grafias), salve como **CSV UTF-8** e rode:
   ```
   python main.py importar "ACAREAÇÕES HRAO.xlsx" --aplicar
   ```
   A importação também grava o **de/para** de nomes (tabela `apelidos`): ao cadastrar, "fulano graca" ou um apelido aprovado vira o nome padrão. Para corrigir um de/para depois, edite a tabela `apelidos` no Table Editor do Supabase.
4. **Criar a admin**: `python main.py criar-admin nome.do.admin` (pede a senha).
5. **Vercel**: Add New → Project → importar o repositório do GitHub, cadastrar as 5 variáveis e fazer o deploy. Cada push na `main` publica sozinho.

## Rodar localmente

```
uvicorn api.index:app --env-file .env --port 8000 --reload
cd frontend && npm install && cd ..
npm run dev --prefix frontend
```
Abrir http://localhost:5173.

## Testes

```
python -m pytest tests
npm test --prefix frontend
```
Para testar gravação no banco, defina `DATABASE_URL_TESTE` (cria e apaga um registro).

## Regras

- **Situação do motorista**: média mensal de reclamações no período — até 1 Ótimo, até 3 Bom, até 5 Regular, acima de 5 Atenção.
- **Backup**: o plano gratuito do Supabase não tem backup para baixar; exporte o `.xlsx` em Registros periodicamente.
