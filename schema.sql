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

-- De/para de nomes: (campo, chave sem acento/caixa/espaços) → nome padrão.
-- A API consulta ao salvar; ajustes manuais pelo Table Editor do Supabase.
create table if not exists apelidos (
  campo text not null check (campo in ('motorista', 'embarcador')),
  chave text not null check (chave <> ''),
  nome text not null check (nome <> ''),
  primary key (campo, chave)
);

alter table apelidos enable row level security;
