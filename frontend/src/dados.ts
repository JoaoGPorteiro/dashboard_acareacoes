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
