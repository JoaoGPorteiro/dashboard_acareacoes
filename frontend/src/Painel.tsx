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
  const [todosEmbarcadores, setTodosEmbarcadores] = useState(false);

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
  const embarcadores = contar(doPeriodo, "embarcador");

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
          <Barras itens={todosEmbarcadores ? embarcadores : contar(doPeriodo, "embarcador", 4)} total={doPeriodo.length} todos={todosEmbarcadores} />
          {embarcadores.length > 4 && (
            <button type="button" className="link ver-todos" aria-expanded={todosEmbarcadores} onClick={() => setTodosEmbarcadores(!todosEmbarcadores)}>
              {todosEmbarcadores ? "Ver menos" : `Ver todos (${embarcadores.length})`}
            </button>
          )}
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
                      <td><span className="pill" data-procedencia={r.procedencia}>{r.procedencia}</span></td>
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

function Barras({ itens, total, todos }: { itens: Item[]; total: number; todos?: boolean }) {
  const max = Math.max(0, ...itens.map((i) => i.valor));
  if (!itens.length) return <p className="estado">Nenhuma acareação neste período.</p>;
  return (
    <ul className={todos ? "bh todos" : "bh"}>
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
