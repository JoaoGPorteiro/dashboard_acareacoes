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
                  <td><span className="pill" data-procedencia={r.procedencia}>{r.procedencia}</span></td>
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
