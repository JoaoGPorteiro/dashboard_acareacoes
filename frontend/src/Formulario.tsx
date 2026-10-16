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
          <input type="date" name="data_reclamacao" required min="2026-01-01" max={hoje()} defaultValue={r?.data_reclamacao ?? hoje()} />
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
