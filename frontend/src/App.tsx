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
