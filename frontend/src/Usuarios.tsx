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
