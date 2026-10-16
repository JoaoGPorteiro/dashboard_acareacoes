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
