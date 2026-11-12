import { describe, expect, it } from "vitest";
import {
  escopoDaSelecao,
  exigeNovaAnalise,
  mensagemErroSegura,
  mesmoEscopo,
  previaAplicavel,
  respostaAindaValida,
} from "./importarDecisoesGlosaPrevia";

const token = "a".repeat(64);
const linhas = [{ linha: 4 }];

describe("importarDecisoesGlosaPrevia — escopo", () => {
  it("converte a seleção da tela em escopo", () => {
    expect(escopoDaSelecao(1, "todos")).toEqual({ estabelecimentoId: 1, convenioId: null });
    expect(escopoDaSelecao(1, "10")).toEqual({ estabelecimentoId: 1, convenioId: 10 });
    expect(escopoDaSelecao(undefined, "10")).toBeNull();
  });

  it("compara hospital e convênio", () => {
    expect(mesmoEscopo({ estabelecimentoId: 1, convenioId: 10 }, { estabelecimentoId: 1, convenioId: 10 })).toBe(true);
    expect(mesmoEscopo({ estabelecimentoId: 1, convenioId: 10 }, { estabelecimentoId: 1, convenioId: null })).toBe(false);
    expect(mesmoEscopo({ estabelecimentoId: 1, convenioId: 10 }, { estabelecimentoId: 2, convenioId: 10 })).toBe(false);
    expect(mesmoEscopo(null, null)).toBe(false);
  });
});

describe("importarDecisoesGlosaPrevia — prévia aplicável", () => {
  const previa = { seq: 1, escopo: { estabelecimentoId: 1, convenioId: 10 }, linhas, tokenPrevia: token, dados: { casados: 2 } };

  it("só vale para o mesmo escopo e a mesma extração", () => {
    expect(previaAplicavel(previa, { estabelecimentoId: 1, convenioId: 10 }, linhas)).toBe(true);
    expect(previaAplicavel(previa, { estabelecimentoId: 2, convenioId: 10 }, linhas)).toBe(false);
    expect(previaAplicavel(previa, { estabelecimentoId: 1, convenioId: null }, linhas)).toBe(false);
    expect(previaAplicavel(previa, { estabelecimentoId: 1, convenioId: 10 }, [{ linha: 4 }])).toBe(false);
    expect(previaAplicavel(previa, null, linhas)).toBe(false);
    expect(previaAplicavel({ ...previa, tokenPrevia: "x" }, { estabelecimentoId: 1, convenioId: 10 }, linhas)).toBe(false);
    expect(previaAplicavel(null, { estabelecimentoId: 1, convenioId: 10 }, linhas)).toBe(false);
  });

  it("descarta resposta atrasada de outra análise ou de outro escopo", () => {
    const e = { estabelecimentoId: 1, convenioId: 10 };
    expect(respostaAindaValida(3, 3, e, e)).toBe(true);
    expect(respostaAindaValida(2, 3, e, e)).toBe(false);
    expect(respostaAindaValida(3, 3, e, { estabelecimentoId: 2, convenioId: 10 })).toBe(false);
  });
});

describe("importarDecisoesGlosaPrevia — mensagens seguras", () => {
  const fallback = "Mensagem genérica";
  it("repassa apenas mensagens de domínio já sanitizadas", () => {
    expect(mensagemErroSegura({ message: "Reanalise a planilha.", data: { code: "CONFLICT" } }, fallback)).toBe("Reanalise a planilha.");
    expect(mensagemErroSegura({ message: "Sem acesso", data: { code: "FORBIDDEN" } }, fallback)).toBe("Sem acesso");
  });
  it("oculta erro interno, de rede ou desconhecido", () => {
    expect(mensagemErroSegura({ message: "ER_DUP_ENTRY host=db", data: { code: "INTERNAL_SERVER_ERROR" } }, fallback)).toBe(fallback);
    expect(mensagemErroSegura(new Error("Failed to fetch"), fallback)).toBe(fallback);
    expect(mensagemErroSegura({ message: "x".repeat(400), data: { code: "BAD_REQUEST" } }, fallback)).toBe(fallback);
    expect(mensagemErroSegura(undefined, fallback)).toBe(fallback);
  });
  it("conflito exige nova análise", () => {
    expect(exigeNovaAnalise({ data: { code: "CONFLICT" } })).toBe(true);
    expect(exigeNovaAnalise({ data: { code: "INTERNAL_SERVER_ERROR" } })).toBe(false);
  });
});
