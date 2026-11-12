import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { TRPCError } from "@trpc/server";

vi.mock("./db", () => ({ getDb: vi.fn() }));

import { getDb } from "./db";
import {
  ImportacaoDecisoesError,
  erroSeguroImportacaoDecisoes,
  importarDecisoesGlosaPlanilha,
  linhasAfetadas,
  type LinhaDecisaoPlanilha,
} from "./importDecisoesGlosaPlanilhaService";

/**
 * Banco sintético em memória com transação real (staging + commit/rollback),
 * suficiente para provar atomicidade, contagem de linhas afetadas e escopo.
 * Sem MySQL, sem rede, sem dados reais.
 */
type Item = {
  id: number; estabelecimentoId: number; convenioId: number; protocolo: string; numeroGuia: string;
  sequencialItem: number | null; codigoItem: string; valorGlosa: string; codigoGlosa: string; tipoLancamento: string | null;
  classificacaoGlosa: string | null; classificacaoMotivo: string | null; motivoAceite: string | null;
};

function criarBanco(itens: Item[]) {
  const estado = { itens: itens.map((i) => ({ ...i })), historico: [] as any[] };
  const falhas = { affectedRows: [] as number[], insertFalhaNa: 0, updateLancaNa: 0 };
  const contadores = { transacoes: 0, commits: 0, rollbacks: 0, selects: 0 };
  let escopoSelect: { estabelecimentoId: number; convenioId: number | null } = { estabelecimentoId: 0, convenioId: null };

  const db = {
    /** O teste informa o escopo consultado (o fake não interpreta SQL). */
    _escopo(e: typeof escopoSelect) { escopoSelect = e; },
    select() {
      return {
        from() {
          return {
            where: async () => {
              contadores.selects++;
              return estado.itens
                .filter((i) => i.estabelecimentoId === escopoSelect.estabelecimentoId)
                .filter((i) => escopoSelect.convenioId === null || i.convenioId === escopoSelect.convenioId)
                .map((i) => ({ ...i }));
            },
          };
        },
      };
    },
    async transaction<T>(fn: (tx: any) => Promise<T>): Promise<T> {
      contadores.transacoes++;
      const staging = { updates: [] as Array<{ id: number; set: any }>, inserts: [] as any[] };
      let nUpdate = 0;
      let nInsert = 0;
      const tx = {
        update() {
          return {
            set(set: any) {
              return {
                where: async () => {
                  nUpdate++;
                  if (falhas.updateLancaNa === nUpdate) throw new Error("ER_LOCK_DEADLOCK: Deadlock found; host=db-interno.safatle:3306 user=app_rw");
                  const affected = falhas.affectedRows[nUpdate - 1] ?? 1;
                  // Simula o item-alvo pela ordem do plano (os testes usam ids crescentes).
                  if (affected > 0) staging.updates.push({ id: nUpdate, set });
                  return [{ affectedRows: affected }, []];
                },
              };
            },
          };
        },
        insert() {
          return {
            values: async (v: any) => {
              nInsert++;
              if (falhas.insertFalhaNa === nInsert) throw new Error("ER_NO_REFERENCED_ROW_2: Cannot add or update a child row: a foreign key constraint fails (`safatle_prod`.`decisoes_glosa`)");
              staging.inserts.push(v);
            },
          };
        },
      };
      try {
        const r = await fn(tx);
        // COMMIT: aplica updates (na ordem dos procedimentoId inseridos) e histórico.
        staging.inserts.forEach((h, idx) => {
          const alvo = estado.itens.find((i) => i.id === h.procedimentoId)!;
          Object.assign(alvo, staging.updates[idx].set);
        });
        estado.historico.push(...staging.inserts);
        contadores.commits++;
        return r;
      } catch (e) {
        contadores.rollbacks++;
        throw e;
      }
    },
  };
  return { db, estado, falhas, contadores };
}

const itemBase = (p: Partial<Item> & { id: number }): Item => ({
  estabelecimentoId: 1, convenioId: 10, protocolo: "P1", numeroGuia: "G1", sequencialItem: null,
  codigoItem: "C1", valorGlosa: "100.00", codigoGlosa: "1801 - x", tipoLancamento: null,
  classificacaoGlosa: "pendente", classificacaoMotivo: null, motivoAceite: null, ...p,
});

const linhas: LinhaDecisaoPlanilha[] = [
  { linha: 4, decisao: "Sim", protocolo: "P1", guiaOperadora: "G1", codigoItem: "C1", valorGlosa: "100", justificativa: "Recorrer" },
  { linha: 5, decisao: "Não", protocolo: "P1", guiaOperadora: "G1", codigoItem: "C2", valorGlosa: "50", justificativa: "Orientação", observacao: "Glosa Aceita Pago conforme Contrato" },
];

const itens = () => [itemBase({ id: 101 }), itemBase({ id: 102, codigoItem: "C2", valorGlosa: "50.00" })];

let banco: ReturnType<typeof criarBanco>;
let consoleError: ReturnType<typeof vi.spyOn>;

function usar(b: ReturnType<typeof criarBanco>) {
  banco = b;
  vi.mocked(getDb).mockResolvedValue(b.db as any);
}

async function previa(estabelecimentoId = 1, convenioId: number | null = 10, ls = linhas) {
  banco.db._escopo({ estabelecimentoId, convenioId });
  return importarDecisoesGlosaPlanilha({ estabelecimentoId, convenioId: convenioId ?? undefined, linhas: ls, aplicar: false, userId: 7 });
}
async function aplicar(tokenPrevia: string | null | undefined, estabelecimentoId = 1, convenioId: number | null = 10, ls = linhas) {
  banco.db._escopo({ estabelecimentoId, convenioId });
  return importarDecisoesGlosaPlanilha({ estabelecimentoId, convenioId: convenioId ?? undefined, linhas: ls, aplicar: true, tokenPrevia, userId: 7 });
}

beforeEach(() => {
  usar(criarBanco(itens()));
  consoleError = vi.spyOn(console, "error").mockImplementation(() => {});
});
afterEach(() => {
  consoleError.mockRestore();
  vi.mocked(getDb).mockReset();
});

describe("importarDecisoesGlosaPlanilha — prévia", () => {
  it("devolve escopo e token sem gravar nada", async () => {
    const r = await previa();
    expect(r).toMatchObject({ casados: 2, aplicados: 0, escopo: { estabelecimentoId: 1, convenioId: 10 } });
    expect(r.tokenPrevia).toMatch(/^[0-9a-f]{64}$/);
    expect(banco.contadores.transacoes).toBe(0);
    expect(banco.estado.historico).toHaveLength(0);
  });

  it("token muda com o convênio e com o hospital (mesma planilha)", async () => {
    usar(criarBanco([...itens(), itemBase({ id: 201, estabelecimentoId: 2 }), itemBase({ id: 202, estabelecimentoId: 2, codigoItem: "C2", valorGlosa: "50.00" })]));
    const a = await previa(1, 10);
    const b = await previa(1, null);
    const c = await previa(2, null);
    expect(new Set([a.tokenPrevia, b.tokenPrevia, c.tokenPrevia]).size).toBe(3);
  });
});

describe("importarDecisoesGlosaPlanilha — confirmação vinculada à prévia", () => {
  it("exige token de prévia para aplicar", async () => {
    await expect(aplicar(undefined)).rejects.toMatchObject({ codigo: "PREVIA_OBRIGATORIA" });
    expect(banco.contadores.selects).toBe(0);
    expect(banco.contadores.transacoes).toBe(0);
  });

  it("recusa prévia de outro convênio aplicada no novo recorte, sem gravar", async () => {
    const p = await previa(1, 10);
    await expect(aplicar(p.tokenPrevia, 1, null)).rejects.toMatchObject({ codigo: "PREVIA_DESATUALIZADA" });
    expect(banco.contadores.transacoes).toBe(0);
    expect(banco.estado.historico).toHaveLength(0);
  });

  it("recusa prévia de outro hospital, mesmo com itens equivalentes, sem gravar", async () => {
    usar(criarBanco([...itens(), itemBase({ id: 201, estabelecimentoId: 2 }), itemBase({ id: 202, estabelecimentoId: 2, codigoItem: "C2", valorGlosa: "50.00" })]));
    const p = await previa(1, 10);
    await expect(aplicar(p.tokenPrevia, 2, 10)).rejects.toMatchObject({ codigo: "PREVIA_DESATUALIZADA" });
    expect(banco.contadores.transacoes).toBe(0);
    expect(banco.estado.itens.every((i) => i.classificacaoGlosa === "pendente")).toBe(true);
  });

  it("recusa quando um item vinculado mudou de versão após a prévia", async () => {
    const p = await previa();
    banco.estado.itens[1].classificacaoGlosa = "recursar";
    await expect(aplicar(p.tokenPrevia)).rejects.toMatchObject({ codigo: "PREVIA_DESATUALIZADA" });
    expect(banco.contadores.transacoes).toBe(0);
  });

  it("recusa quando a planilha confirmada difere da analisada", async () => {
    const p = await previa();
    const outra = [{ ...linhas[0], decisao: "Não" }, linhas[1]];
    await expect(aplicar(p.tokenPrevia, 1, 10, outra)).rejects.toMatchObject({ codigo: "PREVIA_DESATUALIZADA" });
  });
});

describe("importarDecisoesGlosaPlanilha — gravação atômica e persistência real", () => {
  it("aplica decisão + histórico de todos os itens numa única transação", async () => {
    const p = await previa();
    const r = await aplicar(p.tokenPrevia);
    expect(r.aplicados).toBe(2);
    expect(banco.contadores).toMatchObject({ transacoes: 1, commits: 1, rollbacks: 0 });
    expect(banco.estado.historico.map((h) => [h.procedimentoId, h.decisao])).toEqual([[101, "recursar"], [102, "aceitar"]]);
    expect(banco.estado.itens.find((i) => i.id === 102)).toMatchObject({ classificacaoGlosa: "aceitar", motivoAceite: "Glosa Aceita Pago conforme Contrato" });
  });

  it("UPDATE sem linha afetada não conta como aplicado: reverte tudo e informa conflito", async () => {
    const p = await previa();
    banco.falhas.affectedRows = [1, 0];
    const erro = await aplicar(p.tokenPrevia).catch((e) => e);
    expect(erro).toBeInstanceOf(ImportacaoDecisoesError);
    expect(erro.codigo).toBe("CONFLITO_ITEM");
    expect(banco.contadores).toMatchObject({ commits: 0, rollbacks: 1 });
    expect(banco.estado.historico).toHaveLength(0);
    expect(banco.estado.itens.every((i) => i.classificacaoGlosa === "pendente")).toBe(true);
  });

  it("falha no histórico não deixa decisão gravada nem planilha parcialmente aplicada", async () => {
    const p = await previa();
    banco.falhas.insertFalhaNa = 2;
    const erro = await aplicar(p.tokenPrevia).catch((e) => e);
    expect(erro.codigo).toBe("FALHA_GRAVACAO");
    expect(banco.contadores).toMatchObject({ commits: 0, rollbacks: 1 });
    expect(banco.estado.historico).toHaveLength(0);
    expect(banco.estado.itens.every((i) => i.classificacaoGlosa === "pendente")).toBe(true);
    // Mensagem ao cliente é genérica; detalhe técnico só no log do servidor.
    expect(erro.message).not.toMatch(/ER_|constraint|safatle_prod|decisoes_glosa/);
    expect(consoleError).toHaveBeenCalled();
  });

  it("falha no UPDATE posterior reverte os itens anteriores e sanitiza a mensagem", async () => {
    const p = await previa();
    banco.falhas.updateLancaNa = 2;
    const erro = await aplicar(p.tokenPrevia).catch((e) => e);
    expect(erro.codigo).toBe("FALHA_GRAVACAO");
    expect(erro.message).not.toMatch(/Deadlock|host=|user=|3306/);
    expect(banco.estado.historico).toHaveLength(0);
    expect(banco.estado.itens.every((i) => i.classificacaoGlosa === "pendente")).toBe(true);
  });

  it("banco indisponível gera erro seguro", async () => {
    vi.mocked(getDb).mockResolvedValue(null as any);
    await expect(importarDecisoesGlosaPlanilha({ estabelecimentoId: 1, linhas, aplicar: false, userId: 7 })).rejects.toMatchObject({ codigo: "BANCO_INDISPONIVEL" });
  });
});

describe("linhasAfetadas", () => {
  it("lê affectedRows de mysql2 ([header, fields]) e de objeto simples; inválido = 0", () => {
    expect(linhasAfetadas([{ affectedRows: 1 }, []])).toBe(1);
    expect(linhasAfetadas({ affectedRows: 3 })).toBe(3);
    expect(linhasAfetadas([{}])).toBe(0);
    expect(linhasAfetadas(undefined)).toBe(0);
    expect(linhasAfetadas({ affectedRows: "1" })).toBe(0);
  });
});

describe("erroSeguroImportacaoDecisoes", () => {
  it("mapeia erros de domínio para códigos tRPC com mensagem segura", () => {
    expect(erroSeguroImportacaoDecisoes(new ImportacaoDecisoesError("PREVIA_DESATUALIZADA")).code).toBe("CONFLICT");
    expect(erroSeguroImportacaoDecisoes(new ImportacaoDecisoesError("CONFLITO_ITEM")).code).toBe("CONFLICT");
    expect(erroSeguroImportacaoDecisoes(new ImportacaoDecisoesError("PREVIA_OBRIGATORIA")).code).toBe("BAD_REQUEST");
    expect(erroSeguroImportacaoDecisoes(new ImportacaoDecisoesError("FALHA_GRAVACAO")).code).toBe("INTERNAL_SERVER_ERROR");
  });

  it("preserva erros de permissão já tipados", () => {
    const proibido = new TRPCError({ code: "FORBIDDEN", message: "Sem acesso ao estabelecimento" });
    expect(erroSeguroImportacaoDecisoes(proibido)).toBe(proibido);
  });

  it("nunca repassa mensagem de erro inesperado do driver/serviço", () => {
    const e = erroSeguroImportacaoDecisoes(new Error("connect ECONNREFUSED 10.0.0.5:3306 password=segredo"));
    expect(e.code).toBe("INTERNAL_SERVER_ERROR");
    expect(e.message).not.toMatch(/ECONNREFUSED|10\.0\.0\.5|segredo/);
    const interno = erroSeguroImportacaoDecisoes(new TRPCError({ code: "INTERNAL_SERVER_ERROR", message: "ER_BAD_FIELD_ERROR coluna x" }));
    expect(interno.message).not.toMatch(/ER_BAD_FIELD/);
  });
});
