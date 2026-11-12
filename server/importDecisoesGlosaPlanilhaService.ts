/**
 * Importação de decisões de glosa (aceitar / recursar) a partir de planilha Excel
 * de consultoria (ex.: aba "8. Oportunidades Pendentes").
 *
 * Regras de negócio:
 *  - Coluna "VAMOS RECURSAR? (Sim / Não / Avaliar)":
 *      Sim     -> classificação "recursar"
 *      Não     -> classificação "aceitar" (Glosa Aceita)
 *      Avaliar / vazio -> ignorado (permanece como está no portal)
 *  - Coluna "Justificativa / Orientação para Recurso" -> justificativa gravada no item
 *    (classificacaoMotivo). Na Glosa Aceita, o motivoAceite usa a coluna
 *    "Sugestão da consultoria" (motivo real da equipe), com fallback para a justificativa.
 *
 * O vínculo com o item glosado do portal (tabela demonstrativo) é feito por:
 *   Guia (operadora OU prestador) + Código do item (obrigatórios)
 *   Protocolo (quando ambos possuem, precisa ser igual), Seq. do item e Valor glosado
 *   são usados para desempate / pontuação.
 */
import { createHash } from "node:crypto";
import { TRPCError } from "@trpc/server";
import { and, eq, inArray, sql } from "drizzle-orm";
import { getDb } from "./db";
import { decisoesGlosa, demonstrativo } from "../drizzle/schema";

export type DecisaoPlanilha = "aceitar" | "recursar";

export interface LinhaDecisaoPlanilha {
  /** Linha original na planilha (1-based), apenas para relatório */
  linha: number;
  decisao: string;
  protocolo?: string | null;
  guiaPrestador?: string | null;
  guiaOperadora?: string | null;
  sequencialItem?: string | number | null;
  codigoItem?: string | null;
  descricaoItem?: string | null;
  valorGlosa?: string | number | null;
  justificativa?: string | null;
  /** "Sugestão da consultoria" (coluna C) — motivo real da equipe, usado na Glosa Aceita */
  observacao?: string | null;
}

export interface CandidatoDemonstrativo {
  id: number;
  protocolo: string | null;
  numeroGuia: string | null;
  sequencialItem: number | null;
  codigoItem: string | null;
  valorGlosa: string | number | null;
}

export type StatusCasamento = "casado" | "nao_encontrado" | "decisao_invalida";

export interface ResultadoCasamento {
  linha: number;
  status: StatusCasamento;
  decisao: DecisaoPlanilha | null;
  demonstrativoId: number | null;
  pontuacao: number;
  detalhe: string;
}

const s = (v: unknown) => (v === null || v === undefined ? "" : String(v).trim());

/**
 * Monta os textos gravados no item:
 *  - Recursar: motivo = Justificativa / Orientação para Recurso.
 *  - Glosa Aceita: motivoAceite = Sugestão da consultoria (coluna C, motivo real da equipe),
 *    com fallback para a Justificativa; o motivo da classificação guarda as duas juntas.
 */
export function montarMotivosDecisao(
  decisao: DecisaoPlanilha,
  justificativa: unknown,
  observacao: unknown,
): { motivo: string; motivoAceite: string | null } {
  const just = s(justificativa);
  const obs = s(observacao);
  if (decisao === "recursar") {
    return { motivo: just || "Recursar (importado da planilha)", motivoAceite: null };
  }
  const motivoAceite = obs || just || "Glosa aceita (importado da planilha)";
  const motivo = obs && just && obs !== just ? `${obs}\nOrientação: ${just}` : motivoAceite;
  return { motivo, motivoAceite };
}

/** Sim -> recursar, Não -> aceitar; qualquer outro valor -> null (ignorar). */
export function normalizarDecisaoPlanilha(valor: unknown): DecisaoPlanilha | null {
  const v = s(valor)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
  if (v === "sim" || v === "s") return "recursar";
  if (v === "nao" || v === "n") return "aceitar";
  return null;
}

/** Normaliza identificadores numéricos (guia, protocolo, código): só alfanuméricos, sem zeros à esquerda. */
export function normalizarIdentificador(valor: unknown): string {
  return s(valor)
    .replace(/[^0-9a-zA-Z]/g, "")
    .replace(/^0+(?=.)/, "")
    .toUpperCase();
}

function paraNumero(valor: unknown): number | null {
  const t = s(valor);
  if (!t) return null;
  const n = Number(t.includes(",") ? t.replace(/\./g, "").replace(",", ".") : t);
  return Number.isFinite(n) ? n : null;
}

/**
 * Casa as linhas da planilha com os itens glosados do portal.
 * Cada item do portal é usado no máximo uma vez (linhas duplicadas na planilha
 * consomem itens duplicados no portal na ordem de melhor pontuação).
 */
export function casarDecisoesComDemonstrativo(
  linhas: LinhaDecisaoPlanilha[],
  candidatos: CandidatoDemonstrativo[],
): ResultadoCasamento[] {
  // Indexa candidatos por guia normalizada + código normalizado
  const indice = new Map<string, CandidatoDemonstrativo[]>();
  for (const c of candidatos) {
    const chave = `${normalizarIdentificador(c.numeroGuia)}|${normalizarIdentificador(c.codigoItem)}`;
    const lista = indice.get(chave);
    if (lista) lista.push(c);
    else indice.set(chave, [c]);
  }

  const consumidos = new Set<number>();
  return linhas.map((linha): ResultadoCasamento => {
    const decisao = normalizarDecisaoPlanilha(linha.decisao);
    const base = { linha: linha.linha, decisao, demonstrativoId: null, pontuacao: 0 };
    if (!decisao) {
      return { ...base, status: "decisao_invalida", detalhe: `Decisão "${s(linha.decisao)}" ignorada (use Sim ou Não)` };
    }

    const codigo = normalizarIdentificador(linha.codigoItem);
    const guias = Array.from(
      new Set([normalizarIdentificador(linha.guiaOperadora), normalizarIdentificador(linha.guiaPrestador)].filter(Boolean)),
    );
    if (!codigo || guias.length === 0) {
      return { ...base, status: "nao_encontrado", detalhe: "Linha sem guia ou código do item" };
    }

    const protocolo = normalizarIdentificador(linha.protocolo);
    const seq = paraNumero(linha.sequencialItem);
    const valor = paraNumero(linha.valorGlosa);

    let melhor: CandidatoDemonstrativo | null = null;
    let melhorPontuacao = -1;
    for (const guia of guias) {
      for (const c of indice.get(`${guia}|${codigo}`) || []) {
        if (consumidos.has(c.id)) continue;
        const protocoloCand = normalizarIdentificador(c.protocolo);
        if (protocolo && protocoloCand && protocolo !== protocoloCand) continue;
        let pontos = 1;
        if (protocolo && protocoloCand) pontos += 4;
        const valorCand = paraNumero(c.valorGlosa);
        if (valor !== null && valorCand !== null && Math.abs(valor - valorCand) < 0.01) pontos += 3;
        if (seq !== null && c.sequencialItem !== null && Number(c.sequencialItem) === seq) pontos += 2;
        if (pontos > melhorPontuacao || (pontos === melhorPontuacao && melhor && c.id < melhor.id)) {
          melhor = c;
          melhorPontuacao = pontos;
        }
      }
    }

    if (!melhor) {
      return { ...base, status: "nao_encontrado", detalhe: "Item glosado não localizado no portal (guia + código)" };
    }
    consumidos.add(melhor.id);
    return {
      ...base,
      status: "casado",
      demonstrativoId: melhor.id,
      pontuacao: melhorPontuacao,
      detalhe: melhorPontuacao >= 8 ? "Vínculo exato" : melhorPontuacao >= 4 ? "Vínculo provável" : "Vínculo por guia + código",
    };
  });
}

export interface ImportarDecisoesInput {
  estabelecimentoId: number;
  convenioId?: number;
  linhas: LinhaDecisaoPlanilha[];
  aplicar: boolean;
  userId: number;
  /**
   * Token devolvido pela prévia (aplicar=false). Obrigatório para aplicar:
   * vincula a confirmação ao mesmo escopo (hospital/convênio), à mesma planilha
   * e à mesma versão dos itens analisados.
   */
  tokenPrevia?: string | null;
}

export interface ImportarDecisoesResultado {
  totalLinhas: number;
  casados: number;
  naoEncontrados: number;
  ignorados: number;
  aRecursar: number;
  aAceitar: number;
  aplicados: number;
  /** Escopo exato usado na análise (o cliente só habilita a aplicação para este escopo). */
  escopo: { estabelecimentoId: number; convenioId: number | null };
  /** Assinatura da prévia: escopo + planilha + versão dos itens vinculados. */
  tokenPrevia: string;
  resultados: Array<ResultadoCasamento & { codigoItem: string; guia: string; descricaoItem: string; classificacaoAtual: string | null }>;
}

/** Códigos de erro seguros (mensagem pode ser exibida ao usuário). */
export type CodigoErroImportacaoDecisoes =
  | "BANCO_INDISPONIVEL"
  | "PREVIA_OBRIGATORIA"
  | "PREVIA_DESATUALIZADA"
  | "NADA_A_APLICAR"
  | "CONFLITO_ITEM"
  | "FALHA_GRAVACAO";

const MENSAGENS_ERRO: Record<CodigoErroImportacaoDecisoes, string> = {
  BANCO_INDISPONIVEL: "Serviço temporariamente indisponível. Tente novamente em instantes.",
  PREVIA_OBRIGATORIA: "Analise a planilha antes de aplicar as decisões.",
  PREVIA_DESATUALIZADA:
    "A prévia não corresponde mais ao hospital/convênio selecionado ou os itens foram alterados. Reanalise a planilha antes de aplicar.",
  NADA_A_APLICAR: "Nenhum item vinculado para aplicar.",
  CONFLITO_ITEM:
    "Um ou mais itens foram alterados por outra operação durante a aplicação. Nenhuma decisão foi gravada; reanalise a planilha.",
  FALHA_GRAVACAO: "Não foi possível aplicar as decisões. Nenhuma alteração foi gravada.",
};

/** Erro de domínio com mensagem sanitizada (nunca carrega texto do banco/driver). */
export class ImportacaoDecisoesError extends Error {
  readonly codigo: CodigoErroImportacaoDecisoes;
  constructor(codigo: CodigoErroImportacaoDecisoes) {
    super(MENSAGENS_ERRO[codigo]);
    this.name = "ImportacaoDecisoesError";
    this.codigo = codigo;
  }
}

/** Linhas afetadas de um UPDATE (mysql2 devolve [ResultSetHeader, fields]). */
export function linhasAfetadas(resultado: unknown): number {
  const header = Array.isArray(resultado) ? resultado[0] : resultado;
  const n = (header as { affectedRows?: unknown } | null | undefined)?.affectedRows;
  return typeof n === "number" && Number.isFinite(n) ? n : 0;
}

type CandidatoCompleto = CandidatoDemonstrativo & {
  convenioId: number | null;
  codigoGlosa: string | null;
  tipoLancamento: string | null;
  classificacaoGlosa: string | null;
  classificacaoMotivo: string | null;
  motivoAceite: string | null;
};

/**
 * Assinatura determinística da prévia. Muda se mudar o escopo, a planilha
 * ou o estado (classificação/motivos/valor) de qualquer item vinculado.
 */
export function calcularTokenPrevia(
  escopo: { estabelecimentoId: number; convenioId: number | null },
  linhas: LinhaDecisaoPlanilha[],
  casamentos: ResultadoCasamento[],
  porId: Map<number, CandidatoCompleto>,
): string {
  const plano = casamentos
    .filter((r) => r.status === "casado" && r.demonstrativoId && r.decisao)
    .map((r) => {
      const c = porId.get(r.demonstrativoId!)!;
      return [r.linha, c.id, r.decisao, c.classificacaoGlosa ?? null, c.classificacaoMotivo ?? null, c.motivoAceite ?? null, s(c.valorGlosa)];
    });
  const hashLinhas = createHash("sha256").update(JSON.stringify(linhas)).digest("hex");
  return createHash("sha256")
    .update(JSON.stringify({ v: 1, e: escopo.estabelecimentoId, c: escopo.convenioId, l: hashLinhas, p: plano }))
    .digest("hex");
}

const TAMANHO_LOTE = 500;

export async function importarDecisoesGlosaPlanilha(input: ImportarDecisoesInput): Promise<ImportarDecisoesResultado> {
  const db = await getDb();
  if (!db) throw new ImportacaoDecisoesError("BANCO_INDISPONIVEL");
  if (input.aplicar && !input.tokenPrevia) throw new ImportacaoDecisoesError("PREVIA_OBRIGATORIA");

  const escopo = { estabelecimentoId: input.estabelecimentoId, convenioId: input.convenioId ?? null };

  // Busca candidatos pelas guias presentes na planilha (valores originais e normalizados)
  const guias = new Set<string>();
  for (const l of input.linhas) {
    for (const g of [l.guiaOperadora, l.guiaPrestador]) {
      const t = s(g);
      if (t) {
        guias.add(t);
        guias.add(normalizarIdentificador(t));
      }
    }
  }
  const listaGuias = Array.from(guias).filter(Boolean);

  const candidatos: CandidatoCompleto[] = [];
  for (let i = 0; i < listaGuias.length; i += TAMANHO_LOTE) {
    const lote = listaGuias.slice(i, i + TAMANHO_LOTE);
    const condicoes = [
      eq(demonstrativo.estabelecimentoId, input.estabelecimentoId),
      inArray(demonstrativo.numeroGuia, lote),
      sql`COALESCE(${demonstrativo.valorGlosa}, 0) > 0`,
    ];
    if (input.convenioId) condicoes.push(eq(demonstrativo.convenioId, input.convenioId));
    const rows = await db
      .select({
        id: demonstrativo.id,
        protocolo: demonstrativo.protocolo,
        numeroGuia: demonstrativo.numeroGuia,
        sequencialItem: demonstrativo.sequencialItem,
        codigoItem: demonstrativo.codigoItem,
        valorGlosa: demonstrativo.valorGlosa,
        convenioId: demonstrativo.convenioId,
        codigoGlosa: demonstrativo.codigoGlosa,
        tipoLancamento: demonstrativo.tipoLancamento,
        classificacaoGlosa: demonstrativo.classificacaoGlosa,
        classificacaoMotivo: demonstrativo.classificacaoMotivo,
        motivoAceite: demonstrativo.motivoAceite,
      })
      .from(demonstrativo)
      .where(and(...condicoes));
    candidatos.push(...rows);
  }
  // Ordem estável para o desempate por id
  candidatos.sort((a, b) => a.id - b.id);
  const porId = new Map(candidatos.map((c) => [c.id, c]));

  const casamentos = casarDecisoesComDemonstrativo(input.linhas, candidatos);
  const linhaPorNumero = new Map(input.linhas.map((l) => [l.linha, l]));
  const tokenPrevia = calcularTokenPrevia(escopo, input.linhas, casamentos, porId);

  let aplicados = 0;
  if (input.aplicar) {
    // A confirmação só vale para a mesma prévia (escopo + planilha + versão dos itens).
    if (input.tokenPrevia !== tokenPrevia) throw new ImportacaoDecisoesError("PREVIA_DESATUALIZADA");
    const plano = casamentos.filter((r) => r.status === "casado" && r.demonstrativoId && r.decisao);
    if (plano.length === 0) throw new ImportacaoDecisoesError("NADA_A_APLICAR");

    try {
      // Tudo-ou-nada: decisão + histórico de toda a planilha numa única transação.
      aplicados = await db.transaction(async (tx) => {
        let confirmados = 0;
        for (const r of plano) {
          const item = porId.get(r.demonstrativoId!)!;
          const decisao = r.decisao!;
          const linhaOrigem = linhaPorNumero.get(r.linha);
          const { motivo, motivoAceite } = montarMotivosDecisao(decisao, linhaOrigem?.justificativa, linhaOrigem?.observacao);
          const resultado = await tx
            .update(demonstrativo)
            .set({
              classificacaoGlosa: decisao,
              classificacaoConfianca: 100,
              classificacaoMotivo: motivo,
              motivoAceite,
              dataAceite: decisao === "aceitar" ? new Date() : null,
            })
            .where(
              and(
                eq(demonstrativo.id, item.id),
                eq(demonstrativo.estabelecimentoId, input.estabelecimentoId),
                // Concorrência otimista: o item precisa estar como na prévia.
                sql`${demonstrativo.classificacaoGlosa} <=> ${item.classificacaoGlosa ?? null}`,
                sql`${demonstrativo.classificacaoMotivo} <=> ${item.classificacaoMotivo ?? null}`,
                sql`${demonstrativo.motivoAceite} <=> ${item.motivoAceite ?? null}`,
              ),
            );
          if (linhasAfetadas(resultado) !== 1) throw new ImportacaoDecisoesError("CONFLITO_ITEM");
          await tx.insert(decisoesGlosa).values({
            convenioId: item.convenioId || 0,
            codigoGlosa: (s(item.codigoGlosa).match(/^(\d+)/)?.[1] || "0000").slice(0, 20),
            codigoProcedimento: s(item.codigoItem).slice(0, 50) || null,
            tipoProcedimento: item.tipoLancamento ? s(item.tipoLancamento).slice(0, 50) : null,
            decisao,
            valorGlosado: item.valorGlosa !== null && item.valorGlosa !== undefined ? String(item.valorGlosa) : undefined,
            motivoDecisao: motivo,
            procedimentoId: item.id,
            userId: input.userId,
          });
          confirmados++;
        }
        if (confirmados !== plano.length) throw new ImportacaoDecisoesError("CONFLITO_ITEM");
        return confirmados;
      });
    } catch (e) {
      if (e instanceof ImportacaoDecisoesError) throw e;
      // Detalhe técnico só no log do servidor; o cliente recebe mensagem genérica.
      console.error("[importarDecisoesGlosaPlanilha] falha ao gravar (transação revertida):", e);
      throw new ImportacaoDecisoesError("FALHA_GRAVACAO");
    }
  }

  const resultados = casamentos.map((r) => {
    const l = linhaPorNumero.get(r.linha);
    const item = r.demonstrativoId ? porId.get(r.demonstrativoId) : undefined;
    return {
      ...r,
      codigoItem: s(l?.codigoItem),
      guia: s(l?.guiaOperadora) || s(l?.guiaPrestador),
      descricaoItem: s(l?.descricaoItem),
      classificacaoAtual: item ? item.classificacaoGlosa ?? "pendente" : null,
    };
  });

  const casados = casamentos.filter((r) => r.status === "casado");
  return {
    totalLinhas: input.linhas.length,
    casados: casados.length,
    naoEncontrados: casamentos.filter((r) => r.status === "nao_encontrado").length,
    ignorados: casamentos.filter((r) => r.status === "decisao_invalida").length,
    aRecursar: casados.filter((r) => r.decisao === "recursar").length,
    aAceitar: casados.filter((r) => r.decisao === "aceitar").length,
    aplicados,
    escopo,
    tokenPrevia,
    resultados,
  };
}

const CODIGO_TRPC: Record<CodigoErroImportacaoDecisoes, TRPCError["code"]> = {
  BANCO_INDISPONIVEL: "INTERNAL_SERVER_ERROR",
  PREVIA_OBRIGATORIA: "BAD_REQUEST",
  PREVIA_DESATUALIZADA: "CONFLICT",
  NADA_A_APLICAR: "BAD_REQUEST",
  CONFLITO_ITEM: "CONFLICT",
  FALHA_GRAVACAO: "INTERNAL_SERVER_ERROR",
};

/**
 * Converte qualquer erro do serviço em TRPCError com mensagem segura.
 * Erros de autorização/validação já tipados (TRPCError) são preservados;
 * qualquer outro erro é registrado no servidor e devolvido como mensagem genérica.
 */
export function erroSeguroImportacaoDecisoes(e: unknown): TRPCError {
  if (e instanceof ImportacaoDecisoesError) {
    return new TRPCError({ code: CODIGO_TRPC[e.codigo], message: e.message });
  }
  if (e instanceof TRPCError && e.code !== "INTERNAL_SERVER_ERROR") return e;
  console.error("[importarDecisoesPlanilha] erro inesperado:", e);
  return new TRPCError({ code: "INTERNAL_SERVER_ERROR", message: MENSAGENS_ERRO.FALHA_GRAVACAO });
}