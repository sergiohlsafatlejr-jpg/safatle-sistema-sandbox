/**
 * Regras de estado da prévia do "Importar decisões de recurso (Excel)".
 *
 * A prévia só pode ser aplicada no MESMO escopo (hospital + convênio) e para a
 * MESMA extração de planilha em que foi analisada. Qualquer troca de escopo,
 * novo arquivo, cancelamento ou fechamento invalida a prévia e respostas
 * atrasadas de análises anteriores.
 */

export type EscopoImportacaoDecisoes = { estabelecimentoId: number; convenioId: number | null };

export type PreviaImportacaoDecisoes<TDados, TLinhas = unknown> = {
  /** Sequência da análise que gerou esta prévia. */
  seq: number;
  escopo: EscopoImportacaoDecisoes;
  /** Referência da extração analisada (identidade de objeto). */
  linhas: TLinhas;
  tokenPrevia: string;
  dados: TDados;
};

/** Escopo atual da tela; null quando não há hospital selecionado. */
export function escopoDaSelecao(estabelecimentoId: number | null | undefined, convenio: string): EscopoImportacaoDecisoes | null {
  if (!estabelecimentoId || !Number.isFinite(estabelecimentoId)) return null;
  const convenioId = convenio && convenio !== "todos" ? Number.parseInt(convenio, 10) : NaN;
  return { estabelecimentoId, convenioId: Number.isFinite(convenioId) ? convenioId : null };
}

export function mesmoEscopo(a: EscopoImportacaoDecisoes | null | undefined, b: EscopoImportacaoDecisoes | null | undefined): boolean {
  return !!a && !!b && a.estabelecimentoId === b.estabelecimentoId && (a.convenioId ?? null) === (b.convenioId ?? null);
}

/** A prévia continua aplicável para o escopo e a extração exibidos agora? */
export function previaAplicavel<TDados, TLinhas>(
  previa: PreviaImportacaoDecisoes<TDados, TLinhas> | null | undefined,
  escopoAtual: EscopoImportacaoDecisoes | null,
  linhasAtuais: TLinhas | null | undefined,
): previa is PreviaImportacaoDecisoes<TDados, TLinhas> {
  return (
    !!previa &&
    !!linhasAtuais &&
    previa.linhas === linhasAtuais &&
    mesmoEscopo(previa.escopo, escopoAtual) &&
    /^[0-9a-f]{64}$/.test(previa.tokenPrevia)
  );
}

/** Resposta de análise só é aceita se ainda for a análise corrente e o escopo não mudou. */
export function respostaAindaValida(
  seqResposta: number,
  seqAtual: number,
  escopoAnalisado: EscopoImportacaoDecisoes,
  escopoAtual: EscopoImportacaoDecisoes | null,
): boolean {
  return seqResposta === seqAtual && mesmoEscopo(escopoAnalisado, escopoAtual);
}

const CODIGOS_COM_MENSAGEM_SEGURA = new Set(["BAD_REQUEST", "CONFLICT", "FORBIDDEN", "UNAUTHORIZED", "NOT_FOUND", "PRECONDITION_FAILED"]);

/**
 * Mensagem exibível ao usuário. Só repassa o texto de erros tRPC de domínio
 * (já sanitizados pelo servidor); erros internos, de rede ou desconhecidos
 * viram a mensagem genérica informada.
 */
export function mensagemErroSegura(e: unknown, fallback: string): string {
  const err = e as { message?: unknown; data?: { code?: unknown } | null } | null | undefined;
  const code = typeof err?.data?.code === "string" ? err.data.code : null;
  const msg = typeof err?.message === "string" ? err.message.trim() : "";
  if (code && CODIGOS_COM_MENSAGEM_SEGURA.has(code) && msg && msg.length <= 300) return msg;
  return fallback;
}

/** Erro de conflito (prévia desatualizada / item alterado) exige nova análise. */
export function exigeNovaAnalise(e: unknown): boolean {
  return (e as { data?: { code?: unknown } | null } | null | undefined)?.data?.code === "CONFLICT";
}
