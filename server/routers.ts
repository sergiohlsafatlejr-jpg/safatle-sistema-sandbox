import { finalizarCsvBradescoInvalido } from "./bradescoCsvFinalizacao";
import { delegarPlanilhaRetorno } from "./retornoUploadDispatch";
import { assertInternalProcessingReady, dispatchInternalFileProcessing } from "./internalFileDispatch";
import {
  GLOSA_MANUAL_NO_ORIGINAL_MESSAGE,
  GLOSA_MANUAL_ORIGIN_LABEL,
  isGlosaManualOriginName,
} from "../shared/glosaManualOrigin";
import {
  confirmarEnvioLote,
  confirmarEnvioLoteSchema,
} from "./recursoLoteEnvioService";
import { consultarAcoesPublicasLoteXmlPorArquivos } from "./xmlImportacaoLoteService";
import {
  dadosPagamentoManual,
  pagamentoRecursoManualSchema,
} from "./recursoPagamentoManual";
import { erroSeguroImportacaoDecisoes, importarDecisoesGlosaPlanilha } from "./importDecisoesGlosaPlanilhaService";
import { COOKIE_NAME } from "@shared/const";
import { getSessionCookieOptions } from "./_core/cookies";
import { sdk } from "./_core/sdk";
import { systemRouter } from "./_core/systemRouter";
import {
  adminProcedure,
  publicProcedure,
  protectedProcedure,
  reportProcedure,
  trackedProtectedProcedure,
  router,
} from "./_core/trpc";
import { z } from "zod";
import { TRPCError } from "@trpc/server";
import { and, eq, getTableColumns } from "drizzle-orm";
import { nanoid } from "nanoid";
import { storagePut, storageGet } from "./storage";
import { parseFile } from "./parsers";
import type { InsertFaturamentoTiss } from "../drizzle/schema";
import {
  parseExcelRecebimentoTiss,
  parseXmlRecebimentoTiss,
} from "./recebimentoTissParser";
import {
  compararProcedimentos,
  toDivergenciaInsert,
  gerarResumoComparacao,
} from "./comparador";
import * as db from "./db";
import { getDb } from "./db";
import {
  salvarNotificacaoAtendimento,
  salvarNotificacoesAtendimentoEmLote,
  buscarNotificacoesAtendimento,
  getHistoricoNotificacoesAtendimento,
  verificarAtendimentosDoEstabelecimento,
  salvarHistoricoNotificacoesAtendimento,
  listarHistoricoNotificacoesAtendimento,
} from "./db";
import {
  getAtendimentosParadosUnificados,
  getAtendimentosParadosPorEstabelecimento,
  calcularDiasParadoUnificado,
  getKPIsPorTipo,
  getQuantidadePorPlano,
  getQuantidadePorServico,
} from "./atendimentosUnificados";
import { motorRegrasRouter } from "./routers/motorRegrasRouter";
import {
  enviarEmail,
  gerarHtmlNotificacaoAtendimentos,
  verificarConexaoSMTP,
} from "./emailService";
import { padroesProcedimentosRouter } from "./routers/padroesProcedimentosRouter";
import { padroesCobrancaRouter } from "./routers/padroesCobrancaRouter";
import { integradorDadosRouter } from "./routers/integradorDadosRouter";
import { easyLocalRouter } from "./routers/easyLocalRouter";
import { easyLocalIntegradorRouter } from "./routers/easyLocalIntegradorRouter";
import { easyLocalAgendamentoRouter } from "./routers/easyLocalAgendamentoRouter";
import { atendimentosHopDiarioRouter } from "./routers/atendimentosHopDiarioRouter";
import { faturamentoUnificadoRouter } from "./routers/faturamentoUnificadoRouter";
import { hupConciliacaoPilotoRouter } from "./routers/hupConciliacaoPilotoRouter";
import { contasConvenioRouter } from "./routers/contasConvenioRouter";
import { cbhpmRouter } from "./routers/cbhpmRouter";
import { auditoriaRouter } from "./routers/auditoriaRouter";
import { conferenciaRouter } from "./routers/conferenciaRouter";
import { relatorioAtendimentosRouter } from "./routers/relatorioAtendimentosRouter";
import { relatorioCustosRouter } from "./routers/relatorioCustosRouter";
import { getDadosBIFaturadoRecebido } from "./biFaturadoRecebido";
import { getDadosFaturadoRecebidoUnificado } from "./biFaturadoRecebidoUnificado";
import { relatorioFaturamentoRouter } from "./routers/relatorioFaturamentoRouter";
import { faturamentoExternoRouter } from "./routers/faturamentoExternoRouter";
import { nfseRouter } from "./routers/nfseRouter";
import { financeiroRouter } from "./routers/financeiroRouter";
import { contratosRouter } from "./routers/contratosRouter";
import { propostasRouter } from "./routers/propostasRouter";
import * as dbRecebGeral from "./db-recebimentoGeral";
import * as dbConvMap from "./db-convenioMapeamento";
import { auditSystemRouter } from "./routers/auditSystemRouter";
import { contratosConveniosRouter } from "./routers/contratosConveniosRouter";
import { dashboardBiRouter } from "./routers/dashboardBiRouter";
import { dashboardHomeRouter } from "./routers/dashboardHomeRouter";
import { rhRouter } from "./routers/rhRouter";
import { rpaRouter } from "./routers/rpaRouter";
import { relatoriosGlosasBiRouter } from "./routers/relatoriosGlosasBiRouter";
import { medComparaRouter } from "./routers/medComparaRouter";
import { tasyRouter } from "./routers/tasyRouter";
import { tasyImportacaoRouter } from "./routers/tasyImportacaoRouter";
import { fabricaFaturamentoRouter } from "./routers/fabricaFaturamentoRouter";
import { serializeTasyCivilDate } from "./tasyCivilDate";
import {
  buildAtendimentoNormalizationSql,
  normalizeAtendimentoCategory,
  resolveAtendimentoServiceType,
} from "./atendimentoNormalization";
import {
  ATENDIMENTOS_VALOR_SQL,
  escolherMetricaPainel,
  getAtendimentosPaineisValor,
  PAINEIS_VALOR_VAZIOS,
} from "./atendimentosPaineisValor";
import {
  buildAtendimentoDiasParadoSql,
  buildAtendimentoPermanenciaSql,
  resolverSaidaEfetivaAtendimento,
} from "./atendimentosDiasParado";
import { relatorioFisioterapiaRouter } from "./routers/relatorioFisioterapiaRouter";
import { comparativoMesAMesRouter } from "./routers/comparativoMesAMesRouter";
import { fonteRelatoriosRouter } from "./routers/fonteRelatoriosRouter";
import { aceiteAutomaticoRouter } from "./routers/aceiteAutomaticoRouter";
import { faturamentoDigitacaoRouter } from "./routers/faturamentoDigitacaoRouter";
import { previsaoPagamentosImecRouter } from "./routers/previsaoPagamentosImecRouter";
import { faturadoRecebidoUnificadoRouter } from "./routers/faturadoRecebidoUnificadoRouter";
import { warleineFaturamentoRouter } from "./routers/warleineFaturamentoRouter";
import { painelBiPersonalizadoRouter } from "./routers/painelBiPersonalizadoRouter";
import { atendimentosBiRouter } from "./routers/atendimentosBiRouter";
import { relatorioVisitaRouter } from "./routers/relatorioVisitaRouter";
import { passwordAuthRouter } from "./routers/passwordAuthRouter";
import { validateNewPassword } from "./passwordAuthService";
import {
  ACTION_PROFILE_DEFAULTS,
  normalizeBiParent,
  sanitizeCustomModulePermissions,
  type StandardServiceGroup,
} from "../shared/permissions";
import {
  encontrarLinhaCabecalhoProcedimentos,
  detectarFonteProcedimentoPorNome,
  normalizarCodigo,
  normalizarProcedimentos,
  PROCEDIMENTO_FONTES,
} from "./procedimentosTabela";
import { gerarUnimedDbf } from "./unimedDbfExporter";
import { lerEConsolidarArquivosProcedimentos } from "./procedimentosImportacaoLote";
import { retornosRecursosRouter } from "./routers/retornosRecursosRouter";
import { contaContextoMedicoRouter } from "./routers/contaContextoMedicoRouter";
import { prontuarioPdfRouter } from "./routers/prontuarioPdfRouter";

const PASSWORD_SESSION_MS = 12 * 60 * 60 * 1000;
const relatorioLaboratorioProcedure = reportProcedure("relLaboratorio");
const relatorioVisitaXmlProcedure = reportProcedure("relVisitasFaturadasXml");
const relatorioUltrassomProcedure = reportProcedure("relUltrassom");
const relatorioOxUtiProcedure = reportProcedure("relOxUti");
const relatorioOxUtiFaturadoProcedure = reportProcedure("relOxUtiFaturado");
const resumoPacienteOxUtiProcedure = reportProcedure("relResumoPacienteOxUti");

async function refreshPasswordSession(ctx: any, userId: number) {
  const user = await db.getUserById(userId);
  if (!user)
    throw new TRPCError({
      code: "INTERNAL_SERVER_ERROR",
      message: "Conta não encontrada",
    });
  const token = await sdk.createSessionToken(user.openId, {
    name: user.name || "Usuário",
    expiresInMs: PASSWORD_SESSION_MS,
    sessionVersion: Number(user.sessionVersion || 0),
  });
  ctx.res.cookie(COOKIE_NAME, token, {
    ...getSessionCookieOptions(ctx.req),
    maxAge: PASSWORD_SESSION_MS,
  });
}
import { sqlValoresRelatorioOxUti } from "./relatorioOxUtiValores";
import {
  exigirAcessoEstabelecimento,
  exigirConvenioDoEstabelecimento,
  exigirAcessoLote,
  exigirAcessoRecurso,
} from "./recursoGlosaPermissions";

/**
 * Sanitize filename to remove special characters that can cause issues with S3/URLs
 * - Removes accents (normalizes to ASCII)
 * - Replaces + with underscore
 * - Replaces spaces with underscore
 * - Removes other special characters
 * - Ensures only alphanumeric, underscore, hyphen, and dot remain
 */
function sanitizeFilename(filename: string): string {
  // Remove file extension temporarily
  const lastDotIndex = filename.lastIndexOf(".");
  const hasExtension = lastDotIndex > 0;
  const name = hasExtension ? filename.substring(0, lastDotIndex) : filename;
  const ext = hasExtension ? filename.substring(lastDotIndex) : "";

  // Normalize unicode characters (remove accents)
  let sanitized = name.normalize("NFD").replace(/[\u0300-\u036f]/g, "");

  // Replace + and spaces with underscore
  sanitized = sanitized.replace(/[+\s]+/g, "_");

  // Remove any character that is not alphanumeric, underscore, hyphen, or dot
  sanitized = sanitized.replace(/[^a-zA-Z0-9_\-]/g, "");

  // Remove multiple consecutive underscores
  sanitized = sanitized.replace(/_+/g, "_");

  // Remove leading/trailing underscores
  sanitized = sanitized.replace(/^_+|_+$/g, "");

  // Ensure we have a valid filename
  if (!sanitized) {
    sanitized = "arquivo";
  }

  return sanitized + ext;
}

/**
 * Mapeia tipoDespesa do parser para tipoItem legível na tabela faturamento_tiss
 * Conforme padrão TISS ANS: codigoDespesa 1=gás, 2=medicamento, 3=material, 5=diária, 7=taxa
 */
export function mapTipoDespesaParaTipoItem(tipoDespesa?: string): string {
  switch (tipoDespesa) {
    case "gas":
      return "GÁS MEDICINAL";
    case "medicamento":
      return "MEDICAMENTO";
    case "material":
      return "MATERIAL";
    case "diaria":
      return "DIÁRIA";
    case "taxa":
      return "TAXA/ALUGUÉIS";
    case "procedimento":
      return "PROCEDIMENTO";
    case "outros":
      return "OUTROS";
    default:
      return "PROCEDIMENTO";
  }
}

function arquivoEhLoteXmlProtegido(arquivo: {
  s3Key?: string | null;
  s3Url?: string | null;
}): boolean {
  return arquivo.s3Key === "xml-importacao-protegido" && !arquivo.s3Url;
}
function exigirArquivoOriginal(arquivo: { nome?: unknown }): void {
  if (isGlosaManualOriginName(arquivo.nome)) {
    throw new TRPCError({
      code: "BAD_REQUEST",
      message: GLOSA_MANUAL_NO_ORIGINAL_MESSAGE,
    });
  }
}
async function obterArquivoAutorizado(
  user: { id: number; role: string },
  arquivoId: number,
  permissao: "visualizar" | "editar" | "excluir"
) {
  const arquivo = await db.getArquivoById(arquivoId);
  if (!arquivo) {
    throw new TRPCError({
      code: "NOT_FOUND",
      message: "Arquivo não encontrado",
    });
  }
  const estabelecimentoId = Number(arquivo.estabelecimentoId);
  const permitido =
    user.role === "admin" ||
    (Number.isSafeInteger(estabelecimentoId) &&
      estabelecimentoId > 0 &&
      (await db.verificarPermissaoEstabelecimento(
        user.id,
        estabelecimentoId,
        permissao
      )));
  if (!permitido) {
    throw new TRPCError({
      code: "FORBIDDEN",
      message: "Sem acesso ao arquivo selecionado.",
    });
  }
  return arquivo;
}

export const appRouter = router({
  system: systemRouter,
  relatorioAtendimentos: relatorioAtendimentosRouter,
  relatorioCustos: relatorioCustosRouter,
  relatorioFaturamento: relatorioFaturamentoRouter,
  faturamentoExterno: faturamentoExternoRouter,
  nfse: nfseRouter,
  financeiro: financeiroRouter,
  contratos: contratosRouter,
  propostas: propostasRouter,
  auditSystem: auditSystemRouter,
  contratosConvenios: contratosConveniosRouter,
  dashboardBi: dashboardBiRouter,
  dashboardHome: dashboardHomeRouter,
  rh: rhRouter,
  rpa: rpaRouter,
  relatoriosGlosasBi: relatoriosGlosasBiRouter,
  medCompara: medComparaRouter,
  tasy: tasyRouter,
  tasyImportacao: tasyImportacaoRouter,
  relatorioFisioterapia: relatorioFisioterapiaRouter,
  comparativoMesAMes: comparativoMesAMesRouter,
  fonteRelatorios: fonteRelatoriosRouter,
  aceiteAutomatico: aceiteAutomaticoRouter,
  faturamentoDigitacao: faturamentoDigitacaoRouter,
  previsaoPagamentosImec: previsaoPagamentosImecRouter,
  faturadoRecebidoUnificado: faturadoRecebidoUnificadoRouter,
  warleineFaturamento: warleineFaturamentoRouter,
  painelBiPersonalizado: painelBiPersonalizadoRouter,
  atendimentosBi: atendimentosBiRouter,
  relatorioVisita: relatorioVisitaRouter,
  passwordAuth: passwordAuthRouter,
  retornosRecursos: retornosRecursosRouter,
  prontuarioPdf: prontuarioPdfRouter,
  fabricaFaturamento: fabricaFaturamentoRouter,

  auth: router({
    me: publicProcedure.query(({ ctx }) => {
      if (!ctx.user) return null;
      const {
        passwordHash: _passwordHash,
        emailCanonical: _emailCanonical,
        sessionVersion: _sessionVersion,
        ...safeUser
      } = ctx.user;
      return safeUser;
    }),
    logout: publicProcedure.mutation(({ ctx }) => {
      const cookieOptions = getSessionCookieOptions(ctx.req);
      ctx.res.clearCookie(COOKIE_NAME, { ...cookieOptions, maxAge: -1 });
      return { success: true } as const;
    }),

    // Verificar se usuário tem senha definida
    hasPassword: protectedProcedure.query(async ({ ctx }) => {
      return db.hasUserPassword(ctx.user.id);
    }),

    // Alterar senha do usuário
    changePassword: protectedProcedure
      .input(
        z.object({
        currentPassword: z.string().optional(),
        newPassword: z.string().min(1).max(128),
        confirmPassword: z.string(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        const policyError = validateNewPassword(input.newPassword);
        if (policyError)
          throw new TRPCError({ code: "BAD_REQUEST", message: policyError });
        // Validar confirmação de senha
        if (input.newPassword !== input.confirmPassword) {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message: "As senhas não conferem",
          });
        }

        // Verificar se usuário tem senha definida
        const hasPassword = await db.hasUserPassword(ctx.user.id);

        if (hasPassword && !input.currentPassword) {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message: "Informe a senha atual",
          });
        }

        const result = await db.changeUserPassword(
          ctx.user.id,
          input.currentPassword || "",
          input.newPassword
        );

        if (!result.success) {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message: result.message,
          });
        }

        await refreshPasswordSession(ctx, ctx.user.id);
        return { success: true, message: result.message };
      }),

    // Definir senha pela primeira vez (para usuários sem senha)
    setInitialPassword: protectedProcedure
      .input(
        z.object({
        newPassword: z.string().min(1).max(128),
        confirmPassword: z.string(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        const policyError = validateNewPassword(input.newPassword);
        if (policyError)
          throw new TRPCError({ code: "BAD_REQUEST", message: policyError });
        // Validar confirmação de senha
        if (input.newPassword !== input.confirmPassword) {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message: "As senhas não conferem",
          });
        }

        // Verificar se usuário já tem senha
        const hasPassword = await db.hasUserPassword(ctx.user.id);
        if (hasPassword) {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message:
              "Você já possui uma senha definida. Use a opção de alterar senha.",
          });
        }

        await db.setUserPassword(ctx.user.id, input.newPassword);
        await refreshPasswordSession(ctx, ctx.user.id);
        return { success: true, message: "Senha definida com sucesso" };
      }),
  }),

  // ============ CONVENIOS ============
  convenios: router({
    list: protectedProcedure
      .input(
        z
          .object({
        ativo: z.enum(["sim", "nao"]).optional(),
            estabelecimentoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getConvenios(input?.ativo, input?.estabelecimentoId);
      }),

    get: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ input }) => {
        return db.getConvenioById(input.id);
      }),

    create: protectedProcedure
      .input(
        z.object({
          nome: z.string().min(1),
          codigo: z.string().optional(),
          registroAns: z.string().optional(),
          prazoRecursoGlosa: z.number().optional(),
        })
      )
      .mutation(async ({ input }) => {
        return db.createConvenio({
          nome: input.nome,
          codigo: input.codigo || null,
          registroAns: input.registroAns || null,
          prazoRecursoGlosa: input.prazoRecursoGlosa || 30,
        });
      }),

    update: protectedProcedure
      .input(
        z.object({
          id: z.number(),
          nome: z.string().optional(),
          codigo: z.string().optional(),
          registroAns: z.string().optional(),
          prazoRecursoGlosa: z.number().optional(),
          ativo: z.enum(["sim", "nao"]).optional(),
        })
      )
      .mutation(async ({ input }) => {
        const { id, ...data } = input;
        await db.updateConvenio(id, data);
        return { success: true };
      }),

    // Rotas de prestadores por convênio/estabelecimento
    listarPrestadores: protectedProcedure
      .input(z.object({ convenioId: z.number() }))
      .query(async ({ input }) => {
        return db.listarPrestadoresPorConvenio(input.convenioId);
      }),

    getPrestador: protectedProcedure
      .input(
        z.object({
        convenioId: z.number(),
          estabelecimentoId: z.number(),
        })
      )
      .query(async ({ input }) => {
        return db.getPrestadorPorConvenioEstabelecimento(
          input.convenioId,
          input.estabelecimentoId
        );
      }),

    getPrestadorPorCodigo: protectedProcedure
      .input(
        z.object({
        codigoPrestador: z.string(),
          convenioId: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getPrestadorPorCodigo(
          input.codigoPrestador,
          input.convenioId
        );
      }),

    upsertPrestador: protectedProcedure
      .input(
        z.object({
        convenioId: z.number(),
        estabelecimentoId: z.number(),
        codigoPrestador: z.string().min(1),
          nomePrestador: z.string().optional(),
        })
      )
      .mutation(async ({ input }) => {
        return db.upsertPrestadorConvenioEstabelecimento(input);
      }),

    excluirPrestador: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input }) => {
        return db.excluirPrestadorConvenioEstabelecimento(input.id);
      }),

    listarTodosPrestadores: protectedProcedure.query(async () => {
        return db.listarTodosPrestadores();
      }),
  }),

  // ============ ESTABELECIMENTOS ============
  estabelecimentos: router({
    list: protectedProcedure
      .input(z.object({ ativo: z.enum(["sim", "nao"]).optional() }).optional())
      .query(async ({ input }) => {
        return db.getEstabelecimentos(input?.ativo);
      }),

    get: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ input }) => {
        return db.getEstabelecimentoById(input.id);
      }),

    create: protectedProcedure
      .input(
        z.object({
          nome: z.string().min(1),
          cnpj: z.string().optional(),
          cnes: z.string().optional(),
          endereco: z.string().optional(),
        })
      )
      .mutation(async ({ input }) => {
        return db.createEstabelecimento({
          nome: input.nome,
          cnpj: input.cnpj || null,
          cnes: input.cnes || null,
          endereco: input.endereco || null,
        });
      }),

    update: protectedProcedure
      .input(
        z.object({
          id: z.number(),
          nome: z.string().optional(),
          cnpj: z.string().optional(),
          cnes: z.string().optional(),
          endereco: z.string().optional(),
          ativo: z.enum(["sim", "nao"]).optional(),
        })
      )
      .mutation(async ({ input }) => {
        const { id, ...data } = input;
        await db.updateEstabelecimento(id, data);
        return { success: true };
      }),

    delete: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input }) => {
        const result = await db.deleteEstabelecimento(input.id);
        if (!result.success) {
          throw new TRPCError({ code: "BAD_REQUEST", message: result.message });
        }
        return result;
      }),
  }),

  // ============ REGRAS DE CONCILIAÇÃO ============
  regrasConciliacao: router({
    list: protectedProcedure.query(async () => {
        return db.getRegrasConciliacao();
      }),

    get: protectedProcedure
      .input(z.object({ convenioId: z.number() }))
      .query(async ({ input }) => {
        return db.getRegraConciliacaoPorConvenio(input.convenioId);
      }),

    upsert: protectedProcedure
      .input(
        z.object({
          convenioId: z.number(),
          itensNaoEncontrados: z
            .enum(["glosado", "pago", "divergente"])
            .optional(),
          toleranciaValor: z.string().optional(),
          toleranciaPercentual: z.string().optional(),
          usarCodigo: z.enum(["sim", "nao"]).optional(),
          usarGuia: z.enum(["sim", "nao"]).optional(),
          usarData: z.enum(["sim", "nao"]).optional(),
          usarPaciente: z.enum(["sim", "nao"]).optional(),
          formatoRetorno: z
            .enum(["excel_completo", "excel_glosas", "xml_tiss", "csv", "pdf"])
            .optional(),
          prazoRecursoDias: z.number().optional(),
          observacoes: z.string().optional(),
          ativo: z.enum(["sim", "nao"]).optional(),
        })
      )
      .mutation(async ({ input }) => {
        return db.upsertRegraConciliacao(input);
      }),

    delete: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input }) => {
        return db.deleteRegraConciliacao(input.id);
      }),
  }),

  // ============ ARQUIVOS ============
  arquivos: router({
    list: protectedProcedure
      .input(
        z
          .object({
            convenioId: z.number().optional(),
            direcao: z.enum(["enviado", "retornado"]).optional(),
            tipoArquivo: z
              .enum(["xml", "excel", "pdf", "csv", "txt"])
              .optional(),
            status: z.enum(["pendente", "processado", "erro"]).optional(),
            dataInicio: z.string().optional(),
            dataFim: z.string().optional(),
            busca: z.string().optional(),
            estabelecimentoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        const estabelecimentoId = input?.estabelecimentoId;
        if (!Number.isSafeInteger(estabelecimentoId) || !estabelecimentoId) {
          return [];
        }
        if (ctx.user.role !== "admin") {
          const permitido = await db.verificarPermissaoEstabelecimento(
            ctx.user.id,
            estabelecimentoId,
            "visualizar"
          );
          if (!permitido) {
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "Sem acesso ao estabelecimento selecionado.",
            });
          }
        }
        const arquivosList = await db.getArquivos({
          ...input,
          estabelecimentoId,
          dataInicio: input?.dataInicio
            ? new Date(input.dataInicio)
            : undefined,
          dataFim: input?.dataFim ? new Date(input.dataFim) : undefined,
        });

        // Enriquecer com totais e com uma ação sanitizada do lote protegido.
        // A lista não recebe identificação de lote, XML, hashes, URLs ou chaves.
        const arquivoIds = arquivosList.map((a: any) => a.id).filter(Boolean);
        const idsLoteProtegido = arquivosList
          .filter((a: any) => arquivoEhLoteXmlProtegido(a))
          .map((a: any) => a.id)
          .filter(Boolean);
        const totaisPromise: ReturnType<typeof db.getValorTotalPorArquivos> =
          arquivoIds.length
            ? db.getValorTotalPorArquivos(arquivoIds)
            : Promise.resolve({});
        const [totais, acoesLote] = await Promise.all([
          totaisPromise,
          consultarAcoesPublicasLoteXmlPorArquivos({
            arquivoIds: idsLoteProtegido,
            estabelecimentoId,
          }),
        ]);
        return arquivosList.map((a: any) => {
          const { s3Key: _s3Key, s3Url: _s3Url, ...arquivoPublico } = a;
          const loteProtegido = arquivoEhLoteXmlProtegido(a);
          const origemManual = isGlosaManualOriginName(a.nome);
          return {
            ...arquivoPublico,
            nome: origemManual ? GLOSA_MANUAL_ORIGIN_LABEL : a.nome,
            origemManual,
            loteProtegido,
            loteAcao: loteProtegido
              ? acoesLote.get(a.id) ?? "sem_acao_manual"
              : null,
            valorTotalInformado: totais[a.id]?.valorInformado || 0,
            valorTotalLiberado: totais[a.id]?.valorLiberado || 0,
            valorTotalGlosado: totais[a.id]?.valorGlosado || 0,
          };
        });
      }),

    get: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ input, ctx }) => {
        const arquivo = await obterArquivoAutorizado(
          ctx.user,
          input.id,
          "visualizar"
        );
        const { s3Key: _s3Key, s3Url: _s3Url, ...arquivoPublico } = arquivo;
        const estabelecimentoArquivoId = Number(arquivo.estabelecimentoId);
        const loteProtegido =
          arquivoEhLoteXmlProtegido(arquivo) &&
          Number.isSafeInteger(estabelecimentoArquivoId) &&
          estabelecimentoArquivoId > 0;
        const acoesLote = loteProtegido
          ? await consultarAcoesPublicasLoteXmlPorArquivos({
              arquivoIds: [arquivo.id],
              estabelecimentoId: estabelecimentoArquivoId,
            })
          : new Map();
        const origemManual = isGlosaManualOriginName(arquivo.nome);
        return {
          ...arquivoPublico,
          nome: origemManual ? GLOSA_MANUAL_ORIGIN_LABEL : arquivo.nome,
          origemManual,
          loteProtegido,
          loteAcao: loteProtegido
            ? acoesLote.get(arquivo.id) ?? "sem_acao_manual"
            : null,
        };
      }),

    stats: protectedProcedure.query(async ({ ctx }) => {
      return db.getArquivosStats(ctx.user.id);
    }),

    // Descontinuada: XML de envio segue somente pelo lote multipart protegido.
    detectarPrestadores: protectedProcedure
      .input(
        z.object({
          conteudo: z.string(), // Contrato legado, não é mais processado.
        convenioId: z.number().optional(),
        })
      )
      .mutation(async ({ input }) => {
        void input;
        throw new TRPCError({
          code: "BAD_REQUEST",
          message:
            "A detecção automática por XML foi substituída pelo lote protegido. Selecione o estabelecimento e o convênio antes do envio.",
                  });
      }),

    upload: protectedProcedure
      .input(
        z.object({
          nome: z.string(),
          tipoArquivo: z.enum(["xml", "excel", "pdf", "csv", "txt"]),
          direcao: z.enum(["enviado", "retornado"]),
          convenioId: z.number(),
          estabelecimentoId: z.number(), // Estabelecimento associado ao arquivo
          conteudo: z.string(), // Base64 encoded
          dataReferencia: z.string().optional(),
          dataPagamento: z.string().optional(), // Data de pagamento do convênio
        })
      )
      .mutation(async ({ input, ctx }) => {
        // O escopo é autorizado antes de decodificar ou persistir o arquivo.
        // Admin ainda precisa apontar um convênio compatível; demais usuários
        // precisam de permissão de edição na unidade.
        const convenio = await db.getConvenioById(input.convenioId);
        if (!convenio) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Convênio não encontrado",
          });
        }
        if (
          convenio.estabelecimentoId &&
          convenio.estabelecimentoId !== input.estabelecimentoId
        ) {
          throw new TRPCError({
            code: "FORBIDDEN",
            message:
              "O convênio informado não pertence ao estabelecimento selecionado",
          });
        }
        if (ctx.user.role !== "admin") {
          await exigirConvenioDoEstabelecimento(
            ctx.user.id,
            input.convenioId,
            input.estabelecimentoId,
            "editar"
          );
        }

        if (input.tipoArquivo === "xml" && input.direcao === "enviado") {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message:
              "Envie XMLs de guias pelo lote protegido na tela de Upload.",
          });
        }

        // Decode base64 content
        // O namespace da âncora não pode ser usado para criar/substituir documento.
        exigirArquivoOriginal({ nome: input.nome });
        assertInternalProcessingReady();
        const buffer = Buffer.from(input.conteudo, "base64");

        // Validar arquivo vazio
        if (buffer.length === 0) {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message:
              "O arquivo está vazio (0 bytes). Verifique o arquivo e tente novamente.",
          });
        }

        // Sanitize filename to remove special characters
        const sanitizedNome = sanitizeFilename(input.nome);

        // Verificar se já existe um arquivo com mesmo nome, convênio e estabelecimento
        const arquivoExistente = await db.findArquivoExistente({
          nome: input.nome,
          convenioId: input.convenioId,
          estabelecimentoId: input.estabelecimentoId,
          userId: ctx.user.id,
          direcao: input.direcao,
        });

        let arquivoId: number;
        let s3Key: string;
        let url: string;
        let isReimportacao = false;

        if (arquivoExistente) {
          exigirArquivoOriginal(arquivoExistente);
          // Reimportação: atualizar arquivo existente
          console.log(
            "[Upload] Arquivo existente encontrado, atualizando:",
            arquivoExistente.id
          );
          isReimportacao = true;
          arquivoId = arquivoExistente.id;

          // Gerar nova chave S3 para o arquivo atualizado
          const ext =
            input.tipoArquivo === "excel"
              ? "xlsx"
              : input.tipoArquivo === "csv"
                ? "csv"
                : input.tipoArquivo;
          s3Key = `arquivos/${ctx.user.id}/${nanoid()}-${sanitizedNome}.${ext}`;

          // Upload to S3
          const contentType =
            input.tipoArquivo === "xml"
              ? "application/xml"
              : input.tipoArquivo === "excel"
              ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              : input.tipoArquivo === "csv"
              ? "text/csv"
              : input.tipoArquivo === "txt"
              ? "text/plain"
              : "application/pdf";

          const uploadResult = await storagePut(s3Key, buffer, contentType);
          url = uploadResult.url;
          // A classificação do XML e qualquer mutação financeira pertencem ao
          // processor interno. Não antecipar DELETEs em reimportação: XML de
          // honorários puro nunca toca tabelas financeiras. A remoção completa
          // de gerações de fontes em reimportação exige migração posterior à
          // 0140 e permanece explicitamente fora deste fluxo.

          // Atualizar registro do arquivo
          await db.updateArquivo(arquivoId, {
            s3Key,
            s3Url: url,
            tamanho: buffer.length,
            status: "pendente",
            progresso: 0,
            itensProcessados: 0,
            totalItens: null,
            dataReferencia: input.dataReferencia
              ? new Date(input.dataReferencia)
              : null,
            dataPagamento: input.dataPagamento
              ? new Date(input.dataPagamento)
              : null,
          });
        } else {
          // Novo arquivo: criar registro
          const ext =
            input.tipoArquivo === "excel"
              ? "xlsx"
              : input.tipoArquivo === "csv"
                ? "csv"
                : input.tipoArquivo;
          s3Key = `arquivos/${ctx.user.id}/${nanoid()}-${sanitizedNome}.${ext}`;

          // Upload to S3
          const contentType =
            input.tipoArquivo === "xml"
              ? "application/xml"
              : input.tipoArquivo === "excel"
              ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              : input.tipoArquivo === "csv"
              ? "text/csv"
              : input.tipoArquivo === "txt"
              ? "text/plain"
              : "application/pdf";

          const uploadResult = await storagePut(s3Key, buffer, contentType);
          url = uploadResult.url;

          // Create arquivo record
          const result = await db.createArquivo({
            nome: input.nome,
            tipoArquivo: input.tipoArquivo,
            direcao: input.direcao,
            convenioId: input.convenioId,
            estabelecimentoId: input.estabelecimentoId,
            userId: ctx.user.id,
            s3Key,
            s3Url: url,
            tamanho: buffer.length,
            status: "pendente",
            dataReferencia: input.dataReferencia
              ? new Date(input.dataReferencia)
              : null,
            dataPagamento: input.dataPagamento
              ? new Date(input.dataPagamento)
              : null,
          });
          arquivoId = result.id;
        }

        // Todo upload é processado pelo worker interno. Isso mantém a classificação
        // de XML de honorários e o fluxo financeiro em uma única fronteira. O worker
        // decide antes de qualquer DELETE/INSERT financeiro e finaliza o status.
        assertInternalProcessingReady();
        await db.updateArquivoStatus(arquivoId, "processando");
        await db.updateArquivoProgresso(arquivoId, 5, 0);

        // POST /api/internal/process-file: segredo exclusivamente server-side.
        void dispatchInternalFileProcessing({
          arquivoId,
          limparDadosAnteriores: isReimportacao || !(input.direcao === "retornado" && input.tipoArquivo === "excel"),
        }).catch(() => undefined);

        // DTO sanitizado: a URL de storage não é necessária para acompanhar o
        // upload e não deve sair pela API.
        return {
          id: arquivoId,
          reimportado: isReimportacao,
          processandoEmBackground: true,
        };
      }),

    // Endpoint 'reprocessar' unificado está definido mais abaixo (suporta retornado + envio)

    procedimentos: protectedProcedure
      .input(z.object({ arquivoId: z.number() }))
      .query(async ({ input, ctx }) => {
        await obterArquivoAutorizado(ctx.user, input.arquivoId, "visualizar");
        // Buscar de faturamento_tiss em vez da tabela procedimentos (removida)
        return db.getFaturamentoTissByArquivo(input.arquivoId);
      }),

    delete: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input, ctx }) => {
        const arquivo = await obterArquivoAutorizado(
          ctx.user,
          input.id,
          "excluir"
        );
        exigirArquivoOriginal(arquivo);
        if (arquivoEhLoteXmlProtegido(arquivo)) {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message:
              "O XML protegido pertence a um lote coordenado e não pode ser excluído por esta ação.",
          });
              }
        // Delete associated faturamento_tiss
        await db.deleteFaturamentoTissByArquivo(input.id);

        // Delete associated recebimento_tiss (XML retornados)
        const recTiss = await db.deleteRecebimentoTissByArquivo(input.id);

        // Delete associated recebimentos_excel (Excel retornados)
        const recExcel = await db.deleteRecebimentosExcelByArquivo(input.id);

        // Delete associated demonstrativo entries
        const demo = await db.deleteDemonstrativoByArquivo(input.id);

        // Delete associated contas_convenio_itens
                    try {
          const { contasConvenioItens, contasConvenioResumo } = await import(
            "../drizzle/schema"
          );
          const { eq, and, sql: sqlDrizzle } = await import("drizzle-orm");
                      const dbCC = await getDb();
          if (dbCC) {
            // Buscar guias afetadas antes de deletar para recalcular resumos
            const guiasAfetadas = await dbCC
              .selectDistinct({
                numeroConta: contasConvenioItens.numeroConta,
                estabelecimentoId: contasConvenioItens.estabelecimentoId,
              })
              .from(contasConvenioItens)
              .where(eq(contasConvenioItens.arquivoId, input.id));

            // Deletar itens do arquivo
            await dbCC
              .delete(contasConvenioItens)
              .where(eq(contasConvenioItens.arquivoId, input.id));

            // Recalcular resumos das guias afetadas
            for (const guia of guiasAfetadas) {
              const itensRestantes = await dbCC
                .select()
                .from(contasConvenioItens)
                .where(
                          and(
                    eq(contasConvenioItens.numeroConta, guia.numeroConta),
                    eq(
                      contasConvenioItens.estabelecimentoId,
                      guia.estabelecimentoId
                    )
                          )
                        );

              if (itensRestantes.length === 0) {
                // Sem itens restantes: deletar resumo
                await dbCC
                  .delete(contasConvenioResumo)
                  .where(
                    and(
                      eq(contasConvenioResumo.numeroConta, guia.numeroConta),
                      eq(
                        contasConvenioResumo.estabelecimentoId,
                        guia.estabelecimentoId
                      )
                    )
                  );
                          } else {
                // Recalcular totais
                let valorTotal = 0;
                for (const item of itensRestantes) {
                  valorTotal += parseFloat(String(item.valorTotal || 0));
                              }
                await dbCC
                  .update(contasConvenioResumo)
                  .set({
                    totalItens: itensRestantes.length,
                    valorTotal: String(valorTotal.toFixed(2)),
                  })
                  .where(
                    and(
                      eq(contasConvenioResumo.numeroConta, guia.numeroConta),
                      eq(
                        contasConvenioResumo.estabelecimentoId,
                        guia.estabelecimentoId
                      )
                    )
                  );
                            }
                          }
            console.log(
              `[Arquivo Delete] contas_convenio limpo para ${guiasAfetadas.length} guias`
            );
                          }
        } catch (ccError) {
          console.error(
            "[Arquivo Delete] Erro ao limpar contas_convenio:",
            ccError
          );
        }

        // Delete the arquivo record itself
        await db.deleteArquivo(input.id);

        console.log(
          `[Arquivo Delete] Arquivo ${input.id} (${arquivo.nome}) exclu\u00eddo em cascata: recebimentoTiss=${recTiss}, recebimentosExcel=${recExcel}, demonstrativo=${demo}`
        );
        return { success: true };
      }),

    reprocessar: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input, ctx }) => {
        const arquivo = await obterArquivoAutorizado(
          ctx.user,
          input.id,
          "editar"
        );
        exigirArquivoOriginal(arquivo);
        if (arquivoEhLoteXmlProtegido(arquivo)) {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message:
              "O XML protegido é retomado pelo worker do lote, não pelo reprocessamento legado.",
          });
        }

        // Marcar como processando
        assertInternalProcessingReady();
        await db.updateArquivoStatus(input.id, "processando");
        await db.updateArquivoProgresso(input.id, 5, 0);

        // Disparar processamento via rota interna (request HTTP separada com seu próprio timeout)
        // O Cloud Run mantém a instância viva enquanto há requests em andamento
        void dispatchInternalFileProcessing({
          arquivoId: input.id,
          limparDadosAnteriores: !(arquivo.direcao === "retornado" && arquivo.tipoArquivo === "excel" && (arquivo.itensProcessados || 0) === 0 && !arquivo.totalItens),
        }).catch(() => undefined);

        // Retornar imediatamente ao frontend
        return {
          success: true,
          procedimentosCount: 0,
          message: "Reprocessamento iniciado. Acompanhe o progresso na lista.",
        };
      }),

    // Corrigir arquivos travados em processando
    corrigirTravados: protectedProcedure
      .input(z.object({ minutosTimeout: z.number().optional() }).optional())
      .mutation(async ({ input }) => {
        const minutosTimeout = input?.minutosTimeout || 10;
        return db.corrigirArquivosTravados(minutosTimeout);
      }),

    // Verificar arquivos em processamento
    processando: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number().optional() }).optional())
      .query(async ({ input }) => {
        return db.getArquivosProcessando(input?.estabelecimentoId);
      }),

    // Verificar status de um arquivo específico
    status: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ input, ctx }) => {
        const arquivo = await obterArquivoAutorizado(
          ctx.user,
          input.id,
          "visualizar"
        );
        const origemManual = isGlosaManualOriginName(arquivo.nome);
        return {
          id: arquivo.id,
          nome: origemManual ? GLOSA_MANUAL_ORIGIN_LABEL : arquivo.nome,
          origemManual,
          status: arquivo.status,
          progresso: arquivo.progresso || 0,
          itensProcessados: arquivo.itensProcessados || 0,
          totalItens: arquivo.totalItens || 0,
        };
      }),
  }),

  // ============ COMPARACOES ============
  comparacoes: router({
    list: protectedProcedure
      .input(
        z
          .object({
            convenioId: z.number().optional(),
            status: z.enum(["pendente", "concluida", "erro"]).optional(),
            dataInicio: z.string().optional(),
            dataFim: z.string().optional(),
            estabelecimentoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        return db.getComparacoes({
          ...input,
          estabelecimentoId: input?.estabelecimentoId,
          dataInicio: input?.dataInicio
            ? new Date(input.dataInicio)
            : undefined,
          dataFim: input?.dataFim ? new Date(input.dataFim) : undefined,
        });
      }),

    get: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ input }) => {
        const comparacao = await db.getComparacaoById(input.id);
        if (!comparacao) return null;

        const divergencias = await db.getDivergenciasByComparacaoId(input.id);
        return { ...comparacao, divergencias };
      }),

    stats: protectedProcedure.query(async ({ ctx }) => {
      return db.getComparacoesStats(ctx.user.id);
    }),

    criar: trackedProtectedProcedure
      .input(
        z.object({
          arquivoEnviadoId: z.number(),
          arquivoRetornadoId: z.number(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        // Get arquivos
        const [arquivoEnviado, arquivoRetornado] = await Promise.all([
          db.getArquivoById(input.arquivoEnviadoId),
          db.getArquivoById(input.arquivoRetornadoId),
        ]);

        if (!arquivoEnviado || !arquivoRetornado) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Arquivo não encontrado",
          });
        }

        if (arquivoEnviado.convenioId !== arquivoRetornado.convenioId) {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message: "Os arquivos devem ser do mesmo convênio",
          });
        }

        // Create comparacao record
        const { id: comparacaoId } = await db.createComparacao({
          arquivoEnviadoId: input.arquivoEnviadoId,
          arquivoRetornadoId: input.arquivoRetornadoId,
          convenioId: arquivoEnviado.convenioId,
          userId: ctx.user.id,
          status: "pendente",
        });

        // Get itens faturados (de faturamento_tiss)
        const [procEnviados, procRetornados] = await Promise.all([
          db.getFaturamentoTissByArquivo(input.arquivoEnviadoId),
          db.getFaturamentoTissByArquivo(input.arquivoRetornadoId),
        ]);

        // Run comparison
        const resultado = compararProcedimentos(procEnviados, procRetornados);

        // Save divergencias
        if (resultado.divergencias.length > 0) {
          const divergenciasToInsert = resultado.divergencias.map(d =>
            toDivergenciaInsert(d, comparacaoId)
          );
          await db.createDivergencias(divergenciasToInsert);
        }

        // Update comparacao with results
        await db.updateComparacao(comparacaoId, {
          status: "concluida",
          totalItensEnviados: resultado.totalItensEnviados,
          totalItensRetornados: resultado.totalItensRetornados,
          totalDivergencias: resultado.totalDivergencias,
          valorTotalEnviado: resultado.valorTotalEnviado.toString(),
          valorTotalRetornado: resultado.valorTotalRetornado.toString(),
          diferencaValor: resultado.diferencaValor.toString(),
        });

        return {
          id: comparacaoId,
          ...resultado,
          resumo: gerarResumoComparacao(resultado),
        };
      }),

    divergencias: protectedProcedure
      .input(z.object({ comparacaoId: z.number() }))
      .query(async ({ input }) => {
        return db.getDivergenciasByComparacaoId(input.comparacaoId);
      }),

    resolverDivergencia: protectedProcedure
      .input(
        z.object({
          id: z.number(),
          resolvido: z.enum(["sim", "nao"]),
        })
      )
      .mutation(async ({ input }) => {
        await db.updateDivergenciaResolvido(input.id, input.resolvido);
        return { success: true };
      }),
  }),

  // ============ INSIGHTS DE IA ============
  insightsIA: router({
    // Gerar insights para um arquivo
    gerar: protectedProcedure
      .input(
        z.object({
          arquivoId: z.number(),
          estabelecimentoId: z.number(),
          convenioId: z.number().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const insights = await db.gerarInsightsContaIA({
          arquivoId: input.arquivoId,
          estabelecimentoId: input.estabelecimentoId,
          convenioId: input.convenioId,
          userId: ctx.user.id!,
        });
        return insights;
      }),

    // Listar insights de um arquivo ou estabelecimento
    list: protectedProcedure
      .input(
        z.object({
          arquivoId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
          convenioId: z.number().optional(),
          status: z
            .enum(["pendente", "aceito", "rejeitado", "ignorado"])
            .optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getInsightsIA(input);
      }),

    // Atualizar status de um insight
    atualizarStatus: protectedProcedure
      .input(
        z.object({
          id: z.number(),
          status: z.enum(["pendente", "aceito", "rejeitado", "ignorado"]),
          feedback: z.string().optional(),
        })
      )
      .mutation(async ({ input }) => {
        const success = await db.atualizarStatusInsightIA(
          input.id,
          input.status,
          input.feedback
        );
        return { success };
      }),

    // Aprender padrões de um convênio/estabelecimento
    aprenderPadroes: protectedProcedure
      .input(
        z.object({
          estabelecimentoId: z.number(),
          convenioId: z.number().optional(),
        })
      )
      .mutation(async ({ input }) => {
        const resultado = await db.analisarPadroesCobranca({
          estabelecimentoId: input.estabelecimentoId,
          convenioId: input.convenioId,
        });
        return resultado;
      }),

    // Listar padrões aprendidos
    padroes: protectedProcedure
      .input(
        z.object({
          estabelecimentoId: z.number(),
          convenioId: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getPadroesCobranca(input);
      }),

    // Métricas de acurácia dos insights
    metricas: protectedProcedure
      .input(
        z.object({
          estabelecimentoId: z.number().optional(),
          convenioId: z.number().optional(),
          dataInicio: z.date().optional(),
          dataFim: z.date().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getMetricasAcuraciaIA(input);
      }),

    // Verificar e notificar insights críticos
    verificarCriticos: protectedProcedure
      .input(
        z.object({
          estabelecimentoId: z.number(),
          limiarImpacto: z.number().optional(),
          limiarConfianca: z.number().optional(),
        })
      )
      .mutation(async ({ input }) => {
        return db.verificarENotificarInsightsCriticos(input);
      }),

    // Processar insights automaticamente após importação
    processarAutomatico: protectedProcedure
      .input(
        z.object({
          arquivoId: z.number(),
          estabelecimentoId: z.number(),
          convenioId: z.number().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        return db.processarInsightsAutomaticos({
          ...input,
          userId: ctx.user.id!,
        });
      }),
  }),

  // ============ CODIGOS PROCEDIMENTOS (ADMIN) ============
  codigosProcedimentos: router({
    list: protectedProcedure
      .input(z.object({ ativo: z.enum(["sim", "nao"]).optional() }).optional())
      .query(async ({ input }) => {
        return db.getCodigosProcedimentos(input?.ativo);
      }),

    create: trackedProtectedProcedure
      .input(
        z.object({
          codigo: z.string().min(1),
          descricao: z.string().min(1),
          valorReferencia: z.number().optional(),
          categoria: z.string().optional(),
        })
      )
      .mutation(async ({ input }) => {
        return db.createCodigoProcedimento({
          codigo: input.codigo,
          descricao: input.descricao,
          valorReferencia: input.valorReferencia?.toString() || null,
          categoria: input.categoria || null,
        });
      }),

    update: trackedProtectedProcedure
      .input(
        z.object({
          id: z.number(),
          codigo: z.string().optional(),
          descricao: z.string().optional(),
          valorReferencia: z.number().optional(),
          categoria: z.string().optional(),
          ativo: z.enum(["sim", "nao"]).optional(),
        })
      )
      .mutation(async ({ input }) => {
        const { id, valorReferencia, ...rest } = input;
        await db.updateCodigoProcedimento(id, {
          ...rest,
          valorReferencia: valorReferencia?.toString(),
        });
        return { success: true };
      }),

    delete: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input }) => {
        await db.deleteCodigoProcedimento(input.id);
        return { success: true };
      }),
  }),

  // ============ CAMPOS COMPARACAO (ADMIN) ============
  camposComparacao: router({
    list: protectedProcedure
      .input(z.object({ ativo: z.enum(["sim", "nao"]).optional() }).optional())
      .query(async ({ input }) => {
        return db.getCamposComparacao(input?.ativo);
      }),

    create: protectedProcedure
      .input(
        z.object({
          nome: z.string().min(1),
          campoOrigem: z.string().min(1),
          obrigatorio: z.enum(["sim", "nao"]).optional(),
          tolerancia: z.number().optional(),
        })
      )
      .mutation(async ({ input }) => {
        return db.createCampoComparacao({
          nome: input.nome,
          campoOrigem: input.campoOrigem,
          obrigatorio: input.obrigatorio || "nao",
          tolerancia: input.tolerancia?.toString() || null,
        });
      }),

    update: protectedProcedure
      .input(
        z.object({
          id: z.number(),
          nome: z.string().optional(),
          campoOrigem: z.string().optional(),
          ativo: z.enum(["sim", "nao"]).optional(),
          obrigatorio: z.enum(["sim", "nao"]).optional(),
          tolerancia: z.number().optional(),
        })
      )
      .mutation(async ({ input }) => {
        const { id, tolerancia, ...rest } = input;
        await db.updateCampoComparacao(id, {
          ...rest,
          tolerancia: tolerancia?.toString(),
        });
        return { success: true };
      }),
  }),

  // ============ ITENS MANUAIS ============
  itensManuals: router({
    list: protectedProcedure
      .input(
        z
          .object({
            arquivoId: z.number().optional(),
            comparacaoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        return db.getItensManuals({
          ...input,
          userId: ctx.user.id,
        });
      }),

    create: trackedProtectedProcedure
      .input(
        z.object({
          arquivoId: z.number().optional(),
          comparacaoId: z.number().optional(),
          codigo: z.string().min(1),
          descricao: z.string().optional(),
          quantidade: z.number().optional(),
          valorUnitario: z.number().optional(),
          valorTotal: z.number().optional(),
          observacao: z.string().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        return db.createItemManual({
          arquivoId: input.arquivoId || null,
          comparacaoId: input.comparacaoId || null,
          userId: ctx.user.id,
          codigo: input.codigo,
          descricao: input.descricao || null,
          quantidade: input.quantidade || 1,
          valorUnitario: input.valorUnitario?.toString() || null,
          valorTotal: input.valorTotal?.toString() || null,
          observacao: input.observacao || null,
        });
      }),

    delete: trackedProtectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input }) => {
        await db.deleteItemManual(input.id);
        return { success: true };
      }),
  }),

  // ============ PROCEDIMENTOS ============
  procedimentos: router({
    list: protectedProcedure
      .input(
        z
          .object({
            arquivoId: z.number().optional(),
            convenioId: z.number().optional(),
            estabelecimentoId: z.number().optional(),
            search: z.string().optional(),
            nomeMedico: z.string().optional(),
            crmMedico: z.string().optional(),
            codigoPrestadorExecutante: z.string().optional(),
            statusGlosa: z
              .enum(["todos", "pago", "glosado", "parcial"])
              .optional(),
            apenasRetornados: z.boolean().optional(),
            direcaoArquivo: z.enum(["enviado", "retornado"]).optional(),
            mesReferencia: z.number().min(1).max(12).optional(),
            anoReferencia: z.number().min(2000).max(2100).optional(),
            page: z.number().default(1),
            pageSize: z.number().default(20),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        return db.getProcedimentosPaginated({
          arquivoId: input?.arquivoId,
          convenioId: input?.convenioId,
          estabelecimentoId: input?.estabelecimentoId,
          search: input?.search,
          nomeMedico: input?.nomeMedico,
          crmMedico: input?.crmMedico,
          codigoPrestadorExecutante: input?.codigoPrestadorExecutante,
          statusGlosa: input?.statusGlosa,
          apenasRetornados: input?.apenasRetornados,
          direcaoArquivo: input?.direcaoArquivo,
          mesReferencia: input?.mesReferencia,
          anoReferencia: input?.anoReferencia,
          page: input?.page || 1,
          pageSize: input?.pageSize || 20,
        });
      }),

    // Listar prestadores executantes únicos para filtro
    listarPrestadoresExecutantes: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
            convenioId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.listarPrestadoresExecutantes({
          estabelecimentoId: input?.estabelecimentoId,
          convenioId: input?.convenioId,
        });
      }),
  }),

  // ============ DASHBOARD ============
  dashboard: router({
    resumo: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number().optional() }).optional())
      .query(async ({ input, ctx }) => {
        const estabelecimentoId = input?.estabelecimentoId;
        const [arquivosStats, comparacoesStats] = await Promise.all([
          db.getArquivosStats(ctx.user.id, estabelecimentoId),
          db.getComparacoesStats(ctx.user.id, estabelecimentoId),
        ]);

        return {
          arquivos: arquivosStats,
          comparacoes: comparacoesStats,
        };
      }),

    ultimasComparacoes: protectedProcedure
      .input(
        z
          .object({
        limit: z.number().default(5),
            estabelecimentoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        const comparacoes = await db.getComparacoes({
          estabelecimentoId: input?.estabelecimentoId,
        });
        return comparacoes.slice(0, input?.limit || 5);
      }),

    ultimosArquivos: protectedProcedure
      .input(
        z
          .object({
        limit: z.number().default(10),
            estabelecimentoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        const arquivos = await db.getArquivos({
          estabelecimentoId: input?.estabelecimentoId,
        });
        return arquivos.slice(0, input?.limit || 10);
      }),
  }),

  // ============ FATURAMENTO ============
  faturamento: router({
    porConvenio: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        mesReferencia: z.number().optional(),
        anoReferencia: z.number().optional(),
        codigoPrestadorExecutante: z.string().optional(),
          })
          .optional()
      )
      .query(async ({ ctx, input }) => {
        return db.getFaturamentoPorConvenio(
          ctx.user.id,
          input?.estabelecimentoId,
          input?.mesReferencia,
          input?.anoReferencia,
          input?.codigoPrestadorExecutante
        );
      }),

    porMes: protectedProcedure
      .input(
        z
          .object({
          convenioId: z.number().optional(),
          meses: z.number().default(12),
          estabelecimentoId: z.number().optional(),
          anoReferencia: z.number().optional(),
          codigoPrestadorExecutante: z.string().optional(),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        return db.getFaturamentoPorMes(
          ctx.user.id,
          input?.convenioId,
          input?.meses || 12,
          input?.estabelecimentoId,
          input?.anoReferencia,
          input?.codigoPrestadorExecutante
        );
      }),

    resumoGeral: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        mesReferencia: z.number().optional(),
        anoReferencia: z.number().optional(),
        codigoPrestadorExecutante: z.string().optional(),
          })
          .optional()
      )
      .query(async ({ ctx, input }) => {
        return db.getResumoGeral(
          ctx.user.id,
          input?.estabelecimentoId,
          input?.mesReferencia,
          input?.anoReferencia,
          input?.codigoPrestadorExecutante
        );
      }),
  }),

  // ============ ANÁLISE DE GLOSA ============
  glosa: router({
    porConvenio: protectedProcedure
      .input(
        z
          .object({
        convenioId: z.number().optional(),
            estabelecimentoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ ctx, input }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input?.estabelecimentoId,
          "visualizar"
        );
        return db.getGlosaPorConvenio(
          ctx.user.id,
          input?.estabelecimentoId,
          input?.convenioId
        );
      }),

    porProcedimento: protectedProcedure
      .input(
        z
          .object({
          convenioId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
          limit: z.number().default(20),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input?.estabelecimentoId,
          "visualizar"
        );
        return db.getGlosaPorProcedimento(
          ctx.user.id,
          input?.convenioId,
          input?.limit || 20,
          input?.estabelecimentoId
        );
      }),

    tendencia: protectedProcedure
      .input(
        z
          .object({
          convenioId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
          meses: z.number().default(12),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input?.estabelecimentoId,
          "visualizar"
        );
        return db.getTendenciaGlosa(
          ctx.user.id,
          input?.convenioId,
          input?.meses || 12,
          input?.estabelecimentoId
        );
      }),

    resumo: protectedProcedure
      .input(
        z
          .object({
        convenioId: z.number().optional(),
            estabelecimentoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ ctx, input }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input?.estabelecimentoId,
          "visualizar"
        );
        return db.getResumoGlosa(
          ctx.user.id,
          input?.estabelecimentoId,
          input?.convenioId
        );
      }),

    // Itens glosados com filtros avançados
    itensGlosados: protectedProcedure
      .input(
        z
          .object({
          convenioId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
          dataReferenciaInicio: z.date().optional(),
          dataReferenciaFim: z.date().optional(),
          dataPagamento: z.string().optional(),
          tipo: z.string().optional(),
          codigoGlosa: z.string().optional(),
          motivoGlosa: z.string().optional(),
            classificacao: z
              .enum(["todos", "pendente", "aceitar", "recursar"])
              .optional(),
          search: z.string().optional(),
          page: z.number().default(1),
          pageSize: z.number().default(50),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input?.estabelecimentoId,
          "visualizar"
        );
        return db.getItensGlosados({
          convenioId: input?.convenioId,
          estabelecimentoId: input?.estabelecimentoId,
          dataReferenciaInicio: input?.dataReferenciaInicio,
          dataReferenciaFim: input?.dataReferenciaFim,
          dataPagamento: input?.dataPagamento,
          tipo: input?.tipo,
          codigoGlosa: input?.codigoGlosa,
          motivoGlosa: input?.motivoGlosa,
          classificacao: input?.classificacao,
          search: input?.search,
          page: input?.page || 1,
          pageSize: input?.pageSize || 50,
        });
      }),

    // Itens glosados aceitos (separados)
    itensGlosadosAceitos: protectedProcedure
      .input(
        z
          .object({
          convenioId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
          dataReferenciaInicio: z.date().optional(),
          dataReferenciaFim: z.date().optional(),
          dataPagamento: z.string().optional(),
          search: z.string().optional(),
          page: z.number().default(1),
          pageSize: z.number().default(50),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input?.estabelecimentoId,
          "visualizar"
        );
        return db.getItensGlosadosAceitos({
          convenioId: input?.convenioId,
          estabelecimentoId: input?.estabelecimentoId,
          dataReferenciaInicio: input?.dataReferenciaInicio,
          dataReferenciaFim: input?.dataReferenciaFim,
          dataPagamento: input?.dataPagamento,
          search: input?.search,
          page: input?.page || 1,
          pageSize: input?.pageSize || 50,
        });
      }),
  }),

  // ============ RECURSOS DE GLOSA ============
  recursos: router({
    list: protectedProcedure
      .input(
        z
          .object({
          convenioId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
          status: z.string().optional(),
          prioridade: z.string().optional(),
          dataInicio: z.date().optional(),
          dataFim: z.date().optional(),
          busca: z.string().optional(),
          page: z.number().default(1),
          limit: z.number().default(20),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input?.estabelecimentoId,
          "visualizar"
        );
        return db.getRecursosGlosa(
          ctx.user.id,
          {
            convenioId: input?.convenioId,
            estabelecimentoId: input?.estabelecimentoId,
            status: input?.status,
            prioridade: input?.prioridade,
            dataInicio: input?.dataInicio,
            dataFim: input?.dataFim,
            busca: input?.busca,
          },
          input?.page || 1,
          input?.limit || 20
        );
      }),

    byId: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ input, ctx }) => {
        return exigirAcessoRecurso(ctx.user.id, input.id, "visualizar");
      }),

    create: trackedProtectedProcedure
      .input(
        z.object({
          divergenciaId: z.number().optional(),
          estabelecimentoId: z.number(),
          convenioId: z.number(),
          codigoProcedimento: z.string().optional(),
          descricaoProcedimento: z.string().optional(),
          guiaNumero: z.string().optional(),
          pacienteNome: z.string().optional(),
          valorCobrado: z.string().optional(),
          valorGlosado: z.string().optional(),
          motivoGlosaConvenio: z.string().optional(),
          justificativaRecurso: z.string(),
          prioridade: z
            .enum(["baixa", "media", "alta", "urgente"])
            .default("media"),
          dataGlosa: z.date().optional(),
          dataPrazoResposta: z.date().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirConvenioDoEstabelecimento(
          ctx.user.id,
          input.convenioId,
          input.estabelecimentoId,
          "editar"
        );
        const id = await db.createRecursoGlosa({
          ...input,
          userId: ctx.user.id,
          status: "rascunho",
        });
        return { id };
      }),

    update: trackedProtectedProcedure
      .input(
        z.object({
          id: z.number(),
          data: z.object({
            codigoProcedimento: z.string().optional(),
            descricaoProcedimento: z.string().optional(),
            guiaNumero: z.string().optional(),
            pacienteNome: z.string().optional(),
            valorCobrado: z.string().optional(),
            valorGlosado: z.string().optional(),
            valorRecuperado: z.string().optional(),
            motivoGlosaConvenio: z.string().optional(),
            justificativaRecurso: z.string().optional(),
            prioridade: z
              .enum(["baixa", "media", "alta", "urgente"])
              .optional(),
            status: z
              .enum([
              "rascunho",
              "pendente_envio",
              "enviado",
              "em_analise",
              "deferido",
              "deferido_parcial",
              "indeferido",
                "cancelado",
              ])
              .optional(),
            dataGlosa: z.date().optional(),
            dataEnvioRecurso: z.date().optional(),
            dataPrazoResposta: z.date().optional(),
            dataResposta: z.date().optional(),
            protocoloRecurso: z.string().optional(),
            respostaConvenio: z.string().optional(),
          }),
          descricao: z.string().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoRecurso(ctx.user.id, input.id, "editar");
        await db.updateRecursoGlosa(
          input.id,
          ctx.user.id,
          input.data,
          input.descricao
        );
        return { success: true };
      }),

    delete: trackedProtectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoRecurso(ctx.user.id, input.id, "excluir");
        await db.deleteRecursoGlosa(input.id);
        return { success: true };
      }),

    addHistorico: protectedProcedure
      .input(
        z.object({
          recursoId: z.number(),
          tipo: z.enum([
            "criacao",
            "edicao",
            "envio",
            "resposta_convenio",
            "anexo_adicionado",
            "status_alterado",
            "comentario",
            "lembrete",
          ]),
          descricao: z.string(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoRecurso(ctx.user.id, input.recursoId, "editar");
        await db.addHistoricoRecurso({
          recursoId: input.recursoId,
          userId: ctx.user.id,
          tipo: input.tipo,
          descricao: input.descricao,
        });
        return { success: true };
      }),

    estatisticas: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number().optional() }).optional())
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input?.estabelecimentoId,
          "visualizar"
        );
        return db.getEstatisticasRecursos(
          ctx.user.id,
          input?.estabelecimentoId
        );
      }),

    enviar: protectedProcedure
      .input(
        z.object({
          id: z.number(),
          protocoloRecurso: z.string().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoRecurso(ctx.user.id, input.id, "editar");
        await db.updateRecursoGlosa(
          input.id,
          ctx.user.id,
          {
            status: "enviado",
            dataEnvioRecurso: new Date(),
            protocoloRecurso: input.protocoloRecurso,
          },
          "Recurso enviado ao convênio"
        );

        // Atualizar status dos procedimentos vinculados ao recurso para "enviado"
        await db.atualizarStatusProcedimentosPorRecurso(
          input.id,
          "recurso_enviado"
        );

        return { success: true };
      }),

    registrarResposta: protectedProcedure
      .input(
        z.object({
          id: z.number(),
          status: z.enum(["deferido", "deferido_parcial", "indeferido"]),
          respostaConvenio: z.string(),
          valorRecuperado: z.string().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoRecurso(ctx.user.id, input.id, "editar");
        await db.updateRecursoGlosa(
          input.id,
          ctx.user.id,
          {
            status: input.status,
            dataResposta: new Date(),
            respostaConvenio: input.respostaConvenio,
            valorRecuperado: input.valorRecuperado,
          },
          `Resposta do convênio: ${input.status}`
        );

        // Registrar no histórico de contestações para aprendizado da IA
        const recurso = await db.getRecursoById(input.id);
        if (recurso) {
          await db.registrarHistoricoContestacao({
            recursoId: input.id,
            convenioId: recurso.convenioId,
            userId: ctx.user.id,
            codigoGlosa:
              recurso.motivoGlosaConvenio?.match(/\d{4}/)?.[0] || "0000",
            descricaoGlosa: recurso.motivoGlosaConvenio || "",
            codigoProcedimento: recurso.codigoProcedimento || undefined,
            descricaoProcedimento: recurso.descricaoProcedimento || undefined,
            valorGlosado: recurso.valorGlosado || undefined,
            valorRecuperado: input.valorRecuperado || undefined,
            argumentoUtilizado: recurso.justificativaRecurso,
            argumentoOrigem: "manual",
            resultado: input.status,
            argumentoEfetivo:
              input.status === "deferido"
                ? "sim"
                : input.status === "deferido_parcial"
                  ? "parcial"
                  : "nao",
          });

          // Atualizar status dos procedimentos vinculados ao recurso
          const novoStatusProcedimento =
            input.status === "deferido" || input.status === "deferido_parcial"
            ? "recurso_deferido"
            : "recurso_indeferido";

          await db.atualizarStatusProcedimentosPorRecurso(
            input.id,
            novoStatusProcedimento
          );
        }

        return { success: true };
      }),

    // Atualizar valor recebido e data de pagamento de um item de recurso
    atualizarPagamentoItem: protectedProcedure
      .input(pagamentoRecursoManualSchema)
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoRecurso(ctx.user.id, input.recursoId, "editar");
        const updateData = dadosPagamentoManual(input);

        await db.updateRecursoGlosa(
          input.recursoId,
          ctx.user.id,
          updateData,
          "Atualização de pagamento do recurso"
        );

        // Recalcular o total do lote após atualização do item
        const recurso = await db.getRecursoById(input.recursoId);
        if (recurso && recurso.loteId) {
          await db.recalcularTotalLote(recurso.loteId);
        }

        return { success: true };
      }),

    // Sugerir argumento com IA baseado no histórico
    sugerirArgumentoIA: protectedProcedure
      .input(
        z.object({
          codigoGlosa: z.string(),
          convenioId: z.number(),
          estabelecimentoId: z.number(),
          codigoProcedimento: z.string().optional(),
          valorGlosado: z.string().optional(),
          dataReferencia: z.string().regex(/^\d{4}-\d{2}-\d{2}$/).optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirConvenioDoEstabelecimento(
          ctx.user.id,
          input.convenioId,
          input.estabelecimentoId,
          "visualizar"
        );
        return db.sugerirArgumentoComIA({
          codigoGlosa: input.codigoGlosa,
          convenioId: input.convenioId,
          estabelecimentoId: input.estabelecimentoId,
          codigoProcedimento: input.codigoProcedimento,
          valorGlosado: input.valorGlosado,
          userId: ctx.user.id,
          dataReferencia: input.dataReferencia,
        });
      }),

    // Listar histórico de contestações
    historicoContestacoes: protectedProcedure
      .input(
        z.object({
          estabelecimentoId: z.number(),
          convenioId: z.number().optional(),
          codigoGlosa: z.string().optional(),
          resultado: z.string().optional(),
          page: z.number().default(1),
          limit: z.number().default(20),
        })
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input.estabelecimentoId,
          "visualizar"
        );
        if (input.convenioId) {
          await exigirConvenioDoEstabelecimento(
            ctx.user.id,
            input.convenioId,
            input.estabelecimentoId,
            "visualizar"
          );
        }
        return db.getHistoricoContestacoes({
          userId: ctx.user.id,
          estabelecimentoId: input.estabelecimentoId,
          convenioId: input.convenioId,
          codigoGlosa: input.codigoGlosa,
          resultado: input.resultado,
          page: input.page || 1,
          limit: input.limit || 20,
        });
      }),

    // Obter estatísticas de contestações por código de glosa
    estatisticasPorCodigo: protectedProcedure
      .input(
        z.object({
          codigoGlosa: z.string(),
          estabelecimentoId: z.number(),
          convenioId: z.number().optional(),
        })
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input.estabelecimentoId,
          "visualizar"
        );
        if (input.convenioId) {
          await exigirConvenioDoEstabelecimento(
            ctx.user.id,
            input.convenioId,
            input.estabelecimentoId,
            "visualizar"
          );
        }
        return db.getEstatisticasContestacaoPorCodigo({
          codigoGlosa: input.codigoGlosa,
          convenioId: input.convenioId,
          userId: ctx.user.id,
          estabelecimentoId: input.estabelecimentoId,
        });
      }),

    // Obter melhores argumentos para um código de glosa
    melhoresArgumentos: protectedProcedure
      .input(
        z.object({
          codigoGlosa: z.string(),
          estabelecimentoId: z.number(),
          convenioId: z.number().optional(),
          limit: z.number().default(5),
        })
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input.estabelecimentoId,
          "visualizar"
        );
        if (input.convenioId) {
          await exigirConvenioDoEstabelecimento(
            ctx.user.id,
            input.convenioId,
            input.estabelecimentoId,
            "visualizar"
          );
        }
        return db.getMelhoresArgumentos({
          codigoGlosa: input.codigoGlosa,
          convenioId: input.convenioId,
          userId: ctx.user.id,
          estabelecimentoId: input.estabelecimentoId,
          limit: input.limit,
        });
      }),

    // Criar recursos em lote
    createBatch: protectedProcedure
      .input(
        z.object({
          itens: z.array(
            z.object({
              procedimentoId: z.number().optional(), // ID do procedimento para marcar como "recurso criado"
              convenioId: z.number(),
              estabelecimentoId: z.number().optional(), // ID do estabelecimento
              codigoProcedimento: z.string().optional(),
              descricaoProcedimento: z.string().optional(),
              guiaNumero: z.string().optional(),
              pacienteNome: z.string().optional(),
              pacienteCarteirinha: z.string().optional(),
              valorCobrado: z.string().optional(),
              valorGlosado: z.string().optional(),
              motivoGlosaConvenio: z.string().optional(),
            })
          ),
          estabelecimentoId: z.number().optional(), // ID do estabelecimento (para todos os itens)
          justificativaRecurso: z.string(),
          prioridade: z
            .enum(["baixa", "media", "alta", "urgente"])
            .default("media"),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const itensComEscopo = input.itens.map(item => ({
          item,
          estabelecimentoId: item.estabelecimentoId || input.estabelecimentoId,
        }));
        if (
          itensComEscopo.some(({ estabelecimentoId }) => !estabelecimentoId)
        ) {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message: "Selecione o estabelecimento dos recursos",
          });
        }

        for (const { item, estabelecimentoId } of itensComEscopo) {
          await exigirConvenioDoEstabelecimento(
            ctx.user.id,
            item.convenioId,
            estabelecimentoId,
            "editar"
          );

          if (item.procedimentoId) {
            const demonstrativo = await db.getDemonstrativoById(
              item.procedimentoId
            );
            if (
              !demonstrativo ||
              demonstrativo.estabelecimentoId !== estabelecimentoId ||
              demonstrativo.convenioId !== item.convenioId
            ) {
              throw new TRPCError({
                code: "FORBIDDEN",
                message:
                  "Item do demonstrativo não pertence ao estabelecimento e convênio selecionados",
              });
            }
          }
        }

        const resultados: {
          id: number;
          codigo: string;
          procedimentoId?: number;
        }[] = [];
        const erros: { codigo: string; erro: string }[] = [];
        const procedimentosParaAtualizar: number[] = [];

        for (const item of input.itens) {
          try {
            // Verificar se já existe recurso para este procedimento
            if (item.procedimentoId) {
              const recursoExistente = await db.verificarRecursoExistente(
                item.procedimentoId
              );
              if (recursoExistente) {
                erros.push({
                  codigo: item.codigoProcedimento || "",
                  erro: "Já existe recurso criado para este item",
                });
                continue;
              }
            }

            const id = await db.createRecursoGlosa({
              ...item,
              estabelecimentoId:
                item.estabelecimentoId || input.estabelecimentoId,
              justificativaRecurso: input.justificativaRecurso,
              prioridade: input.prioridade,
              userId: ctx.user.id,
              status: "rascunho",
            });

            resultados.push({
              id,
              codigo: item.codigoProcedimento || "",
              procedimentoId: item.procedimentoId,
            });

            // Marcar procedimento como "recurso criado"
            if (item.procedimentoId) {
              procedimentosParaAtualizar.push(item.procedimentoId);
            }
          } catch (error: any) {
            erros.push({
              codigo: item.codigoProcedimento || "",
              erro: error.message,
            });
          }
        }

        // Atualizar status dos procedimentos em lote
        if (procedimentosParaAtualizar.length > 0) {
          await db.marcarProcedimentosComRecurso(procedimentosParaAtualizar);
        }

        return {
          sucesso: resultados.length,
          falhas: erros.length,
          resultados,
          erros,
        };
      }),
    // Classificar glosa (aceitar ou recursar)
    classificarGlosa: protectedProcedure
      .input(
        z.object({
          procedimentoId: z.number(), // Na verdade é o demonstrativoId
          classificacao: z.enum(["aceitar", "recursar"]),
          motivo: z.string().optional(),
          motivoAceite: z.string().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        // Buscar item do demonstrativo
        const item = await db.getDemonstrativoById(input.procedimentoId);
        if (!item) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Item não encontrado",
          });
        }
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          item.estabelecimentoId,
          "editar"
        );

        const codigoGlosa = (item.codigoGlosa || "").match(/^(\d+)/)?.[1] || "";

        // Registrar decisão para aprendizado
        await db.registrarDecisaoGlosa({
          convenioId: item.convenioId || 0,
          codigoGlosa: codigoGlosa || "0000",
          codigoProcedimento: item.codigoItem || "",
          tipoProcedimento: item.tipoLancamento || null,
          decisao: input.classificacao,
          valorGlosado: item.valorGlosa || undefined,
          motivoDecisao: input.motivo || null,
          procedimentoId: input.procedimentoId,
          userId: ctx.user.id,
        });

        // Atualizar classificação na tabela demonstrativo
        await db.atualizarClassificacaoDemonstrativo(
          input.procedimentoId,
          input.classificacao,
          100,
          input.motivo || "Classificado manualmente pelo usuário",
          input.classificacao === "aceitar" ? input.motivoAceite : undefined
        );

        return { success: true };
      }),

    // Importar decisões (Sim = recursar / Não = glosa aceita) de planilha Excel
    importarDecisoesPlanilha: protectedProcedure
      .input(
        z.object({
          estabelecimentoId: z.number(),
          convenioId: z.number().optional(),
          aplicar: z.boolean().default(false),
          tokenPrevia: z.string().regex(/^[0-9a-f]{64}$/).nullish(),
          linhas: z
            .array(
              z.object({
                linha: z.number(),
                decisao: z.string().max(50),
                protocolo: z.string().max(100).nullish(),
                guiaPrestador: z.string().max(100).nullish(),
                guiaOperadora: z.string().max(100).nullish(),
                sequencialItem: z.union([z.string().max(50), z.number()]).nullish(),
                codigoItem: z.string().max(100).nullish(),
                descricaoItem: z.string().max(1000).nullish(),
                valorGlosa: z.union([z.string().max(50), z.number()]).nullish(),
                justificativa: z.string().max(10000).nullish(),
                observacao: z.string().max(10000).nullish(),
              })
            )
            .min(1)
            .max(20000),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input.estabelecimentoId,
          input.aplicar ? "editar" : "visualizar"
        );
        if (input.convenioId) {
          await exigirConvenioDoEstabelecimento(
            ctx.user.id,
            input.convenioId,
            input.estabelecimentoId,
            input.aplicar ? "editar" : "visualizar"
          );
        }
        try {
          return await importarDecisoesGlosaPlanilha({
            estabelecimentoId: input.estabelecimentoId,
            convenioId: input.convenioId,
            linhas: input.linhas,
            aplicar: input.aplicar,
            tokenPrevia: input.tokenPrevia,
            userId: ctx.user.id,
          });
        } catch (e) {
          throw erroSeguroImportacaoDecisoes(e);
        }
      }),

    // Sugerir classificação com base no histórico
    sugerirClassificacao: protectedProcedure
      .input(
        z.object({
          codigoGlosa: z.string(),
          convenioId: z.number(),
          estabelecimentoId: z.number(),
          codigoProcedimento: z.string().optional(),
        })
      )
      .query(async ({ input, ctx }) => {
        await exigirConvenioDoEstabelecimento(
          ctx.user.id,
          input.convenioId,
          input.estabelecimentoId,
          "visualizar"
        );
        return db.sugerirClassificacaoGlosa(
          input.codigoGlosa,
          input.convenioId,
          input.estabelecimentoId,
          input.codigoProcedimento
        );
      }),

    // Classificar glosas automaticamente
    classificarAutomaticamente: protectedProcedure
      .input(
        z.object({
          estabelecimentoId: z.number(),
          convenioId: z.number().optional(),
          limiteConfianca: z.number().min(50).max(100).default(70),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input.estabelecimentoId,
          "editar"
        );
        if (input.convenioId) {
          await exigirConvenioDoEstabelecimento(
            ctx.user.id,
            input.convenioId,
            input.estabelecimentoId,
            "editar"
          );
        }
        return db.classificarGlosasAutomaticamente(
          input.estabelecimentoId,
          input.convenioId,
          input.limiteConfianca || 70
        );
      }),

    // Buscar histórico de decisões para um código de glosa
    historicoDecisoes: protectedProcedure
      .input(
        z.object({
          codigoGlosa: z.string(),
          estabelecimentoId: z.number(),
          convenioId: z.number().optional(),
        })
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input.estabelecimentoId,
          "visualizar"
        );
        if (input.convenioId) {
          await exigirConvenioDoEstabelecimento(
            ctx.user.id,
            input.convenioId,
            input.estabelecimentoId,
            "visualizar"
          );
        }
        return db.buscarHistoricoDecisoes(
          input.codigoGlosa,
          input.convenioId,
          input.estabelecimentoId
        );
      }),

    // ============ LOTES DE RECURSOS ============
    listarLotes: protectedProcedure
      .input(
        z
          .object({
          estabelecimentoId: z.number().optional(),
          convenioId: z.number().optional(),
            status: z
              .enum([
                "rascunho",
                "pendente_envio",
                "enviado",
                "em_analise",
                "respondido",
                "finalizado",
              ])
              .optional(),
          search: z.string().optional(),
          page: z.number().default(1),
          limit: z.number().default(20),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input?.estabelecimentoId,
          "visualizar"
        );
        return db.listarLotesRecurso({
          estabelecimentoId: input?.estabelecimentoId,
          convenioId: input?.convenioId,
          status: input?.status,
          search: input?.search,
          page: input?.page || 1,
          limit: input?.limit || 20,
        });
      }),

    resumoLotes: protectedProcedure
      .input(
        z
          .object({
          estabelecimentoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input?.estabelecimentoId,
          "visualizar"
        );
        return db.getResumoLotesRecurso(input?.estabelecimentoId);
      }),

    atualizarLote: protectedProcedure
      .input(
        z.object({
          loteId: z.number(),
          status: z
            .enum([
              "rascunho",
              "pendente_envio",
              "enviado",
              "em_analise",
              "respondido",
              "finalizado",
            ])
            .optional(),
          protocoloEnvio: z.string().optional(),
          dataEnvio: z.date().optional(),
          dataPrazoPagamento: z.date().optional(),
          dataResposta: z.date().optional(),
          valorTotalRecuperado: z.string().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const lote = await exigirAcessoLote(
          ctx.user.id,
          input.loteId,
          "editar"
        );
        if (
          input.status === "enviado" ||
          input.dataEnvio !== undefined ||
          input.protocoloEnvio !== undefined ||
          (["rascunho", "pendente_envio"].includes(lote.status) &&
            input.status &&
            !["rascunho", "pendente_envio"].includes(input.status))
        ) {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message:
              "Use Confirmar envio, informando protocolo e data, antes de iniciar o acompanhamento.",
          });
        }
        if (
          !["rascunho", "pendente_envio"].includes(lote.status) &&
          input.status &&
          ["rascunho", "pendente_envio"].includes(input.status)
        ) {
          throw new TRPCError({
            code: "PRECONDITION_FAILED",
            message:
              "O lote já está em acompanhamento; seu status não pode regredir por esta operação.",
          });
        }
        return db.atualizarLoteRecurso(input);
      }),

    confirmarEnvioLote: protectedProcedure
      .input(confirmarEnvioLoteSchema)
      .mutation(async ({ input, ctx }) => {
        const lote = await exigirAcessoLote(
          ctx.user.id,
          input.loteId,
          "editar"
        );
        return confirmarEnvioLote(input, ctx.user.id, {
          estabelecimentoId: lote.estabelecimentoId,
          convenioId: lote.convenioId,
        });
      }),

    detalhesLote: protectedProcedure
      .input(z.object({ loteId: z.number() }))
      .query(async ({ input, ctx }) => {
        await exigirAcessoLote(ctx.user.id, input.loteId, "visualizar");
        return db.getDetalhesLoteRecurso(input.loteId);
      }),

    // ============ RECURSOS AGRUPADOS POR CONVÊNIO ============
    agrupadosPorConvenio: protectedProcedure
      .input(
        z
          .object({
          estabelecimentoId: z.number().optional(),
          status: z.string().optional(),
          dataPagamento: z.string().optional(),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input?.estabelecimentoId,
          "visualizar"
        );
        return db.getRecursosAgrupadosPorConvenio({
          estabelecimentoId: input?.estabelecimentoId,
          status: input?.status,
          dataPagamento: input?.dataPagamento,
        });
      }),

    datasPagamentoDisponiveis: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number().optional() }).optional())
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input?.estabelecimentoId,
          "visualizar"
        );
        const { sql } = await import("drizzle-orm");
        const database = await db.getDb();
        if (!database || !input?.estabelecimentoId) return [];
        const [rows] = await database.execute(
          sql.raw(`
          SELECT DISTINCT data_pagamento as dataPagamento, COUNT(*) as qtd
          FROM demonstrativo
          WHERE estabelecimentoId = ${input.estabelecimentoId}
            AND recurso_status = 'recurso_criado'
            AND recurso_id IS NULL
            AND valor_glosa > 0
            AND data_pagamento IS NOT NULL
          GROUP BY data_pagamento
          ORDER BY data_pagamento DESC
        `)
        );
        return (rows as any[]).map(r => ({
          dataPagamento: r.dataPagamento,
          qtd: Number(r.qtd),
        }));
      }),

    recursosDoConvenio: protectedProcedure
      .input(
        z.object({
          convenioId: z.number(),
          estabelecimentoId: z.number().optional(),
          status: z.string().optional(),
          apenasNaoEnviados: z.boolean().optional(),
          dataPagamento: z.string().optional(),
        })
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input.estabelecimentoId,
          "visualizar"
        );
        return db.getRecursosDoConvenio({
          convenioId: input.convenioId,
          estabelecimentoId: input.estabelecimentoId,
          status: input.status,
          apenasNaoEnviados: input.apenasNaoEnviados,
          dataPagamento: input.dataPagamento,
        });
      }),

    criarLote: protectedProcedure
      .input(
        z.object({
          convenioId: z.number(),
          estabelecimentoId: z.number(),
          recursosIds: z.array(z.number()),
          demonstrativoIds: z.array(z.number()).optional(),
          conciliacaoIds: z.array(z.number()).optional(),
          descricao: z.string().optional(),
          protocoloEnvio: z.string().optional(),
          dataInicioConta: z.string().optional(),
          dataFimConta: z.string().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirConvenioDoEstabelecimento(
          ctx.user.id,
          input.convenioId,
          input.estabelecimentoId,
          "editar"
        );
        return db.criarLoteRecurso({
          convenioId: input.convenioId,
          estabelecimentoId: input.estabelecimentoId,
          userId: ctx.user.id,
          recursosIds: input.recursosIds,
          demonstrativoIds: input.demonstrativoIds,
          conciliacaoIds: input.conciliacaoIds,
          descricao: input.descricao,
          protocoloEnvio: input.protocoloEnvio,
          dataInicioConta: input.dataInicioConta,
          dataFimConta: input.dataFimConta,
        });
      }),

    enviarLote: protectedProcedure
      .input(
        z.object({
          loteId: z.number(),
          protocoloEnvio: z.string().optional(),
          dataPrazoPagamento: z.date().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoLote(ctx.user.id, input.loteId, "editar");
        throw new TRPCError({
          code: "BAD_REQUEST",
          message:
            "Use Confirmar envio para registrar o protocolo, a data e a confirmação do envio externo.",
        });
      }),

    // ============ EXPORTAÇÃO DE RECURSOS ============
    exportar: protectedProcedure
      .input(
        z.object({
          formato: z.enum(["excel", "pdf"]),
          convenioId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
          status: z.string().optional(),
          dataInicio: z.date().optional(),
          dataFim: z.date().optional(),
          ids: z.array(z.number()).optional(), // IDs específicos para exportar
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoEstabelecimento(
          ctx.user.id,
          input.estabelecimentoId,
          "visualizar"
        );
        const recursos = await db.getRecursosParaExportacao(ctx.user.id, {
          convenioId: input.convenioId,
          estabelecimentoId: input.estabelecimentoId,
          status: input.status,
          dataInicio: input.dataInicio,
          dataFim: input.dataFim,
          ids: input.ids,
        });

        if (recursos.length === 0) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Nenhum recurso encontrado para exportação",
          });
        }

        // Formatar dados para exportação
        const dadosExportacao = recursos.map((r, index) => ({
          numero: index + 1,
          convenio: r.convenioNome,
          estabelecimento: r.estabelecimentoNome,
          guia: r.guiaNumero || "N/A",
          paciente: r.pacienteNome || "N/A",
          codigoProcedimento: r.codigoProcedimento || "N/A",
          descricaoProcedimento: r.descricaoProcedimento || "N/A",
          valorCobrado: r.valorCobrado ? parseFloat(r.valorCobrado) : 0,
          valorGlosado: r.valorGlosado ? parseFloat(r.valorGlosado) : 0,
          valorRecuperado: r.valorRecuperado
            ? parseFloat(r.valorRecuperado)
            : 0,
          motivoGlosa: r.motivoGlosaConvenio || "N/A",
          justificativaRecurso: r.justificativaRecurso || "N/A",
          status: r.status,
          prioridade: r.prioridade,
          protocolo: r.protocoloRecurso || "N/A",
          dataGlosa: r.dataGlosa
            ? new Date(r.dataGlosa).toLocaleDateString("pt-BR")
            : "N/A",
          dataEnvio: r.dataEnvioRecurso
            ? new Date(r.dataEnvioRecurso).toLocaleDateString("pt-BR")
            : "N/A",
          dataResposta: r.dataResposta
            ? new Date(r.dataResposta).toLocaleDateString("pt-BR")
            : "N/A",
          respostaConvenio: r.respostaConvenio || "N/A",
          dataCriacao: new Date(r.createdAt).toLocaleDateString("pt-BR"),
        }));

        // Calcular totais
        const totais = {
          totalRecursos: recursos.length,
          totalValorCobrado: dadosExportacao.reduce(
            (sum, r) => sum + r.valorCobrado,
            0
          ),
          totalValorGlosado: dadosExportacao.reduce(
            (sum, r) => sum + r.valorGlosado,
            0
          ),
          totalValorRecuperado: dadosExportacao.reduce(
            (sum, r) => sum + r.valorRecuperado,
            0
          ),
        };

        return {
          dados: dadosExportacao,
          totais,
          formato: input.formato,
          dataExportacao: new Date().toISOString(),
        };
      }),

    // Exportar recursos de um lote específico
    exportarLote: protectedProcedure
      .input(
        z.object({
          loteId: z.number(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoLote(ctx.user.id, input.loteId, "visualizar");
        const dados = await db.getRecursosLoteParaExportacao(input.loteId);
        if (!dados) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Lote não encontrado",
          });
        }
        if (dados.recursos.length === 0) {
          throw new TRPCError({
            code: "PRECONDITION_FAILED",
            message:
              "Os itens deste lote histórico não estão mais vinculados ao lote. A planilha não foi gerada para evitar um arquivo vazio.",
          });
        }
        return dados;
      }),

    // Anexar PDF de protocolo ao lote
    anexarProtocoloLote: protectedProcedure
      .input(
        z.object({
          loteId: z.number(),
          anexoUrl: z.string(),
          anexoKey: z.string(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoLote(ctx.user.id, input.loteId, "editar");
        const sucesso = await db.atualizarAnexoProtocoloLote(
          input.loteId,
          input.anexoUrl,
          input.anexoKey
        );
        if (!sucesso) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: "Erro ao anexar protocolo",
          });
        }
        return { sucesso: true };
      }),

    // Exportar XML ANS/TISS para recurso de glosa (múltiplos XMLs por protocolo)
    exportarXml: protectedProcedure
      .input(
        z.object({
          loteId: z.number(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoLote(ctx.user.id, input.loteId, "visualizar");
        const xmls = await db.gerarXmlsRecursoGlosaPorProtocolo(input.loteId);
        if (!xmls || xmls.length === 0) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Lote não encontrado ou sem recursos",
          });
        }

        // Validar a saída do Portal antes de qualquer persistência.
        const { validarXmlRecursoGlosaDoPortal } = await import(
          "./tiss/recursoGlosaValidation"
        );
        const timestamp = Date.now();
        const resultados = [];

        for (let i = 0; i < xmls.length; i++) {
          const xmlItem = xmls[i];
          const validacao = await validarXmlRecursoGlosaDoPortal(xmlItem.xml);
          if (!validacao.valido) {
            throw new TRPCError({
              code: "PRECONDITION_FAILED",
              message:
                "O XML de recurso gerado não atende ao padrão TISS e não foi exportado.",
            });
          }
          const protSuffix =
            xmls.length > 1 ? `_protocolo_${xmlItem.protocolo}` : "";
          const s3Key = `recursos-glosa/lote-${input.loteId}/recurso_glosa_lote_${input.loteId}${protSuffix}_${timestamp}.xml`;
          const { url: xmlUrl } = await storagePut(
            s3Key,
            Buffer.from(xmlItem.xml, "utf-8"),
            "application/xml"
          );

          resultados.push({
            protocolo: xmlItem.protocolo,
            xml: xmlItem.xml,
            xmlUrl,
            validacao,
            guias: xmlItem.guias,
            valorTotalRecursado: xmlItem.valorTotalRecursado,
          });
        }

        // Atualizar lote com URL do primeiro XML (ou todos)
        if (resultados.length === 1) {
          await db.atualizarXmlLote(
            input.loteId,
            resultados[0].xmlUrl,
            `recursos-glosa/lote-${input.loteId}/recurso_glosa_lote_${input.loteId}_${timestamp}.xml`
          );
        } else {
          // Salvar URLs de todos os XMLs separados por vírgula
          const allUrls = resultados.map(r => r.xmlUrl).join(",");
          await db.atualizarXmlLote(
            input.loteId,
            allUrls,
            `recursos-glosa/lote-${input.loteId}/`
          );
        }

        // Retorno compatível: xml do primeiro + array completo
        return {
          xml: resultados[0].xml,
          validacao: resultados[0].validacao,
          xmlUrl: resultados[0].xmlUrl,
          // Novos campos para múltiplos protocolos
          xmlsPorProtocolo: resultados,
          totalProtocolos: resultados.length,
        };
      }),

    // Exportar todos os XMLs de um lote em um arquivo ZIP
    exportarXmlZip: protectedProcedure
      .input(
        z.object({
          loteId: z.number(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoLote(ctx.user.id, input.loteId, "visualizar");
        const xmls = await db.gerarXmlsRecursoGlosaPorProtocolo(input.loteId);
        if (!xmls || xmls.length === 0) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Lote não encontrado ou sem recursos",
          });
        }

        const archiver = (await import("archiver")).default;
        const { PassThrough } = await import("stream");

        // Criar ZIP em memória
        const chunks: Buffer[] = [];
        const passThrough = new PassThrough();
        passThrough.on("data", (chunk: Buffer) => chunks.push(chunk));

        const archive = archiver("zip", { zlib: { level: 9 } });
        archive.pipe(passThrough);

        // Validar a saída do Portal antes de criar ou persistir o ZIP.
        const { validarXmlRecursoGlosaDoPortal } = await import(
          "./tiss/recursoGlosaValidation"
        );
        const resultados: Array<{
          protocolo: string;
          valido: boolean;
          erros: number;
        }> = [];

        for (const xmlItem of xmls) {
          const validacao = await validarXmlRecursoGlosaDoPortal(xmlItem.xml);
          if (!validacao.valido) {
            throw new TRPCError({
              code: "PRECONDITION_FAILED",
              message:
                "O XML de recurso gerado não atende ao padrão TISS e o ZIP não foi exportado.",
            });
          }
          const fileName =
            xmls.length > 1
            ? `recurso_glosa_lote_${input.loteId}_protocolo_${xmlItem.protocolo}.xml`
            : `recurso_glosa_lote_${input.loteId}.xml`;
          archive.append(xmlItem.xml, { name: fileName });
          resultados.push({
            protocolo: xmlItem.protocolo,
            valido: validacao.valido,
            erros: validacao.erros?.length || 0,
          });
        }

        await archive.finalize();

        // Aguardar todos os chunks
        await new Promise<void>(resolve => passThrough.on("end", resolve));
        const zipBuffer = Buffer.concat(chunks);

        // Salvar ZIP no S3
        const timestamp = Date.now();
        const s3Key = `recursos-glosa/lote-${input.loteId}/recurso_glosa_lote_${input.loteId}_todos_${timestamp}.zip`;
        const { url: zipUrl } = await storagePut(
          s3Key,
          zipBuffer,
          "application/zip"
        );

        return {
          zipUrl,
          totalXmls: xmls.length,
          resultados,
        };
      }),

    // Validar XML sem gerar/salvar
    validarXml: protectedProcedure
      .input(
        z.object({
          loteId: z.number(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        await exigirAcessoLote(ctx.user.id, input.loteId, "visualizar");
        const xml = await db.gerarXmlRecursoGlosa(input.loteId);
        if (!xml) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Lote não encontrado ou sem recursos",
          });
        }
        const { validarXmlRecursoGlosaDoPortal } = await import(
          "./tiss/recursoGlosaValidation"
        );
        const validacao = await validarXmlRecursoGlosaDoPortal(xml);
        return { validacao };
      }),

    // Listar histórico de XMLs gerados para um lote
    historicoXml: protectedProcedure
      .input(
        z.object({
          loteId: z.number(),
        })
      )
      .query(async ({ input, ctx }) => {
        await exigirAcessoLote(ctx.user.id, input.loteId, "visualizar");
        return db.getHistoricoXmlLote(input.loteId);
      }),
  }),

  // ============ CONCILIAÇÃO AUTOMÁTICA ============
  conciliacao: router({
    // Buscar dados para conciliação por convênio
    porConvenio: protectedProcedure
      .input(
        z.object({
          convenioId: z.number(),
          estabelecimentoId: z.number().optional(),
          dataInicio: z.string().optional(),
          dataFim: z.string().optional(),
          mesReferencia: z.number().min(1).max(12).optional(),
          anoReferencia: z.number().min(2000).max(2100).optional(),
          pagina: z.number().min(1).optional().default(1),
          itensPorPagina: z.number().min(10).max(500).optional().default(100),
        })
      )
      .query(async ({ input, ctx }) => {
        return db.getConciliacaoPorConvenio({
          convenioId: input.convenioId,
          userId: ctx.user.id,
          estabelecimentoId: input.estabelecimentoId,
          dataInicio: input.dataInicio ? new Date(input.dataInicio) : undefined,
          dataFim: input.dataFim ? new Date(input.dataFim) : undefined,
          mesReferencia: input.mesReferencia,
          anoReferencia: input.anoReferencia,
          pagina: input.pagina,
          itensPorPagina: input.itensPorPagina,
        });
      }),

    // Resumo da conciliação por convênio
    resumo: protectedProcedure
      .input(
        z
          .object({
          convenioId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
          mesReferencia: z.number().min(1).max(12).optional(),
          anoReferencia: z.number().min(2000).max(2100).optional(),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        return db.getResumoConciliacao({
          convenioId: input?.convenioId,
          userId: ctx.user.id,
          estabelecimentoId: input?.estabelecimentoId,
        });
      }),

    // Buscar conciliação agrupada por conta
    agrupadaPorConta: protectedProcedure
      .input(
        z.object({
          convenioId: z.number(),
          estabelecimentoId: z.number().optional(),
          mesReferencia: z.number().min(1).max(12),
          anoReferencia: z.number().min(2000).max(2100),
          pagina: z.number().min(1).optional().default(1),
          itensPorPagina: z.number().min(10).max(200).optional().default(50),
        })
      )
      .query(async ({ input, ctx }) => {
        return db.getConciliacaoAgrupadaPorConta({
          convenioId: input.convenioId,
          userId: ctx.user.id,
          estabelecimentoId: input.estabelecimentoId,
          mesReferencia: input.mesReferencia,
          anoReferencia: input.anoReferencia,
          pagina: input.pagina,
          itensPorPagina: input.itensPorPagina,
        });
      }),

    // Buscar detalhes de uma conta específica
    detalhesConta: protectedProcedure
      .input(
        z.object({
          convenioId: z.number(),
          guiaNumero: z.string(),
          estabelecimentoId: z.number().optional(),
          mesReferencia: z.number().min(1).max(12).optional(),
          anoReferencia: z.number().min(2000).max(2100).optional(),
        })
      )
      .query(async ({ input, ctx }) => {
        return db.getDetalhesConta({
          convenioId: input.convenioId,
          guiaNumero: input.guiaNumero,
          userId: ctx.user.id,
          estabelecimentoId: input.estabelecimentoId,
          mesReferencia: input.mesReferencia,
          anoReferencia: input.anoReferencia,
        });
      }),

    // Buscar itens não recebidos (pendentes de pagamento)
    naoRecebidos: protectedProcedure
      .input(
        z.object({
          convenioId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
          mesReferencia: z.number().min(1).max(12).optional(),
          anoReferencia: z.number().min(2000).max(2100).optional(),
          pagina: z.number().min(1).optional().default(1),
          itensPorPagina: z.number().min(10).max(500).optional().default(50),
        })
      )
      .query(async ({ input, ctx }) => {
        return db.getItensNaoRecebidos({
          convenioId: input.convenioId,
          userId: ctx.user.id,
          estabelecimentoId: input.estabelecimentoId,
          mesReferencia: input.mesReferencia,
          anoReferencia: input.anoReferencia,
          pagina: input.pagina,
          itensPorPagina: input.itensPorPagina,
        });
      }),
  }),

  // ============ TENDÊNCIAS DE GLOSA ============
  tendencias: router({
    // Tendências de glosa por convênio
    porConvenio: protectedProcedure
      .input(
        z
          .object({
          convenioId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
          meses: z.number().optional().default(6),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        return db.getTendenciasGlosa({
          userId: ctx.user.id,
          convenioId: input?.convenioId,
          estabelecimentoId: input?.estabelecimentoId,
          meses: input?.meses || 6,
        });
      }),

    // Tendência geral (todos os convênios)
    geral: protectedProcedure
      .input(
        z
          .object({
          estabelecimentoId: z.number().optional(),
          meses: z.number().optional().default(6),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        return db.getTendenciaGeral({
          userId: ctx.user.id,
          estabelecimentoId: input?.estabelecimentoId,
          meses: input?.meses || 6,
        });
      }),
  }),

  // ============ REPASSE MÉDICO ============
  repasse: router({
    list: protectedProcedure
      .input(
        z
          .object({
          convenioId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
          dataInicio: z.string().optional(),
          dataFim: z.string().optional(),
          search: z.string().optional(),
          page: z.number().optional().default(1),
          pageSize: z.number().optional().default(50),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        return db.getRepasseData({
          userId: ctx.user.id,
          convenioId: input?.convenioId,
          estabelecimentoId: input?.estabelecimentoId,
          dataInicio: input?.dataInicio
            ? new Date(input.dataInicio)
            : undefined,
          dataFim: input?.dataFim ? new Date(input.dataFim) : undefined,
          search: input?.search,
          page: input?.page || 1,
          pageSize: input?.pageSize || 50,
        });
      }),

    // Exportar todos os dados sem paginação
    exportAll: protectedProcedure
      .input(
        z
          .object({
          convenioId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
          dataInicio: z.string().optional(),
          dataFim: z.string().optional(),
          search: z.string().optional(),
          })
          .optional()
      )
      .query(async ({ input, ctx }) => {
        return db.getRepasseData({
          userId: ctx.user.id,
          convenioId: input?.convenioId,
          estabelecimentoId: input?.estabelecimentoId,
          dataInicio: input?.dataInicio
            ? new Date(input.dataInicio)
            : undefined,
          dataFim: input?.dataFim ? new Date(input.dataFim) : undefined,
          search: input?.search,
          page: 1,
          pageSize: 100000, // Buscar todos os registros
        });
      }),
  }),

  // ============ TABELAS DE PREÇOS ============
  tabelasPreco: router({
    // Listar tabelas de preços
    list: protectedProcedure
      .input(
        z
          .object({
          convenioId: z.number().optional(),
          estabelecimentoId: z.number().optional(),
            tipo: z
              .enum(["diarias", "mat_med", "taxas", "procedimentos"])
              .optional(),
          procedimentoFonte: z.enum(PROCEDIMENTO_FONTES).optional(),
          codigo: z.string().optional(),
          nome: z.string().optional(),
          apenasVigentes: z.boolean().optional().default(true),
          page: z.number().optional().default(1),
          limit: z.number().optional().default(50),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getTabelasPreco({
          convenioId: input?.convenioId,
          estabelecimentoId: input?.estabelecimentoId,
          tipo: input?.tipo,
          procedimentoFonte: input?.procedimentoFonte,
          codigo: input?.codigo,
          nome: input?.nome,
          apenasVigentes: input?.apenasVigentes,
          page: input?.page,
          limit: input?.limit,
        });
      }),

    exportarUnimedOpcoes: protectedProcedure
      .input(
        z.object({
        convenioId: z.number().int().positive(),
        estabelecimentoId: z.number().int().positive(),
        })
      )
      .query(async ({ input }) => {
        return db.getOpcoesExportacaoUnimed(
          input.convenioId,
          input.estabelecimentoId
        );
      }),

    exportarUnimedDbf: protectedProcedure
      .input(
        z.object({
        convenioId: z.number().int().positive(),
        estabelecimentoId: z.number().int().positive(),
          vigencia: z
            .string()
            .regex(
              /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/,
              "Vigência inválida."
            ),
          tipoBeneficiario: z
            .string()
            .min(1, "Selecione o tipo de beneficiário."),
        })
      )
      .mutation(async ({ input }) => {
        const rows = await db.getLinhasExportacaoUnimed(input);
        if (!rows.length) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message:
              "Nenhum procedimento encontrado para os filtros selecionados.",
          });
        }

        const conteudo = await gerarUnimedDbf(rows);
        const beneficiarioSlug = input.tipoBeneficiario
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "")
          .replace(/[^a-zA-Z0-9]+/g, "")
          .replace(/^\s+|\s+$/g, "");
        const nomeArquivo =
          beneficiarioSlug.toUpperCase() === "REDEPADRAO"
          ? "UnimedRedePadrao.DBF"
          : `Unimed${beneficiarioSlug || "Procedimentos"}.DBF`;

        return {
          nomeArquivo,
          conteudo: conteudo.toString("base64"),
          totalItens: rows.length,
        };
      }),

    // Buscar item por ID
    getById: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ input }) => {
        return db.getTabelaPrecoById(input.id);
      }),

    // Criar novo item
    create: trackedProtectedProcedure
      .input(
        z.object({
          convenioId: z.number(),
          estabelecimentoId: z.number().optional(), // Estabelecimento específico ou null para global
          tipo: z.enum(["diarias", "mat_med", "taxas", "procedimentos"]),
          codigo: z.string(),
          nome: z.string(),
          valor: z.string(),
          vigenciaInicio: z.string(),
          procedimentoFonte: z.enum(PROCEDIMENTO_FONTES).optional(),
          porte: z.string().optional(),
          numeroAuxiliares: z.number().int().nonnegative().optional(),
          tipoBeneficiario: z.string().optional(),
          unidade: z.string().optional(),
          observacao: z.string().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const vigenciaInicio = new Date(input.vigenciaInicio);

        const id = await db.createTabelaPreco({
          convenioId: input.convenioId,
          estabelecimentoId: input.estabelecimentoId || null,
          tipo: input.tipo,
          codigo: input.codigo,
          nome: input.nome,
          valor: input.valor,
          vigenciaInicio,
          procedimentoFonte:
            input.tipo === "procedimentos"
              ? input.procedimentoFonte || "site"
              : null,
          porte: input.tipo === "procedimentos" ? input.porte || null : null,
          numeroAuxiliares:
            input.tipo === "procedimentos"
              ? (input.numeroAuxiliares ?? null)
              : null,
          tipoBeneficiario:
            input.tipo === "procedimentos"
              ? input.tipoBeneficiario || null
              : null,
          unidade: input.unidade,
          observacao: input.observacao,
        });

        // Registrar histórico de criação
        if (id) {
          await db.registrarHistoricoPreco({
            tabelaPrecoId: id,
            userId: ctx.user.id,
            tipoAlteracao: "criacao",
            valorNovo: input.valor,
            vigenciaInicioNovo: vigenciaInicio,
            nomeNovo: input.nome,
            codigoNovo: input.codigo,
          });
        }

        return { id };
      }),

    // Atualizar item
    update: trackedProtectedProcedure
      .input(
        z.object({
          id: z.number(),
          codigo: z.string().optional(),
          nome: z.string().optional(),
          valor: z.string().optional(),
          vigenciaInicio: z.string().optional(),
          procedimentoFonte: z.enum(PROCEDIMENTO_FONTES).optional(),
          porte: z.string().optional(),
          numeroAuxiliares: z.number().int().nonnegative().optional(),
          tipoBeneficiario: z.string().optional(),
          unidade: z.string().optional(),
          observacao: z.string().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const { id, ...data } = input;

        // Buscar dados anteriores para o histórico
        const itemAnterior = await db.getTabelaPrecoById(id);

        const updateData: any = {};
        if (data.codigo) updateData.codigo = data.codigo;
        if (data.nome) updateData.nome = data.nome;
        if (data.valor) updateData.valor = data.valor;
        if (data.vigenciaInicio)
          updateData.vigenciaInicio = new Date(data.vigenciaInicio);
        if (data.procedimentoFonte !== undefined)
          updateData.procedimentoFonte = data.procedimentoFonte;
        if (data.porte !== undefined) updateData.porte = data.porte;
        if (data.numeroAuxiliares !== undefined)
          updateData.numeroAuxiliares = data.numeroAuxiliares;
        if (data.tipoBeneficiario !== undefined)
          updateData.tipoBeneficiario = data.tipoBeneficiario;
        if (data.unidade !== undefined) updateData.unidade = data.unidade;
        if (data.observacao !== undefined)
          updateData.observacao = data.observacao;

        const success = await db.updateTabelaPreco(id, updateData);

        // Registrar histórico de edição
        if (success && itemAnterior) {
          await db.registrarHistoricoPreco({
            tabelaPrecoId: id,
            userId: ctx.user.id,
            tipoAlteracao: "edicao",
            // Valores anteriores
            valorAnterior: itemAnterior.valor,
            vigenciaInicioAnterior: itemAnterior.vigenciaInicio,
            nomeAnterior: itemAnterior.nome,
            codigoAnterior: itemAnterior.codigo,
            // Valores novos
            valorNovo: data.valor || itemAnterior.valor,
            vigenciaInicioNovo: data.vigenciaInicio
              ? new Date(data.vigenciaInicio)
              : itemAnterior.vigenciaInicio,
            nomeNovo: data.nome || itemAnterior.nome,
            codigoNovo: data.codigo || itemAnterior.codigo,
          });
        }

        return { success };
      }),

    // Excluir item (soft delete)
    delete: trackedProtectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input, ctx }) => {
        // Buscar dados anteriores para o histórico
        const itemAnterior = await db.getTabelaPrecoById(input.id);

        const success = await db.deleteTabelaPreco(input.id);

        // Registrar histórico de exclusão
        if (success && itemAnterior) {
          await db.registrarHistoricoPreco({
            tabelaPrecoId: input.id,
            userId: ctx.user.id,
            tipoAlteracao: "exclusao",
            valorAnterior: itemAnterior.valor,
            vigenciaInicioAnterior: itemAnterior.vigenciaInicio,
            nomeAnterior: itemAnterior.nome,
            codigoAnterior: itemAnterior.codigo,
          });
        }

        return { success };
      }),

    // Contar itens por tipo
    contarPorTipo: protectedProcedure
      .input(
        z.object({
          convenioId: z.number(),
          estabelecimentoId: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.contarItensTabelaPreco(
          input.convenioId,
          input.estabelecimentoId
        );
      }),

    // Buscar histórico de alterações de um item
    getHistorico: protectedProcedure
      .input(z.object({ tabelaPrecoId: z.number() }))
      .query(async ({ input }) => {
        return db.getHistoricoPreco(input.tabelaPrecoId);
      }),

    // Buscar histórico de alterações por convênio (relatório)
    getHistoricoConvenio: protectedProcedure
      .input(
        z.object({
          convenioId: z.number(),
          tipo: z.string().optional(),
          dataInicio: z.string().optional(),
          dataFim: z.string().optional(),
          userId: z.number().optional(),
          tipoAlteracao: z
            .enum(["criacao", "edicao", "exclusao", "importacao"])
            .optional(),
          page: z.number().optional().default(1),
          limit: z.number().optional().default(50),
        })
      )
      .query(async ({ input }) => {
        return db.getHistoricoPrecosPorConvenio(input.convenioId, {
          tipo: input.tipo,
          dataInicio: input.dataInicio ? new Date(input.dataInicio) : undefined,
          dataFim: input.dataFim ? new Date(input.dataFim) : undefined,
          userId: input.userId,
          tipoAlteracao: input.tipoAlteracao,
          page: input.page,
          limit: input.limit,
        });
      }),

    // Importar tabela de preços (Excel, CSV, DBF)
    importar: protectedProcedure
      .input(
        z.object({
          convenioId: z.number(),
          estabelecimentoId: z.number().optional(), // Estabelecimento específico ou null para global
          tipo: z.enum(["diarias", "mat_med", "taxas", "procedimentos"]),
          nomeArquivo: z.string(),
          formatoArquivo: z.enum(["excel", "csv", "dbf"]),
          procedimentoFonte: z.enum(PROCEDIMENTO_FONTES).optional(),
          conteudo: z.string(), // Base64 encoded file content
          substituirExistentes: z.boolean().optional().default(false),
        })
      )
      .mutation(async ({ input, ctx }) => {
        // Registrar importação
        const importacaoId = await db.createImportacaoTabela({
          convenioId: input.convenioId,
          userId: ctx.user.id,
          tipo: input.tipo,
          nomeArquivo: input.nomeArquivo,
          formatoArquivo: input.formatoArquivo,
          status: "processando",
        });

        if (!importacaoId) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: "Erro ao registrar importação",
          });
        }

        try {
          // Decodificar conteúdo Base64 e ler o primeiro conjunto de dados.
          const buffer = Buffer.from(input.conteudo, "base64");
          let rawItems: Array<Record<string, unknown>> = [];
          const fontePeloNome =
            input.tipo === "procedimentos"
            ? detectarFonteProcedimentoPorNome(input.nomeArquivo)
            : undefined;
          let procedimentoFonte =
            input.tipo === "procedimentos"
            ? fontePeloNome || input.procedimentoFonte || "site"
            : undefined;

          if (input.formatoArquivo === "excel") {
            const XLSX = await import("xlsx");
            const workbook = XLSX.read(buffer, { type: "buffer" });
            const sheetNames =
              input.tipo === "procedimentos"
              ? workbook.SheetNames
              : workbook.SheetNames.slice(0, 1);
            let encontrouCabecalho = false;
            const fontesCandidatas =
              input.tipo === "procedimentos"
                ? ([
                  fontePeloNome,
                  input.procedimentoFonte,
                  "dados_completo",
                  "porte_anestesico",
                  "preco_portes",
                  "site",
                  ].filter(Boolean) as Array<
                    (typeof PROCEDIMENTO_FONTES)[number]
                  >)
              : [];

            for (const sheetName of sheetNames) {
              const sheet = workbook.Sheets[sheetName];
              const rows = XLSX.utils.sheet_to_json(sheet, {
                header: 1,
                defval: null,
                raw: true,
              }) as unknown[][];
              const fonteDaAba = fontesCandidatas.find(
                fonte => encontrarLinhaCabecalhoProcedimentos(rows, fonte) >= 0
              );
              const headerRow = fonteDaAba
                ? encontrarLinhaCabecalhoProcedimentos(rows, fonteDaAba)
                : -1;

              if (headerRow >= 0) {
                encontrouCabecalho = true;
                procedimentoFonte = fonteDaAba;
                rawItems = XLSX.utils.sheet_to_json(sheet, {
                  range: headerRow,
                  defval: null,
                  raw: true,
                }) as Array<Record<string, unknown>>;
                break;
              }
            }

            if (input.tipo === "procedimentos" && !encontrouCabecalho) {
              throw new Error(
                `Não foi encontrada uma aba com as colunas do formato ${procedimentoFonte}. Verifique se o arquivo contém codigo, descricao, porte, numero_auxiliares, valor e tipo_de_beneficiario.`
              );
            }

            if (input.tipo !== "procedimentos" && !rawItems.length) {
              const sheet = workbook.Sheets[workbook.SheetNames[0]];
              rawItems = XLSX.utils.sheet_to_json(sheet) as Array<
                Record<string, unknown>
              >;
            }
          } else if (input.formatoArquivo === "csv") {
            const { parse } = await import("csv-parse/sync");
            rawItems = parse(buffer.toString("utf-8"), {
              columns: true,
              skip_empty_lines: true,
              bom: true,
              relax_column_count: true,
            }) as Array<Record<string, unknown>>;
          } else if (input.formatoArquivo === "dbf") {
            const DBFFile = await import("dbffile");
            const tempPath = `/tmp/import_${Date.now()}_${Math.random().toString(36).slice(2)}.dbf`;
            const fs = await import("fs/promises");
            try {
              await fs.writeFile(tempPath, buffer);
              const dbf = await DBFFile.DBFFile.open(tempPath);
              rawItems = (await dbf.readRecords()) as Array<
                Record<string, unknown>
              >;
            } finally {
              await fs.rm(tempPath, { force: true });
            }
          }

          let tabelaItems: any[] = [];
          let importErrors: string[] = [];

          if (input.tipo === "procedimentos") {
            const fonte = procedimentoFonte || "site";
            if (fonte === "preco_portes" || fonte === "porte_anestesico") {
              throw new Error(
                "As planilhas Preço de Portes e Porte Anestésico são apenas tabelas de lookup. Importe os quatro arquivos pelo lote consolidado; nenhum registro auxiliar será salvo."
              );
            }
            let precoPortes = new Map<string, number>();
            let portesAnestesicos = new Map<string, number>();

            // O formato Dados Completos usa valores simbólicos como 3C,
            // resolvidos contra a tabela de Preço de Portes do próprio Safatle.
            if (fonte === "dados_completo") {
              const referencia = await db.getTabelasPreco({
                convenioId: input.convenioId,
                estabelecimentoId: input.estabelecimentoId,
                tipo: "procedimentos",
                procedimentoFonte: "preco_portes",
                apenasVigentes: false,
                page: 1,
                limit: 100000,
              });
              precoPortes = new Map(
                referencia.items.map(item => [
                  normalizarCodigo(item.codigo, "preco_portes"),
                  Number(item.valor),
                ])
              );

              const referenciaAnestesica = await db.getTabelasPreco({
                convenioId: input.convenioId,
                estabelecimentoId: input.estabelecimentoId,
                tipo: "procedimentos",
                procedimentoFonte: "porte_anestesico",
                apenasVigentes: false,
                page: 1,
                limit: 100000,
              });
              portesAnestesicos = new Map(
                referenciaAnestesica.items.map(item => [
                  `AN${normalizarCodigo(item.codigo, "porte_anestesico")}`,
                  Number(item.valor),
                ])
              );
            }

            const normalized = normalizarProcedimentos(
              rawItems,
              fonte,
              precoPortes,
              portesAnestesicos
            );
            const portesNaoResolvidos = normalized.errors.some(
              error =>
                error.includes("Porte ") &&
                error.includes("não encontrado na tabela de Preço de Portes")
            );
            if (fonte === "dados_completo" && portesNaoResolvidos) {
              throw new Error(
                precoPortes.size === 0
                ? "ATENÇÃO: importe primeiro a planilha de Preço de Portes para este convênio e estabelecimento. Ela é necessária para converter valores como 3C, 4A e 5B."
                  : "A tabela de Preço de Portes não possui todos os códigos necessários para esta planilha. Confira os códigos de porte e importe a referência correta antes de tentar novamente."
              );
            }
            const anestesicosNaoResolvidos = normalized.errors.some(
              error =>
                error.includes("Porte anestésico") &&
                error.includes("não encontrado")
            );
            if (fonte === "dados_completo" && anestesicosNaoResolvidos) {
              throw new Error(
                portesAnestesicos.size === 0
                ? "ATENÇÃO: importe primeiro a planilha de Porte Anestésico para este convênio e estabelecimento. Ela é necessária para converter portes como 5 em AN5 (R$ 806,41)."
                  : "A tabela de Porte Anestésico não possui todos os portes numéricos necessários para esta planilha."
              );
            }
            importErrors = normalized.errors;
            if (!normalized.items.length) {
              throw new Error(
                importErrors[0] ||
                  "Nenhum registro válido foi encontrado no arquivo."
              );
            }

            tabelaItems = normalized.items.map(item => ({
              convenioId: input.convenioId,
              estabelecimentoId: input.estabelecimentoId || null,
              tipo: "procedimentos" as const,
              procedimentoFonte: fonte,
              codigo: item.codigo,
              nome: item.nome,
              valor: item.valor,
              vigenciaInicio: new Date(),
              porte: item.porte || null,
              numeroAuxiliares: item.numeroAuxiliares ?? null,
              tipoBeneficiario: item.tipoBeneficiario || null,
              valorPorteAnestesico: item.valorPorteAnestesico || null,
              unidade: null,
              observacao: null,
            }));
          } else {
            const findValue = (
              item: Record<string, unknown>,
              keys: string[]
            ): unknown => {
              for (const key of keys) {
                if (item[key] !== undefined) return item[key];
                if (item[`${key} `] !== undefined) return item[`${key} `];
                if (item[` ${key}`] !== undefined) return item[` ${key}`];
              }
              return undefined;
            };

            const parseDate = (value: unknown): Date => {
              if (value instanceof Date && !Number.isNaN(value.getTime()))
                return value;
              const text = String(value ?? "").trim();
              const brMatch = text.match(
                /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/
              );
              if (brMatch)
                return new Date(
                  Number(brMatch[3]),
                  Number(brMatch[2]) - 1,
                  Number(brMatch[1])
                );
              const isoMatch = text.match(
                /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/
              );
              if (isoMatch)
                return new Date(
                  Number(isoMatch[1]),
                  Number(isoMatch[2]) - 1,
                  Number(isoMatch[3])
                );
              const serial = Number(text);
              if (Number.isFinite(serial) && serial > 1 && serial < 100000) {
                return new Date(
                  new Date(1899, 11, 30).getTime() +
                    serial * 24 * 60 * 60 * 1000
                );
              }
              const parsed = new Date(text);
              return Number.isNaN(parsed.getTime()) ? new Date() : parsed;
            };

            tabelaItems = rawItems
              .map(item => {
                const codigo = findValue(item, [
                  "codigo",
                  "Codigo",
                  "CODIGO",
                  "cod",
                  "Cod",
                  "COD",
                  "code",
                  "Code",
                  "CODE",
                ]);
                const nome = findValue(item, [
                  "nome",
                  "Nome",
                  "NOME",
                  "descricao",
                  "Descricao",
                  "DESCRICAO",
                  "name",
                  "Name",
                  "NAME",
                  "desc",
                  "Desc",
                  "DESC",
                ]);
                const valor = parseValor(
                  findValue(item, [
                    "valor",
                    "Valor",
                    "VALOR",
                    "preco",
                    "Preco",
                    "PRECO",
                    "price",
                    "Price",
                    "PRICE",
                    "vl",
                    "Vl",
                    "VL",
                  ])
                );
              const vigenciaInicioRaw = findValue(item, [
                  "vigencia_inicio",
                  "Vigencia_inicio",
                  "VIGENCIA_INICIO",
                  "inicio_vigencia",
                  "Inicio_vigencia",
                  "INICIO_VIGENCIA",
                  "vigenciaInicio",
                  "VigenciaInicio",
                  "VIGENCIAINICIO",
                  "data_inicio",
                  "Data_inicio",
                  "DATA_INICIO",
                  "dataInicio",
                  "DataInicio",
                  "DATAINICIO",
                  "inicio",
                  "Inicio",
                  "INICIO",
              ]);
              return {
                convenioId: input.convenioId,
                estabelecimentoId: input.estabelecimentoId || null,
                tipo: input.tipo,
                codigo: String(codigo ?? "").trim(),
                nome: String(nome ?? "").trim(),
                valor: valor.toFixed(2),
                vigenciaInicio: parseDate(vigenciaInicioRaw),
                  unidade: findValue(item, [
                    "unidade",
                    "Unidade",
                    "UNIDADE",
                    "unit",
                    "Unit",
                    "UNIT",
                  ]),
                  observacao: findValue(item, [
                    "observacao",
                    "Observacao",
                    "OBSERVACAO",
                    "obs",
                    "Obs",
                    "OBS",
                  ]),
              };
              })
              .filter(item => item.codigo && item.nome);
          }

          // Só desativar os registros anteriores depois que o arquivo foi lido
          // e validado; assim um erro não deixa a tabela sem dados ativos.
          if (input.substituirExistentes) {
            await db.desativarTabelasPreco(
              input.convenioId,
              input.tipo,
              input.tipo === "procedimentos"
                ? input.procedimentoFonte || "site"
                : undefined
            );
          }

          // Inserir em lote
          const resultado = await db.createTabelasPrecoEmLote(tabelaItems);

          // Atualizar status da importação
          await db.updateImportacaoTabela(importacaoId, {
            status: "concluido",
            totalItens: rawItems.length,
            itensImportados: resultado.inseridos,
            itensErro: resultado.erros + importErrors.length,
            mensagemErro: importErrors.length
              ? importErrors.slice(0, 20).join(" | ")
              : null,
          });

          return {
            success: true,
            importacaoId,
            totalItens: rawItems.length,
            itensImportados: resultado.inseridos,
            itensErro: resultado.erros + importErrors.length,
          };
        } catch (error: any) {
          // Atualizar status com erro
          await db.updateImportacaoTabela(importacaoId, {
            status: "erro",
            mensagemErro: error.message,
          });

          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao processar arquivo: ${error.message}`,
          });
        }
      }),

    /**
     * Importa os quatro arquivos Unimed em memória e só então publica o
     * resultado consolidado. Preço de Portes e Porte Anestésico são lookup;
     * nunca são persistidos como registros independentes.
     */
    importarLoteProcedimentos: protectedProcedure
      .input(
        z.object({
        convenioId: z.number(),
        estabelecimentoId: z.number().optional(),
          arquivos: z
            .array(
              z.object({
          nomeArquivo: z.string().min(1),
          formatoArquivo: z.enum(["excel", "csv", "dbf"]),
          conteudo: z.string().min(1),
              })
            )
            .length(4),
        substituirExistentes: z.boolean().default(false),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const importacaoId = await db.createImportacaoTabela({
          convenioId: input.convenioId,
          userId: ctx.user.id,
          tipo: "procedimentos",
          nomeArquivo:
            `Lote Unimed: ${input.arquivos.map(arquivo => arquivo.nomeArquivo).join(" + ")}`.slice(
              0,
              255
            ),
          formatoArquivo: "excel",
          status: "processando",
        });

        if (!importacaoId) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: "Erro ao registrar o lote de importação.",
          });
        }

        try {
          const fontes = input.arquivos.map(arquivo =>
            detectarFonteProcedimentoPorNome(arquivo.nomeArquivo)
          );
          if (fontes.some(fonte => !fonte)) {
            throw new Error(
              "Um ou mais arquivos não têm nome reconhecido. Use os quatro nomes padrão da Unimed."
            );
          }
          const fontesValidas = fontes.filter(
            (fonte): fonte is (typeof PROCEDIMENTO_FONTES)[number] => !!fonte
          );
          const duplicadas = fontesValidas.filter(
            (fonte, index) => fontesValidas.indexOf(fonte) !== index
          );
          if (
            duplicadas.length > 0 ||
            new Set(fontesValidas).size !== PROCEDIMENTO_FONTES.length
          ) {
            throw new Error(
              "O lote deve conter exatamente um arquivo de Preço de Portes, um de Porte Anestésico, um de Dados Completos e um da Tabela Unimed Site."
            );
          }

          const consolidado = await lerEConsolidarArquivosProcedimentos(
            input.arquivos
          );
          if (consolidado.items.length === 0) {
            throw new Error(
              "Nenhum procedimento válido foi encontrado após a consolidação dos quatro arquivos."
            );
          }

          const vigenciaInicio = new Date();
          const tabelaItems = consolidado.items.map(item => ({
            convenioId: input.convenioId,
            estabelecimentoId: input.estabelecimentoId || null,
            tipo: "procedimentos" as const,
            procedimentoFonte: item.procedimentoFonte,
            codigo: item.codigo,
            nome: item.nome,
            valor: item.valor,
            vigenciaInicio,
            porte: item.porte || null,
            numeroAuxiliares: item.numeroAuxiliares ?? null,
            tipoBeneficiario: item.tipoBeneficiario || null,
            valorPorteAnestesico: item.valorPorteAnestesico || null,
            unidade: null,
            observacao: null,
          }));

          const persistido = await db.publicarProcedimentosConsolidados(
            tabelaItems,
            input.convenioId,
            input.estabelecimentoId,
            input.substituirExistentes
          );

          await db.updateImportacaoTabela(importacaoId, {
            status: "concluido",
            totalItens: consolidado.totalEntrada,
            itensImportados: persistido.inseridos,
            itensErro: 0,
            mensagemErro: null,
          });

          return {
            success: true,
            importacaoId,
            totalItens: consolidado.totalEntrada,
            itensImportados: persistido.inseridos,
            itensErro: 0,
            itensDuplicadosSite: consolidado.totalDuplicadosSite,
            auxiliaresRemovidos: persistido.auxiliaresRemovidos,
          };
        } catch (error: any) {
          await db.updateImportacaoTabela(importacaoId, {
            status: "erro",
            mensagemErro: error.message,
          });
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao consolidar os quatro arquivos: ${error.message}`,
          });
        }
      }),

    // Listar histórico de importações
    historicoImportacoes: protectedProcedure
      .input(z.object({ convenioId: z.number().optional() }).optional())
      .query(async ({ input }) => {
        return db.getImportacoesTabela(input?.convenioId);
      }),
  }),

  // ==================== REGRAS DE NEGÓCIO ====================
  regrasNegocio: router({
    // Listar regras
    list: protectedProcedure
      .input(
        z
          .object({
        convenioId: z.number().optional(),
        estabelecimentoId: z.number().optional(),
        ativo: z.enum(["sim", "nao"]).optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getRegrasNegocio(input || {});
      }),

    // Obter regra por ID
    getById: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ input }) => {
        return db.getRegraNegocioById(input.id);
      }),

    // Criar regra
    create: protectedProcedure
      .input(
        z.object({
        convenioId: z.number().optional(),
        estabelecimentoId: z.number().optional(),
        nome: z.string().min(1),
        descricao: z.string().optional(),
        codigoProcedimentoPrincipal: z.string().min(1),
        descricaoProcedimentoPrincipal: z.string().optional(),
          tipoVerificacao: z.enum([
            "deve_conter",
            "nao_deve_conter",
            "pode_conter",
            "quantidade_minima",
            "quantidade_maxima",
          ]),
          acaoInconsistencia: z.enum([
            "alerta",
            "bloquear",
            "sugerir_adicao",
            "sugerir_remocao",
          ]),
        prioridade: z.number().optional(),
          itens: z
            .array(
              z.object({
          codigoItem: z.string().min(1),
          descricaoItem: z.string().optional(),
                tipoItem: z.enum([
                  "procedimento",
                  "taxa",
                  "material",
                  "medicamento",
                  "diaria",
                  "outros",
                ]),
          quantidadeMinima: z.number().optional(),
          quantidadeMaxima: z.number().optional(),
          valorEsperado: z.string().optional(),
          toleranciaValor: z.string().optional(),
          obrigatorio: z.enum(["sim", "nao"]).optional(),
              })
            )
            .optional(),
        })
      )
      .mutation(async ({ input }) => {
        const { itens, ...regraData } = input;

        // Criar regra
        const regraId = await db.createRegraNegocio(regraData as any);
        if (!regraId) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: "Erro ao criar regra",
          });
        }

        // Adicionar itens
        if (itens && itens.length > 0) {
          for (const item of itens) {
            await db.addItemRegraNegocio({
              regraId,
              ...item,
            } as any);
          }
        }

        return { id: regraId };
      }),

    // Atualizar regra
    update: protectedProcedure
      .input(
        z.object({
        id: z.number(),
        nome: z.string().optional(),
        descricao: z.string().optional(),
        codigoProcedimentoPrincipal: z.string().optional(),
        descricaoProcedimentoPrincipal: z.string().optional(),
          tipoVerificacao: z
            .enum([
              "deve_conter",
              "nao_deve_conter",
              "pode_conter",
              "quantidade_minima",
              "quantidade_maxima",
            ])
            .optional(),
          acaoInconsistencia: z
            .enum(["alerta", "bloquear", "sugerir_adicao", "sugerir_remocao"])
            .optional(),
        prioridade: z.number().optional(),
        ativo: z.enum(["sim", "nao"]).optional(),
        })
      )
      .mutation(async ({ input }) => {
        const { id, ...data } = input;
        return db.updateRegraNegocio(id, data as any);
      }),

    // Excluir regra
    delete: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input }) => {
        return db.deleteRegraNegocio(input.id);
      }),

    // Adicionar item à regra
    addItem: protectedProcedure
      .input(
        z.object({
        regraId: z.number(),
        codigoItem: z.string().min(1),
        descricaoItem: z.string().optional(),
          tipoItem: z.enum([
            "procedimento",
            "taxa",
            "material",
            "medicamento",
            "diaria",
            "outros",
          ]),
        quantidadeMinima: z.number().optional(),
        quantidadeMaxima: z.number().optional(),
        valorEsperado: z.string().optional(),
        toleranciaValor: z.string().optional(),
        obrigatorio: z.enum(["sim", "nao"]).optional(),
        })
      )
      .mutation(async ({ input }) => {
        return db.addItemRegraNegocio(input as any);
      }),

    // Remover item da regra
    removeItem: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input }) => {
        return db.removeItemRegraNegocio(input.id);
      }),

    // Preview dos padrões Tasy para gerar regras
    previewPadroesTasy: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        convenioId: z.number().optional(),
        competencia: z.string().optional(),
        frequenciaMinima: z.number().default(10),
        })
      )
      .query(async ({ input }) => {
        const { analisarPadroesCobrancaTasy } = await import("./db");
        const analise = await analisarPadroesCobrancaTasy({
          estabelecimentoId: input.estabelecimentoId,
          convenioId: input.convenioId,
          competencia: input.competencia,
        });

        // Filtrar padrões com frequência mínima e que tenham itens associados
        const padroesElegiveis = analise.padroesComposicao
          .filter(
            p =>
              p.frequenciaPercentual >= input.frequenciaMinima &&
              p.itensAssociados.length > 0
          )
          .map(p => ({
            codigoProcedimento: p.codigoProcedimento,
            descricao: p.descricao,
            tipo: p.tipo,
            frequenciaPercentual: p.frequenciaPercentual,
            totalOcorrencias: p.totalOcorrencias,
            contasComItem: p.contasComItem,
            quantidadeMedia: p.quantidadeMedia,
            quantidadeMin: p.quantidadeMin,
            quantidadeMax: p.quantidadeMax,
            valorUnitarioMedio: p.valorUnitarioMedio,
            valorTotalMedio: p.valorTotalMedio,
            itensAssociados: p.itensAssociados
              .filter(a => a.frequencia >= input.frequenciaMinima)
              .map(a => ({
                codigo: a.codigo,
                descricao: a.descricao,
                tipo: a.tipo,
                frequencia: a.frequencia,
                quantidadeMedia: a.quantidadeMedia,
                valorUnitarioMedio: a.valorUnitarioMedio,
              })),
          }))
          .filter(p => p.itensAssociados.length > 0);

        return {
          padroesElegiveis,
          totalContas: analise.totalContas,
          totalItens: analise.totalItens,
          totalPadroes: padroesElegiveis.length,
        };
      }),

    // Gerar regras de negócio a partir dos padrões Tasy
    gerarRegrasDosPadroesTasy: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        convenioId: z.number().optional(),
        competencia: z.string().optional(),
        frequenciaMinima: z.number().default(10),
        padroesSelecionados: z.array(z.string()).optional(), // códigos dos procedimentos a incluir
        })
      )
      .mutation(async ({ input }) => {
        const { analisarPadroesCobrancaTasy } = await import("./db");
        const analise = await analisarPadroesCobrancaTasy({
          estabelecimentoId: input.estabelecimentoId,
          convenioId: input.convenioId,
          competencia: input.competencia,
        });

        const padroesElegiveis = analise.padroesComposicao
          .filter(
            p =>
              p.frequenciaPercentual >= input.frequenciaMinima &&
              p.itensAssociados.length > 0
          )
          .filter(
            p =>
              !input.padroesSelecionados ||
              input.padroesSelecionados.includes(p.codigoProcedimento)
          )
          .filter(
            p =>
              p.itensAssociados.filter(
                a => a.frequencia >= input.frequenciaMinima
              ).length > 0
          );

        let criadas = 0;
        let itensAdicionados = 0;

        const mapTipoItem = (
          tipo: string
        ):
          | "procedimento"
          | "taxa"
          | "material"
          | "medicamento"
          | "diaria"
          | "outros" => {
          const t = tipo.toUpperCase();
          if (
            t.includes("PROC") ||
            t === "P" ||
            t === "C" ||
            t === "O" ||
            t === "01"
          )
            return "procedimento";
          if (t.includes("TAXA")) return "taxa";
          if (t.includes("MAT")) return "material";
          if (t.includes("MED") || t.includes("MEDICAMENTO"))
            return "medicamento";
          if (t.includes("DIAR")) return "diaria";
          return "outros";
        };

        for (const padrao of padroesElegiveis) {
          const itensAssociadosFiltrados = padrao.itensAssociados.filter(
            a => a.frequencia >= input.frequenciaMinima
          );

          if (itensAssociadosFiltrados.length === 0) continue;

          // Criar a regra
          const regraId = await db.createRegraNegocio({
            convenioId: input.convenioId || null,
            estabelecimentoId: input.estabelecimentoId,
            nome: `Padrão Tasy: ${padrao.descricao || padrao.codigoProcedimento}`,
            descricao: `Regra gerada automaticamente a partir de ${padrao.contasComItem} contas do Tasy. Frequência: ${padrao.frequenciaPercentual}%. Valor médio: R$ ${padrao.valorTotalMedio.toFixed(2)}`,
            codigoProcedimentoPrincipal: padrao.codigoProcedimento,
            descricaoProcedimentoPrincipal: padrao.descricao || null,
            tipoVerificacao: "deve_conter",
            acaoInconsistencia: "alerta",
            prioridade:
              padrao.frequenciaPercentual >= 50
                ? 8
                : padrao.frequenciaPercentual >= 25
                  ? 5
                  : 3,
          } as any);

          if (!regraId) continue;
          criadas++;

          // Adicionar itens associados
          for (const item of itensAssociadosFiltrados) {
            await db.addItemRegraNegocio({
              regraId,
              codigoItem: item.codigo,
              descricaoItem: item.descricao || null,
              tipoItem: mapTipoItem(item.tipo),
              quantidadeMinima: Math.max(1, Math.floor(item.quantidadeMedia)),
              quantidadeMaxima: Math.ceil(item.quantidadeMedia * 2),
              valorEsperado:
                item.valorUnitarioMedio > 0
                  ? String(item.valorUnitarioMedio.toFixed(2))
                  : null,
              toleranciaValor:
                item.valorUnitarioMedio > 0
                  ? String((item.valorUnitarioMedio * 0.2).toFixed(2))
                  : null,
              obrigatorio: item.frequencia >= 70 ? "sim" : "nao",
            } as any);
            itensAdicionados++;
          }
        }

        return {
          regrasCriadas: criadas,
          itensAdicionados,
          totalPadroesAnalisados: padroesElegiveis.length,
        };
      }),
  }),

  // ==================== ALERTAS DE DIVERGÊNCIA ====================
  alertas: router({
    // Listar alertas por arquivo
    byArquivo: protectedProcedure
      .input(z.object({ arquivoId: z.number() }))
      .query(async ({ input }) => {
        return db.getAlertasByArquivo(input.arquivoId);
      }),

    // Listar alertas pendentes
    pendentes: protectedProcedure
      .input(
        z
          .object({
        convenioId: z.number().optional(),
        tipoAlerta: z.string().optional(),
        severidade: z.string().optional(),
        page: z.number().optional(),
        pageSize: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getAlertasPendentes(input || {});
      }),

    // Resumo de alertas por arquivo
    resumo: protectedProcedure
      .input(z.object({ arquivoId: z.number() }))
      .query(async ({ input }) => {
        return db.getResumoAlertasArquivo(input.arquivoId);
      }),

    // Atualizar status do alerta
    updateStatus: protectedProcedure
      .input(
        z.object({
        id: z.number(),
          status: z.enum([
            "pendente",
            "analisando",
            "corrigido",
            "ignorado",
            "aceito",
          ]),
        observacao: z.string().optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        return db.updateAlertaStatus(
          input.id,
          input.status,
          ctx.user.id,
          input.observacao
        );
      }),

    // Executar comparação com tabela de preços
    compararComTabela: protectedProcedure
      .input(z.object({ arquivoId: z.number() }))
      .mutation(async ({ input }) => {
        return db.compararComTabelaPrecos(input.arquivoId);
      }),

    // Verificar regras de negócio
    verificarRegras: protectedProcedure
      .input(z.object({ arquivoId: z.number() }))
      .mutation(async ({ input }) => {
        return db.verificarRegrasNegocio(input.arquivoId);
      }),

    // Sugerir itens faltantes (IA)
    sugerirItens: protectedProcedure
      .input(z.object({ arquivoId: z.number() }))
      .query(async ({ input }) => {
        return db.sugerirItensFaltantes(input.arquivoId);
      }),

    // Atualizar padrões de conta
    atualizarPadroes: protectedProcedure
      .input(
        z.object({
        convenioId: z.number(),
        estabelecimentoId: z.number().optional(),
        })
      )
      .mutation(async ({ input }) => {
        await db.atualizarPadroesContas(
          input.convenioId,
          input.estabelecimentoId
        );
        return { success: true };
      }),

    // Salvar validação no histórico
    salvarValidacao: protectedProcedure
      .input(
        z.object({
        arquivoId: z.number(),
        convenioId: z.number(),
        totalItens: z.number(),
        divergenciasPreco: z.number(),
        violacoesRegras: z.number(),
        sugestoesIA: z.number(),
        valorDiferenca: z.number(),
        detalhes: z.any(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        return db.salvarHistoricoValidacao({
          ...input,
          userId: ctx.user.id,
        });
      }),

    // Listar histórico de validações
    historicoValidacoes: protectedProcedure
      .input(
        z
          .object({
        convenioId: z.number().optional(),
        arquivoId: z.number().optional(),
        limit: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.listarHistoricoValidacoes(input || {});
      }),

    // Buscar validação por ID
    getValidacao: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ input }) => {
        return db.getHistoricoValidacao(input.id);
      }),

    // Salvar histórico de validação XML
    salvarHistoricoXml: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        nomeArquivo: z.string(),
        dataProcessamento: z.date(),
        totalContas: z.number(),
        contasValidas: z.number(),
        contasInvalidas: z.number(),
        scoreConformidadeMedio: z.number(),
        resultadoCompleto: z.any(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        return db.salvarHistoricoValidacaoXml({
          ...input,
          usuarioId: ctx.user.id,
        });
      }),

    // Listar histórico de validações XML
    listarHistoricoXml: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        usuarioId: z.number().optional(),
        dataInicio: z.date().optional(),
        dataFim: z.date().optional(),
        limit: z.number().optional(),
        offset: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ ctx, input }) => {
        const filtros = {
          ...input,
          estabelecimentoId: input?.estabelecimentoId,
        };
        return db.listarHistoricoValidacoesXml(filtros);
      }),

    // Obter detalhes de uma validação XML
    obterDetalhesXml: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ input }) => {
        return db.obterDetalhesValidacaoXml(input.id);
      }),

    // Obter estatísticas de validações XML
    obterEstatisticasXml: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return db.obterEstatisticasValidacaoXml(input.estabelecimentoId);
      }),
  }),

  // ============ PERMISSÕES E DASHBOARD CONSOLIDADO ============
  permissoes: router({
    // Buscar permissões do usuário atual
    minhasPermissoes: protectedProcedure.query(async ({ ctx }) => {
      return db.getPermissoesUsuario(ctx.user.id);
    }),

    // Buscar estabelecimentos permitidos para o usuário
    estabelecimentosPermitidos: protectedProcedure.query(async ({ ctx }) => {
      // Admins têm acesso a todos
      if (ctx.user.role === "admin") {
        return db.getEstabelecimentos("sim");
      }
      const idsPermitidos = await db.getEstabelecimentosPermitidos(ctx.user.id);
      if (idsPermitidos.length === 0) {
        return [];
      }
      const todosEst = await db.getEstabelecimentos("sim");
      return todosEst.filter(e => idsPermitidos.includes(e.id));
    }),

    // Verificar se usuário é gestor
    verificarGestor: protectedProcedure.query(async ({ ctx }) => {
      return db.verificarSeGestor(ctx.user.id);
    }),

    // Listar somente unidades que o usuário pode administrar
    estabelecimentosGerenciaveis: protectedProcedure.query(async ({ ctx }) => {
      return db.getEstabelecimentosGerenciaveis(ctx.user.id);
    }),

    // Verificar permissão específica para um estabelecimento
    verificarPermissao: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
          tipoPermissao: z
            .enum(["visualizar", "editar", "excluir", "gerenciar"])
            .optional(),
        })
      )
      .query(async ({ ctx, input }) => {
        return db.verificarPermissaoEstabelecimento(
          ctx.user.id,
          input.estabelecimentoId,
          input.tipoPermissao || "visualizar"
        );
      }),

    // Listar usuários de um estabelecimento (apenas gestores)
    usuariosEstabelecimento: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ ctx, input }) => {
        const podeGerenciar = await db.verificarPermissaoEstabelecimento(
          ctx.user.id,
          input.estabelecimentoId,
          "gerenciar"
        );
        if (!podeGerenciar && ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message:
              "Você não tem permissão para gerenciar este estabelecimento",
          });
        }
        return db.getUsuariosEstabelecimento(input.estabelecimentoId);
      }),

    // Criar ou atualizar permissão (apenas gestores/admin)
    upsertPermissao: protectedProcedure
      .input(
        z.object({
        userId: z.number(),
        estabelecimentoId: z.number(),
          grupoServico: z
            .enum([
              "administrador",
              "faturista",
              "recurso_glosa",
              "gestor",
              "visualizador",
              "usuario_tasy",
            ])
            .optional(),
        grupoId: z.number().nullable().optional(),
        podeVisualizar: z.enum(["sim", "nao"]).optional(),
        podeEditar: z.enum(["sim", "nao"]).optional(),
        podeExcluir: z.enum(["sim", "nao"]).optional(),
        podeGerenciar: z.enum(["sim", "nao"]).optional(),
        acessoDashboard: z.enum(["sim", "nao"]).optional(),
        acessoArquivos: z.enum(["sim", "nao"]).optional(),
        acessoComparacoes: z.enum(["sim", "nao"]).optional(),
        acessoFaturamento: z.enum(["sim", "nao"]).optional(),
        acessoTabelasPreco: z.enum(["sim", "nao"]).optional(),
        acessoAnaliseGlosa: z.enum(["sim", "nao"]).optional(),
        acessoDicionarioGlosas: z.enum(["sim", "nao"]).optional(),
        acessoRecursosGlosa: z.enum(["sim", "nao"]).optional(),
        acessoConvenios: z.enum(["sim", "nao"]).optional(),
        acessoRegrasNegocio: z.enum(["sim", "nao"]).optional(),
        acessoProdutividade: z.enum(["sim", "nao"]).optional(),
        acessoEstabelecimentos: z.enum(["sim", "nao"]).optional(),
        acessoPermissoes: z.enum(["sim", "nao"]).optional(),
        acessoImportacaoTasy: z.enum(["sim", "nao"]).optional(),
        acessoContasFaturadas: z.enum(["sim", "nao"]).optional(),
        acessoRelatoriosTasy: z.enum(["sim", "nao"]).optional(),
        acessoRelatoriosBi: z.enum(["sim", "nao"]).optional(),
        acessoConciliacaoContasPagas: z.enum(["sim", "nao"]).optional(),
        acessoRecebimentosXml: z.enum(["sim", "nao"]).optional(),
        acessoRecebimentosExcel: z.enum(["sim", "nao"]).optional(),
        acessoDemonstrativo: z.enum(["sim", "nao"]).optional(),
        acessoContaConvenio: z.enum(["sim", "nao"]).optional(),
        acessoRecursos: z.enum(["sim", "nao"]).optional(),
        acessoAtendimentos: z.enum(["sim", "nao"]).optional(),
        acessoAtendimentosFaturar: z.enum(["sim", "nao"]).optional(),
        acessoRelFaturadoRecebido: z.enum(["sim", "nao"]).optional(),
        acessoRelPrevisaoPagamentosImec: z.enum(["sim", "nao"]).optional(),
        acessoRelLaboratorio: z.enum(["sim", "nao"]).optional(),
        acessoRelUltrassom: z.enum(["sim", "nao"]).optional(),
        acessoRelVisitaHospitalar: z.enum(["sim", "nao"]).optional(),
        acessoRelVisitasFaturadasXml: z.enum(["sim", "nao"]).optional(),
        acessoRelOxUti: z.enum(["sim", "nao"]).optional(),
        acessoRelOxUtiFaturado: z.enum(["sim", "nao"]).optional(),
        acessoRelResumoPacienteOxUti: z.enum(["sim", "nao"]).optional(),
        acessoRelFisioterapia: z.enum(["sim", "nao"]).optional(),
        acessoRelComparativoMesAMes: z.enum(["sim", "nao"]).optional(),
        acessoRelRecebimentoGeral: z.enum(["sim", "nao"]).optional(),
        acessoRelFaturamento: z.enum(["sim", "nao"]).optional(),
        acessoRelAtendimentos: z.enum(["sim", "nao"]).optional(),
        acessoRelCustos: z.enum(["sim", "nao"]).optional(),
        acessoRelNaoRecebidos: z.enum(["sim", "nao"]).optional(),
        acessoRelPrevisaoGlosa: z.enum(["sim", "nao"]).optional(),
        acessoRelatoriosGlosasBi: z.enum(["sim", "nao"]).optional(),
        acessoFaturamentoExterno: z.enum(["sim", "nao"]).optional(),
        // Módulos Safatle
        acessoPainelExecutivo: z.enum(["sim", "nao"]).optional(),
        acessoVisaoGeral: z.enum(["sim", "nao"]).optional(),
        acessoFinanceiro: z.enum(["sim", "nao"]).optional(),
        acessoContratos: z.enum(["sim", "nao"]).optional(),
        acessoPropostas: z.enum(["sim", "nao"]).optional(),
        acessoRecursosHumanos: z.enum(["sim", "nao"]).optional(),
        acessoAtendimentosConsolidados: z.enum(["sim", "nao"]).optional(),
        acessoNfseConsolidado: z.enum(["sim", "nao"]).optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        const podeGerenciar = await db.verificarPermissaoEstabelecimento(
          ctx.user.id,
          input.estabelecimentoId,
          "gerenciar"
        );
        if (!podeGerenciar && ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message:
              "Você não tem permissão para gerenciar este estabelecimento",
          });
        }
        if (ctx.user.role !== "admin") {
          const alvo = await db.getUserById(input.userId);
          if (alvo?.role === "admin") {
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "Gestores não podem alterar administradores do sistema",
            });
          }
        }
        if (input.grupoId) {
          const grupo = await db.buscarGrupoServicoPorId(input.grupoId);
          if (
            !grupo ||
            grupo.ativo !== "sim" ||
            (grupo.estabelecimentoId &&
              grupo.estabelecimentoId !== input.estabelecimentoId)
          ) {
            throw new TRPCError({
              code: "BAD_REQUEST",
              message: "Grupo personalizado inválido para este estabelecimento",
            });
          }
        }

        const anterior = await db.getPermissaoUsuarioEstabelecimento(
          input.userId,
          input.estabelecimentoId
        );
        const normalizado = normalizeBiParent({
          ...input,
          acessoRecursosHumanos:
            input.acessoRecursosHumanos ??
            (anterior as any)?.acessoRecursosHumanos,
        });
        const resultado = await db.upsertPermissaoEstabelecimento(normalizado);
        const usuarioAfetado = await db.getUserById(input.userId);
        await db.registrarLogAuditoria({
          usuarioId: ctx.user.id,
          usuarioNome: ctx.user.name,
          usuarioAfetadoId: input.userId,
          usuarioAfetadoNome: usuarioAfetado?.name || "Usuário",
          estabelecimentoId: input.estabelecimentoId,
          tipoAcao: anterior ? "alterar_permissao" : "criar_permissao",
          descricao: anterior
            ? "Permissões de acesso atualizadas"
            : "Permissão de acesso criada",
          valoresAnteriores: anterior || undefined,
          valoresNovos: normalizado,
        });
        return resultado;
      }),

    // Obter módulos permitidos por grupo de serviço
    modulosPorGrupo: protectedProcedure
      .input(z.object({ grupoServico: z.string() }))
      .query(({ input }) => {
        return db.getModulosPermitidosPorGrupo(input.grupoServico);
      }),

    // Listar todos os usuários do sistema (para seleção)
    listarUsuarios: protectedProcedure.query(async ({ ctx }) => {
      if (ctx.user.role === "admin") {
        return db.listarTodosUsuarios();
      }
      const gerenciaveis = await db.getEstabelecimentosGerenciaveis(
        ctx.user.id
      );
      if (gerenciaveis.length === 0) {
        throw new TRPCError({
          code: "FORBIDDEN",
          message: "Apenas gestores podem listar usuários",
        });
      }
      return db.listarUsuariosPorEstabelecimentos(
        gerenciaveis.map(item => item.id)
      );
    }),

    // Listar permissões de todos os usuários para o Safatle
    listarPermissoesSafatle: protectedProcedure.query(async ({ ctx }) => {
      if (ctx.user.role !== "admin") {
        throw new TRPCError({
          code: "FORBIDDEN",
          message: "Apenas administradores podem ver permissões Safatle",
        });
      }
      const allUsers = await db.listarTodosUsuarios();
      const results: any[] = [];
      for (const user of allUsers) {
        const perms = await db.getPermissoesUsuario(user.id);
        const safatlePerm = perms.find(
          (p: any) => p.estabelecimentoId === 2160001
        );
        results.push({
          userId: user.id,
          userName: user.name,
          userEmail: user.email,
          userRole: user.role,
          permissions: safatlePerm || null,
        });
      }
      return results;
    }),

    // Remover permissão (apenas gestores/admin)
    removerPermissao: protectedProcedure
      .input(
        z.object({
        userId: z.number(),
        estabelecimentoId: z.number(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        const podeGerenciar = await db.verificarPermissaoEstabelecimento(
          ctx.user.id,
          input.estabelecimentoId,
          "gerenciar"
        );
        if (!podeGerenciar && ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message:
              "Você não tem permissão para gerenciar este estabelecimento",
          });
        }
        if (ctx.user.role !== "admin") {
          const alvo = await db.getUserById(input.userId);
          if (alvo?.role === "admin") {
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "Gestores não podem alterar administradores do sistema",
            });
          }
        }
        const anterior = await db.getPermissaoUsuarioEstabelecimento(
          input.userId,
          input.estabelecimentoId
        );
        const usuarioAfetado = await db.getUserById(input.userId);
        await db.deletePermissaoEstabelecimento(
          input.userId,
          input.estabelecimentoId
        );
        await db.registrarLogAuditoria({
          usuarioId: ctx.user.id,
          usuarioNome: ctx.user.name,
          usuarioAfetadoId: input.userId,
          usuarioAfetadoNome: usuarioAfetado?.name || "Usuário",
          estabelecimentoId: input.estabelecimentoId,
          tipoAcao: "remover_permissao",
          descricao: "Acesso ao estabelecimento removido",
          valoresAnteriores: anterior || undefined,
        });
        return { success: true };
      }),

    // Criar novo usuário (apenas admin/gestor)
    criarUsuario: protectedProcedure
      .input(
        z.object({
        name: z.string().min(2),
        email: z.string().email(),
        role: z.enum(["admin", "user", "tasy_user"]).optional(),
          grupoServico: z
            .enum([
              "administrador",
              "faturista",
              "recurso_glosa",
              "gestor",
              "visualizador",
              "usuario_tasy",
            ])
            .optional(),
        grupoId: z.number().optional(),
        estabelecimentosIds: z.array(z.number()).optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        if (ctx.user.role !== "admin") {
          const isGestor = await db.verificarSeGestor(ctx.user.id);
          if (!isGestor) {
            throw new TRPCError({
              code: "FORBIDDEN",
              message:
                "Apenas administradores ou gestores podem criar usuários",
            });
          }
          if (input.role === "admin") {
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "Gestores não podem criar administradores do sistema",
            });
          }
          const gerenciaveis = await db.getEstabelecimentosGerenciaveis(
            ctx.user.id
          );
          const idsGerenciaveis = new Set(gerenciaveis.map(item => item.id));
          if (
            (input.estabelecimentosIds || []).some(
              id => !idsGerenciaveis.has(id)
            )
          ) {
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "Você só pode atribuir estabelecimentos que administra",
            });
          }
        }

        if (input.grupoId) {
          const grupo = await db.buscarGrupoServicoPorId(input.grupoId);
          const estabelecimentosDestino = input.estabelecimentosIds || [];
          const grupoDisponivelEmTodos =
            grupo?.ativo === "sim" &&
            estabelecimentosDestino.every(
              estabelecimentoId =>
                grupo.estabelecimentoId === null ||
                grupo.estabelecimentoId === estabelecimentoId
          );
          if (!grupoDisponivelEmTodos) {
            throw new TRPCError({
              code: "BAD_REQUEST",
              message:
                "Grupo personalizado inválido, inativo ou indisponível para um dos estabelecimentos",
            });
          }
        }

        try {
          const userId = await db.criarUsuario(input);

          // Criar permissões para múltiplos estabelecimentos
          if (
            input.estabelecimentosIds &&
            input.estabelecimentosIds.length > 0
          ) {
            const perfil: StandardServiceGroup =
              input.role === "tasy_user"
              ? "usuario_tasy"
                : input.grupoServico || "visualizador";
            const acoes = ACTION_PROFILE_DEFAULTS[perfil];
            for (const estabelecimentoId of input.estabelecimentosIds) {
              await db.criarPermissaoEstabelecimento({
                usuarioId: userId,
                estabelecimentoId,
                grupoId: input.grupoId,
                grupoServico: perfil,
                podeVisualizar: acoes.podeVisualizar === "sim",
                podeEditar: acoes.podeEditar === "sim",
                podeExcluir: acoes.podeExcluir === "sim",
                podeGerenciar: acoes.podeGerenciar === "sim",
              });
            }
          }

          // Registrar no log de auditoria
          const estabelecimentosNomes =
            input.estabelecimentosIds && input.estabelecimentosIds.length > 0
            ? ` com acesso a ${input.estabelecimentosIds.length} estabelecimento(s)`
            : "";
          await db.registrarLogAuditoria({
            usuarioId: ctx.user.id,
            usuarioNome: ctx.user.name,
            usuarioAfetadoId: userId,
            usuarioAfetadoNome: input.name,
            tipoAcao: "criar_usuario",
            descricao: `Usuário ${input.name} (${input.email}) criado${estabelecimentosNomes}`,
            valoresNovos: {
              name: input.name,
              email: input.email,
              role: input.role || "user",
              estabelecimentosIds: input.estabelecimentosIds,
              grupoId: input.grupoId,
            },
          });

          return { success: true, userId };
        } catch (error: any) {
          throw new TRPCError({
            code: "CONFLICT",
            message: error.message || "Erro ao criar usuário",
          });
        }
      }),

    // Listar grupos de serviço
    listarGrupos: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number().optional() }))
      .query(async ({ ctx, input }) => {
        if (ctx.user.role !== "admin") {
          if (
            !input.estabelecimentoId ||
            !(await db.verificarPermissaoEstabelecimento(
              ctx.user.id,
              input.estabelecimentoId,
              "gerenciar"
            ))
          ) {
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "Selecione uma unidade que você administra",
            });
          }
        }
        return db.listarGruposServico(input.estabelecimentoId);
      }),

    // Criar grupo de serviço personalizado
    criarGrupo: protectedProcedure
      .input(
        z.object({
        nome: z.string().min(2),
        descricao: z.string().optional(),
        cor: z.string().optional(),
        icone: z.string().optional(),
          permissoesPadrao: z
            .record(z.string(), z.enum(["sim", "nao"]))
            .optional(),
        estabelecimentoId: z.number().optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        if (ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Apenas administradores podem criar grupos",
          });
        }

        const grupoId = await db.criarGrupoServico({
          ...input,
          permissoesPadrao: sanitizeCustomModulePermissions(
            input.permissoesPadrao
          ),
          criadoPor: ctx.user.id,
        });

        // Registrar no log de auditoria
        await db.registrarLogAuditoria({
          usuarioId: ctx.user.id,
          usuarioNome: ctx.user.name,
          usuarioAfetadoId: ctx.user.id,
          usuarioAfetadoNome: ctx.user.name,
          estabelecimentoId: input.estabelecimentoId,
          tipoAcao: "criar_grupo",
          descricao: `Grupo de serviço "${input.nome}" criado`,
          valoresNovos: input,
        });

        return { success: true, grupoId };
      }),

    // Atualizar grupo de serviço
    atualizarGrupo: protectedProcedure
      .input(
        z.object({
        id: z.number(),
        nome: z.string().min(2).optional(),
        descricao: z.string().optional(),
        cor: z.string().optional(),
        icone: z.string().optional(),
          permissoesPadrao: z
            .record(z.string(), z.enum(["sim", "nao"]))
            .optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        if (ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Apenas administradores podem atualizar grupos",
          });
        }

        const grupoAnterior = await db.buscarGrupoServicoPorId(input.id);
        const { id, ...dadosInput } = input;
        const dados = {
          ...dadosInput,
          ...(dadosInput.permissoesPadrao
            ? {
                permissoesPadrao: sanitizeCustomModulePermissions(
                  dadosInput.permissoesPadrao
                ),
              }
            : {}),
        };
        await db.atualizarGrupoServico(id, dados);

        // Registrar no log de auditoria
        await db.registrarLogAuditoria({
          usuarioId: ctx.user.id,
          usuarioNome: ctx.user.name,
          usuarioAfetadoId: ctx.user.id,
          usuarioAfetadoNome: ctx.user.name,
          tipoAcao: "alterar_grupo",
          descricao: `Grupo de serviço "${grupoAnterior?.nome}" atualizado`,
          valoresAnteriores: grupoAnterior,
          valoresNovos: dados,
        });

        return { success: true };
      }),

    // Excluir grupo de serviço
    excluirGrupo: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ ctx, input }) => {
        if (ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Apenas administradores podem excluir grupos",
          });
        }

        const grupo = await db.buscarGrupoServicoPorId(input.id);
        await db.excluirGrupoServico(input.id);

        // Registrar no log de auditoria
        await db.registrarLogAuditoria({
          usuarioId: ctx.user.id,
          usuarioNome: ctx.user.name,
          usuarioAfetadoId: ctx.user.id,
          usuarioAfetadoNome: ctx.user.name,
          tipoAcao: "excluir_grupo",
          descricao: `Grupo de serviço "${grupo?.nome}" excluído`,
          valoresAnteriores: grupo,
        });

        return { success: true };
      }),

    // Listar logs de auditoria
    logsAuditoria: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number().optional(),
        usuarioId: z.number().optional(),
        usuarioAfetadoId: z.number().optional(),
        tipoAcao: z.string().optional(),
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
        limite: z.number().optional(),
        offset: z.number().optional(),
        })
      )
      .query(async ({ ctx, input }) => {
        if (ctx.user.role !== "admin") {
          if (
            !input.estabelecimentoId ||
            !(await db.verificarPermissaoEstabelecimento(
              ctx.user.id,
              input.estabelecimentoId,
              "gerenciar"
            ))
          ) {
            throw new TRPCError({
              code: "FORBIDDEN",
              message:
                "Você só pode consultar a auditoria das unidades que administra",
            });
          }
        }

        return db.listarLogsAuditoria({
          ...input,
          dataInicio: input.dataInicio ? new Date(input.dataInicio) : undefined,
          dataFim: input.dataFim ? new Date(input.dataFim) : undefined,
        });
      }),

    // Listar estabelecimentos de um usuário específico
    listarEstabelecimentosUsuario: protectedProcedure
      .input(z.object({ userId: z.number() }))
      .query(async ({ ctx, input }) => {
        const estabelecimentosUsuario = await db.getEstabelecimentosUsuario(
          input.userId
        );
        if (ctx.user.role !== "admin") {
          const gerenciaveis = await db.getEstabelecimentosGerenciaveis(
            ctx.user.id
          );
          const idsGerenciaveis = new Set(gerenciaveis.map(item => item.id));
          if (idsGerenciaveis.size === 0)
            throw new TRPCError({ code: "FORBIDDEN" });
          return estabelecimentosUsuario.filter(item =>
            idsGerenciaveis.has(item.id)
          );
        }
        return estabelecimentosUsuario;
      }),

    // Atualizar estabelecimentos de um usuário (adicionar/remover)
    atualizarEstabelecimentosUsuario: protectedProcedure
      .input(
        z.object({
        userId: z.number(),
        estabelecimentosIds: z.array(z.number()),
        grupoId: z.number().optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        let idsGerenciaveis: Set<number> | null = null;
        if (ctx.user.role !== "admin") {
          const alvo = await db.getUserById(input.userId);
          if (alvo?.role === "admin")
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "Gestores não podem alterar administradores",
            });
          const gerenciaveis = await db.getEstabelecimentosGerenciaveis(
            ctx.user.id
          );
          idsGerenciaveis = new Set(gerenciaveis.map(item => item.id));
          if (input.estabelecimentosIds.some(id => !idsGerenciaveis!.has(id))) {
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "Você só pode alterar unidades que administra",
            });
          }
        }

        // Buscar estabelecimentos atuais do usuário
        const estabelecimentosAtuais = await db.getEstabelecimentosUsuario(
          input.userId
        );
        const idsAtuais = estabelecimentosAtuais.map(e => e.id);
        const idsAtuaisEditaveis = idsGerenciaveis
          ? idsAtuais.filter(id => idsGerenciaveis!.has(id))
          : idsAtuais;

        // Identificar o que adicionar e remover
        const idsParaAdicionar = input.estabelecimentosIds.filter(
          id => !idsAtuais.includes(id)
        );
        const idsParaRemover = idsAtuaisEditaveis.filter(
          id => !input.estabelecimentosIds.includes(id)
        );

        // Remover permissões
        for (const estabelecimentoId of idsParaRemover) {
          await db.deletePermissaoEstabelecimento(
            input.userId,
            estabelecimentoId
          );
        }

        // Adicionar permissões
        for (const estabelecimentoId of idsParaAdicionar) {
          await db.criarPermissaoEstabelecimento({
            usuarioId: input.userId,
            estabelecimentoId,
            grupoId: input.grupoId,
            podeVisualizar: true,
            podeEditar: false,
            podeExcluir: false,
            podeGerenciar: false,
          });
        }

        // Buscar informações do usuário para o log
        const usuario = await db.getUserById(input.userId);

        // Registrar no log de auditoria
        await db.registrarLogAuditoria({
          usuarioId: ctx.user.id,
          usuarioNome: ctx.user.name,
          usuarioAfetadoId: input.userId,
          usuarioAfetadoNome: usuario?.name || "Usuário",
          tipoAcao: "editar_estabelecimentos",
          descricao: `Estabelecimentos atualizados: ${idsParaAdicionar.length} adicionado(s), ${idsParaRemover.length} removido(s)`,
          valoresAnteriores: { estabelecimentosIds: idsAtuais },
          valoresNovos: { estabelecimentosIds: input.estabelecimentosIds },
        });

        return {
          success: true,
          adicionados: idsParaAdicionar.length,
          removidos: idsParaRemover.length,
        };
      }),

    // Conceder acesso a todos os estabelecimentos (apenas admin)
    concederAcessoTotal: protectedProcedure
      .input(
        z.object({
        userId: z.number(),
          permissoes: z
            .object({
          podeVisualizar: z.enum(["sim", "nao"]).optional(),
          podeEditar: z.enum(["sim", "nao"]).optional(),
          podeExcluir: z.enum(["sim", "nao"]).optional(),
          podeGerenciar: z.enum(["sim", "nao"]).optional(),
            })
            .optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        if (ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Apenas administradores podem conceder acesso total",
          });
        }
        await db.concederAcessoTodosEstabelecimentos(
          input.userId,
          input.permissoes
        );
        return { success: true };
      }),

    // Excluir usuário do sistema
    excluirUsuario: protectedProcedure
      .input(z.object({ userId: z.number() }))
      .mutation(async ({ ctx, input }) => {
        if (ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Apenas administradores podem excluir usuários",
          });
        }
        const result = await db.deleteUsuario(input.userId, ctx.user.id);
        if (!result.success) {
          throw new TRPCError({ code: "BAD_REQUEST", message: result.message });
        }
        return result;
      }),
  }),

  // ============ MÉTRICAS DE PRODUTIVIDADE ============
  produtividade: router({
    // Buscar métricas de produtividade de classificação de glosas
    metricas: protectedProcedure
      .input(
        z
          .object({
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
          })
          .optional()
      )
      .query(async ({ ctx, input }) => {
        return db.getMetricasProdutividade({
          userId: ctx.user.id,
          dataInicio: input?.dataInicio
            ? new Date(input.dataInicio)
            : undefined,
          dataFim: input?.dataFim ? new Date(input.dataFim) : undefined,
        });
      }),

    // Buscar métricas de envio de XML
    metricasEnvioXML: protectedProcedure
      .input(
        z
          .object({
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
        estabelecimentoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getMetricasEnvioXML({
          dataInicio: input?.dataInicio
            ? new Date(input.dataInicio)
            : undefined,
          dataFim: input?.dataFim ? new Date(input.dataFim) : undefined,
          estabelecimentoId: input?.estabelecimentoId,
        });
      }),

    // Buscar métricas de importação via banco (integrador de dados)
    metricasImportacaoBanco: protectedProcedure
      .input(
        z
          .object({
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
        estabelecimentoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getMetricasImportacaoBanco({
          dataInicio: input?.dataInicio
            ? new Date(input.dataInicio)
            : undefined,
          dataFim: input?.dataFim ? new Date(input.dataFim) : undefined,
          estabelecimentoId: input?.estabelecimentoId,
        });
      }),
  }),

  // ============ DASHBOARD CONSOLIDADO ============
  dashboardConsolidado: router({
    // Buscar dados consolidados de todos os estabelecimentos
    dados: protectedProcedure.query(async ({ ctx }) => {
      const dados = await db.getDadosConsolidados(ctx.user.id);
      if (!dados) {
        throw new TRPCError({
          code: "FORBIDDEN",
          message:
            "Você não tem permissão para visualizar o dashboard consolidado",
        });
      }
      return dados;
    }),

    // Buscar comparativo de glosas entre estabelecimentos
    comparativoGlosas: protectedProcedure
      .input(
        z
          .object({
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
          })
          .optional()
      )
      .query(async ({ ctx, input }) => {
        let periodo: { inicio: Date; fim: Date } | undefined;
        if (input?.dataInicio && input?.dataFim) {
          periodo = {
            inicio: new Date(input.dataInicio),
            fim: new Date(input.dataFim),
          };
        }
        const comparativo = await db.getComparativoGlosasEstabelecimentos(
          ctx.user.id,
          periodo
        );
        if (!comparativo) {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Você não tem permissão para visualizar o comparativo",
          });
        }
        return comparativo;
      }),

    // Buscar dados de recursos de glosa consolidados de todos os estabelecimentos
    recursadoConsolidado: protectedProcedure
      .input(
        z
          .object({
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
          })
          .optional()
      )
      .query(async ({ ctx }) => {
        const database = await getDb();
        if (!database)
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: "Banco de dados não disponível",
          });

        const { recursosGlosa, lotesRecurso, convenios, estabelecimentos } =
          await import("../drizzle/schema");
        const { sql, eq, and, gte, lte, inArray, isNotNull, count, sum } =
          await import("drizzle-orm");

        // Buscar todos os estabelecimentos que o usuário tem acesso
        const estabList = await db.getEstabelecimentos();
        const estabMap = new Map(estabList.map(e => [e.id, e.nome]));
        const estabIds = estabList.map(e => e.id);

        // KPIs gerais de recursos de glosa
        const [kpisResult] = await database
          .select({
            totalRecursos: count(),
            valorTotalGlosado: sum(recursosGlosa.valorGlosado),
            valorTotalRecursado: sum(recursosGlosa.valorCobrado),
            valorTotalRecuperado: sum(recursosGlosa.valorRecuperado),
            valorTotalRecebido: sum(recursosGlosa.valorRecebido),
          })
          .from(recursosGlosa)
          .where(
            estabIds.length > 0
              ? inArray(recursosGlosa.estabelecimentoId, estabIds)
              : sql`1=1`
          );

        // KPIs por status
        const statusCounts = await database
          .select({
            status: recursosGlosa.status,
            quantidade: count(),
            valorGlosado: sum(recursosGlosa.valorGlosado),
            valorRecuperado: sum(recursosGlosa.valorRecuperado),
          })
          .from(recursosGlosa)
          .where(
            estabIds.length > 0
              ? inArray(recursosGlosa.estabelecimentoId, estabIds)
              : sql`1=1`
          )
          .groupBy(recursosGlosa.status);

        // Agrupar por estabelecimento
        const porEstab = await database
          .select({
            estabelecimentoId: recursosGlosa.estabelecimentoId,
            totalRecursos: count(),
            valorGlosado: sum(recursosGlosa.valorGlosado),
            valorRecuperado: sum(recursosGlosa.valorRecuperado),
            valorRecebido: sum(recursosGlosa.valorRecebido),
          })
          .from(recursosGlosa)
          .where(
            and(
              estabIds.length > 0
                ? inArray(recursosGlosa.estabelecimentoId, estabIds)
                : sql`1=1`,
            isNotNull(recursosGlosa.estabelecimentoId)
            )
          )
          .groupBy(recursosGlosa.estabelecimentoId);

        // Agrupar por convênio (top 10)
        const porConvenio = await database
          .select({
            convenioId: recursosGlosa.convenioId,
            totalRecursos: count(),
            valorGlosado: sum(recursosGlosa.valorGlosado),
            valorRecuperado: sum(recursosGlosa.valorRecuperado),
          })
          .from(recursosGlosa)
          .where(
            estabIds.length > 0
              ? inArray(recursosGlosa.estabelecimentoId, estabIds)
              : sql`1=1`
          )
          .groupBy(recursosGlosa.convenioId)
          .orderBy(sql`SUM(${recursosGlosa.valorGlosado}) DESC`)
          .limit(10);

        // Buscar nomes dos convênios
        const convIds = porConvenio.map(c => c.convenioId).filter(Boolean);
        let convMap = new Map<number, string>();
        if (convIds.length > 0) {
          const convList = await database
            .select({ id: convenios.id, nome: convenios.nome })
            .from(convenios)
            .where(inArray(convenios.id, convIds));
          convMap = new Map(convList.map(c => [c.id, c.nome]));
        }

        // Lotes de recurso - resumo
        const [lotesKpis] = await database
          .select({
            totalLotes: count(),
            valorTotalGlosado: sum(lotesRecurso.valorTotalGlosado),
            valorTotalRecursado: sum(lotesRecurso.valorTotalRecursado),
            valorTotalRecuperado: sum(lotesRecurso.valorTotalRecuperado),
            valorTotalRecebido: sum(lotesRecurso.valorTotalRecebido),
          })
          .from(lotesRecurso)
          .where(
            estabIds.length > 0
              ? inArray(lotesRecurso.estabelecimentoId, estabIds)
              : sql`1=1`
          );

        const valorGlosado = Number(kpisResult?.valorTotalGlosado || 0);
        const valorRecuperado = Number(kpisResult?.valorTotalRecuperado || 0);
        const valorRecebido = Number(kpisResult?.valorTotalRecebido || 0);
        const taxaRecuperacao =
          valorGlosado > 0 ? (valorRecuperado / valorGlosado) * 100 : 0;

        return {
          kpis: {
            totalRecursos: Number(kpisResult?.totalRecursos || 0),
            valorTotalGlosado: valorGlosado,
            valorTotalRecursado: Number(kpisResult?.valorTotalRecursado || 0),
            valorTotalRecuperado: valorRecuperado,
            valorTotalRecebido: valorRecebido,
            taxaRecuperacao: Math.round(taxaRecuperacao * 100) / 100,
          },
          porStatus: statusCounts.map(s => ({
            status: s.status,
            quantidade: Number(s.quantidade),
            valorGlosado: Number(s.valorGlosado || 0),
            valorRecuperado: Number(s.valorRecuperado || 0),
          })),
          porEstabelecimento: porEstab.map(e => ({
            id: e.estabelecimentoId,
            nome:
              estabMap.get(e.estabelecimentoId!) ||
              `Estab. ${e.estabelecimentoId}`,
            totalRecursos: Number(e.totalRecursos),
            valorGlosado: Number(e.valorGlosado || 0),
            valorRecuperado: Number(e.valorRecuperado || 0),
            valorRecebido: Number(e.valorRecebido || 0),
            taxaRecuperacao:
              Number(e.valorGlosado || 0) > 0
                ? Math.round(
                    (Number(e.valorRecuperado || 0) /
                      Number(e.valorGlosado || 0)) *
                      10000
                  ) / 100
                : 0,
          })),
          porConvenio: porConvenio.map(c => ({
            convenioId: c.convenioId,
            convenioNome:
              convMap.get(c.convenioId) || `Convênio ${c.convenioId}`,
            totalRecursos: Number(c.totalRecursos),
            valorGlosado: Number(c.valorGlosado || 0),
            valorRecuperado: Number(c.valorRecuperado || 0),
          })),
          lotes: {
            totalLotes: Number(lotesKpis?.totalLotes || 0),
            valorTotalGlosado: Number(lotesKpis?.valorTotalGlosado || 0),
            valorTotalRecursado: Number(lotesKpis?.valorTotalRecursado || 0),
            valorTotalRecuperado: Number(lotesKpis?.valorTotalRecuperado || 0),
            valorTotalRecebido: Number(lotesKpis?.valorTotalRecebido || 0),
          },
        };
      }),

    // Buscar atendimentos consolidados de todos os estabelecimentos
    atendimentosConsolidados: protectedProcedure.query(async ({ ctx }) => {
      if (ctx.user.role !== "admin") {
        throw new TRPCError({
          code: "FORBIDDEN",
          message: "Visão consolidada disponível somente para administradores.",
        });
      }
      const { getAtendimentosParadosUnificados, calcularDiasParadoUnificado } =
        await import("./atendimentosUnificados");
      const estabelecimentosList = await db.getEstabelecimentos();
      const estabMap = new Map(estabelecimentosList.map(e => [e.id, e.nome]));

      const todos = await getAtendimentosParadosUnificados();

      // KPIs gerais
      const total = todos.length;
      const totalDias = todos.reduce(
        (sum, a) =>
          sum +
          calcularDiasParadoUnificado(
            a.data_entrada,
            a.data_saida || undefined,
            a.origemSistema,
            a.dtEntrega,
            a.dtEtapa
          ),
        0
      );
      const mediaDias = total > 0 ? Math.round(totalDias / total) : 0;
      const valorTotal = todos.reduce(
        (sum, a) => sum + Number(a.valorConta || 0),
        0
      );

      // Agrupar por estabelecimento
      const porEstab: Record<
        number,
        {
          nome: string;
          total: number;
          valorTotal: number;
          semConta: number;
          aFaturar: number;
          internacao: number;
          ambulatorio: number;
          exame: number;
          outros: number;
        }
      > = {};
      todos.forEach(a => {
        const eid = a.estabelecimentoId;
        if (!porEstab[eid]) {
          porEstab[eid] = {
            nome: estabMap.get(eid) || `Estab. ${eid}`,
            total: 0,
            valorTotal: 0,
            semConta: 0,
            aFaturar: 0,
            internacao: 0,
            ambulatorio: 0,
            exame: 0,
            outros: 0,
          };
        }
        const e = porEstab[eid];
        e.total++;
        e.valorTotal += Number(a.valorConta || 0);
        const desc = (a.descricao_atendimento || "").toLowerCase();
        if (desc.includes("sem_conta") || desc.includes("sem conta"))
          e.semConta++;
        else e.aFaturar++;
        const tipo = (a.tipo_atendimento || "").toLowerCase();
        if (tipo.includes("internac")) e.internacao++;
        else if (tipo.includes("ambulat")) e.ambulatorio++;
        else if (tipo.includes("exame")) e.exame++;
        else e.outros++;
      });

      // Agrupar por convênio (top 15)
      const porConvenio: Record<string, number> = {};
      todos.forEach(a => {
        const conv = a.convenio || "Sem Convênio";
        porConvenio[conv] = (porConvenio[conv] || 0) + 1;
      });
      const topConvenios = Object.entries(porConvenio)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15)
        .map(([convenio, quantidade]) => ({ convenio, quantidade }));

      // Agrupar por sistema de origem
      const porOrigem: Record<string, number> = {};
      todos.forEach(a => {
        const orig = a.origemSistema || "Desconhecido";
        porOrigem[orig] = (porOrigem[orig] || 0) + 1;
      });

      return {
        kpis: { total, mediaDias, valorTotal },
        porEstabelecimento: Object.entries(porEstab).map(([id, data]) => ({
          id: Number(id),
          ...data,
        })),
        topConvenios,
        porOrigem: Object.entries(porOrigem).map(([origem, quantidade]) => ({
          origem,
          quantidade,
        })),
      };
    }),
  }),

  // ============ MOTIVOS DE GLOSA ============
  motivosGlosa: router({
    // Listar motivos de glosa
    list: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        tipoOrigem: z.enum(["tiss", "personalizado"]).optional(),
        ativo: z.enum(["sim", "nao"]).optional(),
        search: z.string().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getMotivosGlosa(input || {});
      }),

    // Buscar motivo por código (combinado banco + TISS)
    porCodigo: protectedProcedure
      .input(
        z.object({
        codigo: z.string(),
        estabelecimentoId: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getMotivoGlosaCombinado(
          input.codigo,
          input.estabelecimentoId
        );
      }),

    // Listar grupos disponíveis
    grupos: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getGruposMotivosGlosa(input?.estabelecimentoId);
      }),

    // Criar novo motivo de glosa
    criar: protectedProcedure
      .input(
        z.object({
        codigo: z.string().min(1, "Código é obrigatório"),
        grupo: z.string().min(1, "Grupo é obrigatório"),
        descricao: z.string().min(1, "Descrição é obrigatória"),
          descricaoSimplificada: z
            .string()
            .min(1, "Descrição simplificada é obrigatória"),
        argumentoContestacao: z.string().optional(),
        acoesRecomendadas: z.array(z.string()).optional(),
        documentosSugeridos: z.array(z.string()).optional(),
        dificuldadeReversao: z.number().min(1).max(5).optional(),
        probabilidadeSucesso: z.number().min(0).max(100).optional(),
        estabelecimentoId: z.number().optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        // Verificar se já existe um motivo com esse código
        const existente = await db.getMotivoGlosaPorCodigo(
          input.codigo,
          input.estabelecimentoId
        );
        if (existente) {
          throw new TRPCError({
            code: "CONFLICT",
            message: `Já existe um motivo de glosa com o código ${input.codigo}`,
          });
        }

        await db.criarMotivoGlosa({
          ...input,
          tipoOrigem: "personalizado",
          criadoPor: ctx.user.id,
        });

        return { success: true };
      }),

    // Atualizar motivo de glosa
    atualizar: protectedProcedure
      .input(
        z.object({
        id: z.number(),
        codigo: z.string().optional(),
        grupo: z.string().optional(),
        descricao: z.string().optional(),
        descricaoSimplificada: z.string().optional(),
        argumentoContestacao: z.string().optional(),
        acoesRecomendadas: z.array(z.string()).optional(),
        documentosSugeridos: z.array(z.string()).optional(),
        dificuldadeReversao: z.number().min(1).max(5).optional(),
        probabilidadeSucesso: z.number().min(0).max(100).optional(),
        ativo: z.enum(["sim", "nao"]).optional(),
        })
      )
      .mutation(async ({ input }) => {
        const { id, ...data } = input;
        await db.atualizarMotivoGlosa(id, data);
        return { success: true };
      }),

    // Excluir motivo de glosa
    excluir: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input }) => {
        await db.excluirMotivoGlosa(input.id);
        return { success: true };
      }),

    // Popular tabela com códigos TISS oficiais
    seedTiss: protectedProcedure.mutation(async () => {
        const { seedMotivosGlosa } = await import("./seedMotivosGlosa");
        return seedMotivosGlosa();
      }),
  }),

  // ============ ANÁLISE INTELIGENTE DE CONTAS (IA) ============
  ia: router({
    // Obter estatísticas por código de procedimento
    estatisticasPorCodigo: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        convenioId: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getEstatisticasPorCodigo(
          input.estabelecimentoId,
          input.convenioId
        );
      }),

    // Identificar contas com valores fora da média (outliers)
    contasOutliers: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        convenioId: z.number().optional(),
        limiteDesvio: z.number().default(2),
        })
      )
      .query(async ({ input }) => {
        return db.getContasOutliers(
          input.estabelecimentoId,
          input.convenioId,
          input.limiteDesvio
        );
      }),

    // Detectar padrões de erro por funcionário
    padroesErroPorFuncionario: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        })
      )
      .query(async ({ input }) => {
        return db.getPadroesErroPorFuncionario(input.estabelecimentoId);
      }),

    // Buscar motivos de glosa por funcionário
    motivosGlosaPorFuncionario: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        userId: z.number(),
        })
      )
      .query(async ({ input }) => {
        return db.getMotivosGlosaPorFuncionario(
          input.estabelecimentoId,
          input.userId
        );
      }),

    // Calcular risco de glosa para contas
    riscoGlosa: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        arquivoId: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.calcularRiscoGlosa(input.estabelecimentoId, input.arquivoId);
      }),

    // Gerar alertas e recomendações da IA
    alertas: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        })
      )
      .query(async ({ input }) => {
        return db.gerarAlertasIA(input.estabelecimentoId);
      }),
  }),

  // ============ REGRAS DE IA ============
  regrasIA: router({
    // Listar todas as regras de IA
    list: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getRegrasIA(input?.estabelecimentoId);
      }),

    // Buscar regra por ID
    get: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ input }) => {
        return db.getRegraIAById(input.id);
      }),

    // Buscar regra por código
    getPorCodigo: protectedProcedure
      .input(
        z.object({
        codigo: z.string(),
        estabelecimentoId: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getRegraIAPorCodigo(input.codigo, input.estabelecimentoId);
      }),

    // Criar nova regra
    create: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number().optional(),
        codigo: z.string().min(1),
        nome: z.string().min(1),
        descricao: z.string().optional(),
          categoria: z.enum([
            "outlier",
            "padrao_erro",
            "risco_glosa",
            "tendencia",
            "comparacao",
          ]),
          tipoAlerta: z.enum(["critico", "alerta", "info"]).optional(),
          parametros: z
            .object({
          limiteDesvioAbaixo: z.number().optional(),
          limiteDesvioAcima: z.number().optional(),
          minimoContasHistorico: z.number().optional(),
          periodoAnalise: z.number().optional(),
          taxaGlosaMinima: z.number().optional(),
          minimoProcedimentos: z.number().optional(),
          periodoMeses: z.number().optional(),
          scoreRiscoMinimo: z.number().optional(),
          historicoMinimoContas: z.number().optional(),
          variacaoMinima: z.number().optional(),
          periodoComparacao: z.number().optional(),
          maxResultados: z.number().optional(),
            })
            .optional(),
        prioridade: z.number().optional(),
          ativo: z.enum(["sim", "nao"]).optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        return db.createRegraIA({
          ...input,
          criadoPor: ctx.user.id,
        });
      }),

    // Atualizar regra existente
    update: protectedProcedure
      .input(
        z.object({
        id: z.number(),
        nome: z.string().optional(),
        descricao: z.string().optional(),
          tipoAlerta: z.enum(["critico", "alerta", "info"]).optional(),
          parametros: z
            .object({
          limiteDesvioAbaixo: z.number().optional(),
          limiteDesvioAcima: z.number().optional(),
          minimoContasHistorico: z.number().optional(),
          periodoAnalise: z.number().optional(),
          taxaGlosaMinima: z.number().optional(),
          minimoProcedimentos: z.number().optional(),
          periodoMeses: z.number().optional(),
          scoreRiscoMinimo: z.number().optional(),
          historicoMinimoContas: z.number().optional(),
          variacaoMinima: z.number().optional(),
          periodoComparacao: z.number().optional(),
          maxResultados: z.number().optional(),
            })
            .optional(),
        prioridade: z.number().optional(),
          ativo: z.enum(["sim", "nao"]).optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const { id, ...data } = input;
        return db.updateRegraIA(id, {
          ...data,
          atualizadoPor: ctx.user.id,
        });
      }),

    // Excluir regra
    delete: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input }) => {
        return db.deleteRegraIA(input.id);
      }),

    // Ativar/desativar regra
    toggle: protectedProcedure
      .input(
        z.object({
        id: z.number(),
          ativo: z.enum(["sim", "nao"]),
        })
      )
      .mutation(async ({ input }) => {
        return db.toggleRegraIA(input.id, input.ativo);
      }),

    // Restaurar regra para valores padrão
    restaurarPadrao: protectedProcedure
      .input(
        z.object({
        codigo: z.string(),
        estabelecimentoId: z.number().optional(),
        })
      )
      .mutation(async ({ input }) => {
        return db.restaurarRegraPadrao(input.codigo, input.estabelecimentoId);
      }),
  }),

  // ============ IMPORTAÇÃO DE DADOS DO TASY ============
  importacaoTasy: router({
    // Criar nova importação (simples, sem upload)
    criar: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        nomeArquivo: z.string(),
        tamanhoArquivo: z.number().optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const importacao = await db.createImportacaoTasy({
          estabelecimentoId: input.estabelecimentoId,
          userId: ctx.user.id,
          nomeArquivo: input.nomeArquivo,
          tamanhoArquivo: input.tamanhoArquivo,
          status: "aguardando",
          progresso: 0,
        });
        return { id: importacao.id };
      }),

    // Criar nova importação (upload de arquivo SQLite)
    upload: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        nomeArquivo: z.string(),
        tamanhoArquivo: z.number().optional(),
        conteudoBase64: z.string(), // Arquivo SQLite em base64
        })
      )
      .mutation(async ({ input, ctx }) => {
        // Cria o registro de importação
        const importacao = await db.createImportacaoTasy({
          estabelecimentoId: input.estabelecimentoId,
          userId: ctx.user.id,
          nomeArquivo: input.nomeArquivo,
          tamanhoArquivo: input.tamanhoArquivo,
          status: "aguardando",
          progresso: 0,
        });

        // Salva o arquivo no S3 para processamento posterior
        const s3Key = `tasy-imports/${input.estabelecimentoId}/${importacao.id}/${sanitizeFilename(input.nomeArquivo)}`;
        const fileBuffer = Buffer.from(input.conteudoBase64, "base64");
        const { url } = await storagePut(
          s3Key,
          fileBuffer,
          "application/x-sqlite3"
        );

        // Atualiza com a URL do S3
        await db.updateImportacaoTasy(importacao.id, {
          s3Key,
          s3Url: url,
        });

        return { id: importacao.id, s3Key, s3Url: url };
      }),

    // Processar importação (lê o SQLite e insere os dados)
    processar: protectedProcedure
      .input(
        z.object({
        importacaoId: z.number(),
          dados: z.array(
            z.object({
          atendimento: z.string(),
          nrInternoConta: z.string().optional(),
          sequencia: z.string().optional(),
          dataFaturado: z.string().optional(),
          guia: z.string().optional(),
          convenio: z.string().optional(),
          paciente: z.string().optional(),
          dataConta: z.string().optional(),
          codigo: z.string().optional(),
          codigoConvenio: z.string().optional(),
          descricao: z.string().optional(),
          quantidade: z.number().optional(),
          unidade: z.string().optional(),
          valorUnitario: z.number().optional(),
          valorTotal: z.number().optional(),
          setor: z.string().optional(),
          protocolo: z.string().optional(),
          statusProtocolo: z.string().optional(),
              tipo: z.enum(["MATERIAL", "HONORARIO"]),
          medico: z.string().optional(),
          funcaoMedico: z.string().optional(),
          crm: z.string().optional(),
          valorMedico: z.number().optional(),
            })
          ),
        })
      )
      .mutation(async ({ input }) => {
        const importacao = await db.getImportacaoTasyById(input.importacaoId);
        if (!importacao) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Importação não encontrada",
          });
        }

        // Atualiza status para processando
        await db.updateImportacaoTasy(input.importacaoId, {
          status: "processando",
          totalRegistros: input.dados.length,
        });

        // Prepara os registros para inserção
        const registros = input.dados.map(d => ({
          estabelecimentoId: importacao.estabelecimentoId,
          importacaoId: input.importacaoId,
          atendimento: d.atendimento,
          nrInternoConta: d.nrInternoConta || null,
          sequencia: d.sequencia || null,
          dataFaturado: d.dataFaturado ? new Date(d.dataFaturado) : null,
          guia: d.guia || null,
          convenio: d.convenio || null,
          paciente: d.paciente || null,
          dataConta: d.dataConta ? new Date(d.dataConta) : null,
          codigo: d.codigo || null,
          codigoConvenio: d.codigoConvenio || null,
          descricao: d.descricao || null,
          quantidade: d.quantidade?.toString() || null,
          unidade: d.unidade || null,
          valorUnitario: d.valorUnitario?.toString() || null,
          valorTotal: d.valorTotal?.toString() || null,
          setor: d.setor || null,
          protocolo: d.protocolo || null,
          statusProtocolo: d.statusProtocolo || null,
          tipo: d.tipo,
          medico: d.medico || null,
          funcaoMedico: d.funcaoMedico || null,
          crm: d.crm || null,
          valorMedico: d.valorMedico?.toString() || null,
        }));

        // Insere em lotes
        const resultado = await db.insertDadosTasyBatch(
          registros as any,
          importacao.estabelecimentoId
        );

        // Conta por tipo
        const totalMateriais = input.dados.filter(
          d => d.tipo === "MATERIAL"
        ).length;
        const totalHonorarios = input.dados.filter(
          d => d.tipo === "HONORARIO"
        ).length;

        // Calcula datas
        const datas = input.dados
          .filter(d => d.dataFaturado)
          .map(d => new Date(d.dataFaturado!));
        const dataInicio =
          datas.length > 0
            ? new Date(Math.min(...datas.map(d => d.getTime())))
            : null;
        const dataFim =
          datas.length > 0
            ? new Date(Math.max(...datas.map(d => d.getTime())))
            : null;

        // Atualiza estatísticas
        const status = resultado.erros > 0 ? "concluido_parcial" : "concluido";
        await db.updateImportacaoTasy(input.importacaoId, {
          status,
          progresso: 100,
          registrosImportados: resultado.inseridos,
          registrosIgnorados: resultado.ignorados,
          registrosErro: resultado.erros,
          totalMateriais,
          totalHonorarios,
          dataInicio,
          dataFim,
        });

        return {
          success: true,
          inseridos: resultado.inseridos,
          ignorados: resultado.ignorados,
          erros: resultado.erros,
        };
      }),

    // Listar importações
    list: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        limite: z.number().optional().default(50),
        })
      )
      .query(async ({ input }) => {
        return db.getImportacoesTasy(input.estabelecimentoId, input.limite);
      }),

    // Buscar importação por ID
    byId: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ input }) => {
        return db.getImportacaoTasyById(input.id);
      }),

    // Excluir importação e seus dados
    delete: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input }) => {
        return db.deleteImportacaoTasy(input.id);
      }),

    // Buscar dados importados
    dados: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
        convenio: z.string().optional(),
          tipo: z.enum(["MATERIAL", "HONORARIO"]).optional(),
        atendimento: z.string().optional(),
        guia: z.string().optional(),
        limite: z.number().optional().default(100),
        offset: z.number().optional().default(0),
        })
      )
      .query(async ({ input }) => {
        const filtros = {
          dataInicio: input.dataInicio ? new Date(input.dataInicio) : undefined,
          dataFim: input.dataFim ? new Date(input.dataFim) : undefined,
          convenio: input.convenio,
          tipo: input.tipo,
          atendimento: input.atendimento,
          guia: input.guia,
          limite: input.limite,
          offset: input.offset,
        };
        return db.getDadosTasy(input.estabelecimentoId, filtros);
      }),

    // Contar dados importados
    count: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
        convenio: z.string().optional(),
          tipo: z.enum(["MATERIAL", "HONORARIO"]).optional(),
        })
      )
      .query(async ({ input }) => {
        const filtros = {
          dataInicio: input.dataInicio ? new Date(input.dataInicio) : undefined,
          dataFim: input.dataFim ? new Date(input.dataFim) : undefined,
          convenio: input.convenio,
          tipo: input.tipo,
        };
        return db.countDadosTasy(input.estabelecimentoId, filtros);
      }),

    // Estatísticas gerais
    estatisticas: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return db.getEstatisticasTasy(input.estabelecimentoId);
      }),

    // Dados agrupados por convênio
    porConvenio: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return db.getDadosTasyPorConvenio(input.estabelecimentoId);
      }),

    // Upload automático via API (para script Python)
    uploadAutomatico: publicProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        nomeArquivo: z.string(),
        tamanhoArquivo: z.number().optional(),
        apiKey: z.string(), // Chave de API para autenticação
          dados: z.array(
            z.object({
          atendimento: z.string(),
          nrInternoConta: z.string().optional(),
          sequencia: z.string().optional(),
          dataFaturado: z.string().optional(),
          guia: z.string().optional(),
          convenio: z.string().optional(),
          paciente: z.string().optional(),
          dataConta: z.string().optional(),
          codigo: z.string().optional(),
          codigoConvenio: z.string().optional(),
          descricao: z.string().optional(),
          quantidade: z.number().optional(),
          unidade: z.string().optional(),
          valorUnitario: z.number().optional(),
          valorTotal: z.number().optional(),
          setor: z.string().optional(),
          protocolo: z.string().optional(),
          statusProtocolo: z.string().optional(),
              tipo: z.enum(["MATERIAL", "HONORARIO"]),
          medico: z.string().optional(),
          funcaoMedico: z.string().optional(),
          crm: z.string().optional(),
          valorMedico: z.number().optional(),
            })
          ),
        })
      )
      .mutation(async ({ input }) => {
        // Valida a chave de API
        const apiKeyValida = await db.validarApiKey(
          input.apiKey,
          input.estabelecimentoId
        );
        if (!apiKeyValida) {
          throw new TRPCError({
            code: "UNAUTHORIZED",
            message:
              "Chave de API inválida ou sem permissão para este estabelecimento",
          });
        }

        // Cria o registro de importação
        const importacao = await db.createImportacaoTasy({
          estabelecimentoId: input.estabelecimentoId,
          userId: apiKeyValida.userId,
          nomeArquivo: input.nomeArquivo,
          tamanhoArquivo: input.tamanhoArquivo,
          status: "processando",
          progresso: 0,
          totalRegistros: input.dados.length,
        });

        // Prepara os registros para inserção
        const registros = input.dados.map(d => ({
          estabelecimentoId: input.estabelecimentoId,
          importacaoId: importacao.id,
          atendimento: d.atendimento,
          nrInternoConta: d.nrInternoConta || null,
          sequencia: d.sequencia || null,
          dataFaturado: d.dataFaturado ? new Date(d.dataFaturado) : null,
          guia: d.guia || null,
          convenio: d.convenio || null,
          paciente: d.paciente || null,
          dataConta: d.dataConta ? new Date(d.dataConta) : null,
          codigo: d.codigo || null,
          codigoConvenio: d.codigoConvenio || null,
          descricao: d.descricao || null,
          quantidade: d.quantidade?.toString() || null,
          unidade: d.unidade || null,
          valorUnitario: d.valorUnitario?.toString() || null,
          valorTotal: d.valorTotal?.toString() || null,
          setor: d.setor || null,
          protocolo: d.protocolo || null,
          statusProtocolo: d.statusProtocolo || null,
          tipo: d.tipo,
          medico: d.medico || null,
          funcaoMedico: d.funcaoMedico || null,
          crm: d.crm || null,
          valorMedico: d.valorMedico?.toString() || null,
        }));

        // Insere em lotes
        const resultado = await db.insertDadosTasyBatch(
          registros as any,
          input.estabelecimentoId
        );

        // Conta por tipo
        const totalMateriais = input.dados.filter(
          d => d.tipo === "MATERIAL"
        ).length;
        const totalHonorarios = input.dados.filter(
          d => d.tipo === "HONORARIO"
        ).length;

        // Calcula datas
        const datas = input.dados
          .filter(d => d.dataFaturado)
          .map(d => new Date(d.dataFaturado!));
        const dataInicio =
          datas.length > 0
            ? new Date(Math.min(...datas.map(d => d.getTime())))
            : null;
        const dataFim =
          datas.length > 0
            ? new Date(Math.max(...datas.map(d => d.getTime())))
            : null;

        // Atualiza estatísticas
        const status = resultado.erros > 0 ? "concluido_parcial" : "concluido";
        await db.updateImportacaoTasy(importacao.id, {
          status,
          progresso: 100,
          registrosImportados: resultado.inseridos,
          registrosIgnorados: resultado.ignorados,
          registrosErro: resultado.erros,
          totalMateriais,
          totalHonorarios,
          dataInicio,
          dataFim,
        });

        return {
          success: true,
          importacaoId: importacao.id,
          inseridos: resultado.inseridos,
          ignorados: resultado.ignorados,
          erros: resultado.erros,
        };
      }),

    // Conciliação Tasy x XML
    conciliar: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        arquivoId: z.number(),
        })
      )
      .mutation(async ({ input }) => {
        return db.compararTasyComXML(input.estabelecimentoId, input.arquivoId);
      }),

    // Resumo de conciliação por convênio
    resumoConciliacao: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return db.getResumoConciliacaoTasyData(input.estabelecimentoId);
      }),

    // Buscar dados para conciliação
    dadosParaConciliacao: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
        convenio: z.string().optional(),
        guia: z.string().optional(),
        atendimento: z.string().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getDadosTasyParaConciliacao(input.estabelecimentoId, {
          dataInicio: input.dataInicio ? new Date(input.dataInicio) : undefined,
          dataFim: input.dataFim ? new Date(input.dataFim) : undefined,
          convenio: input.convenio,
          guia: input.guia,
          atendimento: input.atendimento,
        });
      }),

    // Marcar dados como processados
    marcarProcessados: protectedProcedure
      .input(
        z.object({
        ids: z.array(z.number()),
        procedimentoId: z.number(),
        })
      )
      .mutation(async ({ input }) => {
        return db.marcarDadosTasyProcessados(input.ids, input.procedimentoId);
      }),

    // Validar dados do Tasy com regras de negócio
    validarComRegras: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
        convenio: z.string().optional(),
        atendimento: z.string().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.validarDadosTasyComRegras(input.estabelecimentoId, {
          dataInicio: input.dataInicio ? new Date(input.dataInicio) : undefined,
          dataFim: input.dataFim ? new Date(input.dataFim) : undefined,
          convenio: input.convenio,
          atendimento: input.atendimento,
        });
      }),

    // Resumo de validação por convênio
    resumoValidacao: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return db.getResumoValidacaoTasyPorConvenio(input.estabelecimentoId);
      }),

    // Comparativo entre dois períodos
    comparativo: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        periodo1Mes: z.number().min(1).max(12),
        periodo1Ano: z.number().min(2000).max(2100),
        periodo2Mes: z.number().min(1).max(12),
        periodo2Ano: z.number().min(2000).max(2100),
          agrupamento: z
            .enum(["convenio", "setor", "medico", "tipo"])
            .default("convenio"),
        })
      )
      .query(async ({ input }) => {
        return db.getDadosTasyComparativo(
          input.estabelecimentoId,
          { mes: input.periodo1Mes, ano: input.periodo1Ano },
          { mes: input.periodo2Mes, ano: input.periodo2Ano },
          input.agrupamento
        );
      }),

    // ============ CONTAS PAGAS E ITENS PAGOS ============

    // Processar contas pagas do SQLite
    processarContasPagas: protectedProcedure
      .input(
        z.object({
        importacaoId: z.number(),
          dados: z.array(
            z.object({
          dataRetorno: z.string().optional(),
          seqRetornoGeral: z.string().optional(),
          titulo: z.string().optional(),
          guia: z.string().optional(),
          nrSeqConta: z.string().optional(),
          nrConta: z.string().optional(),
          convenio: z.string().optional(),
          nrProtocolo: z.string().optional(),
          dataRecebimento: z.string().optional(),
          pagoConta: z.number().optional(),
          glosaConta: z.number().optional(),
            })
          ),
        })
      )
      .mutation(async ({ input }) => {
        const importacao = await db.getImportacaoTasyById(input.importacaoId);
        if (!importacao) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Importação não encontrada",
          });
        }

        const registros = input.dados.map(d => ({
          estabelecimentoId: importacao.estabelecimentoId,
          importacaoId: input.importacaoId,
          dataRetorno: d.dataRetorno ? new Date(d.dataRetorno) : null,
          seqRetornoGeral: d.seqRetornoGeral || null,
          titulo: d.titulo || null,
          guia: d.guia || null,
          nrSeqConta: d.nrSeqConta || null,
          nrConta: d.nrConta || null,
          convenio: d.convenio || null,
          nrProtocolo: d.nrProtocolo || null,
          dataRecebimento: d.dataRecebimento
            ? new Date(d.dataRecebimento)
            : null,
          pagoConta: d.pagoConta?.toString() || null,
          glosaConta: d.glosaConta?.toString() || null,
        }));

        const resultado = await db.insertContasPagasTasyBatch(
          registros as any,
          importacao.estabelecimentoId
        );

        return {
          success: true,
          inseridos: resultado.inseridos,
          erros: resultado.erros,
        };
      }),

    // Processar itens pagos do SQLite
    processarItensPagos: protectedProcedure
      .input(
        z.object({
        importacaoId: z.number(),
          dados: z.array(
            z.object({
          titulo: z.string().optional(),
          guia: z.string().optional(),
          nrSeqConta: z.string().optional(),
          conta: z.string().optional(),
          nrProtocolo: z.string().optional(),
          dataRecebimento: z.string().optional(),
          glosaItem: z.number().optional(),
          qndGlosaItem: z.number().optional(),
          motivoGlosa: z.string().optional(),
          procedimento: z.string().optional(),
          material: z.string().optional(),
          setor: z.string().optional(),
            })
          ),
        })
      )
      .mutation(async ({ input }) => {
        const importacao = await db.getImportacaoTasyById(input.importacaoId);
        if (!importacao) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Importação não encontrada",
          });
        }

        const registros = input.dados.map(d => ({
          estabelecimentoId: importacao.estabelecimentoId,
          importacaoId: input.importacaoId,
          titulo: d.titulo || null,
          guia: d.guia || null,
          nrSeqConta: d.nrSeqConta || null,
          conta: d.conta || null,
          nrProtocolo: d.nrProtocolo || null,
          dataRecebimento: d.dataRecebimento
            ? new Date(d.dataRecebimento)
            : null,
          glosaItem: d.glosaItem?.toString() || null,
          qndGlosaItem: d.qndGlosaItem?.toString() || null,
          motivoGlosa: d.motivoGlosa || null,
          procedimento: d.procedimento || null,
          material: d.material || null,
          setor: d.setor || null,
        }));

        const resultado = await db.insertItensPagosTasyBatch(
          registros as any,
          importacao.estabelecimentoId
        );

        return {
          success: true,
          inseridos: resultado.inseridos,
          erros: resultado.erros,
        };
      }),

    // Buscar contas pagas
    contasPagas: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
        convenio: z.string().optional(),
        guia: z.string().optional(),
        limite: z.number().optional().default(100),
        offset: z.number().optional().default(0),
        })
      )
      .query(async ({ input }) => {
        return db.getContasPagasTasy(input.estabelecimentoId, {
          dataInicio: input.dataInicio ? new Date(input.dataInicio) : undefined,
          dataFim: input.dataFim ? new Date(input.dataFim) : undefined,
          convenio: input.convenio,
          guia: input.guia,
          limite: input.limite,
          offset: input.offset,
        });
      }),

    // Buscar itens pagos
    itensPagos: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        guia: z.string().optional(),
        conta: z.string().optional(),
        limite: z.number().optional().default(100),
        offset: z.number().optional().default(0),
        })
      )
      .query(async ({ input }) => {
        return db.getItensPagosTasy(input.estabelecimentoId, {
          guia: input.guia,
          conta: input.conta,
          limite: input.limite,
          offset: input.offset,
        });
      }),

    // Conciliação completa Tasy (faturado x pago x glosado)
    conciliacaoCompleta: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
        convenio: z.string().optional(),
        guia: z.string().optional(),
        atendimento: z.string().optional(),
        mesAno: z.string().optional(), // Formato: "2024-01" para filtrar por mês/ano
        })
      )
      .query(async ({ input }) => {
        return db.getConciliacaoTasyCompleta(input.estabelecimentoId, {
          dataInicio: input.dataInicio ? new Date(input.dataInicio) : undefined,
          dataFim: input.dataFim ? new Date(input.dataFim) : undefined,
          convenio: input.convenio,
          guia: input.guia,
          atendimento: input.atendimento,
          mesAno: input.mesAno,
        });
      }),

    // Buscar meses/anos disponíveis para filtro (da tabela dadosTasy)
    mesesDisponiveis: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return db.getMesesDisponiveisTasy(input.estabelecimentoId);
      }),

    // Buscar convênios das contas pagas
    conveniosContasPagas: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return db.getConveniosContasPagasTasy(input.estabelecimentoId);
      }),

    // ============ NOVAS ROTAS PARA TABELAS SEPARADAS ============

    // Processar procedimentos do Tasy (tabela separada)
    processarProcedimentos: protectedProcedure
      .input(
        z.object({
        importacaoId: z.number(),
          dados: z.array(
            z.object({
          atendimento: z.string(),
          nrInternoConta: z.string().optional(),
          guia: z.string().optional(),
          sequencia: z.string().optional(),
          dataFaturado: z.string().optional(),
          convenio: z.string().optional(),
          paciente: z.string().optional(),
          dataConta: z.string().optional(),
          codigo: z.string().optional(),
          codigoConvenio: z.string().optional(),
          descricao: z.string().optional(),
          quantidade: z.number().optional(),
          unidade: z.string().optional(),
          valorUnitario: z.number().optional(),
          valorTotal: z.number().optional(),
          setor: z.string().optional(),
          protocolo: z.string().optional(),
          statusProtocolo: z.string().optional(),
          medico: z.string().optional(),
          funcaoMedico: z.string().optional(),
          crm: z.string().optional(),
          valorMedico: z.number().optional(),
            })
          ),
        })
      )
      .mutation(async ({ input }) => {
        const importacao = await db.getImportacaoTasyById(input.importacaoId);
        if (!importacao) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Importação não encontrada",
          });
        }

        const registros = input.dados.map(d => ({
          estabelecimentoId: importacao.estabelecimentoId,
          importacaoId: input.importacaoId,
          atendimento: d.atendimento,
          nrInternoConta: d.nrInternoConta || null,
          guia: d.guia || null,
          sequencia: d.sequencia || null,
          dataFaturado: d.dataFaturado ? new Date(d.dataFaturado) : null,
          convenio: d.convenio || null,
          paciente: d.paciente || null,
          dataConta: d.dataConta ? new Date(d.dataConta) : null,
          codigo: d.codigo || null,
          codigoConvenio: d.codigoConvenio || null,
          descricao: d.descricao || null,
          quantidade: d.quantidade?.toString() || null,
          unidade: d.unidade || null,
          valorUnitario: d.valorUnitario?.toString() || null,
          valorTotal: d.valorTotal?.toString() || null,
          setor: d.setor || null,
          protocolo: d.protocolo || null,
          statusProtocolo: d.statusProtocolo || null,
          medico: d.medico || null,
          funcaoMedico: d.funcaoMedico || null,
          crm: d.crm || null,
          valorMedico: d.valorMedico?.toString() || null,
        }));

        const resultado = await db.insertProcedimentosTasyBatch(
          registros as any,
          importacao.estabelecimentoId
        );

        return {
          success: true,
          inseridos: resultado.inseridos,
          erros: resultado.erros,
        };
      }),

    // Processar materiais/medicamentos do Tasy (tabela separada)
    processarMatMed: protectedProcedure
      .input(
        z.object({
        importacaoId: z.number(),
          dados: z.array(
            z.object({
          atendimento: z.string(),
          nrInternoConta: z.string().optional(),
          guia: z.string().optional(),
          sequencia: z.string().optional(),
          dataFaturado: z.string().optional(),
          convenio: z.string().optional(),
          paciente: z.string().optional(),
          dataConta: z.string().optional(),
          codigo: z.string().optional(),
          codigoConvenio: z.string().optional(),
          descricao: z.string().optional(),
          quantidade: z.number().optional(),
          unidade: z.string().optional(),
          valorUnitario: z.number().optional(),
          valorTotal: z.number().optional(),
          setor: z.string().optional(),
          protocolo: z.string().optional(),
          statusProtocolo: z.string().optional(),
              tipoItem: z.enum(["material", "medicamento"]).optional(),
            })
          ),
        })
      )
      .mutation(async ({ input }) => {
        const importacao = await db.getImportacaoTasyById(input.importacaoId);
        if (!importacao) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Importação não encontrada",
          });
        }

        const registros = input.dados.map(d => ({
          estabelecimentoId: importacao.estabelecimentoId,
          importacaoId: input.importacaoId,
          atendimento: d.atendimento,
          nrInternoConta: d.nrInternoConta || null,
          guia: d.guia || null,
          sequencia: d.sequencia || null,
          dataFaturado: d.dataFaturado ? new Date(d.dataFaturado) : null,
          convenio: d.convenio || null,
          paciente: d.paciente || null,
          dataConta: d.dataConta ? new Date(d.dataConta) : null,
          codigo: d.codigo || null,
          codigoConvenio: d.codigoConvenio || null,
          descricao: d.descricao || null,
          quantidade: d.quantidade?.toString() || null,
          unidade: d.unidade || null,
          valorUnitario: d.valorUnitario?.toString() || null,
          valorTotal: d.valorTotal?.toString() || null,
          setor: d.setor || null,
          protocolo: d.protocolo || null,
          statusProtocolo: d.statusProtocolo || null,
          tipoItem: d.tipoItem || "material",
        }));

        const resultado = await db.insertMatMedTasyBatch(
          registros as any,
          importacao.estabelecimentoId
        );

        return {
          success: true,
          inseridos: resultado.inseridos,
          erros: resultado.erros,
        };
      }),

    // Gerar tabela contas_tasy unificada (junção de procedimentos + mat_med)
    gerarContasUnificadas: protectedProcedure
      .input(
        z.object({
        importacaoId: z.number(),
        })
      )
      .mutation(async ({ input }) => {
        const importacao = await db.getImportacaoTasyById(input.importacaoId);
        if (!importacao) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Importação não encontrada",
          });
        }

        const resultado = await db.gerarContasTasyUnificadas(
          importacao.estabelecimentoId,
          input.importacaoId
        );

        return {
          success: true,
          contasCriadas: resultado.contasCriadas,
          itensCriados: resultado.itensCriados,
        };
      }),

    // Buscar contas unificadas
    contasUnificadas: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        convenio: z.string().optional(),
        guia: z.string().optional(),
        nrInternoConta: z.string().optional(),
        status: z.string().optional(),
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
        mesReferencia: z.number().optional(), // 1-12
        anoReferencia: z.number().optional(), // Ex: 2025
        limite: z.number().optional().default(100),
        offset: z.number().optional().default(0),
        })
      )
      .query(async ({ input }) => {
        console.log("[contasUnificadas] Input:", JSON.stringify(input));
        const result = await db.getContasTasy(input.estabelecimentoId, {
          convenio: input.convenio,
          guia: input.guia,
          nrInternoConta: input.nrInternoConta,
          status: input.status,
          dataInicio: input.dataInicio ? new Date(input.dataInicio) : undefined,
          dataFim: input.dataFim ? new Date(input.dataFim) : undefined,
          mesReferencia: input.mesReferencia,
          anoReferencia: input.anoReferencia,
          limite: input.limite,
          offset: input.offset,
        });
        console.log("[contasUnificadas] Result count:", result?.length || 0);
        return result;
      }),

    // Buscar itens de uma conta unificada
    itensContaUnificada: protectedProcedure
      .input(z.object({ contaTasyId: z.number() }))
      .query(async ({ input }) => {
        return db.getItensContaTasy(input.contaTasyId);
      }),

    // Buscar procedimentos separados
    procedimentosSeparados: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        importacaoId: z.number().optional(),
        convenio: z.string().optional(),
        guia: z.string().optional(),
        atendimento: z.string().optional(),
        limite: z.number().optional().default(100),
        offset: z.number().optional().default(0),
        })
      )
      .query(async ({ input }) => {
        return db.getProcedimentosTasy(input.estabelecimentoId, {
          importacaoId: input.importacaoId,
          convenio: input.convenio,
          guia: input.guia,
          atendimento: input.atendimento,
          limite: input.limite,
          offset: input.offset,
        });
      }),

    // Buscar mat/med separados
    matMedSeparados: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        importacaoId: z.number().optional(),
        convenio: z.string().optional(),
        guia: z.string().optional(),
        atendimento: z.string().optional(),
          tipoItem: z.enum(["material", "medicamento"]).optional(),
        limite: z.number().optional().default(100),
        offset: z.number().optional().default(0),
        })
      )
      .query(async ({ input }) => {
        return db.getMatMedTasy(input.estabelecimentoId, {
          importacaoId: input.importacaoId,
          convenio: input.convenio,
          guia: input.guia,
          atendimento: input.atendimento,
          tipoItem: input.tipoItem,
          limite: input.limite,
          offset: input.offset,
        });
      }),

    // Limpar dados de uma importação (para reimportação)
    limparImportacao: protectedProcedure
      .input(z.object({ importacaoId: z.number() }))
      .mutation(async ({ input }) => {
        await db.limparDadosImportacaoTasy(input.importacaoId);
        return { success: true };
      }),
  }),

  // ============ DASHBOARDS SALVOS ============
  dashboards: router({
    // Listar dashboards salvos do usuário
    listar: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ ctx, input }) => {
        return db.listarDashboardsSalvos(ctx.user.id, input.estabelecimentoId);
      }),

    // Buscar dashboard por ID
    buscar: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ ctx, input }) => {
        return db.getDashboardPorId(input.id, ctx.user.id);
      }),

    // Salvar novo dashboard
    salvar: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        nome: z.string().min(1).max(255),
        descricao: z.string().optional(),
        configuracao: z.object({
          tipoGrafico: z.string(),
          agrupamento: z.string(),
          colunasSelecionadas: z.array(z.string()),
          filtros: z.object({
            mes: z.number().optional(),
            ano: z.number().optional(),
            convenio: z.string().optional(),
            tipo: z.string().optional(),
            setor: z.string().optional(),
          }),
        }),
        comparativoAtivo: z.boolean().optional(),
        periodo1Mes: z.number().optional(),
        periodo1Ano: z.number().optional(),
        periodo2Mes: z.number().optional(),
        periodo2Ano: z.number().optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        return db.salvarDashboard(ctx.user.id, input.estabelecimentoId, input);
      }),

    // Atualizar dashboard existente
    atualizar: protectedProcedure
      .input(
        z.object({
        id: z.number(),
        nome: z.string().min(1).max(255).optional(),
        descricao: z.string().optional(),
          configuracao: z
            .object({
          tipoGrafico: z.string(),
          agrupamento: z.string(),
          colunasSelecionadas: z.array(z.string()),
          filtros: z.object({
            mes: z.number().optional(),
            ano: z.number().optional(),
            convenio: z.string().optional(),
            tipo: z.string().optional(),
            setor: z.string().optional(),
          }),
            })
            .optional(),
        comparativoAtivo: z.boolean().optional(),
        periodo1Mes: z.number().optional(),
        periodo1Ano: z.number().optional(),
        periodo2Mes: z.number().optional(),
        periodo2Ano: z.number().optional(),
        favorito: z.boolean().optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        const { id, ...dados } = input;
        return db.atualizarDashboard(id, ctx.user.id, dados);
      }),

    // Excluir dashboard
    excluir: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ ctx, input }) => {
        return db.excluirDashboard(input.id, ctx.user.id);
      }),

    // Compartilhar dashboard com outro usuário
    compartilhar: protectedProcedure
      .input(
        z.object({
        dashboardId: z.number(),
        usuarioId: z.number(),
          permissao: z.enum(["visualizar", "editar"]).default("visualizar"),
        })
      )
      .mutation(async ({ ctx, input }) => {
        return db.compartilharDashboard(
          input.dashboardId,
          ctx.user.id,
          input.usuarioId,
          input.permissao
        );
      }),

    // Listar dashboards compartilhados comigo
    compartilhadosComigo: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number().optional() }))
      .query(async ({ ctx, input }) => {
        return db.listarDashboardsCompartilhadosComigo(
          ctx.user.id,
          input.estabelecimentoId
        );
      }),

    // Listar compartilhamentos de um dashboard
    listarCompartilhamentos: protectedProcedure
      .input(z.object({ dashboardId: z.number() }))
      .query(async ({ ctx, input }) => {
        return db.listarCompartilhamentosDashboard(
          input.dashboardId,
          ctx.user.id
        );
      }),

    // Remover compartilhamento
    removerCompartilhamento: protectedProcedure
      .input(z.object({ compartilhamentoId: z.number() }))
      .mutation(async ({ ctx, input }) => {
        return db.removerCompartilhamentoDashboard(
          input.compartilhamentoId,
          ctx.user.id
        );
      }),

    // Listar usuários para compartilhamento
    usuariosParaCompartilhamento: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ ctx, input }) => {
        return db.listarUsuariosParaCompartilhamento(
          input.estabelecimentoId,
          ctx.user.id
        );
      }),
  }),

  // Router de Alertas de Variação
  alertasVariacao: router({
    // Criar alerta
    criar: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        nome: z.string().min(1).max(255),
          tipoAlerta: z.enum(["queda", "aumento", "ambos"]).default("queda"),
        percentualLimite: z.number().min(1).max(100).default(20),
          metrica: z
            .enum(["faturamento", "quantidade", "glosa"])
            .default("faturamento"),
          agrupamento: z.string().default("convenio"),
          notificarEmail: z.enum(["sim", "nao"]).default("nao"),
        })
      )
      .mutation(async ({ ctx, input }) => {
        return db.criarAlertaVariacao({
          ...input,
          userId: ctx.user.id,
        });
      }),

    // Listar alertas
    listar: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number().optional() }))
      .query(async ({ ctx, input }) => {
        return db.listarAlertasVariacao(ctx.user.id, input.estabelecimentoId);
      }),

    // Atualizar alerta
    atualizar: protectedProcedure
      .input(
        z.object({
        id: z.number(),
        nome: z.string().min(1).max(255).optional(),
          tipoAlerta: z.enum(["queda", "aumento", "ambos"]).optional(),
        percentualLimite: z.number().min(1).max(100).optional(),
          metrica: z.enum(["faturamento", "quantidade", "glosa"]).optional(),
        agrupamento: z.string().optional(),
          ativo: z.enum(["sim", "nao"]).optional(),
          notificarEmail: z.enum(["sim", "nao"]).optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        const { id, ...dados } = input;
        return db.atualizarAlertaVariacao(id, ctx.user.id, dados);
      }),

    // Excluir alerta
    excluir: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ ctx, input }) => {
        return db.excluirAlertaVariacao(input.id, ctx.user.id);
      }),

    // Verificar alertas (executar manualmente)
    verificar: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .mutation(async ({ input }) => {
        return db.verificarAlertasVariacao(input.estabelecimentoId);
      }),

    // Listar histórico de alertas disparados
    historico: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        limite: z.number().default(50),
        apenasNaoVisualizados: z.boolean().default(false),
        })
      )
      .query(async ({ input }) => {
        return db.listarHistoricoAlertasVariacao(
          input.estabelecimentoId,
          input.limite,
          input.apenasNaoVisualizados
        );
      }),

    // Marcar alerta como visualizado
    marcarVisualizado: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input }) => {
        return db.marcarAlertaVisualizado(input.id);
      }),
  }),

  // ============ HISTÓRICO CONCILIAÇÃO TASY ============
  historicoConciliacao: router({
    // Salvar resultado da conciliação
    salvar: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        convenioId: z.number().optional(),
        mesReferencia: z.number().optional(),
        anoReferencia: z.number().optional(),
        totalContas: z.number(),
        contasOk: z.number(),
        contasComGlosa: z.number(),
        contasDivergentes: z.number(),
        contasNaoEncontradas: z.number(),
        valorTotalTasy: z.number(),
        valorTotalPago: z.number(),
        valorTotalGlosado: z.number(),
        valorDiferenca: z.number(),
        percentualGlosa: z.number(),
        percentualRecebido: z.number(),
        observacoes: z.string().optional(),
          itens: z.array(
            z.object({
          contaTasyId: z.number(),
          nrInternoConta: z.string(),
          guia: z.string(),
          paciente: z.string(),
          dataInternacao: z.string().optional(),
          valorTasy: z.number(),
          valorPago: z.number(),
          valorGlosado: z.number(),
          valorDiferenca: z.number(),
              statusConciliacao: z.enum([
                "ok",
                "glosa",
                "divergente",
                "nao_encontrado",
              ]),
          totalProcedimentos: z.number(),
          totalMatMed: z.number(),
          demonstrativoItemId: z.number().optional(),
            })
          ),
        })
      )
      .mutation(async ({ ctx, input }) => {
        const { itens, ...dados } = input;

        // Função para converter data de forma segura
        const parseDate = (dateStr: string | undefined): Date | undefined => {
          if (!dateStr || dateStr === "-" || dateStr === "") return undefined;
          try {
            // Tentar converter formato DD/MM/YYYY para Date
            if (dateStr.includes("/")) {
              const [day, month, year] = dateStr.split("/");
              const date = new Date(
                parseInt(year),
                parseInt(month) - 1,
                parseInt(day)
              );
              if (!isNaN(date.getTime())) return date;
            }
            // Tentar converter ISO ou outros formatos
            const date = new Date(dateStr);
            if (!isNaN(date.getTime())) return date;
            return undefined;
          } catch {
            return undefined;
          }
        };

        const result = await db.salvarResultadoConciliacao(
          { ...dados, userId: ctx.user.id },
          itens.map(item => ({
            ...item,
            dataInternacao: parseDate(item.dataInternacao),
          })),
          new Map() // Detalhes vazios por enquanto
        );
        return result;
      }),

    // Listar histórico de conciliações
    listar: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        convenioId: z.number().optional(),
        mesReferencia: z.number().optional(),
        anoReferencia: z.number().optional(),
        limite: z.number().optional().default(50),
        offset: z.number().optional().default(0),
        })
      )
      .query(async ({ input }) => {
        return db.listarHistoricoConciliacoes(input.estabelecimentoId, input);
      }),

    // Buscar detalhes de uma conciliação
    detalhes: protectedProcedure
      .input(
        z.object({
        resultadoId: z.number(),
        statusConciliacao: z.string().optional(),
        busca: z.string().optional(),
        limite: z.number().optional().default(50),
        offset: z.number().optional().default(0),
        })
      )
      .query(async ({ input }) => {
        const { resultadoId, ...filtros } = input;
        return db.getDetalhesConciliacao(resultadoId, filtros);
      }),

    // Buscar detalhes dos itens de uma conta conciliada
    detalhesItem: protectedProcedure
      .input(z.object({ itemConciliacaoId: z.number() }))
      .query(async ({ input }) => {
        return db.getDetalhesItemConciliacao(input.itemConciliacaoId);
      }),

    // Excluir conciliação
    excluir: protectedProcedure
      .input(z.object({ resultadoId: z.number() }))
      .mutation(async ({ input }) => {
        const success = await db.excluirConciliacao(input.resultadoId);
        return { success };
      }),

    // Buscar evolução das conciliações
    evolucao: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        meses: z.number().optional().default(6),
        })
      )
      .query(async ({ input }) => {
        return db.getEvolucaoConciliacoes(input.estabelecimentoId, input.meses);
      }),
  }),

  // ============ FATURADO TASY ============
  faturadoTasy: router({
    // Listar registros com filtros
    list: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        competencia: z.string().optional(),
        convenio: z.string().optional(),
          tipoItem: z.enum(["PROC/TAXA", "MAT/MED"]).optional(),
        protocolo: z.string().optional(),
        atend: z.string().optional(),
        conta: z.string().optional(),
        cdItem: z.string().optional(),
        descricao: z.string().optional(),
        comGlosa: z.boolean().optional(),
        comPagamento: z.boolean().optional(),
        limite: z.number().optional(),
        offset: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getFaturadoTasy(input.estabelecimentoId, input);
      }),

    // Contar registros
    count: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        competencia: z.string().optional(),
        convenio: z.string().optional(),
          tipoItem: z.enum(["PROC/TAXA", "MAT/MED"]).optional(),
        comGlosa: z.boolean().optional(),
        comPagamento: z.boolean().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.contarFaturadoTasy(input.estabelecimentoId, input);
      }),

    // Buscar estatísticas
    estatisticas: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        competencia: z.string().optional(),
        convenio: z.string().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getEstatisticasFaturadoTasy(input.estabelecimentoId, input);
      }),

    // Listar competências disponíveis
    competencias: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return db.listarCompetenciasFaturadoTasy(input.estabelecimentoId);
      }),

    // Listar convênios disponíveis
    convenios: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return db.listarConveniosFaturadoTasy(input.estabelecimentoId);
      }),

    // Resumo por tipo de item
    resumoPorTipo: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        competencia: z.string().optional(),
        convenio: z.string().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getResumoPorTipoFaturadoTasy(input.estabelecimentoId, input);
      }),

    // Itens glosados
    itensGlosados: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        competencia: z.string().optional(),
        convenio: z.string().optional(),
        limite: z.number().optional(),
        offset: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getItensGlosadosFaturadoTasy(input.estabelecimentoId, input);
      }),

    // Dados para Relatório BI
    dadosBI: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        competencia: z.string().optional(),
        convenio: z.string().optional(),
          tipoItem: z.enum(["PROC/TAXA", "MAT/MED"]).optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getDadosBIFaturadoTasy(input.estabelecimentoId, input);
      }),

    // Importar dados do Excel ou SQLite
    importar: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number().optional(),
          dados: z.array(
            z.object({
          sequencia: z.string().optional(),
          convenio: z.string().optional(),
          competencia: z.string().optional(),
          protocolo: z.string().optional(),
          setor: z.string().optional(),
          atend: z.string().optional(),
          conta: z.string().optional(),
          profExec: z.string().optional(),
          cdMotivoExcConta: z.string().optional(),
          dsComplMotivoExcon: z.string().optional(),
          tipoItem: z.string().optional(), // Aceita qualquer string para SQLite
          cdItem: z.string().optional(),
          cdItemTuss: z.string().optional(),
          dtItem: z.string().optional(),
          descricao: z.string().optional(),
          qtd: z.number().optional(),
          vlFaturado: z.number().optional(),
          aReceber: z.number().optional(),
          vlPago: z.number().optional(),
          vlGlosa: z.number().optional(),
          motivoGlosa: z.string().optional(),
          retorno: z.string().optional(),
          dtPgto: z.string().optional(),
            })
          ),
        importacaoId: z.number(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        // Busca o estabelecimentoId da importação
        let estabelecimentoId: number = input.estabelecimentoId || 0;

        // Sempre busca da importação para garantir consistência
        const importacoes = await db.getImportacoesTasy(0); // Busca todas
        const importacao = importacoes.find(i => i.id === input.importacaoId);
        if (importacao && importacao.estabelecimentoId) {
          estabelecimentoId = importacao.estabelecimentoId;
        }

        if (!estabelecimentoId || estabelecimentoId === 0) {
          throw new TRPCError({
            code: "BAD_REQUEST",
            message: "EstabelecimentoId não encontrado na importação",
          });
        }

        // Função para normalizar tipoItem
        const normalizarTipoItem = (
          tipo: string | undefined
        ): "PROC/TAXA" | "MAT/MED" => {
          if (!tipo) return "PROC/TAXA";
          const tipoUpper = tipo.toUpperCase();
          // Verifica se é mat/med
          if (
            tipoUpper.includes("MAT") ||
            tipoUpper.includes("MED") ||
            tipoUpper.includes("MATERIAL") ||
            tipoUpper.includes("MEDICAMENTO")
          ) {
            return "MAT/MED";
          }
          // Caso contrário, é proc/taxa
          return "PROC/TAXA";
        };

        const registros = input.dados.map(d => ({
          estabelecimentoId: estabelecimentoId,
          importacaoId: input.importacaoId,
          sequencia: d.sequencia || null,
          convenio: d.convenio || null,
          competencia: d.competencia || null,
          protocolo: d.protocolo || null,
          setor: d.setor || null,
          atend: d.atend || null,
          conta: d.conta || null,
          profExec: d.profExec || null,
          cdMotivoExcConta: d.cdMotivoExcConta || null,
          dsComplMotivoExcon: d.dsComplMotivoExcon || null,
          tipoItem: normalizarTipoItem(d.tipoItem),
          cdItem: d.cdItem || null,
          cdItemTuss: d.cdItemTuss || null,
          dtItem: d.dtItem ? new Date(d.dtItem) : null,
          descricao: d.descricao || null,
          qtd: d.qtd ? String(d.qtd) : null,
          vlFaturado: d.vlFaturado ? String(d.vlFaturado) : null,
          aReceber: d.aReceber ? String(d.aReceber) : null,
          vlPago: d.vlPago ? String(d.vlPago) : null,
          vlGlosa: d.vlGlosa ? String(d.vlGlosa) : null,
          motivoGlosa: d.motivoGlosa || null,
          retorno: d.retorno || null,
          dtPgto: d.dtPgto ? new Date(d.dtPgto) : null,
        }));

        const resultado = await db.insertFaturadoTasyBatch(
          registros as any,
          estabelecimentoId || 0
        );
        return resultado;
      }),

    // Excluir por importação
    excluirPorImportacao: protectedProcedure
      .input(z.object({ importacaoId: z.number() }))
      .mutation(async ({ input }) => {
        const count = await db.excluirFaturadoTasyPorImportacao(
          input.importacaoId
        );
        return { excluidos: count };
      }),

    // Dados para Relatório Tasy (formato compatível com importacaoTasy.dados)
    dadosRelatorio: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
        competencia: z.string().optional(), // Competência específica (ex: "2025-12-01 00:00:00")
        convenio: z.string().optional(),
          tipo: z.enum(["MATERIAL", "HONORARIO"]).optional(),
        limite: z.number().optional().default(10000),
        offset: z.number().optional().default(0),
        })
      )
      .query(async ({ input }) => {
        const filtros = {
          dataInicio: input.dataInicio ? new Date(input.dataInicio) : undefined,
          dataFim: input.dataFim ? new Date(input.dataFim) : undefined,
          competencia: input.competencia, // Competência específica
          convenio: input.convenio,
          tipo: input.tipo,
          limite: input.limite,
          offset: input.offset,
        };
        return db.getFaturadoTasyParaRelatorio(
          input.estabelecimentoId,
          filtros
        );
      }),

    // Conciliação de contas - agrupado por conta
    conciliacao: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        competencia: z.string().optional(),
        convenio: z.string().optional(),
        conta: z.string().optional(),
          status: z.enum(["pago", "parcial", "glosado", "pendente"]).optional(),
        limite: z.number().optional().default(500),
        offset: z.number().optional().default(0),
        })
      )
      .query(async ({ input }) => {
        return db.getConciliacaoFaturadoTasy(input.estabelecimentoId, {
          competencia: input.competencia,
          convenio: input.convenio,
          conta: input.conta,
          status: input.status,
          limite: input.limite,
          offset: input.offset,
        });
      }),

    // Itens de uma conta específica
    itensPorConta: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        conta: z.string(),
        competencia: z.string().optional(),
        convenio: z.string().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getItensFaturadoTasyPorConta(
          input.estabelecimentoId,
          input.conta,
          {
          competencia: input.competencia,
          convenio: input.convenio,
          }
        );
      }),

    // Competências disponíveis para filtros
    competenciasDisponiveis: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return db.getCompetenciasFaturadoTasy(input.estabelecimentoId);
      }),
  }),

  // ============ RELATÓRIOS ============

  // ============ RELATÓRIOS BI ============
  relatoriosBI: router({
    // Buscar dados consolidados para BI
    dados: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesReferencia: z.number().optional(),
        anoReferencia: z.number().optional(),
        convenioId: z.number().optional(),
        tipo: z.string().optional(),
        setor: z.string().optional(),
        paciente: z.string().optional(),
        procedimento: z.string().optional(),
        codigoPrestadorExecutante: z.string().optional(),
        receberHospital: z.string().optional(), // 'S' = Hospital, 'N' = Terceiros
        })
      )
      .query(async ({ input }) => {
        return db.getDadosBI(input);
      }),

    // Buscar opções de filtro
    opcoesFiltro: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return db.getOpcoesFiltroBi(input.estabelecimentoId);
      }),

    // Itens por categoria
    itemsPorCategoria: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesReferencia: z.number().optional(),
        anoReferencia: z.number().optional(),
        convenioId: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getItemsPorCategoria(input);
      }),

    // Glosas por motivo
    glosasPorMotivo: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesReferencia: z.number().optional(),
        anoReferencia: z.number().optional(),
        convenioId: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getGlosasPorMotivo(input);
      }),

    // Performance por médico
    performanceMedico: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesReferencia: z.number().optional(),
        anoReferencia: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getPerformanceMedico(input);
      }),

    // BI Faturado x Recebido (faturamento_unificado + recebimentos_excel)
    faturadoRecebido: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        competencias: z.array(z.string()).optional(),
        convenios: z.array(z.string()).optional(),
        tipoItem: z.string().optional(),
        setor: z.string().optional(),
        })
      )
      .query(async ({ input }) => {
        return getDadosBIFaturadoRecebido(input);
      }),

    // BI Faturado x Recebido UNIFICADO (todas as fontes + filtros avançados)
    faturadoRecebidoUnificado: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        competencias: z.array(z.string()).optional(),
        convenios: z.array(z.string()).optional(),
        tipoItem: z.string().optional(),
        setor: z.string().optional(),
        origemSistema: z.string().optional(),
        prestador: z.string().optional(),
        })
      )
      .query(async ({ input }) => {
        return getDadosFaturadoRecebidoUnificado(input);
      }),
  }),

  // ============ RECEBIMENTO TISS ============
  recebimentoTiss: router({
    // Importar dados de Excel para recebimento_tiss
    importarExcel: protectedProcedure
      .input(
        z.object({
        arquivoId: z.number(),
        estabelecimentoId: z.number(),
        conteudo: z.string(), // Base64 encoded
        })
      )
      .mutation(async ({ input }) => {
        try {
          const buffer = Buffer.from(input.conteudo, "base64");

          const result = await parseExcelRecebimentoTiss(
            buffer,
            input.arquivoId,
            input.estabelecimentoId
          );

          if (!result.success) {
            return {
              success: false,
              error: result.error,
              totalImportados: 0,
            };
          }

          // Inserir itens no banco de dados
          const totalImportados = await db.insertRecebimentoTiss(result.items);

          return {
            success: true,
            totalImportados,
            totalLinhas: result.totalRows,
          };
        } catch (error) {
          console.error("[recebimentoTiss.importarExcel] Erro:", error);
          return {
            success: false,
            error:
              error instanceof Error
                ? error.message
                : "Erro ao importar arquivo",
            totalImportados: 0,
          };
        }
      }),

    // Listar itens de recebimento_tiss
    list: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        arquivoId: z.number().optional(),
        convenioId: z.number().optional(),
        numeroProtocolo: z.string().optional(),
        numeroGuia: z.string().optional(),
        beneficiario: z.string().optional(),
        search: z.string().optional(),
        dataInicio: z.string().optional(),
        dataFim: z.string().optional(),
        situacaoItem: z.string().optional(),
            statusGlosa: z
              .enum(["todos", "pago", "glosado", "parcial"])
              .optional(),
        codigoPrestadorExecutante: z.string().optional(),
        mesReferencia: z.number().min(1).max(12).optional(),
        anoReferencia: z.number().min(2000).max(2100).optional(),
        page: z.number().optional(),
        pageSize: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getRecebimentoTiss(input);
      }),

    // Estatísticas de recebimento
    stats: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        arquivoId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getRecebimentoTissStats(input || {});
      }),

    // Excluir itens por arquivo
    excluirPorArquivo: protectedProcedure
      .input(z.object({ arquivoId: z.number() }))
      .mutation(async ({ input }) => {
        const deleted = await db.deleteRecebimentoTissByArquivo(
          input.arquivoId
        );
        return { success: true, deleted };
      }),
  }),

  // ============ FATURAMENTO TISS ============
  faturamentoTiss: router({
    // Listar itens de faturamento_tiss com filtros e paginação
    list: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        convenioId: z.number().optional(),
        arquivoId: z.number().optional(),
        search: z.string().optional(),
        mesReferencia: z.number().min(1).max(12).optional(),
        anoReferencia: z.number().min(2000).max(2100).optional(),
        page: z.number().optional(),
        pageSize: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        const result = await db.getFaturamentoTiss(input || {});
        return result;
      }),

    // Resumo de faturamento
    resumo: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        convenioId: z.number().optional(),
        mesReferencia: z.number().min(1).max(12).optional(),
        anoReferencia: z.number().min(2000).max(2100).optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getFaturamentoTissResumo(input || {});
      }),

    // Buscar itens individuais de uma guia específica (sem agrupamento)
    itensGuia: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number().optional(),
        numeroGuiaPrestador: z.string(),
        numeroLote: z.string().optional(),
        convenioId: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getFaturamentoTissItensGuia(input);
      }),

    // Buscar guias com múltiplos lotes (altas administrativas)
    guiasMultiplosLotes: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        convenioId: z.number().optional(),
        mesReferencia: z.number().min(1).max(12).optional(),
        anoReferencia: z.number().min(2000).max(2100).optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getGuiasMultiplosLotes(input || {});
      }),

    // Excluir itens por arquivo
    excluirPorArquivo: protectedProcedure
      .input(z.object({ arquivoId: z.number() }))
      .mutation(async ({ input }) => {
        const deleted = await db.deleteFaturamentoTissByArquivo(
          input.arquivoId
        );

        // Limpar contas_convenio_itens também
        try {
          const { contasConvenioItens: cciT, contasConvenioResumo: ccrT } =
            await import("../drizzle/schema");
          const { eq, and } = await import("drizzle-orm");
          const dbExcl = await getDb();
          if (dbExcl) {
            const guiasAfetadas = await dbExcl
              .selectDistinct({
              numeroConta: cciT.numeroConta,
                estabelecimentoId: cciT.estabelecimentoId,
              })
              .from(cciT)
              .where(eq(cciT.arquivoId, input.arquivoId));

            await dbExcl
              .delete(cciT)
              .where(eq(cciT.arquivoId, input.arquivoId));

            for (const guia of guiasAfetadas) {
              const restantes = await dbExcl
                .select()
                .from(cciT)
                .where(
                  and(
                    eq(cciT.numeroConta, guia.numeroConta),
                    eq(cciT.estabelecimentoId, guia.estabelecimentoId)
                  )
              );
              if (restantes.length === 0) {
                await dbExcl
                  .delete(ccrT)
                  .where(
                    and(
                      eq(ccrT.numeroConta, guia.numeroConta),
                      eq(ccrT.estabelecimentoId, guia.estabelecimentoId)
                    )
                );
              } else {
                let vt = 0;
                for (const i of restantes)
                  vt += parseFloat(String(i.valorTotal || 0));
                await dbExcl
                  .update(ccrT)
                  .set({
                    totalItens: restantes.length,
                    valorTotal: String(vt.toFixed(2)),
                  })
                  .where(
                    and(
                      eq(ccrT.numeroConta, guia.numeroConta),
                      eq(ccrT.estabelecimentoId, guia.estabelecimentoId)
                    )
                  );
              }
            }
          }
        } catch (e) {
          console.error("[excluirPorArquivo] Erro contas_convenio:", e);
        }

        return { success: true, deleted };
      }),
  }),

  // ============ RECEBIMENTOS EXCEL ============
  // ============ DEMONSTRATIVO ============
  demonstrativo: router({
    // Listar contas do demonstrativo (agrupadas por guia)
    contas: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        convenioId: z.number().optional(),
        arquivoId: z.number().optional(),
        mesReferencia: z.number().min(1).max(12).optional(),
        anoReferencia: z.number().min(2000).max(2100).optional(),
        dataPagamento: z.string().optional(),
        search: z.string().optional(),
            statusGlosa: z
              .enum(["todos", "pago", "glosado", "parcial"])
              .optional(),
        page: z.number().optional(),
        pageSize: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getDemonstrativoContas(input || {});
      }),

    // Buscar itens de uma conta específica
    itensPorGuia: protectedProcedure
      .input(
        z.object({
        numeroGuia: z.string(),
        protocolo: z.string().optional(),
        convenioId: z.number().optional(),
        arquivoId: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return db.getDemonstrativoItensPorGuia(input);
      }),

    // Buscar itens glosados para recurso
    itensGlosados: protectedProcedure
      .input(
        z
          .object({
        convenioId: z.number().optional(),
        arquivoId: z.number().optional(),
        mesReferencia: z.number().min(1).max(12).optional(),
        anoReferencia: z.number().min(2000).max(2100).optional(),
        dataPagamento: z.string().optional(),
        search: z.string().optional(),
        page: z.number().optional(),
        pageSize: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getDemonstrativoItensGlosados(input || {});
      }),

    // Resumo do demonstrativo
    resumo: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        convenioId: z.number().optional(),
        mesReferencia: z.number().min(1).max(12).optional(),
        anoReferencia: z.number().min(2000).max(2100).optional(),
        dataPagamento: z.string().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getDemonstrativoResumo(input || {});
      }),

    // Competências disponíveis (datas de referência do upload)
    competencias: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        convenioId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getDemonstrativoCompetencias(input || {});
      }),

    // Datas de pagamento disponíveis
    datasPagamento: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        convenioId: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getDemonstrativoDatasPagamento(input || {});
      }),

    sincronizar: protectedProcedure
      .input(
        z.object({
        arquivoId: z.number(),
        })
      )
      .mutation(async ({ input }) => {
        try {
          const { syncDemonstrativoByArquivo } = await import(
            "./syncDemonstrativo"
          );
          const result = await syncDemonstrativoByArquivo(
            input.arquivoId,
            "excel"
          );
          return result;
        } catch (error) {
          console.error("[Demonstrativo.Sincronizar] Erro:", error);
          return {
            success: false,
            total: 0,
            error: error instanceof Error ? error.message : "Erro desconhecido",
          };
        }
      }),
  }),

  recebimentosExcel: router({
    // Importar dados de Excel para recebimentos_excel
    importarExcel: protectedProcedure
      .input(
        z.object({
        arquivoId: z.number(),
        estabelecimentoId: z.number(),
        })
      )
      .mutation(async ({ input }) => {
        const { parseExcelRecebimentosExcel } = await import(
          "./recebimentosExcelParser"
        );

        // Buscar arquivo
        const arquivo = await db.getArquivoById(input.arquivoId);
        if (!arquivo) {
          throw new Error("Arquivo não encontrado");
        }

        exigirArquivoOriginal(arquivo);
        // Baixar conteúdo do S3
        assertInternalProcessingReady();
        const downloadController = new AbortController();
        const downloadTimer = setTimeout(() => downloadController.abort(), 60_000);
        let buffer: Buffer;
        try {
          const response = await fetch(arquivo.s3Url, { signal: downloadController.signal });
          if (!response.ok) throw new Error("Download recusado.");
          buffer = Buffer.from(await response.arrayBuffer());
        } catch {
          await db.updateArquivoStatus(input.arquivoId, "erro").catch(() => undefined);
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: "Não foi possível baixar o arquivo para processamento.",
          });
        } finally {
          clearTimeout(downloadTimer);
        }

        // Parsear Excel
        const records = parseExcelRecebimentosExcel(buffer, input.arquivoId);

        if (records.length === 0) {
          return {
            success: false,
            message: "Nenhum registro encontrado no arquivo",
            imported: 0,
          };
        }

        // Inserir no banco
        const imported = await db.insertRecebimentosExcelBatch(records);

        // Atualizar status do arquivo
        await db.updateArquivo(input.arquivoId, {
          status: "processado",
          totalItens: records.length,
          itensProcessados: records.length,
          progresso: 100,
        });

        return { success: true, imported, total: records.length };
      }),

    // Listar itens com filtros e paginação
    list: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        convenioId: z.number().optional(),
        arquivoId: z.number().optional(),
        situacaoItem: z.string().optional(),
        search: z.string().optional(),
        mesReferencia: z.number().min(1).max(12).optional(),
        anoReferencia: z.number().min(2000).max(2100).optional(),
        page: z.number().optional(),
        limit: z.number().optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getRecebimentosExcel(input || {});
      }),

    // Resumo de recebimentos
    resumo: protectedProcedure
      .input(
        z
          .object({
        estabelecimentoId: z.number().optional(),
        convenioId: z.number().optional(),
        mesReferencia: z.number().min(1).max(12).optional(),
        anoReferencia: z.number().min(2000).max(2100).optional(),
          })
          .optional()
      )
      .query(async ({ input }) => {
        return db.getRecebimentosExcelResumo(input || {});
      }),

    // Excluir itens por arquivo
    excluirPorArquivo: protectedProcedure
      .input(z.object({ arquivoId: z.number() }))
      .mutation(async ({ input }) => {
        const deleted = await db.deleteRecebimentosExcelByArquivo(
          input.arquivoId
        );
        return { success: true, deleted };
      }),
  }),

  // ===== ATENDIMENTOS (banco interno com fallback para PostgreSQL externo) =====
  atendimentos: router({
    listar: reportProcedure("atendimentos")
      .input(z.object({ estabelecimentoId: z.number().int().positive() }))
      .query(async ({ input }) => {
        try {
          const estabId = input.estabelecimentoId;
          // Buscar da tabela atendimentos_unificados filtrado por estabelecimento
          const dbInstance = await getDb();
          if (dbInstance) {
            const { atendimentos: atendimentosUnificados } = await import(
              "../drizzle/schema-integracao"
            );
            const { eq, and, notLike, or, isNull, sql } = await import(
              "drizzle-orm"
            );
            // Filtrar por estabelecimento e excluir os que são "A_FATURAR" (esses vão para outra tela)
            let conditions = [
              eq(atendimentosUnificados.estabelecimentoId, estabId),
            ];
            conditions.push(eq(atendimentosUnificados.ativoNaFonte, true));
            conditions.push(
              notLike(
                atendimentosUnificados.descricao_atendimento,
                "%A_FATURAR%"
              )
            );
            // Para dados TASY: mostrar apenas os que NÃO têm protocolo (nomeProtocolo NULL ou vazio)
            // Os que têm protocolo vão para a tela "Atendimentos Sem Protocolo"
            conditions.push(
              or(
                sql`${atendimentosUnificados.origemSistema} NOT IN ('tasy', 'tasy_hemolabor')`,
                isNull(atendimentosUnificados.nomeProtocolo),
                eq(atendimentosUnificados.nomeProtocolo, "")
              )!
            );
            const dadosInternos = await dbInstance
              .select()
              .from(atendimentosUnificados)
              .where(and(...conditions));

            // Buscar motivos de notificação do MySQL interno
            const numatends = dadosInternos
              .map(d => d.numero_atendimento || "")
              .filter(n => n !== "");
            let motivosMap: Record<string, string> = {};
            try {
              const notificacoesMySQL = await buscarNotificacoesAtendimento(
                numatends,
                estabId
              );
              for (const [numatend, notif] of Object.entries(
                notificacoesMySQL
              )) {
                motivosMap[numatend] = notif.motivo;
              }
            } catch (err) {
              console.error(
                "[atendimentos.listar] Erro ao buscar motivos de notificação do MySQL:",
                err
              );
            }

            return dadosInternos.map(d => ({
              numatend: d.numero_atendimento || "",
              nomeplaco: d.convenio || "",
              nomepac: d.paciente || "",
              carater: d.caracter_atendimento || "",
              datatend: d.data_entrada
                ? new Date(d.data_entrada).toISOString()
                : "",
              datasai: d.data_saida
                ? new Date(d.data_saida).toISOString()
                : null,
              tipoatend: d.tipo_atendimento || "",
              tipoatendimentodescricao: d.descricao_atendimento || "",
              codserv: d.codigo_servico || "",
              procprin: d.codigo_procedimento || "",
              codcc_destino: d.destino_conta || "",
              motivo: motivosMap[d.numero_atendimento || ""] || null,
              diasParado:
                d.origemSistema === "tasy" ||
                d.origemSistema === "tasy_hemolabor"
                ? calcularDiasParadoUnificado(
                      d.data_entrada
                        ? new Date(d.data_entrada).toISOString()
                        : null,
                      d.data_saida
                        ? new Date(d.data_saida).toISOString()
                        : null,
                      d.origemSistema || "",
                    d.dtEntrega ? new Date(d.dtEntrega).toISOString() : null,
                    d.dtEtapa ? new Date(d.dtEtapa).toISOString() : null
                  )
                  : calcularDiasParado(
                      d.data_saida ? new Date(d.data_saida).toISOString() : null
                    ),
              origemSistema: d.origemSistema,
              // Novos campos Tasy
              dsCategoria: d.dsCategoria || "",
              dsPlano: d.dsPlano || "",
              competencia: d.competencia || "",
              referencia: d.referencia || "",
              protTasy: d.protTasy || "",
              nomeProtocolo: d.nomeProtocolo || "",
              protConv: d.protConv || "",
              dtEntrega: d.dtEntrega
                ? new Date(d.dtEntrega).toISOString()
                : null,
              protStatus: d.protStatus || "",
              titulo: d.titulo || "",
              dtTitulo: d.dtTitulo ? new Date(d.dtTitulo).toISOString() : null,
              dataVencimento: d.dataVencimento
                ? new Date(d.dataVencimento).toISOString()
                : null,
              dsSetorEntrada: d.dsSetorEntrada || "",
              dsSetorLeito: d.dsSetorLeito || "",
              etapaConta: d.etapaConta || "",
              setorEtapa: d.setorEtapa || "",
              dtEtapa: d.dtEtapa ? new Date(d.dtEtapa).toISOString() : null,
              userEtapa: d.userEtapa || "",
              motivoDevolucao: d.motivoDevolucao || "",
              conta: d.conta || "",
              autorizacao: d.autorizacao || "",
              valorConta: d.valorConta ? String(d.valorConta) : "",
              matricula: d.matricula || "",
              sexo: d.sexo || "",
              idade: d.idade || "",
              medicoResp: d.medicoResp || "",
              crm: d.crm || "",
              dsMotivoAlta: d.dsMotivoAlta || "",
              dataInicio: d.dataInicio || "",
              dataFim: d.dataFim || "",
              codServico: d.codServico || "",
              centroCusto: d.centroCusto || "",
            }));
          }
          // Fallback: banco interno indisponível, retornar vazio
          return [];
        } catch (err: any) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao buscar atendimentos: ${err.message}`,
          });
        }
      }),

    // Listar atendimentos TASY que TÊM protocolo preenchido (para tela Atendimentos Sem Protocolo)
    listarComProtocolo: reportProcedure("atendimentos")
      .input(z.object({ estabelecimentoId: z.number().int().positive() }))
      .query(async ({ input }) => {
        try {
          const estabId = input.estabelecimentoId;
          const dbInstance = await getDb();
          if (dbInstance) {
            const { atendimentos: atendimentosUnificados } = await import(
              "../drizzle/schema-integracao"
            );
            const { eq, and, notLike, isNotNull, ne, or } = await import(
              "drizzle-orm"
            );
            let conditions = [
              eq(atendimentosUnificados.estabelecimentoId, estabId),
            ];
            conditions.push(eq(atendimentosUnificados.ativoNaFonte, true));
            // Apenas TASY com nomeProtocolo preenchido
            conditions.push(
              or(
                eq(atendimentosUnificados.origemSistema, "tasy"),
                eq(atendimentosUnificados.origemSistema, "tasy_hemolabor")
              )!
            );
            conditions.push(isNotNull(atendimentosUnificados.nomeProtocolo));
            conditions.push(ne(atendimentosUnificados.nomeProtocolo, ""));
            conditions.push(
              notLike(
                atendimentosUnificados.descricao_atendimento,
                "%A_FATURAR%"
              )
            );
            const dadosInternos = await dbInstance
              .select()
              .from(atendimentosUnificados)
              .where(and(...conditions));

            // Buscar motivos de notificação
            const numatends = dadosInternos
              .map(d => d.numero_atendimento || "")
              .filter(n => n !== "");
            let motivosMap: Record<string, string> = {};
            try {
              const notificacoesMySQL = await buscarNotificacoesAtendimento(
                numatends,
                estabId
              );
              for (const [numatend, notif] of Object.entries(
                notificacoesMySQL
              )) {
                motivosMap[numatend] = notif.motivo;
              }
            } catch (err) {
              console.error(
                "[atendimentos.listarComProtocolo] Erro ao buscar motivos:",
                err
              );
            }

            return dadosInternos.map(d => ({
              numatend: d.numero_atendimento || "",
              nomeplaco: d.convenio || "",
              nomepac: d.paciente || "",
              carater: d.caracter_atendimento || "",
              datatend: d.data_entrada
                ? new Date(d.data_entrada).toISOString()
                : "",
              datasai: d.data_saida
                ? new Date(d.data_saida).toISOString()
                : null,
              tipoatend: d.tipo_atendimento || "",
              tipoatendimentodescricao: d.descricao_atendimento || "",
              codserv: d.codigo_servico || "",
              procprin: d.codigo_procedimento || "",
              codcc_destino: d.destino_conta || "",
              motivo: motivosMap[d.numero_atendimento || ""] || null,
              diasParado: calcularDiasParadoUnificado(
                d.data_entrada ? new Date(d.data_entrada).toISOString() : null,
                d.data_saida ? new Date(d.data_saida).toISOString() : null,
                d.origemSistema || "",
                d.dtEntrega ? new Date(d.dtEntrega).toISOString() : null,
                d.dtEtapa ? new Date(d.dtEtapa).toISOString() : null
              ),
              origemSistema: d.origemSistema,
              dsCategoria: d.dsCategoria || "",
              dsPlano: d.dsPlano || "",
              competencia: d.competencia || "",
              referencia: d.referencia || "",
              protTasy: d.protTasy || "",
              nomeProtocolo: d.nomeProtocolo || "",
              protConv: d.protConv || "",
              dtEntrega: d.dtEntrega
                ? new Date(d.dtEntrega).toISOString()
                : null,
              protStatus: d.protStatus || "",
              titulo: d.titulo || "",
              dtTitulo: d.dtTitulo ? new Date(d.dtTitulo).toISOString() : null,
              dataVencimento: d.dataVencimento
                ? new Date(d.dataVencimento).toISOString()
                : null,
              dsSetorEntrada: d.dsSetorEntrada || "",
              dsSetorLeito: d.dsSetorLeito || "",
              etapaConta: d.etapaConta || "",
              setorEtapa: d.setorEtapa || "",
              dtEtapa: d.dtEtapa ? new Date(d.dtEtapa).toISOString() : null,
              userEtapa: d.userEtapa || "",
              motivoDevolucao: d.motivoDevolucao || "",
              conta: d.conta || "",
              autorizacao: d.autorizacao || "",
              valorConta: d.valorConta ? String(d.valorConta) : "",
              matricula: d.matricula || "",
              sexo: d.sexo || "",
              idade: d.idade || "",
              medicoResp: d.medicoResp || "",
              crm: d.crm || "",
              dsMotivoAlta: d.dsMotivoAlta || "",
              dataInicio: d.dataInicio || "",
              dataFim: d.dataFim || "",
              codServico: d.codServico || "",
              centroCusto: d.centroCusto || "",
            }));
          }
          return [];
        } catch (err: any) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao buscar atendimentos com protocolo: ${err.message}`,
          });
        }
      }),

    registrarNotificacao: reportProcedure("atendimentos")
      .input(
        z.object({
        estabelecimentoId: z.number().int().positive(),
        numatend: z.string(),
        observacao: z.string(),
          notificacoes: z.array(
            z.object({
          motivo: z.string(),
          setor: z.string(),
          medico: z.string(),
            })
          ),
        })
      )
      .mutation(async ({ input, ctx }) => {
        try {
          const pertence = await verificarAtendimentosDoEstabelecimento(
            input.estabelecimentoId,
            [input.numatend]
          );
          if (!pertence)
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "Atendimento não pertence à unidade selecionada.",
            });
          // Salvar no MySQL interno (fonte principal)
          const mysqlId = await salvarNotificacaoAtendimento({
            numatend: input.numatend,
            estabelecimentoId: input.estabelecimentoId,
            observacao: input.observacao,
            usuario: ctx.user?.name || "Usu\u00e1rio",
            itens: input.notificacoes.map(n => ({
              motivo: n.motivo,
              setor: n.setor,
              medico: n.medico,
            })),
          });

          return { success: true, id: mysqlId };
        } catch (err: any) {
          if (err instanceof TRPCError) throw err;
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao registrar notifica\u00e7\u00e3o: ${err.message}`,
          });
        }
      }),
    registrarNotificacaoEmLote: reportProcedure("atendimentos")
      .input(
        z.object({
        estabelecimentoId: z.number().int().positive(),
          atendimentos: z.array(
            z.object({
          numatend: z.string(),
          nomepac: z.string(),
            })
          ),
        observacao: z.string(),
          notificacoes: z.array(
            z.object({
          motivo: z.string(),
          setor: z.string(),
          medico: z.string(),
            })
          ),
        })
      )
      .mutation(async ({ input, ctx }) => {
        try {
          const pertence = await verificarAtendimentosDoEstabelecimento(
            input.estabelecimentoId,
            input.atendimentos.map(atendimento => atendimento.numatend)
          );
          if (!pertence)
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "O lote contém atendimento de outra unidade.",
            });
          // Salvar no MySQL interno (fonte principal)
          const dadosLote = input.atendimentos.map(a => ({
            numatend: a.numatend,
            estabelecimentoId: input.estabelecimentoId,
            observacao: input.observacao,
            usuario: ctx.user?.name || "Usu\u00e1rio",
            itens: input.notificacoes.map(n => ({
              motivo: n.motivo,
              setor: n.setor,
              medico: n.medico,
            })),
          }));

          const resultadoMySQL =
            await salvarNotificacoesAtendimentoEmLote(dadosLote);

          return { success: true, ids: [], count: resultadoMySQL.salvos };
        } catch (err: any) {
          if (err instanceof TRPCError) throw err;
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao registrar notifica\u00e7\u00f5es em lote: ${err.message}`,
          });
        }
      }),

    // Salvar histórico de notificação com dados completos dos atendimentos
    salvarHistorico: reportProcedure("atendimentos")
      .input(
        z.object({
        estabelecimentoId: z.number().int().positive(),
        qtdAtendimentos: z.number(),
        observacao: z.string(),
          atendimentos: z.array(
            z.object({
          numatend: z.string(),
          nomepac: z.string(),
          nomeplaco: z.string(),
          datatend: z.string(),
          datasai: z.string().nullable(),
          diasParado: z.number(),
          tipoatendimentodescricao: z.string(),
          codserv: z.string(),
            })
          ),
          notificacoes: z.array(
            z.object({
          motivo: z.string(),
          setor: z.string(),
          medico: z.string(),
            })
          ),
        })
      )
      .mutation(async ({ input, ctx }) => {
        try {
          const pertence = await verificarAtendimentosDoEstabelecimento(
            input.estabelecimentoId,
            input.atendimentos.map(atendimento => atendimento.numatend)
          );
          if (!pertence)
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "O histórico contém atendimento de outra unidade.",
            });
          const id = await salvarHistoricoNotificacoesAtendimento({
            estabelecimentoId: input.estabelecimentoId,
            qtdAtendimentos: input.qtdAtendimentos,
            observacao: input.observacao,
            usuario: ctx.user?.name || "Usuário",
            atendimentos: input.atendimentos,
            notificacoes: input.notificacoes,
          });
          return { success: true, id };
        } catch (err: any) {
          if (err instanceof TRPCError) throw err;
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao salvar histórico: ${err.message}`,
          });
        }
      }),

    // Listar histórico de notificações
    listarHistorico: reportProcedure("atendimentos")
      .input(z.object({ estabelecimentoId: z.number().int().positive() }))
      .query(async ({ input }) => {
        try {
          return listarHistoricoNotificacoesAtendimento(
            input.estabelecimentoId
          );
        } catch (err: any) {
           console.error("[atendimentos.listarHistorico] Erro:", err.message);
          return [];
        }
      }),
    listarParadosUnificados: reportProcedure("atendimentos")
      .input(z.object({ estabelecimentoId: z.number().int().positive() }))
      .query(async ({ input }) => {
        try {
          const dados = await getAtendimentosParadosPorEstabelecimento(
            input.estabelecimentoId
          );
          return dados.map(d => ({
            ...d,
            diasParado: calcularDiasParadoUnificado(
              d.data_entrada,
              d.data_saida,
              d.origemSistema,
              d.dtEntrega,
              d.dtEtapa
            ),
          }));
        } catch (err: any) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao buscar atendimentos parados unificados: ${err.message}`,
          });
        }
      }),

    listarPaginado: reportProcedure("atendimentos")
      .input(
        z.object({
        estabelecimentoId: z.number().int().positive(),
        page: z.number().min(1).default(1),
        pageSize: z.number().min(10).max(200).default(50),
        paineisValor: z.boolean().default(false),
        origemSistema: z.string().optional(),
        tipo: z.string().optional(),
        convenio: z.string().optional(),
        etapa: z.string().optional(),
          mesEntrada: z
            .string()
            .regex(/^(?:\d{4}-(?:0[1-9]|1[0-2])|__SEM_ENTRADA__)$/)
            .optional(),
        setor: z.string().trim().min(1).max(255).optional(),
          permanencia: z
            .enum([
              "0-14 DIAS",
              "15-30 DIAS",
              "31-60 DIAS",
              "61-90 DIAS",
              "90+ DIAS",
              "SEM REFERÊNCIA",
            ])
            .optional(),
        protocolo: z.string().optional(),
          situacao: z
            .enum(["pendentes", "finalizadas", "todos"])
            .default("pendentes"),
        ano: z.string().optional(),
        mes: z.string().optional(),
        busca: z.string().optional(),
        descricao: z.string().optional(),
        sortColumn: z.string().default("data_entrada"),
        sortOrder: z.enum(["asc", "desc"]).default("desc"),
        })
      )
      .query(async ({ input }) => {
        try {
          const dbInstance = await getDb();
          if (!dbInstance) {
            return {
              items: [],
              total: 0,
              page: input.page,
              pageSize: input.pageSize,
              totalPages: 0,
              paineisValor: input.paineisValor
                ? PAINEIS_VALOR_VAZIOS
                : undefined,
              aggregations: {
                tipos: [],
                servicos: [],
                convenios: [],
                etapas: [],
                origens: [],
                protocolos: [],
                anos: [],
                totalValor: 0,
              },
            };
          }
          const { atendimentos: atendimentosUnificados } = await import(
            "../drizzle/schema-integracao"
          );
          const {
            eq,
            and,
            notLike,
            or,
            isNull,
            like,
            sql,
            desc: descOrd,
            asc: ascOrd,
            count,
          } = await import("drizzle-orm");
          const { normalizedCategorySql, normalizedServiceSql } =
            buildAtendimentoNormalizationSql({
            origemSistema: atendimentosUnificados.origemSistema,
            tipoAtendimento: atendimentosUnificados.tipo_atendimento,
              descricaoAtendimento:
                atendimentosUnificados.descricao_atendimento,
            dsCategoria: atendimentosUnificados.dsCategoria,
            codigoServico: atendimentosUnificados.codigo_servico,
            dsSetorEntrada: atendimentosUnificados.dsSetorEntrada,
          });
          const { diasParadoSql, referenciaPermanenciaSql } =
            buildAtendimentoDiasParadoSql({
            origemSistema: atendimentosUnificados.origemSistema,
            dataEntrada: atendimentosUnificados.data_entrada,
            dataSaida: atendimentosUnificados.data_saida,
            categoriaAtendimento: normalizedCategorySql,
            dtEntrega: atendimentosUnificados.dtEntrega,
            dtEtapa: atendimentosUnificados.dtEtapa,
          });
          const permanenciaSql = buildAtendimentoPermanenciaSql(
            referenciaPermanenciaSql
          );

          // Condições base (mesmas da procedure listar)
          const baseConditions: any[] = [];
          baseConditions.push(
            eq(
              atendimentosUnificados.estabelecimentoId,
              input.estabelecimentoId
            )
          );
          baseConditions.push(eq(atendimentosUnificados.ativoNaFonte, true));
          baseConditions.push(
            notLike(atendimentosUnificados.descricao_atendimento, "%A_FATURAR%")
          );
          baseConditions.push(
            or(
              sql`${atendimentosUnificados.origemSistema} NOT IN ('tasy', 'tasy_hemolabor')`,
              sql`UPPER(TRIM(COALESCE(${atendimentosUnificados.convenio}, ''))) != 'PARTICULAR'`
            )!
          );
          // Não filtrar por data_saida IS NULL - para WARLEINE/EASYVISION a query externa já traz apenas os atendimentos relevantes
          // Para TASY, a situação pendente/finalizada é escolhida explicitamente pelo usuário.

          // Condições de filtro
          const filterConditions: any[] = [...baseConditions];
          if (input.origemSistema && input.origemSistema !== "all") {
            filterConditions.push(
              eq(atendimentosUnificados.origemSistema, input.origemSistema)
            );
          }
          if (input.convenio) {
            filterConditions.push(
              sql`${ATENDIMENTOS_VALOR_SQL.convenio} = ${input.convenio}`
            );
          }
          if (input.etapa) {
            filterConditions.push(
              sql`${ATENDIMENTOS_VALOR_SQL.etapa} = ${input.etapa}`
            );
          }
          if (input.situacao === "pendentes") {
            filterConditions.push(
              or(
                sql`${atendimentosUnificados.origemSistema} NOT IN ('tasy', 'tasy_hemolabor')`,
                sql`${atendimentosUnificados.titulo} IS NULL OR TRIM(${atendimentosUnificados.titulo}) = ''`
              )!
            );
          } else if (input.situacao === "finalizadas") {
            filterConditions.push(
              sql`${atendimentosUnificados.origemSistema} IN ('tasy', 'tasy_hemolabor')
                  AND ${atendimentosUnificados.titulo} IS NOT NULL
                  AND TRIM(${atendimentosUnificados.titulo}) != ''`
            );
          }
          if (input.descricao) {
            filterConditions.push(
              sql`${normalizedServiceSql} = ${input.descricao}`
            );
          }
          if (input.protocolo && input.protocolo !== "all") {
            if (
              input.protocolo === "null" ||
              input.protocolo === "__sem_protocolo__"
            ) {
              filterConditions.push(
                or(
                  isNull(atendimentosUnificados.nomeProtocolo),
                  eq(atendimentosUnificados.nomeProtocolo, "")
                )!
              );
            } else if (input.protocolo === "com_protocolo") {
              filterConditions.push(
                sql`${atendimentosUnificados.nomeProtocolo} IS NOT NULL AND ${atendimentosUnificados.nomeProtocolo} != ''`
              );
            } else {
              filterConditions.push(
                eq(atendimentosUnificados.nomeProtocolo, input.protocolo)
              );
            }
          }
          if (input.ano || input.mes) {
            if (input.ano && input.mes) {
              const compVal = `${input.ano}/${input.mes}`;
              filterConditions.push(
                or(
                  eq(atendimentosUnificados.competencia, compVal),
                  and(
                    or(
                      isNull(atendimentosUnificados.competencia),
                      eq(atendimentosUnificados.competencia, ""),
                      eq(atendimentosUnificados.competencia, "NULL")
                    )!,
                    sql`YEAR(${atendimentosUnificados.data_entrada}) = ${parseInt(input.ano)}`,
                    sql`MONTH(${atendimentosUnificados.data_entrada}) = ${parseInt(input.mes)}`
                  )!
                )!
              );
            } else if (input.ano) {
              filterConditions.push(
                or(
                  like(atendimentosUnificados.competencia, `${input.ano}/%`),
                  and(
                    or(
                      isNull(atendimentosUnificados.competencia),
                      eq(atendimentosUnificados.competencia, ""),
                      eq(atendimentosUnificados.competencia, "NULL")
                    )!,
                    sql`YEAR(${atendimentosUnificados.data_entrada}) = ${parseInt(input.ano)}`
                  )!
                )!
              );
            } else if (input.mes) {
              filterConditions.push(
                or(
                  like(atendimentosUnificados.competencia, `%/${input.mes}`),
                  and(
                    or(
                      isNull(atendimentosUnificados.competencia),
                      eq(atendimentosUnificados.competencia, ""),
                      eq(atendimentosUnificados.competencia, "NULL")
                    )!,
                    sql`MONTH(${atendimentosUnificados.data_entrada}) = ${parseInt(input.mes)}`
                  )!
                )!
              );
            }
          }
          if (input.busca) {
            const term = `%${input.busca}%`;
            filterConditions.push(
              or(
                like(atendimentosUnificados.numero_atendimento, term),
                like(atendimentosUnificados.paciente, term),
                like(atendimentosUnificados.convenio, term),
                like(atendimentosUnificados.matricula, term),
                like(atendimentosUnificados.conta, term),
                like(atendimentosUnificados.medicoResp, term),
                like(atendimentosUnificados.descricao_atendimento, term),
                like(atendimentosUnificados.codigo_servico, term),
                like(atendimentosUnificados.etapaConta, term),
                like(atendimentosUnificados.setorEtapa, term),
                like(atendimentosUnificados.userEtapa, term)
              )!
            );
          }
          if (input.mesEntrada) {
            filterConditions.push(
              sql`${ATENDIMENTOS_VALOR_SQL.mesEntrada} = ${input.mesEntrada}`
            );
          }
          if (input.setor) {
            filterConditions.push(
              sql`${ATENDIMENTOS_VALOR_SQL.setor} = ${input.setor}`
            );
          }
          if (input.permanencia) {
            filterConditions.push(
              sql`${permanenciaSql} = ${input.permanencia}`
            );
          }

          const aggregationWhereClause = and(...filterConditions);
          if (input.tipo) {
            const normalizedType = input.tipo.toUpperCase();
            const canonicalType =
              normalizedType === "AMBULAT"
                ? "AMBULATORIO"
                : normalizedType === "PRONTO"
                  ? "PRONTO SOCORRO"
              : normalizedType;
            filterConditions.push(
              sql`${normalizedCategorySql} = ${canonicalType}`
            );
          }
          const whereClause = and(...filterConditions);
          const baseWhereClause = and(...baseConditions);

          // Mapeamento de colunas para ordenação
          const sortColumnMap: Record<string, any> = {
            data_entrada: atendimentosUnificados.data_entrada,
            paciente: atendimentosUnificados.paciente,
            convenio: atendimentosUnificados.convenio,
            tipo_atendimento: normalizedCategorySql,
            numero_atendimento: atendimentosUnificados.numero_atendimento,
            valorConta: atendimentosUnificados.valorConta,
            diasParado: diasParadoSql,
            etapaConta: atendimentosUnificados.etapaConta,
            setorEtapa: atendimentosUnificados.setorEtapa,
            origemSistema: atendimentosUnificados.origemSistema,
            descricao_atendimento: atendimentosUnificados.descricao_atendimento,
            codigo_servico: normalizedServiceSql,
            dtEntrega: atendimentosUnificados.dtEntrega,
            dtEtapa: atendimentosUnificados.dtEtapa,
            competencia: atendimentosUnificados.competencia,
            medicoResp: atendimentosUnificados.medicoResp,
            conta: atendimentosUnificados.conta,
            matricula: atendimentosUnificados.matricula,
          };
          const sortCol =
            sortColumnMap[input.sortColumn] ||
            atendimentosUnificados.data_entrada;
          const orderFn = input.sortOrder === "asc" ? ascOrd : descOrd;

          // Queries em paralelo
          const [
            totalResult,
            dataResult,
            valorResult,
            valoresInformadosResult,
          ] = await Promise.all([
            dbInstance
              .select({ count: count() })
              .from(atendimentosUnificados)
              .where(whereClause),
            dbInstance
              .select({
                ...getTableColumns(atendimentosUnificados),
                diasParadoCalculado: diasParadoSql.as("diasParadoCalculado"),
              })
              .from(atendimentosUnificados)
              .where(whereClause)
              .orderBy(orderFn(sortCol), orderFn(atendimentosUnificados.id))
              .limit(input.pageSize)
              .offset((input.page - 1) * input.pageSize),
            dbInstance.execute(
              sql`SELECT COALESCE(SUM(CAST(valorConta AS DECIMAL(18,2))), 0) as totalValor FROM atendimentos_unificados WHERE ${whereClause}`
            ),
            dbInstance.execute(
              sql`SELECT COUNT(valorConta) AS valoresInformados FROM atendimentos_unificados WHERE ${whereClause}`
            ),
          ]);

          const total = (totalResult[0] as any)?.count || 0;
          const totalPages = Math.ceil(total / input.pageSize);
          const totalValor = parseFloat(
            (valorResult as any)?.[0]?.[0]?.totalValor ||
              (valorResult as any)?.[0]?.totalValor ||
              "0"
          );
          const valoresInformados =
            parseInt(
              (valoresInformadosResult as any)?.[0]?.[0]?.valoresInformados ||
                (valoresInformadosResult as any)?.[0]?.valoresInformados ||
                "0",
              10
            ) || 0;
          const painelMetric = escolherMetricaPainel(valoresInformados);
          const paineisValor = input.paineisValor
            ? await getAtendimentosPaineisValor(dbInstance, whereClause!, {
                metric: painelMetric,
                expressions: { permanencia: permanenciaSql },
              })
            : undefined;

          // Agregações para filtros (sobre dados base, sem filtros de usuário)
          const [
            tiposR,
            servicosR,
            conveniosR,
            etapasR,
            origensR,
            protocolosR,
            anosR,
          ] = await Promise.all([
            dbInstance.execute(
              sql`SELECT ${normalizedCategorySql} as val, COUNT(*) as cnt FROM atendimentos_unificados WHERE ${aggregationWhereClause} GROUP BY val ORDER BY cnt DESC`
            ),
            dbInstance.execute(
              sql`SELECT ${normalizedServiceSql} as val, COUNT(*) as cnt FROM atendimentos_unificados WHERE ${aggregationWhereClause} GROUP BY val ORDER BY cnt DESC LIMIT 100`
            ),
            dbInstance.execute(
              sql`SELECT convenio as val, COUNT(*) as cnt FROM atendimentos_unificados WHERE ${baseWhereClause} AND convenio IS NOT NULL AND convenio != '' GROUP BY convenio ORDER BY cnt DESC LIMIT 50`
            ),
            dbInstance.execute(
              sql`SELECT etapaConta as val, COUNT(*) as cnt FROM atendimentos_unificados WHERE ${baseWhereClause} AND etapaConta IS NOT NULL AND etapaConta != '' GROUP BY etapaConta ORDER BY cnt DESC LIMIT 50`
            ),
            dbInstance.execute(
              sql`SELECT origemSistema as val, COUNT(*) as cnt FROM atendimentos_unificados WHERE ${baseWhereClause} GROUP BY origemSistema ORDER BY cnt DESC`
            ),
            dbInstance.execute(
              sql`SELECT nomeProtocolo as val, COUNT(*) as cnt FROM atendimentos_unificados WHERE ${baseWhereClause} AND nomeProtocolo IS NOT NULL AND nomeProtocolo != '' GROUP BY nomeProtocolo ORDER BY cnt DESC LIMIT 50`
            ),
            dbInstance.execute(
              sql`SELECT CASE WHEN competencia IS NOT NULL AND competencia != '' AND competencia != 'NULL' AND competencia LIKE '%/%' THEN SUBSTRING_INDEX(competencia, '/', 1) WHEN data_entrada IS NOT NULL THEN CAST(YEAR(data_entrada) AS CHAR) ELSE NULL END as ano, COUNT(*) as cnt FROM atendimentos_unificados WHERE ${baseWhereClause} GROUP BY ano HAVING ano IS NOT NULL ORDER BY ano DESC`
            ),
          ]);

          const parseAgg = (rows: any) => {
            const data =
              Array.isArray(rows) && Array.isArray(rows[0])
                ? rows[0]
                : Array.isArray(rows)
                  ? rows
                  : [];
            return data
              .map((r: any) => ({
                value: r.val || r.ano || "",
                count: parseInt(r.cnt) || 0,
              }))
              .filter((r: any) => r.value);
          };

          // Buscar motivos de notificação para os registros da página
          const numatends = dataResult
            .map(d => d.numero_atendimento || "")
            .filter(n => n !== "");
          let motivosMap: Record<string, string> = {};
          if (numatends.length > 0) {
            try {
              const notificacoesMySQL = await buscarNotificacoesAtendimento(
                numatends,
                input.estabelecimentoId
              );
              for (const [numatend, notif] of Object.entries(
                notificacoesMySQL
              )) {
                motivosMap[numatend] = notif.motivo;
              }
            } catch (err) {
              console.error(
                "[atendimentos.listarPaginado] Erro ao buscar motivos da unidade:",
                err
              );
            }
          }

          // Mapear dados
          const items = dataResult.map(d => {
            const categoriaAtendimento = normalizeAtendimentoCategory({
              origemSistema: d.origemSistema,
              tipoAtendimento: d.tipo_atendimento,
              descricaoAtendimento: d.descricao_atendimento,
              dsCategoria: d.dsCategoria,
            });
            const tipoServico = resolveAtendimentoServiceType({
              origemSistema: d.origemSistema,
              descricaoAtendimento: d.descricao_atendimento,
              codigoServico: d.codigo_servico,
              dsSetorEntrada: d.dsSetorEntrada,
            });
            const isWarleine =
              d.origemSistema?.trim().toLowerCase() === "warleine";
            const isTasy = ["tasy", "tasy_hemolabor"].includes(
              d.origemSistema?.trim().toLowerCase() || ""
            );
            const saidaEfetiva = isWarleine
              ? resolverSaidaEfetivaAtendimento({
                  dataEntrada: d.data_entrada
                    ? new Date(d.data_entrada).toISOString()
                    : null,
                  dataSaida: d.data_saida
                    ? new Date(d.data_saida).toISOString()
                    : null,
                  origemSistema: d.origemSistema,
                  categoriaAtendimento,
                })
              : null;
            return {
            id: d.id,
            origemSistema: d.origemSistema,
            origemId: d.origemId,
            estabelecimentoId: d.estabelecimentoId,
            numero_atendimento: d.numero_atendimento || "",
            codigo_saida: d.codigo_saida || "",
            convenio: d.convenio || "",
            paciente: d.paciente || "",
            caracter_atendimento: d.caracter_atendimento || "",
            data_entrada: d.data_entrada
                ? isTasy
                  ? serializeTasyCivilDate(d.data_entrada) || ""
                  : new Date(d.data_entrada).toISOString()
              : "",
            data_saida: d.data_saida
                ? isTasy
                  ? serializeTasyCivilDate(d.data_saida)
                  : new Date(d.data_saida).toISOString()
              : null,
            tipo_atendimento: categoriaAtendimento,
            categoriaAtendimento,
            tipoAtendimentoOrigem: d.tipo_atendimento || "",
            tipoServico,
            descricao_atendimento: d.descricao_atendimento || "",
            codigo_servico: d.codigo_servico || "",
            codigo_procedimento: d.codigo_procedimento || "",
            procedimentoDescricao: d.procedimentoDescricao || "",
            destino_conta: d.destino_conta || "",
            motivo: motivosMap[d.numero_atendimento || ""] || null,
            // Mesmo DATEDIFF de ordenação/filtros; não recalcular no relógio do Node.
            diasParado: Number(d.diasParadoCalculado ?? 0),
            ...(saidaEfetiva ? saidaEfetiva : {}),
            dsCategoria: d.dsCategoria || "",
            dsPlano: d.dsPlano || "",
            competencia: d.competencia || "",
            referencia: d.referencia || "",
            protTasy: d.protTasy || "",
            nomeProtocolo: d.nomeProtocolo || "",
            protConv: d.protConv || "",
              dtEntrega: d.dtEntrega
                ? new Date(d.dtEntrega).toISOString()
                : null,
            protStatus: d.protStatus || "",
            titulo: d.titulo || "",
            dtTitulo: d.dtTitulo ? new Date(d.dtTitulo).toISOString() : null,
              dataVencimento: d.dataVencimento
                ? new Date(d.dataVencimento).toISOString()
                : null,
            dsSetorEntrada: d.dsSetorEntrada || "",
            dsSetorLeito: d.dsSetorLeito || "",
            etapaConta: d.etapaConta || "",
            setorEtapa: d.setorEtapa || "",
            dtEtapa: d.dtEtapa ? new Date(d.dtEtapa).toISOString() : null,
            userEtapa: d.userEtapa || "",
            motivoDevolucao: d.motivoDevolucao || "",
            conta: d.conta || "",
            autorizacao: d.autorizacao || "",
              valorConta:
                d.valorConta === null || d.valorConta === undefined
                  ? ""
                  : String(d.valorConta),
            statusConta: d.statusConta || "",
            ativoNaFonte: d.ativoNaFonte,
            ultimoLoteId: d.ultimoLoteId,
              origemAtualizadaEm: d.origemAtualizadaEm
                ? new Date(d.origemAtualizadaEm).toISOString()
                : null,
            matricula: d.matricula || "",
            sexo: d.sexo || "",
            idade: d.idade || "",
            medicoResp: d.medicoResp || "",
            crm: d.crm || "",
            dsMotivoAlta: d.dsMotivoAlta || "",
            dataInicio: d.dataInicio || "",
            dataFim: d.dataFim || "",
            codServico: d.codServico || "",
            centroCusto: d.centroCusto || "",
              dataSincronizacao: d.dataSincronizacao
                ? new Date(d.dataSincronizacao).toISOString()
                : null,
            };
          });

          return {
            items,
            total,
            page: input.page,
            pageSize: input.pageSize,
            totalPages,
            paineisValor,
            aggregations: {
              tipos: parseAgg(tiposR),
              servicos: parseAgg(servicosR),
              convenios: parseAgg(conveniosR),
              etapas: parseAgg(etapasR),
              origens: parseAgg(origensR),
              protocolos: parseAgg(protocolosR),
              anos: parseAgg(anosR),
              totalValor,
              valoresInformados,
            },
          };
        } catch (err: any) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao buscar atendimentos paginados: ${err.message}`,
          });
        }
      }),

    getKPIs: adminProcedure.query(async ({ ctx }) => {
        try {
          const kpis = await getKPIsPorTipo();
          return kpis;
        } catch (err: any) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao buscar KPIs: ${err.message}`,
          });
        }
      }),

    getQuantidadePorPlano: adminProcedure.query(async ({ ctx }) => {
        try {
          const dados = await getQuantidadePorPlano();
          return dados;
        } catch (err: any) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao buscar quantidade por plano: ${err.message}`,
          });
        }
      }),

    getQuantidadePorServico: adminProcedure.query(async ({ ctx }) => {
        try {
          const dados = await getQuantidadePorServico();
          return dados;
        } catch (err: any) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao buscar quantidade por servico: ${err.message}`,
          });
        }
      }),

    // Verificar conexão SMTP
    verificarSMTP: adminProcedure.query(async () => {
        return await verificarConexaoSMTP();
      }),

    // Enviar notificação por e-mail
    enviarNotificacaoEmail: reportProcedure("atendimentos")
      .input(
        z.object({
        estabelecimentoId: z.number().int().positive(),
        destinatarioEmail: z.string().email(),
        estabelecimentoNome: z.string(),
        mensagemPersonalizada: z.string().optional(),
          atendimentos: z
            .array(
              z.object({
          paciente: z.string(),
          numatend: z.string(),
          tipoAtendimento: z.string(),
          plano: z.string(),
          diasParado: z.number(),
          dataEntrada: z.string(),
          observacao: z.string().optional(),
              })
            )
            .min(1)
            .max(200),
        })
      )
      .mutation(async ({ input, ctx }) => {
        try {
          const pertence = await verificarAtendimentosDoEstabelecimento(
            input.estabelecimentoId,
            input.atendimentos.map(atendimento => atendimento.numatend)
          );
          if (!pertence) {
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "O e-mail contém atendimento de outra unidade.",
            });
          }
          const html = gerarHtmlNotificacaoAtendimentos({
            estabelecimentoNome: input.estabelecimentoNome,
            totalAtendimentos: input.atendimentos.length,
            atendimentos: input.atendimentos.map(a => ({
              ...a,
              observacao: a.observacao || undefined,
            })),
            mensagemPersonalizada: input.mensagemPersonalizada,
          });

          const resultado = await enviarEmail({
            destinatario: input.destinatarioEmail,
            assunto: `[Portal Safatle] Notificação de ${input.atendimentos.length} Atendimento${input.atendimentos.length > 1 ? "s" : ""} Parado${input.atendimentos.length > 1 ? "s" : ""} - ${input.estabelecimentoNome}`,
            conteudoHtml: html,
            conteudoTexto: `Notificação de ${input.atendimentos.length} atendimentos parados no ${input.estabelecimentoNome}. Acesse o Portal Safatle para mais detalhes.`,
          });

          if (!resultado.ok) {
            throw new Error(
              resultado.error || "Erro desconhecido ao enviar e-mail"
            );
          }

          return { success: true, messageId: resultado.messageId };
        } catch (err: any) {
          if (err instanceof TRPCError) throw err;
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao enviar e-mail: ${err.message}`,
          });
        }
      }),
  }),

  // ===== ATENDIMENTOS A FATURAR (banco interno com fallback para PostgreSQL externo) =====
  atendimentosFaturar: router({
    listar: reportProcedure("atendimentosFaturar")
      .input(z.object({ estabelecimentoId: z.number().int().positive() }))
      .query(async ({ input }) => {
        try {
          const estabId = input.estabelecimentoId;
          // Buscar da tabela atendimentos_a_faturar (tabela própria, separada da unificados)
          const dbInstance = await getDb();
          if (dbInstance) {
            const { atendimentosAFaturar } = await import(
              "../drizzle/schema-integracao"
            );
            const { eq } = await import("drizzle-orm");
            const dadosInternos = await dbInstance
              .select()
              .from(atendimentosAFaturar)
              .where(eq(atendimentosAFaturar.estabelecimentoId, estabId));
            return dadosInternos.map((d: any) => ({
              numatend: d.numatend || "",
              nomeplaco: d.nomeplaco || "",
              nomepac: d.nomepac || "",
              carater: d.carater || "",
              datatend: d.datatend ? new Date(d.datatend).toISOString() : "",
              datasai: d.datasai ? new Date(d.datasai).toISOString() : null,
              tipoatend: d.tipoatend || "",
              tipoatendimentodescricao: d.tipoatendimentodescricao || "",
              codserv: d.codserv || "",
              procprin: d.procprin || "",
              diasParado: calcularDiasParado(
                d.datasai ? new Date(d.datasai).toISOString() : null
              ),
              origemSistema: d.origemSistema || "EASYVISION",
            }));
          }
          // Fallback: banco interno indisponível, retornar vazio
          return [];
        } catch (err: any) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message: `Erro ao buscar atendimentos a faturar: ${err.message}`,
          });
        }
      }),
  }),

  noticias: router({
    listar: publicProcedure.query(async () => {
      try {
        const feeds = [
          {
            url: "https://news.google.com/rss/search?q=faturamento+hospitalar+sa%C3%BAde&hl=pt-BR&gl=BR&ceid=BR:pt-419",
            categoria: "Faturamento Hospitalar",
          },
          {
            url: "https://news.google.com/rss/search?q=gest%C3%A3o+hospitalar+plano+sa%C3%BAde&hl=pt-BR&gl=BR&ceid=BR:pt-419",
            categoria: "Gestão Hospitalar",
          },
          {
            url: "https://news.google.com/rss/search?q=ANS+conv%C3%AAnio+hospital&hl=pt-BR&gl=BR&ceid=BR:pt-419",
            categoria: "Regulação ANS",
          },
        ];

        const allNoticias: Array<{
          titulo: string;
          link: string;
          fonte: string;
          dataPublicacao: string;
          categoria: string;
        }> = [];

        for (const feed of feeds) {
          try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 5000);
            const response = await fetch(feed.url, {
              headers: { "User-Agent": "Mozilla/5.0" },
              signal: controller.signal,
            });
            clearTimeout(timeoutId);
            if (!response.ok) continue;
            const xml = await response.text();

            // Parse RSS XML simples
            const items = xml.split("<item>").slice(1);
            for (const item of items.slice(0, 5)) {
              const titulo =
                item.match(/<title><!\[CDATA\[(.+?)\]\]><\/title>/)?.[1] ||
                item.match(/<title>(.+?)<\/title>/)?.[1] ||
                "";
              const link =
                item.match(/<link>(.+?)<\/link>/)?.[1] ||
                item.match(/<link\/>\s*(.+?)\s*</)?.[1] ||
                "";
              const fonte =
                item.match(/<source[^>]*>(.+?)<\/source>/)?.[1] ||
                item.match(
                  /<source[^>]*><!\[CDATA\[(.+?)\]\]><\/source>/
                )?.[1] ||
                "";
              const pubDate =
                item.match(/<pubDate>(.+?)<\/pubDate>/)?.[1] || "";

              if (titulo && link) {
                allNoticias.push({
                  titulo: titulo
                    .replace(/&amp;/g, "&")
                    .replace(/&lt;/g, "<")
                    .replace(/&gt;/g, ">")
                    .replace(/&quot;/g, '"'),
                  link: link.trim(),
                  fonte: fonte.replace(/&amp;/g, "&"),
                  dataPublicacao: pubDate,
                  categoria: feed.categoria,
                });
              }
            }
          } catch {
            // Ignora feeds que falharem
          }
        }

        // Ordenar por data mais recente e limitar a 12
        allNoticias.sort((a, b) => {
          const da = new Date(a.dataPublicacao).getTime() || 0;
          const db2 = new Date(b.dataPublicacao).getTime() || 0;
          return db2 - da;
        });

        return allNoticias.slice(0, 12);
      } catch (err: any) {
        // Retorna array vazio em vez de lançar erro para não bloquear a página
        console.error(`[Noticias] Erro ao buscar notícias: ${err.message}`);
        return [] as Array<{
          titulo: string;
          link: string;
          fonte: string;
          dataPublicacao: string;
          categoria: string;
        }>;
      }
    }),
  }),

  // ============ MOTOR DE REGRAS ============
  motorRegras: motorRegrasRouter,

  // ============ PADRÕES DE PROCEDIMENTOS ============
  padroesProcedimentos: padroesProcedimentosRouter,

  // ============ PADRÕES DE COBRANÇA POR CONVÊNIO ============
  padroesCobranca: padroesCobrancaRouter,

  // ============ INTEGRADOR DE DADOS ============
  integradorDados: integradorDadosRouter,

  // ============ EASYLOCAL: CONECTIVIDADE SOMENTE LEITURA ============
  easyLocal: easyLocalRouter,
  easyLocalIntegrador: easyLocalIntegradorRouter,
  easyLocalAgendamento: easyLocalAgendamentoRouter,
  atendimentosHopDiario: atendimentosHopDiarioRouter,

  // ============ CONTAS CONVÊNIO (OPERACIONAL) ============
  contasConvenio: contasConvenioRouter,
  contaContextoMedico: contaContextoMedicoRouter,

  // ============ FATURAMENTO UNIFICADO & CONCILIAÇÃO CRUZADA ============
  faturamentoUnificado: faturamentoUnificadoRouter,
  hupConciliacaoPiloto: hupConciliacaoPilotoRouter,

  // ============ CBHPM & TABELAS DE PORTE ============
  cbhpm: cbhpmRouter,

  // ============ AUDITORIA (Falhas de Prontuário, Ajustes, Aprendizado) ============
  auditoria: auditoriaRouter,

  // ============ CONFERÊNCIA PÓS-CORREÇÃO ============
  conferencia: conferenciaRouter,

  // ============ RECEBIMENTO GERAL ============
  recebimentoGeral: router({
    // Importar dados de integ_faturado_x_recebido para recebimento_geral
    importar: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number().optional(),
        limparAntes: z.boolean().optional().default(false),
        })
      )
      .mutation(async ({ ctx, input }) => {
        if (ctx.user.role !== "admin") {
          const hasPermission = await db.verificarPermissaoEstabelecimento(
            ctx.user.id,
            input.estabelecimentoId || 0,
            "gerenciar"
          );
          if (!hasPermission) {
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "Sem permissão para importar dados.",
            });
          }
        }
        return await dbRecebGeral.importarIntegFaturadoRecebido(
          input.estabelecimentoId,
          input.limparAntes || false
        );
      }),

    // Contar registros na recebimento_geral
    contar: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number().optional(),
        })
      )
      .query(async ({ input }) => {
        return {
          total: await dbRecebGeral.contarRecebimentoGeral(
            input.estabelecimentoId
          ),
        };
      }),

    // Listar registros com paginação e filtros
    listar: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        convenio: z.string().optional(),
        mesProducao: z.string().optional(),
        protocolo: z.string().optional(),
        numeroConta: z.string().optional(),
        receberHospital: z.string().optional(), // 'S' = Hospital, 'N' = Terceiros
        page: z.number().optional().default(1),
        pageSize: z.number().optional().default(50),
        })
      )
      .query(async ({ input }) => {
        return await dbRecebGeral.listarRecebimentoGeral(input);
      }),

    // Resumo/estatísticas
    resumo: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        receberHospital: z.string().optional(), // 'S' = Hospital, 'N' = Terceiros
        })
      )
      .query(async ({ input }) => {
        return await dbRecebGeral.resumoRecebimentoGeral(
          input.estabelecimentoId,
          input.receberHospital
        );
      }),

    // Dados agregados para Relatório BI de Recebimento
    dadosBI: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesProducao: z.string().optional(),
        convenio: z.string().optional(),
        setor: z.string().optional(),
        receberHospital: z.string().optional(), // 'S' = Hospital, 'N' = Terceiros
        })
      )
      .query(async ({ input }) => {
        return await dbRecebGeral.dadosRecebimentoBI(input);
      }),
  }),

  // ============ AVISOS INTERNOS ============
  avisosInternos: router({
    listarAtivos: publicProcedure.query(async () => {
      return await db.listarAvisosAtivos();
    }),

    listarTodos: protectedProcedure.query(async ({ ctx }) => {
      if (ctx.user.role !== "admin") {
        throw new TRPCError({
          code: "FORBIDDEN",
          message: "Apenas administradores podem gerenciar avisos.",
        });
      }
      return await db.listarAvisosInternos();
    }),

    criar: protectedProcedure
      .input(
        z.object({
        titulo: z.string().min(1, "Título é obrigatório"),
        conteudo: z.string().min(1, "Conteúdo é obrigatório"),
        tipo: z.enum(["informacao", "alerta", "urgente"]),
        expiraEm: z.string().nullable().optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        if (ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Apenas administradores podem criar avisos.",
          });
        }
        await db.criarAvisoInterno({
          titulo: input.titulo,
          conteudo: input.conteudo,
          tipo: input.tipo,
          criadoPorId: ctx.user.id,
          criadoPorNome: ctx.user.name || "Admin",
          expiraEm: input.expiraEm ? new Date(input.expiraEm) : null,
        });
        return { success: true };
      }),

    editar: protectedProcedure
      .input(
        z.object({
        id: z.number(),
        titulo: z.string().min(1).optional(),
        conteudo: z.string().min(1).optional(),
        tipo: z.enum(["informacao", "alerta", "urgente"]).optional(),
        ativo: z.enum(["sim", "nao"]).optional(),
        expiraEm: z.string().nullable().optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        if (ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Apenas administradores podem editar avisos.",
          });
        }
        const { id, ...data } = input;
        const updateData: any = { ...data };
        if (data.expiraEm !== undefined) {
          updateData.expiraEm = data.expiraEm ? new Date(data.expiraEm) : null;
        }
        await db.editarAvisoInterno(id, updateData);
        return { success: true };
      }),

    excluir: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ ctx, input }) => {
        if (ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Apenas administradores podem excluir avisos.",
          });
        }
        await db.excluirAvisoInterno(input.id);
        return { success: true };
      }),
  }),

  // ==================== MAPEAMENTO DE CONVÊNIOS ====================
  convenioMapeamento: router({
    // Listar convênios não mapeados para um estabelecimento
    naoMapeados: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return await dbConvMap.listarConveniosNaoMapeados(
          input.estabelecimentoId
        );
      }),

    // Listar convênios do Safatle (globais + do estabelecimento)
    conveniosSafatle: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number().optional() }).optional())
      .query(async ({ input }) => {
        return await dbConvMap.listarConveniosSafatle(input?.estabelecimentoId);
      }),

    // Sugerir mapeamentos automáticos por similaridade
    sugerir: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return await dbConvMap.sugerirMapeamentos(input.estabelecimentoId);
      }),

    // Salvar um mapeamento individual
    salvar: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        nomeOrigem: z.string().min(1),
        codigoOrigem: z.string().nullable().optional(),
        convenioId: z.number(),
          fonte: z
            .enum(["tasy", "integracao", "xml", "excel", "manual"])
            .optional(),
        metodoMatch: z.enum(["automatico", "manual"]).optional(),
        confianca: z.number().min(0).max(100).optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        return await dbConvMap.salvarMapeamento({
          ...input,
          criadoPor: ctx.user.id,
        });
      }),

    // Salvar múltiplos mapeamentos em batch
    salvarBatch: protectedProcedure
      .input(
        z.object({
          mapeamentos: z.array(
            z.object({
          estabelecimentoId: z.number(),
          nomeOrigem: z.string().min(1),
          codigoOrigem: z.string().nullable().optional(),
          convenioId: z.number(),
              fonte: z
                .enum(["tasy", "integracao", "xml", "excel", "manual"])
                .optional(),
          metodoMatch: z.enum(["automatico", "manual"]).optional(),
          confianca: z.number().min(0).max(100).optional(),
            })
          ),
        })
      )
      .mutation(async ({ ctx, input }) => {
        return await dbConvMap.salvarMapeamentosBatch(
          input.mapeamentos.map(m => ({ ...m, criadoPor: ctx.user.id }))
        );
      }),

    // Listar todos os mapeamentos de um estabelecimento
    listar: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return await dbConvMap.listarMapeamentos(input.estabelecimentoId);
      }),

    // Remover (desativar) um mapeamento
    remover: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ input }) => {
        await dbConvMap.removerMapeamento(input.id);
        return { success: true };
      }),

    // Aplicar mapeamentos automáticos com alta confiança
    aplicarAutomaticos: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        threshold: z.number().min(0).max(100).optional().default(80),
        })
      )
      .mutation(async ({ ctx, input }) => {
        return await dbConvMap.aplicarMapeamentosAutomaticos(
          input.estabelecimentoId,
          input.threshold,
          ctx.user.id
        );
      }),

    // Resolver convenioId para um nome de convênio do hospital
    resolver: protectedProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        nomeConvenio: z.string(),
        })
      )
      .query(async ({ input }) => {
        const convenioId = await dbConvMap.resolverConvenioId(
          input.estabelecimentoId,
          input.nomeConvenio
        );
        return { convenioId };
      }),

    // Estatísticas de mapeamento
    estatisticas: protectedProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        return await dbConvMap.estatisticasMapeamento(input.estabelecimentoId);
      }),
  }),
  relatorioLaboratorio: router({
    // Lista de convênios disponíveis
    convenios: relatorioLaboratorioProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          const [rows] = await conn.execute(
            `SELECT DISTINCT d.convenio_id, c.nome as convenio_nome
             FROM demonstrativo d
             LEFT JOIN convenios c ON c.id = d.convenio_id
             WHERE d.estabelecimentoId = ?
             AND (REPLACE(d.codigo_item,'.','') LIKE '402%' OR REPLACE(d.codigo_item,'.','') LIKE '403%')
             AND d.convenio_id IS NOT NULL
             ORDER BY c.nome`,
            [input.estabelecimentoId]
          );
          return (rows as any[]).map((r: any) => ({
            id: r.convenio_id,
            nome: r.convenio_nome || `Convênio ${r.convenio_id}`,
          }));
        } finally {
          if (conn) await conn.end().catch(() => {});
        }
      }),

    // Dados do relatório
    dados: relatorioLaboratorioProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesReferencia: z.number().min(1).max(12).optional(),
        anoReferencia: z.number().min(2000).max(2100).optional(),
        convenioId: z.number().optional(),
          statusFiltro: z
            .enum(["todos", "pago", "glosado", "parcial"])
            .optional(),
        })
      )
      .query(async ({ input }) => {
        console.log(
          "[relatorioLaboratorio.dados] Input:",
          JSON.stringify(input)
        );
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);

          // Filtro base: faixa TUSS de laboratório (402xx e 403xx)
          let whereExtra = "";
          const params: any[] = [];
          if (input.mesReferencia && input.anoReferencia) {
            whereExtra += ` AND MONTH(data_referencia) = ? AND YEAR(data_referencia) = ?`;
            params.push(input.mesReferencia, input.anoReferencia);
          }
          if (input.convenioId) {
            whereExtra += ` AND d.convenio_id = ?`;
            params.push(input.convenioId);
          }
          if (input.statusFiltro && input.statusFiltro !== "todos") {
            if (input.statusFiltro === "pago")
              whereExtra += ` AND UPPER(d.situacao_item) = 'PAGO'`;
            else if (input.statusFiltro === "glosado")
              whereExtra += ` AND UPPER(d.situacao_item) = 'GLOSADO'`;
            else if (input.statusFiltro === "parcial")
              whereExtra += ` AND UPPER(d.situacao_item) = 'PARCIAL'`;
          }

          const baseWhere = `d.estabelecimentoId = ? AND (REPLACE(d.codigo_item,'.','') LIKE '402%' OR REPLACE(d.codigo_item,'.','') LIKE '403%')`;
          const baseParams = [input.estabelecimentoId];

          // Queries em paralelo
          const [
            [porCodigo],
            [resumoRows],
            [competencias],
            [porSituacao],
            [porConvenio],
          ] = await Promise.all([
            conn.execute(
              `SELECT d.codigo_item, MIN(d.descricao_item) as descricao, d.situacao_item,
                      COUNT(*) as qtd,
                      SUM(CAST(d.valor_pago AS DECIMAL(12,2))) as total_pago,
                      SUM(CAST(d.valor_glosa AS DECIMAL(12,2))) as total_glosa
               FROM demonstrativo d WHERE ${baseWhere} ${whereExtra}
               GROUP BY d.codigo_item, d.situacao_item ORDER BY total_pago DESC`,
              [...baseParams, ...params]
            ),
            conn.execute(
              `SELECT COUNT(*) as total_itens,
                      COUNT(DISTINCT d.codigo_item) as codigos_unicos,
                      SUM(CAST(d.valor_pago AS DECIMAL(12,2))) as total_pago,
                      SUM(CAST(d.valor_glosa AS DECIMAL(12,2))) as total_glosa,
                      SUM(CAST(d.valor_informado AS DECIMAL(12,2))) as total_informado
               FROM demonstrativo d WHERE ${baseWhere} ${whereExtra}`,
              [...baseParams, ...params]
            ),
            conn.execute(
              `SELECT DISTINCT DATE_FORMAT(d.data_referencia, '%m-%Y') as competencia,
                      MONTH(d.data_referencia) as mes, YEAR(d.data_referencia) as ano,
                      COUNT(*) as total
               FROM demonstrativo d WHERE ${baseWhere}
               GROUP BY DATE_FORMAT(d.data_referencia, '%m-%Y'), MONTH(d.data_referencia), YEAR(d.data_referencia)
               ORDER BY ano DESC, mes DESC`,
              [...baseParams]
            ),
            conn.execute(
              `SELECT d.situacao_item, COUNT(*) as total,
                      SUM(CAST(d.valor_pago AS DECIMAL(12,2))) as vl_pago,
                      SUM(CAST(d.valor_glosa AS DECIMAL(12,2))) as vl_glosa
               FROM demonstrativo d WHERE ${baseWhere} ${whereExtra}
               GROUP BY d.situacao_item`,
              [...baseParams, ...params]
            ),
            conn.execute(
              `SELECT d.convenio_id, c.nome as convenio_nome,
                      COUNT(*) as qtd,
                      COUNT(DISTINCT d.codigo_item) as codigos,
                      SUM(CAST(d.valor_pago AS DECIMAL(12,2))) as total_pago,
                      SUM(CAST(d.valor_glosa AS DECIMAL(12,2))) as total_glosa
               FROM demonstrativo d
               LEFT JOIN convenios c ON c.id = d.convenio_id
               WHERE ${baseWhere} ${whereExtra}
               GROUP BY d.convenio_id, c.nome
               ORDER BY total_pago DESC`,
              [...baseParams, ...params]
            ),
          ]);

          // Itens individuais apenas com filtro de competência
          let itens: any[] = [];
          if (input.mesReferencia && input.anoReferencia) {
            const [itensResult] = await conn.execute(
              `SELECT d.id, d.codigo_item, d.descricao_item, d.situacao_item,
                      CAST(d.valor_pago AS DECIMAL(12,2)) as valor_pago,
                      CAST(d.valor_glosa AS DECIMAL(12,2)) as valor_glosa,
                      CAST(d.valor_informado AS DECIMAL(12,2)) as valor_informado,
                      d.numero_guia, d.nome_beneficiario, d.data_execucao, d.data_referencia, d.data_pagamento,
                      d.codigo_glosa, d.convenio_id, c.nome as convenio_nome
               FROM demonstrativo d
               LEFT JOIN convenios c ON c.id = d.convenio_id
               WHERE ${baseWhere} ${whereExtra}
               ORDER BY d.codigo_item, d.data_referencia DESC`,
              [...baseParams, ...params]
            );
            itens = itensResult as any[];
          }

          const r = (resumoRows as any[])[0] || {};
          console.log(
            "[relatorioLaboratorio.dados] Resumo:",
            JSON.stringify(r)
          );

          return {
            itens,
            porCodigo: porCodigo as any[],
            porSituacao: porSituacao as any[],
            porConvenio: (porConvenio as any[]).map((c: any) => ({
              convenioId: c.convenio_id,
              convenioNome: c.convenio_nome || `Convênio ${c.convenio_id}`,
              qtd: Number(c.qtd),
              codigos: Number(c.codigos),
              totalPago: Number(c.total_pago || 0),
              totalGlosa: Number(c.total_glosa || 0),
            })),
            resumo: {
              totalItens: Number(r.total_itens || 0),
              totalPago: Number(r.total_pago || 0),
              totalGlosa: Number(r.total_glosa || 0),
              totalInformado: Number(r.total_informado || 0),
              codigosUnicos: Number(r.codigos_unicos || 0),
            },
            competencias: (competencias as any[]).map((c: any) => ({
              value: `${c.mes}-${c.ano}`,
              label: `${String(c.mes).padStart(2, "0")}/${c.ano}`,
              total: Number(c.total),
              mes: Number(c.mes),
              ano: Number(c.ano),
            })),
          };
        } catch (err) {
          console.error("[relatorioLaboratorio.dados] ERRO:", err);
          throw err;
    } finally {
          if (conn) await conn.end().catch(() => {});
        }
      }),
  }),

  // ============ RELATÓRIO VISITA (Pronto Socorro) ============
  relatorioVisitaLegacy: router({
    convenios: relatorioLaboratorioProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input, ctx }) => {
        let conn: any = null;
        try {
          const permitido = await db.verificarPermissaoEstabelecimento(
            ctx.user.id,
            input.estabelecimentoId,
            "visualizar"
          );
          if (!permitido)
            throw new TRPCError({
              code: "FORBIDDEN",
              message: "Sem permissão para visualizar este estabelecimento",
            });
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(db.getApplicationDatabaseUrl()!);
          const [rows] = await conn.execute(
            `SELECT DISTINCT d.convenio_id, c.nome as convenio_nome
             FROM demonstrativo d
             LEFT JOIN convenios c ON c.id = d.convenio_id
             WHERE d.estabelecimentoId = ?
             AND (d.codigo_item = '10102019' OR REPLACE(d.codigo_item,'.','') LIKE '10102%')
             AND d.convenio_id IS NOT NULL
             ORDER BY c.nome`,
            [input.estabelecimentoId]
          );
          return (rows as any[]).map((r: any) => ({
            id: r.convenio_id,
            nome: r.convenio_nome || `Convênio ${r.convenio_id}`,
          }));
        } finally {
          if (conn) await conn.end().catch(() => {});
        }
      }),

    dados: relatorioLaboratorioProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesReferencia: z.number().min(1).max(12).optional(),
        anoReferencia: z.number().min(2000).max(2100).optional(),
        convenioId: z.number().optional(),
          statusFiltro: z
            .enum(["todos", "pago", "glosado", "parcial"])
            .optional(),
        })
      )
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          let whereExtra = "";
          const params: any[] = [];
          if (input.mesReferencia && input.anoReferencia) {
            whereExtra += ` AND MONTH(data_referencia) = ? AND YEAR(data_referencia) = ?`;
            params.push(input.mesReferencia, input.anoReferencia);
          }
          if (input.convenioId) {
            whereExtra += ` AND d.convenio_id = ?`;
            params.push(input.convenioId);
          }
          if (input.statusFiltro && input.statusFiltro !== "todos") {
            if (input.statusFiltro === "pago")
              whereExtra += ` AND UPPER(d.situacao_item) = 'PAGO'`;
            else if (input.statusFiltro === "glosado")
              whereExtra += ` AND UPPER(d.situacao_item) = 'GLOSADO'`;
            else if (input.statusFiltro === "parcial")
              whereExtra += ` AND UPPER(d.situacao_item) = 'PARCIAL'`;
          }

          const baseWhere = `d.estabelecimentoId = ? AND (d.codigo_item = '10102019' OR REPLACE(d.codigo_item,'.','') LIKE '10102%')`;
          const baseParams = [input.estabelecimentoId];

          const [
            [porCodigo],
            [resumoRows],
            [competencias],
            [porSituacao],
            [porConvenio],
          ] = await Promise.all([
            conn.execute(
              `SELECT d.codigo_item, MIN(d.descricao_item) as descricao, d.situacao_item,
                      COUNT(*) as qtd, SUM(CAST(d.valor_pago AS DECIMAL(12,2))) as total_pago, SUM(CAST(d.valor_glosa AS DECIMAL(12,2))) as total_glosa
               FROM demonstrativo d WHERE ${baseWhere} ${whereExtra} GROUP BY d.codigo_item, d.situacao_item ORDER BY total_pago DESC`,
              [...baseParams, ...params]
            ),
            conn.execute(
              `SELECT COUNT(*) as total_itens, COUNT(DISTINCT d.codigo_item) as codigos_unicos,
                      SUM(CAST(d.valor_pago AS DECIMAL(12,2))) as total_pago, SUM(CAST(d.valor_glosa AS DECIMAL(12,2))) as total_glosa, SUM(CAST(d.valor_informado AS DECIMAL(12,2))) as total_informado
               FROM demonstrativo d WHERE ${baseWhere} ${whereExtra}`,
              [...baseParams, ...params]
            ),
            conn.execute(
              `SELECT DISTINCT DATE_FORMAT(d.data_referencia, '%m-%Y') as competencia, MONTH(d.data_referencia) as mes, YEAR(d.data_referencia) as ano, COUNT(*) as total
               FROM demonstrativo d WHERE ${baseWhere} GROUP BY DATE_FORMAT(d.data_referencia, '%m-%Y'), MONTH(d.data_referencia), YEAR(d.data_referencia) ORDER BY ano DESC, mes DESC`,
              [...baseParams]
            ),
            conn.execute(
              `SELECT d.situacao_item, COUNT(*) as total, SUM(CAST(d.valor_pago AS DECIMAL(12,2))) as vl_pago, SUM(CAST(d.valor_glosa AS DECIMAL(12,2))) as vl_glosa
               FROM demonstrativo d WHERE ${baseWhere} ${whereExtra} GROUP BY d.situacao_item`,
              [...baseParams, ...params]
            ),
            conn.execute(
              `SELECT d.convenio_id, c.nome as convenio_nome, COUNT(*) as qtd, COUNT(DISTINCT d.codigo_item) as codigos, SUM(CAST(d.valor_pago AS DECIMAL(12,2))) as total_pago, SUM(CAST(d.valor_glosa AS DECIMAL(12,2))) as total_glosa
               FROM demonstrativo d LEFT JOIN convenios c ON c.id = d.convenio_id WHERE ${baseWhere} ${whereExtra} GROUP BY d.convenio_id, c.nome ORDER BY total_pago DESC`,
              [...baseParams, ...params]
            ),
          ]);

          let itens: any[] = [];
          if (input.mesReferencia && input.anoReferencia) {
            const [itensResult] = await conn.execute(
              `SELECT d.id, d.codigo_item, d.descricao_item, d.situacao_item, CAST(d.valor_pago AS DECIMAL(12,2)) as valor_pago, CAST(d.valor_glosa AS DECIMAL(12,2)) as valor_glosa, CAST(d.valor_informado AS DECIMAL(12,2)) as valor_informado, d.numero_guia, d.nome_beneficiario, d.data_execucao, d.data_referencia, d.data_pagamento, d.codigo_glosa, d.convenio_id, c.nome as convenio_nome
               FROM demonstrativo d LEFT JOIN convenios c ON c.id = d.convenio_id WHERE ${baseWhere} ${whereExtra} ORDER BY d.codigo_item, d.data_referencia DESC`,
              [...baseParams, ...params]
            );
            itens = itensResult as any[];
          }

          const r = (resumoRows as any[])[0] || {};
          return {
            itens,
            porCodigo: porCodigo as any[],
            porSituacao: porSituacao as any[],
            porConvenio: (porConvenio as any[]).map((c: any) => ({
              convenioId: c.convenio_id,
              convenioNome: c.convenio_nome || `Convênio ${c.convenio_id}`,
              qtd: Number(c.qtd),
              codigos: Number(c.codigos),
              totalPago: Number(c.total_pago || 0),
              totalGlosa: Number(c.total_glosa || 0),
            })),
            resumo: {
              totalItens: Number(r.total_itens || 0),
              totalPago: Number(r.total_pago || 0),
              totalGlosa: Number(r.total_glosa || 0),
              totalInformado: Number(r.total_informado || 0),
              codigosUnicos: Number(r.codigos_unicos || 0),
            },
            competencias: (competencias as any[]).map((c: any) => ({
              value: `${c.mes}-${c.ano}`,
              label: `${String(c.mes).padStart(2, "0")}/${c.ano}`,
              total: Number(c.total),
              mes: Number(c.mes),
              ano: Number(c.ano),
            })),
          };
        } finally {
          if (conn) await conn.end().catch(() => {});
        }
      }),
  }),

  // ============ RELATÓRIO VISITA XML (Faturamento/Enviado) ============
  relatorioVisitaXml: router({
    convenios: relatorioVisitaXmlProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          let [rows]: any = await conn.execute(
            `SELECT DISTINCT s.convenioId, c.nome as convenio_nome
             FROM staging_faturamento_xml s
             LEFT JOIN convenios c ON c.id = s.convenioId
             WHERE (s.estabelecimentoId = ? OR s.estabelecimento_id = ?)
             AND REPLACE(s.codigo_item, '.', '') = '10102019'
             AND s.convenioId IS NOT NULL
             ORDER BY c.nome`,
            [input.estabelecimentoId, input.estabelecimentoId]
          );
          if (!rows || rows.length === 0) {
            [rows] = await conn.execute(
              `SELECT DISTINCT ft.convenioId, c.nome as convenio_nome
               FROM faturamento_tiss ft
               LEFT JOIN convenios c ON c.id = ft.convenioId
               WHERE ft.estabelecimentoId = ?
               AND REPLACE(ft.codigo_item, '.', '') = '10102019'
               AND ft.convenioId IS NOT NULL
               ORDER BY c.nome`,
              [input.estabelecimentoId]
            );
          }
          return (rows as any[]).map((r: any) => ({
            id: r.convenioId,
            nome: r.convenio_nome || `Convênio ${r.convenioId}`,
          }));
        } finally {
          if (conn) await conn.end().catch(() => {});
        }
      }),

    competencias: relatorioVisitaXmlProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          let [rows]: any = await conn.execute(
            `SELECT competencia, COUNT(*) as total
             FROM staging_faturamento_xml
             WHERE (estabelecimentoId = ? OR estabelecimento_id = ?)
             AND REPLACE(codigo_item, '.', '') = '10102019'
             AND competencia IS NOT NULL
             GROUP BY competencia
             ORDER BY competencia DESC`,
            [input.estabelecimentoId, input.estabelecimentoId]
          );
          if (!rows || rows.length === 0) {
            [rows] = await conn.execute(
              `SELECT competencia, COUNT(*) as total
               FROM faturamento_tiss
               WHERE estabelecimentoId = ?
               AND REPLACE(codigo_item, '.', '') = '10102019'
               AND competencia IS NOT NULL
               GROUP BY competencia
               ORDER BY competencia DESC`,
              [input.estabelecimentoId]
            );
          }
          return (rows as any[]).map((r: any) => {
            const parts = (r.competencia || "").split("/");
            return {
              value: r.competencia,
              label: r.competencia,
              total: Number(r.total),
              ano: parts[0] ? Number(parts[0]) : 0,
              mes: parts[1] ? Number(parts[1]) : 0,
            };
          });
        } finally {
          if (conn) await conn.end().catch(() => {});
        }
      }),

    dados: relatorioVisitaXmlProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        competencia: z.string().optional(),
        convenioId: z.number().optional(),
        medico: z.string().optional(),
        })
      )
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);

          let whereExtra = "";
          const params: any[] = [
            input.estabelecimentoId,
            input.estabelecimentoId,
          ];

          if (input.competencia) {
            whereExtra += ` AND s.competencia = ?`;
            params.push(input.competencia);
          }
          if (input.convenioId) {
            whereExtra += ` AND s.convenioId = ?`;
            params.push(input.convenioId);
          }
          if (input.medico) {
            whereExtra += ` AND s.nome_prof = ?`;
            params.push(input.medico);
          }

          const baseWhere = `(s.estabelecimentoId = ? OR s.estabelecimento_id = ?) AND REPLACE(s.codigo_item, '.', '') = '10102019'`;

          // Queries em paralelo
          let [[resumoRows], [porMedico], [porConvenio], [itens]]: any =
            await Promise.all([
            // Resumo geral
            conn.execute(
              `SELECT COUNT(*) as total_visitas,
                      COUNT(DISTINCT s.numero_guia_prestador) as total_guias,
                      COUNT(DISTINCT s.nome_prof) as total_medicos,
                      SUM(CAST(s.valor_faturado AS DECIMAL(12,2))) as total_faturado
               FROM staging_faturamento_xml s WHERE ${baseWhere} ${whereExtra}`,
              [...params]
            ),
            // Por médico
            conn.execute(
              `SELECT s.nome_prof as medico, s.conselho_prof as crm,
                      COUNT(*) as qtd_visitas,
                      COUNT(DISTINCT s.numero_guia_prestador) as qtd_guias,
                      SUM(CAST(s.valor_faturado AS DECIMAL(12,2))) as total_faturado
               FROM staging_faturamento_xml s WHERE ${baseWhere} ${whereExtra}
               AND s.nome_prof IS NOT NULL AND s.nome_prof != ''
               GROUP BY s.nome_prof, s.conselho_prof ORDER BY qtd_visitas DESC`,
              [...params]
            ),
            // Por convênio
            conn.execute(
              `SELECT s.convenioId, c.nome as convenio_nome,
                      COUNT(*) as qtd_visitas,
                      COUNT(DISTINCT s.numero_guia_prestador) as qtd_guias,
                      SUM(CAST(s.valor_faturado AS DECIMAL(12,2))) as total_faturado
               FROM staging_faturamento_xml s
               LEFT JOIN convenios c ON c.id = s.convenioId
               WHERE ${baseWhere} ${whereExtra}
               GROUP BY s.convenioId, c.nome ORDER BY total_faturado DESC`,
              [...params]
            ),
            // Itens individuais
            conn.execute(
              `SELECT s.id, s.numero_guia_prestador as guia, s.descricao_item as procedimento,
                      CAST(s.valor_faturado AS DECIMAL(12,2)) as valor,
                      s.nome_prof as medico, s.conselho_prof as crm,
                      s.data_execucao as data_execucao,
                      s.competencia,
                      s.convenioId, c.nome as convenio_nome,
                      s.carteira_beneficiario
               FROM staging_faturamento_xml s
               LEFT JOIN convenios c ON c.id = s.convenioId
               WHERE ${baseWhere} ${whereExtra}
               ORDER BY s.data_execucao DESC, s.nome_prof`,
              [...params]
            ),
          ]);

          if (!resumoRows || Number(resumoRows[0]?.total_visitas || 0) === 0) {
            let whereExtraFt = "";
            const paramsFt: any[] = [input.estabelecimentoId];
            if (input.competencia) {
              whereExtraFt += ` AND s.competencia = ?`;
              paramsFt.push(input.competencia);
            }
            if (input.convenioId) {
              whereExtraFt += ` AND s.convenioId = ?`;
              paramsFt.push(input.convenioId);
            }
            if (input.medico) {
              whereExtraFt += ` AND s.nome_prof = ?`;
              paramsFt.push(input.medico);
            }
            const baseWhereFt = `s.estabelecimentoId = ? AND REPLACE(s.codigo_item, '.', '') = '10102019'`;
            [[resumoRows], [porMedico], [porConvenio], [itens]] = await Promise.all([
              conn.execute(
                `SELECT COUNT(*) as total_visitas,
                        COUNT(DISTINCT s.numero_guia_prestador) as total_guias,
                        COUNT(DISTINCT s.nome_prof) as total_medicos,
                        SUM(CAST(s.valor_faturado AS DECIMAL(12,2))) as total_faturado
                 FROM faturamento_tiss s WHERE ${baseWhereFt} ${whereExtraFt}`,
                [...paramsFt]
              ),
              conn.execute(
                `SELECT s.nome_prof as medico, s.conselho_prof as crm,
                        COUNT(*) as qtd_visitas,
                        COUNT(DISTINCT s.numero_guia_prestador) as qtd_guias,
                        SUM(CAST(s.valor_faturado AS DECIMAL(12,2))) as total_faturado
                 FROM faturamento_tiss s WHERE ${baseWhereFt} ${whereExtraFt}
                 AND s.nome_prof IS NOT NULL AND s.nome_prof != ''
                 GROUP BY s.nome_prof, s.conselho_prof ORDER BY qtd_visitas DESC`,
                [...paramsFt]
              ),
              conn.execute(
                `SELECT s.convenioId, c.nome as convenio_nome,
                        COUNT(*) as qtd_visitas,
                        COUNT(DISTINCT s.numero_guia_prestador) as qtd_guias,
                        SUM(CAST(s.valor_faturado AS DECIMAL(12,2))) as total_faturado
                 FROM faturamento_tiss s
                 LEFT JOIN convenios c ON c.id = s.convenioId
                 WHERE ${baseWhereFt} ${whereExtraFt}
                 GROUP BY s.convenioId, c.nome ORDER BY total_faturado DESC`,
                [...paramsFt]
              ),
              conn.execute(
                `SELECT s.id, s.numero_guia_prestador as guia, s.descricao_item as procedimento,
                        CAST(s.valor_faturado AS DECIMAL(12,2)) as valor,
                        s.nome_prof as medico, s.conselho_prof as crm,
                        s.data_execucao as data_execucao,
                        s.competencia,
                        s.convenioId, c.nome as convenio_nome,
                        s.carteira_beneficiario
                 FROM faturamento_tiss s
                 LEFT JOIN convenios c ON c.id = s.convenioId
                 WHERE ${baseWhereFt} ${whereExtraFt}
                 ORDER BY s.data_execucao DESC, s.nome_prof`,
                [...paramsFt]
              ),
            ]);
          }

          const r = (resumoRows as any[])[0] || {};

          return {
            resumo: {
              totalVisitas: Number(r.total_visitas || 0),
              totalGuias: Number(r.total_guias || 0),
              totalMedicos: Number(r.total_medicos || 0),
              totalFaturado: Number(r.total_faturado || 0),
            },
            porMedico: (porMedico as any[]).map((m: any) => ({
              medico: m.medico,
              crm: m.crm,
              qtdVisitas: Number(m.qtd_visitas),
              qtdGuias: Number(m.qtd_guias),
              totalFaturado: Number(m.total_faturado || 0),
            })),
            porConvenio: (porConvenio as any[]).map((c: any) => ({
              convenioId: c.convenioId,
              convenioNome: c.convenio_nome || `Convênio ${c.convenioId}`,
              qtdVisitas: Number(c.qtd_visitas),
              qtdGuias: Number(c.qtd_guias),
              totalFaturado: Number(c.total_faturado || 0),
            })),
            itens: (itens as any[]).map((i: any) => ({
              id: i.id,
              guia: i.guia,
              procedimento: i.procedimento,
              valor: Number(i.valor || 0),
              medico: i.medico,
              crm: i.crm,
              dataExecucao: i.data_execucao,
              competencia: i.competencia,
              convenioId: i.convenioId,
              convenioNome: i.convenio_nome,
              carteira: i.carteira_beneficiario,
            })),
          };
        } catch (err) {
          console.error("[relatorioVisitaXml.dados] ERRO:", err);
          throw err;
        } finally {
          if (conn) await conn.end().catch(() => {});
        }
      }),
  }),

  // ============ RELATÓRIO ULTRASSOM (Pronto Socorro) ============
  relatorioUltrassom: router({
    convenios: relatorioUltrassomProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          const [rows] = await conn.execute(
            `SELECT DISTINCT d.convenio_id, c.nome as convenio_nome
             FROM demonstrativo d
             LEFT JOIN convenios c ON c.id = d.convenio_id
             WHERE d.estabelecimentoId = ?
             AND d.codigo_item IN ('40901122', '40901351', '40901769', '40901033', '40901114', '40901386', '40901203', '40901211', '40901181')
             AND d.convenio_id IS NOT NULL
             ORDER BY c.nome`,
            [input.estabelecimentoId]
          );
          return (rows as any[]).map((r: any) => ({
            id: r.convenio_id,
            nome: r.convenio_nome || `Convênio ${r.convenio_id}`,
          }));
        } finally {
          if (conn) await conn.end().catch(() => {});
        }
      }),

    dados: relatorioUltrassomProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesReferencia: z.number().min(1).max(12).optional(),
        anoReferencia: z.number().min(2000).max(2100).optional(),
        convenioId: z.number().optional(),
          statusFiltro: z
            .enum(["todos", "pago", "glosado", "parcial"])
            .optional(),
        })
      )
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          let whereExtra = "";
          const params: any[] = [];
          if (input.mesReferencia && input.anoReferencia) {
            whereExtra += ` AND MONTH(data_referencia) = ? AND YEAR(data_referencia) = ?`;
            params.push(input.mesReferencia, input.anoReferencia);
          }
          if (input.convenioId) {
            whereExtra += ` AND d.convenio_id = ?`;
            params.push(input.convenioId);
          }
          if (input.statusFiltro && input.statusFiltro !== "todos") {
            if (input.statusFiltro === "pago")
              whereExtra += ` AND UPPER(d.situacao_item) = 'PAGO'`;
            else if (input.statusFiltro === "glosado")
              whereExtra += ` AND UPPER(d.situacao_item) = 'GLOSADO'`;
            else if (input.statusFiltro === "parcial")
              whereExtra += ` AND UPPER(d.situacao_item) = 'PARCIAL'`;
          }

          const baseWhere = `d.estabelecimentoId = ? AND d.codigo_item IN ('40901122', '40901351', '40901769', '40901033', '40901114', '40901386', '40901203', '40901211', '40901181')`;
          const baseParams = [input.estabelecimentoId];

          const [
            [porCodigo],
            [resumoRows],
            [competencias],
            [porSituacao],
            [porConvenio],
          ] = await Promise.all([
            conn.execute(
              `SELECT d.codigo_item, MIN(d.descricao_item) as descricao, d.situacao_item,
                      COUNT(*) as qtd, SUM(CAST(d.valor_pago AS DECIMAL(12,2))) as total_pago, SUM(CAST(d.valor_glosa AS DECIMAL(12,2))) as total_glosa
               FROM demonstrativo d WHERE ${baseWhere} ${whereExtra} GROUP BY d.codigo_item, d.situacao_item ORDER BY total_pago DESC`,
              [...baseParams, ...params]
            ),
            conn.execute(
              `SELECT COUNT(*) as total_itens, COUNT(DISTINCT d.codigo_item) as codigos_unicos,
                      SUM(CAST(d.valor_pago AS DECIMAL(12,2))) as total_pago, SUM(CAST(d.valor_glosa AS DECIMAL(12,2))) as total_glosa, SUM(CAST(d.valor_informado AS DECIMAL(12,2))) as total_informado
               FROM demonstrativo d WHERE ${baseWhere} ${whereExtra}`,
              [...baseParams, ...params]
            ),
            conn.execute(
              `SELECT DISTINCT DATE_FORMAT(d.data_referencia, '%m-%Y') as competencia, MONTH(d.data_referencia) as mes, YEAR(d.data_referencia) as ano, COUNT(*) as total
               FROM demonstrativo d WHERE ${baseWhere} GROUP BY DATE_FORMAT(d.data_referencia, '%m-%Y'), MONTH(d.data_referencia), YEAR(d.data_referencia) ORDER BY ano DESC, mes DESC`,
              [...baseParams]
            ),
            conn.execute(
              `SELECT d.situacao_item, COUNT(*) as total, SUM(CAST(d.valor_pago AS DECIMAL(12,2))) as vl_pago, SUM(CAST(d.valor_glosa AS DECIMAL(12,2))) as vl_glosa
               FROM demonstrativo d WHERE ${baseWhere} ${whereExtra} GROUP BY d.situacao_item`,
              [...baseParams, ...params]
            ),
            conn.execute(
              `SELECT d.convenio_id, c.nome as convenio_nome, COUNT(*) as qtd, COUNT(DISTINCT d.codigo_item) as codigos, SUM(CAST(d.valor_pago AS DECIMAL(12,2))) as total_pago, SUM(CAST(d.valor_glosa AS DECIMAL(12,2))) as total_glosa
               FROM demonstrativo d LEFT JOIN convenios c ON c.id = d.convenio_id WHERE ${baseWhere} ${whereExtra} GROUP BY d.convenio_id, c.nome ORDER BY total_pago DESC`,
              [...baseParams, ...params]
            ),
          ]);

          let itens: any[] = [];
          if (input.mesReferencia && input.anoReferencia) {
            const [itensResult] = await conn.execute(
              `SELECT d.id, d.codigo_item, d.descricao_item, d.situacao_item, CAST(d.valor_pago AS DECIMAL(12,2)) as valor_pago, CAST(d.valor_glosa AS DECIMAL(12,2)) as valor_glosa, CAST(d.valor_informado AS DECIMAL(12,2)) as valor_informado, d.numero_guia, d.nome_beneficiario, d.data_execucao, d.data_referencia, d.data_pagamento, d.codigo_glosa, d.convenio_id, c.nome as convenio_nome
               FROM demonstrativo d LEFT JOIN convenios c ON c.id = d.convenio_id WHERE ${baseWhere} ${whereExtra} ORDER BY d.codigo_item, d.data_referencia DESC`,
              [...baseParams, ...params]
            );
            itens = itensResult as any[];
          }

          const r = (resumoRows as any[])[0] || {};
          return {
            itens,
            porCodigo: porCodigo as any[],
            porSituacao: porSituacao as any[],
            porConvenio: (porConvenio as any[]).map((c: any) => ({
              convenioId: c.convenio_id,
              convenioNome: c.convenio_nome || `Convênio ${c.convenio_id}`,
              qtd: Number(c.qtd),
              codigos: Number(c.codigos),
              totalPago: Number(c.total_pago || 0),
              totalGlosa: Number(c.total_glosa || 0),
            })),
            resumo: {
              totalItens: Number(r.total_itens || 0),
              totalPago: Number(r.total_pago || 0),
              totalGlosa: Number(r.total_glosa || 0),
              totalInformado: Number(r.total_informado || 0),
              codigosUnicos: Number(r.codigos_unicos || 0),
            },
            competencias: (competencias as any[]).map((c: any) => ({
              value: `${c.mes}-${c.ano}`,
              label: `${String(c.mes).padStart(2, "0")}/${c.ano}`,
              total: Number(c.total),
              mes: Number(c.mes),
              ano: Number(c.ano),
            })),
          };
        } finally {
          if (conn) await conn.end().catch(() => {});
        }
      }),
  }),
  // ============================================================
  // RELATÓRIO OX UTI
  // Relatório analítico de diárias de UTI por paciente, item e tipo
  // ============================================================
  relatorioOxUti: router({
    // Meses de competência (data_referencia = mês do faturamento/XML) disponíveis
    mesesDisponiveis: relatorioOxUtiProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          const [rows] = await conn.execute(
            `SELECT DISTINCT DATE_FORMAT(data_referencia, '%Y-%m') as mes_ref
             FROM demonstrativo
             WHERE estabelecimentoId = ? AND data_referencia IS NOT NULL
             ORDER BY mes_ref DESC`,
            [input.estabelecimentoId]
          );
          return (rows as any[]).map((r: any) => r.mes_ref);
        } finally {
          if (conn) await conn.end();
        }
      }),

    // Meses de pagamento (demonstrativo) disponíveis para filtro
    mesesPagamentoDisponiveis: relatorioOxUtiProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          const [rows] = await conn.execute(
            `SELECT DISTINCT DATE_FORMAT(data_pagamento, '%Y-%m') as mes_pag
             FROM demonstrativo
             WHERE estabelecimentoId = ? AND data_pagamento IS NOT NULL
             ORDER BY mes_pag DESC`,
            [input.estabelecimentoId]
          );
          return (rows as any[]).map((r: any) => r.mes_pag);
        } finally {
          if (conn) await conn.end();
        }
      }),

    // Dados principais do relatório Ox UTI
    dados: relatorioOxUtiProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesRef: z.string(),
        convenioId: z.number().optional(),
        mesPagamento: z.string().optional(), // filtro por mês de pagamento do demonstrativo
        })
      )
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);

          // Filtro principal por competência (data_referencia)
          const valores = sqlValoresRelatorioOxUti("d");
          let baseWhere = `WHERE d.estabelecimentoId = ? AND DATE_FORMAT(d.data_referencia, '%Y-%m') = ?`;
          const baseParams: any[] = [input.estabelecimentoId, input.mesRef];
          // Filtro adicional por mês de pagamento (se selecionado)
          if (input.mesPagamento) {
            baseWhere += ` AND DATE_FORMAT(d.data_pagamento, '%Y-%m') = ?`;
            baseParams.push(input.mesPagamento);
          }
          const convenioFilter = input.convenioId
            ? ` AND d.convenio_id = ?`
            : "";
          if (input.convenioId) baseParams.push(input.convenioId);

          // KPIs gerais
          const [kpiRows] = await conn.execute(
            `SELECT
               COUNT(DISTINCT d.numero_guia) as total_guias,
               COUNT(*) as total_itens,
               COALESCE(SUM(${valores.valorInformado}), 0) as total_informado,
               COALESCE(SUM(${valores.valorPago}), 0) as total_pago,
               COALESCE(SUM(${valores.valorGlosa}), 0) as total_glosado,
               COALESCE(SUM(CASE WHEN d.descricao_item REGEXP '^[Dd][[:space:]]*-' OR (d.tipo_lancamento = 'HOS' AND d.descricao_item LIKE '%Di%ria%') THEN d.quantidade ELSE 0 END), 0) as total_diarias
             FROM demonstrativo d ${baseWhere}${convenioFilter}`,
            baseParams
          );
          const kpi = (kpiRows as any[])[0];

          // Por tipo (DIÁRIA / TAXA / MAT_MED / OUTROS) - categoriza pelo prefixo da descrição
          const [tipoRows] = await conn.execute(
            `SELECT
               CASE
                 WHEN d.descricao_item REGEXP '^[Dd][[:space:]]*-' THEN 'DIARIA'
                 WHEN d.descricao_item REGEXP '^[Tt][[:space:]]*-' THEN 'TAXA'
                 WHEN d.descricao_item REGEXP '^MED[[:space:]]*-' THEN 'MED'
                 WHEN d.descricao_item REGEXP '^MAT[[:space:]]*-' THEN 'MAT'
                 WHEN d.tipo_lancamento = 'HOS' AND d.descricao_item LIKE '%Di%ria%' THEN 'DIARIA'
                 WHEN d.tipo_lancamento = 'HOS' AND d.descricao_item LIKE '%Taxa%' THEN 'TAXA'
                 WHEN d.tipo_lancamento = 'MED' THEN 'MED'
                 WHEN d.tipo_lancamento = 'MAT' THEN 'MAT'
                 ELSE 'OUTROS'
               END as categoria,
               COUNT(*) as qtd_itens,
               COALESCE(SUM(${valores.valorInformado}), 0) as total_informado,
               COALESCE(SUM(${valores.valorPago}), 0) as total_pago,
               COALESCE(SUM(${valores.valorGlosa}), 0) as total_glosado,
               COALESCE(SUM(d.quantidade), 0) as total_quantidade
             FROM demonstrativo d ${baseWhere}${convenioFilter}
             GROUP BY categoria ORDER BY total_pago DESC`,
            baseParams
          );

          // Por paciente - agrupado por paciente + lote + protocolo para separar valores
          const [pacienteRows] = await conn.execute(
            `SELECT
               d.nome_beneficiario as paciente,
               d.carteira_beneficiario as carteira,
               COALESCE(d.lote_prestador, '-') as lote_prestador,
               COALESCE(d.protocolo, '-') as protocolo_tiss,
               COUNT(DISTINCT d.numero_guia) as total_guias,
               COUNT(*) as total_itens,
               COALESCE(SUM(CASE WHEN d.descricao_item REGEXP '^[Dd][[:space:]]*-' OR (d.tipo_lancamento = 'HOS' AND d.descricao_item LIKE '%Di%ria%') THEN d.quantidade ELSE 0 END), 0) as qtd_diarias,
               COALESCE(SUM(${valores.valorInformado}), 0) as total_informado,
               COALESCE(SUM(${valores.valorPago}), 0) as total_pago,
               COALESCE(SUM(${valores.valorGlosa}), 0) as total_glosado
             FROM demonstrativo d ${baseWhere}${convenioFilter}
             GROUP BY d.nome_beneficiario, d.carteira_beneficiario, d.lote_prestador, d.protocolo
             ORDER BY total_pago DESC LIMIT 200`,
            baseParams
          );

          // Por item
          const [itemRows] = await conn.execute(
            `SELECT
               d.codigo_item as codigo,
               d.descricao_item as descricao,
               d.tipo_lancamento as tipo_lancamento,
               CASE
                 WHEN d.descricao_item REGEXP '^[Dd][[:space:]]*-' THEN 'DIARIA'
                 WHEN d.descricao_item REGEXP '^[Tt][[:space:]]*-' THEN 'TAXA'
                 WHEN d.descricao_item REGEXP '^MED[[:space:]]*-' THEN 'MED'
                 WHEN d.descricao_item REGEXP '^MAT[[:space:]]*-' THEN 'MAT'
                 WHEN d.tipo_lancamento = 'HOS' AND d.descricao_item LIKE '%Di%ria%' THEN 'DIARIA'
                 WHEN d.tipo_lancamento = 'HOS' AND d.descricao_item LIKE '%Taxa%' THEN 'TAXA'
                 WHEN d.tipo_lancamento = 'MED' THEN 'MED'
                 WHEN d.tipo_lancamento = 'MAT' THEN 'MAT'
                 ELSE 'OUTROS'
               END as categoria,
               COUNT(*) as qtd_ocorrencias,
               COALESCE(SUM(d.quantidade), 0) as total_quantidade,
               COALESCE(SUM(${valores.valorInformado}), 0) as total_informado,
               COALESCE(SUM(${valores.valorPago}), 0) as total_pago,
               COALESCE(SUM(${valores.valorGlosa}), 0) as total_glosado
             FROM demonstrativo d ${baseWhere}${convenioFilter}
             GROUP BY d.codigo_item, d.descricao_item, d.tipo_lancamento
             ORDER BY total_pago DESC LIMIT 300`,
            baseParams
          );

          const totalDiarias = Number(kpi.total_diarias);
          const totalPago = Number(kpi.total_pago);
          const totalInformado = Number(kpi.total_informado);
          const ticketMedio =
            totalDiarias > 0 ? totalPago / totalDiarias : null;
          const percGlosa =
            totalInformado > 0
              ? (Number(kpi.total_glosado) / totalInformado) * 100
              : 0;

          return {
            kpi: {
              totalGuias: Number(kpi.total_guias),
              totalItens: Number(kpi.total_itens),
              totalInformado,
              totalPago,
              totalGlosado: Number(kpi.total_glosado),
              totalDiarias,
              ticketMedio,
              percGlosa,
            },
            porTipo: (tipoRows as any[]).map((r: any) => ({
              categoria: r.categoria as string,
              qtdItens: Number(r.qtd_itens),
              totalInformado: Number(r.total_informado),
              totalPago: Number(r.total_pago),
              totalGlosado: Number(r.total_glosado),
              totalQuantidade: Number(r.total_quantidade),
            })),
            porPaciente: (pacienteRows as any[]).map((r: any) => ({
              paciente: r.paciente as string,
              carteira: r.carteira as string,
              lotePrestador: r.lote_prestador as string,
              protocoloTiss: r.protocolo_tiss as string,
              totalGuias: Number(r.total_guias),
              totalItens: Number(r.total_itens),
              qtdDiarias: Number(r.qtd_diarias),
              totalInformado: Number(r.total_informado),
              totalPago: Number(r.total_pago),
              totalGlosado: Number(r.total_glosado),
              ticketMedioPaciente:
                Number(r.qtd_diarias) > 0
                  ? Number(r.total_pago) / Number(r.qtd_diarias)
                  : null,
            })),
            porItem: (itemRows as any[]).map((r: any) => ({
              codigo: r.codigo as string,
              descricao: r.descricao as string,
              tipoLancamento: r.tipo_lancamento as string,
              categoria: r.categoria as string,
              qtdOcorrencias: Number(r.qtd_ocorrencias),
              totalQuantidade: Number(r.total_quantidade),
              totalInformado: Number(r.total_informado),
              totalPago: Number(r.total_pago),
              totalGlosado: Number(r.total_glosado),
            })),
          };
        } finally {
          if (conn) await conn.end();
        }
      }),

    // Convênios disponíveis para filtro
    convenios: relatorioOxUtiProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          const [rows] = await conn.execute(
            `SELECT DISTINCT c.id, c.nome
             FROM demonstrativo d
             JOIN convenios c ON c.id = d.convenio_id
             WHERE d.estabelecimentoId = ?
             ORDER BY c.nome`,
            [input.estabelecimentoId]
          );
          return (rows as any[]).map((r: any) => ({
            id: Number(r.id),
            nome: r.nome as string,
          }));
        } finally {
          if (conn) await conn.end();
        }
      }),

    // Itens detalhados de um paciente específico para o modal
    itensPorPaciente: relatorioOxUtiProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesRef: z.string(),
        carteiraBeneficiario: z.string(),
        convenioId: z.number().optional(),
        mesPagamento: z.string().optional(),
        })
      )
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);

          const valores = sqlValoresRelatorioOxUti("d");
          const params: any[] = [
            input.estabelecimentoId,
            input.mesRef,
            input.carteiraBeneficiario,
          ];
          const convenioFilter = input.convenioId
            ? " AND d.convenio_id = ?"
            : "";
          if (input.convenioId) params.push(input.convenioId);
          const pagamentoFilter = input.mesPagamento
            ? ` AND DATE_FORMAT(d.data_pagamento, '%Y-%m') = ?`
            : "";
          if (input.mesPagamento) params.push(input.mesPagamento);

          const [rows] = await conn.execute(
            `SELECT
               d.numero_guia,
               d.codigo_item,
               d.descricao_item,
               d.tipo_lancamento,
               CASE
                 WHEN d.descricao_item REGEXP '^[Dd][[:space:]]*-' THEN 'DIARIA'
                 WHEN d.descricao_item REGEXP '^[Tt][[:space:]]*-' THEN 'TAXA'
                 WHEN d.descricao_item REGEXP '^MED[[:space:]]*-' THEN 'MED'
                 WHEN d.descricao_item REGEXP '^MAT[[:space:]]*-' THEN 'MAT'
                 WHEN d.tipo_lancamento = 'HOS' AND d.descricao_item LIKE '%Di%ria%' THEN 'DIARIA'
                 WHEN d.tipo_lancamento = 'HOS' AND d.descricao_item LIKE '%Taxa%' THEN 'TAXA'
                 WHEN d.tipo_lancamento = 'MED' THEN 'MED'
                 WHEN d.tipo_lancamento = 'MAT' THEN 'MAT'
                 ELSE 'OUTROS'
               END as categoria,
               COALESCE(d.quantidade, 0) as quantidade,
               ${valores.valorInformado} as valor_informado,
               ${valores.valorPago} as valor_pago,
               ${valores.valorGlosa} as valor_glosa
             FROM demonstrativo d
             WHERE d.estabelecimentoId = ?
               AND DATE_FORMAT(d.data_referencia, '%Y-%m') = ?
               AND d.carteira_beneficiario = ?${convenioFilter}${pagamentoFilter}
             ORDER BY categoria, d.descricao_item`,
            params
          );

          // Agrupar por categoria
          const porCategoria: Record<
            string,
            {
            categoria: string;
            itens: Array<{
              numeroGuia: string;
              codigo: string;
              descricao: string;
              quantidade: number;
              valorInformado: number;
              valorPago: number;
              valorGlosa: number;
            }>;
            totalInformado: number;
            totalPago: number;
            totalGlosa: number;
            totalQuantidade: number;
            }
          > = {};

          for (const r of rows as any[]) {
            const cat = r.categoria as string;
            if (!porCategoria[cat]) {
              porCategoria[cat] = {
                categoria: cat,
                itens: [],
                totalInformado: 0,
                totalPago: 0,
                totalGlosa: 0,
                totalQuantidade: 0,
              };
            }
            porCategoria[cat].itens.push({
              numeroGuia: r.numero_guia as string,
              codigo: r.codigo_item as string,
              descricao: r.descricao_item as string,
              quantidade: Number(r.quantidade),
              valorInformado: Number(r.valor_informado),
              valorPago: Number(r.valor_pago),
              valorGlosa: Number(r.valor_glosa),
            });
            porCategoria[cat].totalInformado += Number(r.valor_informado);
            porCategoria[cat].totalPago += Number(r.valor_pago);
            porCategoria[cat].totalGlosa += Number(r.valor_glosa);
            porCategoria[cat].totalQuantidade += Number(r.quantidade);
          }

          const ordemCategoria = ["DIARIA", "TAXA", "MED", "MAT", "OUTROS"];
          return Object.values(porCategoria).sort((a, b) => {
            return (
              ordemCategoria.indexOf(a.categoria) -
              ordemCategoria.indexOf(b.categoria)
            );
          });
        } finally {
          if (conn) await conn.end();
        }
      }),
  }),

  // ─── Relatório OX UTI - Faturado (dados do XML enviado) ───────────────────
  relatorioOxUtiFaturado: router({
    // Meses disponíveis nos itens enviados via XML
    mesesDisponiveis: relatorioOxUtiFaturadoProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          const [rows] = await conn.execute(
            `SELECT DISTINCT REPLACE(cci.competencia, '/', '-') as mes_ref
             FROM contas_convenio_itens cci
             WHERE cci.estabelecimentoId = ? AND cci.competencia IS NOT NULL AND cci.competencia != ''
             ORDER BY mes_ref DESC`,
            [input.estabelecimentoId]
          );
          return (rows as any[]).map((r: any) => r.mes_ref as string);
        } finally {
          if (conn) await conn.end();
        }
      }),

    // Convênios disponíveis para filtro
    convenios: relatorioOxUtiFaturadoProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          const [rows] = await conn.execute(
            `SELECT DISTINCT convenio FROM contas_convenio_itens
             WHERE estabelecimentoId = ? AND convenio IS NOT NULL
             ORDER BY convenio`,
            [input.estabelecimentoId]
          );
          return (rows as any[]).map((r: any) => r.convenio as string);
        } finally {
          if (conn) await conn.end();
        }
      }),

    // Dados principais do relatório
    dados: relatorioOxUtiFaturadoProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesRef: z.string(),
        convenio: z.string().optional(),
        })
      )
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);

          // Usa competencia (formato '2026/05') para filtrar por mês de faturamento
          const mesRefFormatado = input.mesRef.replace("-", "/");
          const baseWhere = `WHERE cci.estabelecimentoId = ? AND cci.competencia = ?`;
          const baseParams: any[] = [input.estabelecimentoId, mesRefFormatado];
          const convenioFilter = input.convenio ? ` AND cci.convenio = ?` : "";
          if (input.convenio) baseParams.push(input.convenio);

          // KPIs gerais
          const [kpiRows] = await conn.execute(
            `SELECT
               COUNT(DISTINCT cci.numeroGuia) as total_guias,
               COUNT(*) as total_itens,
               COALESCE(SUM(cci.valorTotal), 0) as total_faturado,
               COALESCE(SUM(CASE WHEN cci.tipoItem = 'DIÁRIA' THEN cci.quantidade ELSE 0 END), 0) as total_diarias,
               COUNT(DISTINCT cci.carteiraBeneficiario) as total_pacientes
             FROM contas_convenio_itens cci ${baseWhere}${convenioFilter}`,
            baseParams
          );
          const kpi = (kpiRows as any[])[0];

          // Por tipo (usando tipoItem direto do XML)
          const [tipoRows] = await conn.execute(
            `SELECT
               COALESCE(cci.tipoItem, 'NÃO INFORMADO') as categoria,
               COUNT(*) as qtd_itens,
               COALESCE(SUM(cci.quantidade), 0) as total_quantidade,
               COALESCE(SUM(cci.valorTotal), 0) as total_faturado,
               COALESCE(SUM(cci.valorUnitario * cci.quantidade), 0) as total_calculado
             FROM contas_convenio_itens cci ${baseWhere}${convenioFilter}
             GROUP BY COALESCE(cci.tipoItem, 'NÃO INFORMADO')
             ORDER BY total_faturado DESC`,
            baseParams
          );

          // Por paciente - agrupado por paciente + guia + lote para separar valores
          const [pacienteRows] = await conn.execute(
            `SELECT
               cci.pacienteNome as paciente,
               cci.carteiraBeneficiario as carteira,
               COALESCE(cci.numeroGuia, '-') as numero_guia,
               COALESCE(cci.numeroLote, '-') as numero_lote,
               COUNT(DISTINCT cci.numeroGuia) as total_guias,
               COUNT(*) as total_itens,
               COALESCE(SUM(CASE WHEN cci.tipoItem = 'DIÁRIA' THEN cci.quantidade ELSE 0 END), 0) as qtd_diarias,
               COALESCE(SUM(cci.valorTotal), 0) as total_faturado,
               GROUP_CONCAT(DISTINCT cci.convenio ORDER BY cci.convenio SEPARATOR ', ') as convenios
             FROM contas_convenio_itens cci ${baseWhere}${convenioFilter}
             GROUP BY cci.pacienteNome, cci.carteiraBeneficiario, cci.numeroGuia, cci.numeroLote
             ORDER BY total_faturado DESC
             LIMIT 200`,
            baseParams
          );

          // Por item
          const [itemRows] = await conn.execute(
            `SELECT
               cci.codigoItem as codigo,
               cci.descricaoItem as descricao,
               COALESCE(cci.tipoItem, 'NÃO INFORMADO') as categoria,
               COUNT(*) as qtd_ocorrencias,
               COALESCE(SUM(cci.quantidade), 0) as total_quantidade,
               COALESCE(SUM(cci.valorTotal), 0) as total_faturado,
               AVG(cci.valorUnitario) as valor_unitario_medio
             FROM contas_convenio_itens cci ${baseWhere}${convenioFilter}
             GROUP BY cci.codigoItem, cci.descricaoItem, COALESCE(cci.tipoItem, 'NÃO INFORMADO')
             ORDER BY total_faturado DESC
             LIMIT 300`,
            baseParams
          );

          const totalFaturado = Number(kpi.total_faturado);
          const totalDiarias = Number(kpi.total_diarias);
          const ticketMedio =
            totalDiarias > 0 ? totalFaturado / totalDiarias : null;

          return {
            kpi: {
              totalGuias: Number(kpi.total_guias),
              totalItens: Number(kpi.total_itens),
              totalFaturado,
              totalDiarias,
              totalPacientes: Number(kpi.total_pacientes),
              ticketMedio,
            },
            porTipo: (tipoRows as any[]).map((r: any) => ({
              categoria: r.categoria as string,
              qtdItens: Number(r.qtd_itens),
              totalQuantidade: Number(r.total_quantidade),
              totalFaturado: Number(r.total_faturado),
            })),
            porPaciente: (pacienteRows as any[]).map((r: any) => ({
              paciente: r.paciente as string,
              carteira: r.carteira as string,
              numeroGuia: r.numero_guia as string,
              numeroLote: r.numero_lote as string,
              totalGuias: Number(r.total_guias),
              totalItens: Number(r.total_itens),
              qtdDiarias: Number(r.qtd_diarias),
              totalFaturado: Number(r.total_faturado),
              convenios: r.convenios as string,
              ticketMedioPaciente:
                Number(r.qtd_diarias) > 0
                  ? Number(r.total_faturado) / Number(r.qtd_diarias)
                  : null,
            })),
            porItem: (itemRows as any[]).map((r: any) => ({
              codigo: r.codigo as string,
              descricao: r.descricao as string,
              categoria: r.categoria as string,
              qtdOcorrencias: Number(r.qtd_ocorrencias),
              totalQuantidade: Number(r.total_quantidade),
              totalFaturado: Number(r.total_faturado),
              valorUnitarioMedio: Number(r.valor_unitario_medio),
            })),
          };
        } finally {
          if (conn) await conn.end();
        }
      }),

    // Meses de pagamento disponíveis (para filtro de data pagamento)
    mesesPagamento: relatorioOxUtiFaturadoProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          const [rows] = await conn.execute(
            `SELECT DISTINCT DATE_FORMAT(re.dataPagto, '%Y-%m') as mes_pag
             FROM recebimentos_excel re
             WHERE re.estabelecimentoId = ? AND re.dataPagto IS NOT NULL
             ORDER BY mes_pag DESC`,
            [input.estabelecimentoId]
          );
          return (rows as any[]).map((r: any) => r.mes_pag as string);
        } finally {
          if (conn) await conn.end();
        }
      }),

    // Cruzamento Faturado x Pago por guia (competência)
    faturadoXPago: relatorioOxUtiFaturadoProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesRef: z.string(),
        convenio: z.string().optional(),
        mesPagamento: z.string().optional(), // YYYY-MM para filtrar por mês de pagamento do demonstrativo
        })
      )
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);

          // Buscar guias faturadas (XML enviado) no mês
          const baseParams: any[] = [input.estabelecimentoId, input.mesRef];
          let convenioFilter = "";
          if (input.convenio) {
            convenioFilter = " AND cci.convenio = ?";
            baseParams.push(input.convenio);
          }

          // Dados do XML enviado (faturado) - agrupado por guia
          const [faturadoRows] = await conn.execute(
            `SELECT
               cci.numeroConta as numero_guia,
               cci.pacienteNome as paciente,
               cci.convenio,
               MIN(cci.dataExecucao) as data_faturado,
               COALESCE(SUM(cci.valorTotal), 0) as valor_faturado,
               COUNT(*) as qtd_itens
             FROM contas_convenio_itens cci
             WHERE cci.estabelecimentoId = ?
               AND DATE_FORMAT(cci.dataExecucao, '%Y-%m') = ?
               ${convenioFilter}
             GROUP BY cci.numeroConta, cci.pacienteNome, cci.convenio
             ORDER BY cci.numeroConta`,
            baseParams
          );

          // Buscar pagamentos recebidos para essas guias
          // Usar recebimentos_excel que tem os dados dos demonstrativos de pagamento
          const guias = (faturadoRows as any[])
            .map((r: any) => r.numero_guia)
            .filter(Boolean);

          let pagamentoMap: Record<
            string,
            {
              dataPagamento: string | null;
              valorRecebido: number;
              valorGlosa: number;
            }
          > = {};

          if (guias.length > 0) {
            // Buscar em lotes de 500 para evitar query muito grande
            const batchSize = 500;
            for (let i = 0; i < guias.length; i += batchSize) {
              const batch = guias.slice(i, i + batchSize);
              const placeholders = batch.map(() => "?").join(",");
              const [recRows] = await conn.execute(
                `SELECT
                   re.numeroGuia as numero_guia,
                   MAX(re.dataPagto) as data_pagamento,
                   COALESCE(SUM(CAST(re.valorPagamento AS DECIMAL(12,2))), 0) as valor_recebido,
                   COALESCE(SUM(CAST(re.valorGlosa AS DECIMAL(12,2))), 0) as valor_glosa
                 FROM recebimentos_excel re
                 WHERE re.numeroGuia IN (${placeholders})
                   AND re.estabelecimentoId = ?${input.mesPagamento ? ` AND DATE_FORMAT(re.dataPagto, '%Y-%m') = ?` : ""}
                 GROUP BY re.numeroGuia`,
                [
                  ...batch,
                  input.estabelecimentoId,
                  ...(input.mesPagamento ? [input.mesPagamento] : []),
                ]
              );
              for (const r of recRows as any[]) {
                pagamentoMap[r.numero_guia] = {
                  dataPagamento: r.data_pagamento
                    ? new Date(r.data_pagamento).toISOString().split("T")[0]
                    : null,
                  valorRecebido: Number(r.valor_recebido),
                  valorGlosa: Number(r.valor_glosa),
                };
              }
            }
          }

          // Montar resultado cruzado
          const resultado = (faturadoRows as any[]).map((r: any) => {
            const pag = pagamentoMap[r.numero_guia];
            const valorFaturado = Number(r.valor_faturado);
            const valorRecebido = pag?.valorRecebido ?? 0;
            const valorGlosa = pag?.valorGlosa ?? 0;
            return {
              numeroGuia: r.numero_guia as string,
              paciente: r.paciente as string,
              convenio: r.convenio as string,
              dataFaturado: r.data_faturado
                ? new Date(r.data_faturado).toISOString().split("T")[0]
                : null,
              dataPagamento: pag?.dataPagamento ?? null,
              valorFaturado,
              valorRecebido,
              valorGlosa,
              diferenca: valorFaturado - valorRecebido,
              status:
                valorRecebido > 0
                  ? valorRecebido >= valorFaturado
                    ? "pago"
                    : "parcial"
                  : "pendente",
              qtdItens: Number(r.qtd_itens),
            };
          });

          // Totais
          const totalFaturado = resultado.reduce(
            (s, r) => s + r.valorFaturado,
            0
          );
          const totalRecebido = resultado.reduce(
            (s, r) => s + r.valorRecebido,
            0
          );
          const totalGlosa = resultado.reduce((s, r) => s + r.valorGlosa, 0);
          const totalPendente = totalFaturado - totalRecebido;
          const guiasPagas = resultado.filter(r => r.status === "pago").length;
          const guiasParciais = resultado.filter(
            r => r.status === "parcial"
          ).length;
          const guiasPendentes = resultado.filter(
            r => r.status === "pendente"
          ).length;

          return {
            resumo: {
              totalGuias: resultado.length,
              totalFaturado,
              totalRecebido,
              totalGlosa,
              totalPendente,
              guiasPagas,
              guiasParciais,
              guiasPendentes,
              percentualRecebido:
                totalFaturado > 0 ? (totalRecebido / totalFaturado) * 100 : 0,
            },
            guias: resultado,
          };
        } finally {
          if (conn) await conn.end();
        }
      }),
  }),

  // Relatório Resumo Paciente Ox UTI - Faturado x Recebido por paciente/guia com categorias
  resumoPacienteOxUti: router({
    mesesDisponiveis: resumoPacienteOxUtiProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          const [rows] = await conn.execute(
            `SELECT DISTINCT REPLACE(cci.competencia, '/', '-') as mes_ref
             FROM contas_convenio_itens cci
             WHERE cci.estabelecimentoId = ? AND cci.competencia IS NOT NULL AND cci.competencia != ''
             ORDER BY mes_ref DESC`,
            [input.estabelecimentoId]
          );
          return (rows as any[]).map((r: any) => r.mes_ref as string);
        } finally {
          if (conn) await conn.end();
        }
      }),

    convenios: resumoPacienteOxUtiProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          const [rows] = await conn.execute(
            `SELECT DISTINCT convenio FROM contas_convenio_itens
             WHERE estabelecimentoId = ? AND convenio IS NOT NULL
             ORDER BY convenio`,
            [input.estabelecimentoId]
          );
          return (rows as any[]).map((r: any) => r.convenio as string);
        } finally {
          if (conn) await conn.end();
        }
      }),

    faturado: resumoPacienteOxUtiProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesRef: z.string(),
        convenio: z.string().optional(),
        })
      )
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);

          const mesRefFormatado = input.mesRef.replace("-", "/");
          let whereClause = `WHERE cci.estabelecimentoId = ? AND cci.competencia = ?`;
          const params: any[] = [input.estabelecimentoId, mesRefFormatado];
          if (input.convenio) {
            whereClause += ` AND cci.convenio = ?`;
            params.push(input.convenio);
          }

          // Query faturado agrupado por carteira + guia + lote
          const [rows] = await conn.execute(
            `SELECT
               cci.carteiraBeneficiario,
               cci.pacienteNome,
               cci.numeroGuia,
               cci.numeroLote,
               cci.convenio,
               DATE_FORMAT(MIN(cci.dataExecucao), '%Y-%m-%d') as data_entrada,
               DATE_FORMAT(MAX(cci.dataExecucao), '%Y-%m-%d') as data_saida,
               SUM(CASE WHEN cci.tipoItem = 'DIÁRIA' THEN cci.quantidade ELSE 0 END) as qtd_diarias,
               SUM(CASE WHEN cci.tipoItem = 'DIÁRIA' THEN cci.valorTotal ELSE 0 END) as vl_diaria,
               SUM(CASE WHEN cci.tipoItem IN ('MEDICAMENTO','MATERIAL') THEN cci.valorTotal ELSE 0 END) as vl_matmed,
               SUM(CASE WHEN cci.tipoItem = 'GÁS MEDICINAL' THEN cci.valorTotal ELSE 0 END) as vl_gas,
               SUM(CASE WHEN cci.tipoItem IN ('TAXA/ALUGUÉIS','TAXA','ALUGUEL') THEN cci.valorTotal ELSE 0 END) as vl_taxa,
               SUM(cci.valorTotal) as vl_total
             FROM contas_convenio_itens cci
             ${whereClause}
             GROUP BY cci.carteiraBeneficiario, cci.pacienteNome, cci.numeroGuia, cci.numeroLote, cci.convenio
             ORDER BY cci.carteiraBeneficiario, cci.numeroGuia, cci.numeroLote`,
            params
          );

          const resultado = (rows as any[]).map((r: any) => ({
            paciente: (r.pacienteNome || r.carteiraBeneficiario) as string,
            carteira: r.carteiraBeneficiario as string,
            guia: r.numeroGuia as string,
            lote: (r.numeroLote || "") as string,
            convenio: r.convenio as string,
            dataEntrada: r.data_entrada,
            dataSaida: r.data_saida,
            qtdDiarias: Number(r.qtd_diarias) || 0,
            vlDiaria: Number(r.vl_diaria) || 0,
            vlMatmed: Number(r.vl_matmed) || 0,
            vlGas: Number(r.vl_gas) || 0,
            vlTaxa: Number(r.vl_taxa) || 0,
            vlTotal: Number(r.vl_total) || 0,
          }));

          const totais = resultado.reduce(
            (acc, r) => ({
            vlDiaria: acc.vlDiaria + r.vlDiaria,
            vlMatmed: acc.vlMatmed + r.vlMatmed,
            vlGas: acc.vlGas + r.vlGas,
            vlTaxa: acc.vlTaxa + r.vlTaxa,
            vlTotal: acc.vlTotal + r.vlTotal,
            qtdDiarias: acc.qtdDiarias + r.qtdDiarias,
            }),
            {
              vlDiaria: 0,
              vlMatmed: 0,
              vlGas: 0,
              vlTaxa: 0,
              vlTotal: 0,
              qtdDiarias: 0,
            }
          );

          return {
            dados: resultado,
            totais: {
              ...totais,
              totalPacientes: new Set(resultado.map(r => r.carteira)).size,
              totalGuias: new Set(resultado.map(r => r.guia)).size,
            },
          };
        } finally {
          if (conn) await conn.end();
        }
      }),

    recebido: resumoPacienteOxUtiProcedure
      .input(
        z.object({
        estabelecimentoId: z.number(),
        mesRef: z.string(),
        convenio: z.string().optional(),
        })
      )
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          console.log(
            "[ResumoPacienteOxUti.recebido] Input:",
            JSON.stringify(input)
          );
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);

          // Primeiro buscar as guias da competência selecionada no faturamento
          const mesRefFormatado = input.mesRef.replace("-", "/");
          let fatWhere = `WHERE cci.estabelecimentoId = ? AND cci.competencia = ?`;
          const fatParams: any[] = [input.estabelecimentoId, mesRefFormatado];
          if (input.convenio) {
            fatWhere += ` AND cci.convenio = ?`;
            fatParams.push(input.convenio);
          }
          const [guiaRows] = await conn.execute(
            `SELECT DISTINCT cci.numeroGuia FROM contas_convenio_itens cci ${fatWhere} AND cci.numeroGuia IS NOT NULL`,
            fatParams
          );
          const guias = (guiaRows as any[]).map((r: any) => r.numeroGuia);
          console.log("[recebido] Guias:", guias.length, guias.slice(0, 3));
          if (guias.length === 0) {
            return {
              dados: [],
              totais: {
                vlDiaria: 0,
                vlMatmed: 0,
                vlOutros: 0,
                vlTotal: 0,
                vlGlosa: 0,
                totalPacientes: 0,
                totalGuias: 0,
              },
            };
          }

          // Buscar recebimentos dessas guias (usar conn.query para evitar problema com IN e prepared statements)
          const guiasEscaped = guias
            .map((g: string) => conn.escape(g))
            .join(",");
          let whereClause = `WHERE re.estabelecimentoId = ${conn.escape(input.estabelecimentoId)} AND re.numero_guia IN (${guiasEscaped})`;

          const [rows] = await conn.query(
            `SELECT
               re.beneficiario,
               re.nome_beneficiario,
               re.numero_guia,
               re.protocolo_tiss,
               re.lote_prestador,
               DATE_FORMAT(re.data_pagto, '%Y-%m-%d') as data_pagamento,
               SUM(CASE WHEN re.tipo_lancamento = 'HOS' THEN re.valor_pagamento ELSE 0 END) as vl_diaria,
               SUM(CASE WHEN re.tipo_lancamento IN ('MED','MAT') THEN re.valor_pagamento ELSE 0 END) as vl_matmed,
               SUM(CASE WHEN re.tipo_lancamento NOT IN ('HOS','MED','MAT') THEN re.valor_pagamento ELSE 0 END) as vl_outros,
               SUM(re.valor_pagamento) as vl_total,
               SUM(re.valor_glosa) as vl_glosa
             FROM recebimentos_excel re
             ${whereClause}
             GROUP BY re.beneficiario, re.nome_beneficiario, re.numero_guia, re.protocolo_tiss, re.lote_prestador, DATE_FORMAT(re.data_pagto, '%Y-%m-%d')
             ORDER BY re.beneficiario, re.numero_guia, re.protocolo_tiss`
          );

          console.log("[recebido] Rows:", (rows as any[]).length);
          const resultado = (rows as any[]).map((r: any) => ({
            paciente: (r.nome_beneficiario || r.beneficiario || "") as string,
            carteira: (r.beneficiario || "") as string,
            guia: (r.numero_guia || "") as string,
            protocolo: (r.protocolo_tiss || "") as string,
            lote: (r.lote_prestador || "") as string,
            convenio: "" as string,
            dataPagamento: r.data_pagamento,
            vlDiaria: Number(r.vl_diaria) || 0,
            vlMatmed: Number(r.vl_matmed) || 0,
            vlOutros: Number(r.vl_outros) || 0,
            vlTotal: Number(r.vl_total) || 0,
            vlGlosa: Number(r.vl_glosa) || 0,
          }));

          const totais = resultado.reduce(
            (acc, r) => ({
            vlDiaria: acc.vlDiaria + r.vlDiaria,
            vlMatmed: acc.vlMatmed + r.vlMatmed,
            vlOutros: acc.vlOutros + r.vlOutros,
            vlTotal: acc.vlTotal + r.vlTotal,
            vlGlosa: acc.vlGlosa + r.vlGlosa,
            }),
            { vlDiaria: 0, vlMatmed: 0, vlOutros: 0, vlTotal: 0, vlGlosa: 0 }
          );

          return {
            dados: resultado,
            totais: {
              ...totais,
              totalPacientes: new Set(resultado.map(r => r.carteira)).size,
              totalGuias: new Set(resultado.map(r => r.guia)).size,
            },
          };
        } finally {
          if (conn) await conn.end();
        }
      }),

    mesesPagamento: resumoPacienteOxUtiProcedure
      .input(z.object({ estabelecimentoId: z.number() }))
      .query(async ({ input }) => {
        let conn: any = null;
        try {
          const mysql2 = await import("mysql2/promise");
          conn = await mysql2.createConnection(process.env.DATABASE_URL!);
          const [rows] = await conn.execute(
            `SELECT DISTINCT DATE_FORMAT(dataPagto, '%Y-%m') as mes_ref
             FROM recebimentos_excel
             WHERE estabelecimentoId = ? AND dataPagto IS NOT NULL
             ORDER BY mes_ref DESC`,
            [input.estabelecimentoId]
          );
          return (rows as any[]).map((r: any) => r.mes_ref as string);
        } finally {
          if (conn) await conn.end();
        }
      }),
  }),
});

function calcularDiasParado(datasai: string | null): number {
  if (!datasai) return 0;
  const dataSaida = new Date(datasai);
  dataSaida.setHours(0, 0, 0, 0);
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  const diffTime = hoje.getTime() - dataSaida.getTime();
  const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
  return Math.max(0, diffDays);
}

export type AppRouter = typeof appRouter;
