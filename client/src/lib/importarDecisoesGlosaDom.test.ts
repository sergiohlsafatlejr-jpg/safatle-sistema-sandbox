import { execFileSync } from "node:child_process";
import { homedir } from "node:os";
import { resolve } from "node:path";
import { expect, it } from "vitest";

it("DOM real: prévia de decisões vinculada ao escopo, cancelamento limpo e erros sanitizados (Issue #187)", () => {
  const root = resolve(import.meta.dirname, "../../..");
  const output = execFileSync(process.execPath, ["client/src/lib/importarDecisoesGlosaDom.harness.cjs"], {
    cwd: root,
    env: {
      PATH: process.env.PATH,
      HOME: process.env.HOME || homedir(),
      USERPROFILE: process.env.USERPROFILE || homedir(),
      SYSTEMROOT: process.env.SYSTEMROOT,
      TEMP: process.env.TEMP,
      TMP: process.env.TMP,
      LANG: process.env.LANG || "C.UTF-8",
      DATABASE_URL: "",
      TZ: "America/Sao_Paulo",
      ...(process.env.DOSSIE_DOM_CHROMIUM ? { DOSSIE_DOM_CHROMIUM: process.env.DOSSIE_DOM_CHROMIUM } : {}),
    },
    encoding: "utf8",
    timeout: 160_000,
    maxBuffer: 8_000_000,
  });
  for (const cenario of [
    "build_offline_componente_real",
    "troca_hospital_invalida_previa_e_reanalise_usa_novo_escopo",
    "troca_convenio_invalida_previa",
    "resposta_atrasada_de_outro_escopo_descartada",
    "cancelar_e_reabrir_sem_estado_antigo",
    "erro_interno_sanitizado_no_toast_e_dom",
    "conflito_exige_nova_analise",
    "aplicacao_sem_confirmacao_total_nao_e_sucesso",
    "sucesso_confirma_escopo_e_token_da_previa",
    "importarDecisoesGlosaDom",
  ]) {
    expect(output).toContain(`PASS ${cenario}`);
  }
  expect(output).not.toContain("FAIL ");
}, 170_000);
