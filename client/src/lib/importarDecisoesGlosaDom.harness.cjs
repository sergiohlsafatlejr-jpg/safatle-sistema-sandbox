'use strict';
/**
 * DOM real (Chromium headless) do ImportarDecisoesGlosaDialog — Issue #187.
 * Componente, Radix Dialog/Select e React 19 reais; somente tRPC e sonner são
 * substituídos por dublês offline controláveis. Sem backend, banco, rede ou dados reais.
 */
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const assert = require('node:assert/strict');
const { createRequire } = require('node:module');
const { pathToFileURL } = require('node:url');
const root = path.resolve(__dirname, '../../..');
const req = createRequire(path.join(root, 'package.json'));
let browser, temporary;
const timeout = setTimeout(() => {
  console.error('FAIL DOM_TOTAL_TIMEOUT_150000');
  browser?.process()?.kill('SIGKILL');
  if (temporary) fs.rmSync(temporary, { recursive: true, force: true });
  process.exit(1);
}, 150000);

function executable(puppeteer) {
  const candidates = [];
  if (process.env.DOSSIE_DOM_CHROMIUM) candidates.push(path.resolve(process.env.DOSSIE_DOM_CHROMIUM));
  try { candidates.push(puppeteer.executablePath()); } catch {}
  for (const dir of (process.env.PATH || '').split(path.delimiter))
    for (const file of ['chromium', 'chromium-browser', 'google-chrome', 'chrome', 'chrome.exe', 'headless_shell']) candidates.push(path.join(dir, file));
  candidates.push('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe');
  const found = candidates.find((f) => { try { return fs.statSync(f).isFile(); } catch { return false; } });
  if (!found) throw new Error('CHROMIUM_EXECUTABLE_REQUIRED: instale Chromium ou defina DOSSIE_DOM_CHROMIUM; não há skip.');
  return found;
}

const MOCKS = String.raw`
import { useCallback, useState } from "react";
const w = window as any;
w.audit = w.audit || { calls: [], pending: [], toasts: [] };
export const toast = {
  success: (m: string) => w.audit.toasts.push(["success", m]),
  error: (m: string) => w.audit.toasts.push(["error", m]),
  warning: (m: string) => w.audit.toasts.push(["warning", m]),
  info: (m: string) => w.audit.toasts.push(["info", m]),
};
export const Toaster = () => null;
function useMutation() {
  const [state, setState] = useState<{ data: any; isPending: boolean }>({ data: undefined, isPending: false });
  const mutateAsync = useCallback((input: any) => {
    setState((s) => ({ ...s, isPending: true }));
    return new Promise((resolve, reject) => {
      w.audit.calls.push(JSON.parse(JSON.stringify(input)));
      w.audit.pending.push({
        input,
        resolve: (d: any) => { setState({ data: d, isPending: false }); resolve(d); },
        reject: (e: any) => { setState((s) => ({ ...s, isPending: false })); reject(e); },
      });
    });
  }, []);
  const reset = useCallback(() => setState({ data: undefined, isPending: false }), []);
  return { ...state, mutateAsync, reset };
}
export const trpc = { recursos: { importarDecisoesPlanilha: { useMutation } } };
`;

const APP = (dialogPath) => String.raw`
import { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import { ImportarDecisoesGlosaDialog } from ${JSON.stringify(dialogPath)};
const w = window as any;
w.audit = w.audit || { calls: [], pending: [], toasts: [] };
w.audit.imported = 0;
// Resposta sintética do servidor para a chamada i (escopo/token derivados do input).
w.audit.respond = (i: number, override: any = {}) => {
  const p = w.audit.pending[i];
  const n = p.input.linhas.length;
  const token = (String(p.input.estabelecimentoId) + "-" + String(p.input.convenioId ?? "x")).padEnd(64, "0").replace(/[^0-9a-f]/g, "f").slice(0, 64);
  p.resolve({
    totalLinhas: n, casados: n, naoEncontrados: 0, ignorados: 0,
    aRecursar: p.input.linhas.filter((l: any) => l.decisao === "Sim").length,
    aAceitar: p.input.linhas.filter((l: any) => l.decisao !== "Sim").length,
    aplicados: p.input.aplicar ? n : 0,
    escopo: { estabelecimentoId: p.input.estabelecimentoId, convenioId: p.input.convenioId ?? null },
    tokenPrevia: token,
    resultados: p.input.linhas.map((l: any, k: number) => ({ linha: l.linha, status: "casado", decisao: l.decisao === "Sim" ? "recursar" : "aceitar", demonstrativoId: 100 + k, pontuacao: 8, detalhe: "Vínculo exato", codigoItem: l.codigoItem, guia: l.guiaOperadora, descricaoItem: "", classificacaoAtual: "pendente" })),
    ...override,
  });
};
w.audit.fail = (i: number, err: any) => w.audit.pending[i].reject(err);
function App() {
  const [estab, setEstab] = useState(1);
  const [open, setOpen] = useState(true);
  useEffect(() => { w.audit.setEstab = setEstab; w.audit.setOpen = setOpen; w.audit.isOpen = () => open; }, [open]);
  return (
    <ImportarDecisoesGlosaDialog
      open={open}
      onOpenChange={setOpen}
      estabelecimentoId={estab}
      convenios={[{ id: 10, nome: "Convenio Sintetico A" }, { id: 20, nome: "Convenio Sintetico B" }]}
      convenioInicial="10"
      onImportado={() => { w.audit.imported++; }}
    />
  );
}
createRoot(document.getElementById("root")!).render(<App />);
w.audit.ready = true;
`;

async function build(out) {
  const { build } = await import(pathToFileURL(req.resolve('vite')).href);
  const dialog = path.join(root, 'client/src/components/glosas/ImportarDecisoesGlosaDialog.tsx').replace(/\\/g, '/');
  fs.writeFileSync(path.join(out, 'dom-mocks.tsx'), MOCKS);
  fs.writeFileSync(path.join(out, 'dom-app.tsx'), APP(dialog));
  fs.writeFileSync(path.join(out, 'index.html'), '<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>Teste offline — dados fictícios</title></head><body><div id="root"></div><script type="module" src="/dom-app.tsx"></script></body></html>\n');
  const mocks = path.join(out, 'dom-mocks.tsx');
  await build({
    configFile: false, root: out, base: './', esbuild: { jsx: 'automatic' }, logLevel: 'warn',
    plugins: [{ name: 'offline-trpc-sonner', enforce: 'pre', resolveId(source) { if (source === '@/lib/trpc' || source === 'sonner' || /client[\\/]src[\\/]lib[\\/]trpc(\.tsx?)?$/.test(source)) return mocks; } }],
    resolve: {
      alias: [
        { find: /^@\//, replacement: path.join(root, 'client/src').replace(/\\/g, '/') + '/' },
        ...['react', 'react/jsx-runtime', 'react-dom', 'react-dom/client'].map((p) => ({ find: new RegExp('^' + p.replace(/[/]/g, '\\/') + '$'), replacement: req.resolve(p) })),
      ],
      dedupe: ['react', 'react-dom'],
    },
    build: { outDir: path.join(out, 'dist'), emptyOutDir: true, minify: false, sourcemap: false, modulePreload: false },
  });
  const dist = path.join(out, 'dist');
  let html = fs.readFileSync(path.join(dist, 'index.html'), 'utf8');
  html = html.replace(/<script type="module" crossorigin src="\.\/([^"]+)"><\/script>/g, (_, f) => '<script type="module">' + fs.readFileSync(path.join(dist, f), 'utf8').replace(/<\/script/gi, '<\\/script') + '</script>');
  html = html.replace(/<link rel="stylesheet" crossorigin href="\.\/([^"]+)">/g, (_, f) => '<style>' + fs.readFileSync(path.join(dist, f), 'utf8') + '</style>');
  assert(!/<script[^>]+src="\.\//.test(html), 'assets devem estar inline');
  fs.writeFileSync(path.join(out, 'offline.html'), html);
}

function planilha(file) {
  const XLSX = req('xlsx');
  const rows = [
    ['Oportunidades Pendentes de Recurso (dados fictícios)'],
    ['Grupo', 'VAMOS RECURSAR? (Sim / Não / Avaliar)', 'Sugestão da consultoria', 'Nº Protocolo', 'Nº Guia Prestador', 'Nº Guia Operadora', 'Seq Item', 'Código Procedimento', 'Descrição do Item', 'Vl. Glosado (R$)', 'Códigos Glosa', 'Justificativa / Orientação para Recurso'],
    ['A', 'Sim', 'Recursar', '9001', '11', '7001', '1', '10101012', 'ITEM SINTETICO 1', '100', '1801', 'Justificativa sintética 1'],
    ['A', 'Não', 'Glosa Aceita Pago conforme Contrato', '9001', '11', '7001', '2', '10101020', 'ITEM SINTETICO 2', '50', '1801', 'Orientação sintética 2'],
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), '8. Oportunidades Pendentes');
  XLSX.writeFile(wb, file);
}

(async () => {
  const { default: puppeteer } = await import(pathToFileURL(req.resolve('puppeteer')).href);
  assert.match(req('react/package.json').version, /^19\./, 'React 19 real obrigatório');
  temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'decisoes-glosa-dom-'));
  const out = path.join(temporary, 'build');
  fs.mkdirSync(out);
  await build(out);
  const xlsx = path.join(temporary, 'decisoes-sinteticas.xlsx');
  planilha(xlsx);
  console.log('PASS build_offline_componente_real');

  browser = await puppeteer.launch({ executablePath: executable(puppeteer), headless: true, pipe: true, userDataDir: path.join(temporary, 'profile'), args: ['--no-sandbox', '--disable-dev-shm-usage', '--disable-background-networking', '--no-first-run', '--host-resolver-rules=MAP * ~NOTFOUND', '--disable-extensions'] });
  const requests = [], errors = [];
  let page;
  const settle = () => page.evaluate(() => new Promise((r) => setTimeout(r, 40)));
  async function fresh() {
    if (page) await page.close();
    page = await browser.newPage();
    page.setDefaultTimeout(10000);
    await page.setRequestInterception(true);
    page.on('request', (r) => { if (/^(file:|data:|about:|blob:)/.test(r.url())) void r.continue(); else { requests.push(r.url()); void r.abort('blockedbyclient'); } });
    page.on('pageerror', (e) => errors.push(e.message));
    await page.evaluateOnNewDocument(() => { const deny = () => { throw new Error('OFFLINE'); }; window.fetch = async () => deny(); window.WebSocket = deny; window.XMLHttpRequest = deny; });
    await page.goto(pathToFileURL(path.join(out, 'offline.html')).href, { waitUntil: 'load' });
    await page.waitForFunction(() => window.audit && window.audit.ready && document.querySelector('[role=dialog]'));
    await settle();
  }
  const calls = () => page.evaluate(() => window.audit.calls);
  const toasts = () => page.evaluate(() => window.audit.toasts);
  const has = (sel) => page.evaluate((s) => !!document.querySelector(s), sel);
  async function botao(texto) {
    return page.evaluate((t) => { const b = [...document.querySelectorAll('[role=dialog] button')].find((x) => x.textContent.includes(t)); return b ? { disabled: b.disabled, text: b.textContent } : null; }, texto);
  }
  async function clicar(texto) {
    const handle = await page.evaluateHandle((t) => [...document.querySelectorAll('[role=dialog] button')].find((x) => x.textContent.includes(t)), texto);
    assert(handle.asElement(), 'botão ausente: ' + texto);
    await handle.asElement().click();
    await settle();
  }
  async function enviarPlanilha() {
    const input = await page.$('#importar-decisoes-arquivo');
    const antes = (await calls()).length;
    await input.uploadFile(xlsx);
    try { await page.waitForFunction((n) => window.audit.calls.length >= n, {}, antes + 1); }
    catch (e) { const diag = await page.evaluate(() => ({ calls: window.audit.calls.length, toasts: window.audit.toasts, files: document.querySelector('#importar-decisoes-arquivo')?.files?.length, disabled: document.querySelector('#importar-decisoes-arquivo')?.disabled, texto: document.querySelector('[role=dialog]')?.textContent.slice(0, 300) })); throw new Error('upload sem analise: ' + JSON.stringify(diag)); }
    await settle();
  }
  async function responder(i, override) { await page.evaluate((i, o) => window.audit.respond(i, o), i, override || {}); await settle(); }
  async function falhar(i, err) { await page.evaluate((i, e) => window.audit.fail(i, e), i, err); await settle(); }
  async function previaPronta() {
    await enviarPlanilha();
    await responder(0);
    assert.equal(await has('[data-testid=previa-decisoes]'), true, 'prévia exibida');
    assert.equal((await botao('Aplicar')).disabled, false, 'Aplicar habilitado com prévia válida');
  }
  async function escolherConvenio(nome) {
    await page.click('[role=dialog] [role=combobox]');
    await page.waitForSelector('[role=option]');
    const opt = await page.evaluateHandle((n) => [...document.querySelectorAll('[role=option]')].find((o) => o.textContent.includes(n)), nome);
    await opt.asElement().click();
    await settle();
  }
  const semAplicacao = async () => assert.equal((await calls()).filter((c) => c.aplicar).length, 0, 'nenhuma aplicação enviada');

  const results = [];
  async function test(name, fn) {
    try { await fresh(); await fn(); results.push(name); console.log('PASS ' + name); }
    catch (e) { console.error('FAIL ' + name + ': ' + (e.stack || e.message).split('\n').slice(0, 4).join(' | ')); results.push('FAIL'); }
  }

  await test('troca_hospital_invalida_previa_e_reanalise_usa_novo_escopo', async () => {
    await previaPronta();
    await page.evaluate(() => window.audit.setEstab(2)); await settle();
    assert.equal(await has('[data-testid=previa-decisoes]'), false, 'prévia antiga oculta');
    assert.equal(await has('[data-testid=previa-desatualizada]'), true, 'aviso de reanálise');
    assert.equal((await botao('Aplicar')).disabled, true, 'Aplicar bloqueado após troca de hospital');
    await semAplicacao();
    await clicar('Reanalisar');
    const c = await calls();
    assert.deepEqual([c[1].estabelecimentoId, c[1].convenioId, c[1].aplicar], [2, 10, false]);
    await responder(1);
    await clicar('Aplicar');
    const ap = (await calls())[2];
    assert.deepEqual([ap.aplicar, ap.estabelecimentoId, ap.convenioId], [true, 2, 10]);
    assert.match(ap.tokenPrevia, /^2f10/, 'token da prévia do novo escopo');
  });

  await test('troca_convenio_invalida_previa', async () => {
    await previaPronta();
    await escolherConvenio('Convenio Sintetico B');
    assert.equal((await botao('Aplicar')).disabled, true, 'Aplicar bloqueado após troca de convênio');
    assert.equal(await has('[data-testid=previa-desatualizada]'), true);
    await escolherConvenio('Todos os convênios');
    assert.equal((await botao('Aplicar')).disabled, true, 'Aplicar segue bloqueado em "Todos"');
    await semAplicacao();
  });

  await test('resposta_atrasada_de_outro_escopo_descartada', async () => {
    await enviarPlanilha();
    await page.evaluate(() => window.audit.setEstab(2)); await settle();
    await responder(0);
    assert.equal(await has('[data-testid=previa-decisoes]'), false, 'resposta antiga não vira prévia');
    assert.equal((await botao('Aplicar')).disabled, true);
    await semAplicacao();
  });

  await test('cancelar_e_reabrir_sem_estado_antigo', async () => {
    await previaPronta();
    await escolherConvenio('Convenio Sintetico B');
    await clicar('Cancelar');
    assert.equal(await has('[role=dialog]'), false, 'modal fechado');
    await page.evaluate(() => window.audit.setOpen(true)); await settle();
    await page.waitForSelector('[role=dialog]');
    assert.equal(await has('[data-testid=previa-decisoes]'), false, 'sem prévia antiga');
    assert.equal(await page.evaluate(() => document.querySelector('[role=dialog]').textContent.includes('decisoes-sinteticas.xlsx')), false, 'sem arquivo antigo');
    assert.equal(await page.evaluate(() => document.querySelector('[role=dialog] [role=combobox]').textContent.includes('Convenio Sintetico A')), true, 'convênio volta ao inicial');
    assert.equal((await botao('Aplicar')).disabled, true);
    await semAplicacao();
  });

  await test('erro_interno_sanitizado_no_toast_e_dom', async () => {
    await previaPronta();
    await clicar('Aplicar');
    await falhar(1, { message: 'ER_NO_REFERENCED_ROW_2: fk decisoes_glosa host=db-interno user=app_rw', data: { code: 'INTERNAL_SERVER_ERROR' } });
    const t = await toasts();
    assert.deepEqual(t.at(-1), ['error', 'Não foi possível aplicar as decisões. Nenhuma alteração foi gravada.']);
    assert.equal(JSON.stringify(t).includes('ER_'), false);
    assert.equal(await page.evaluate(() => document.body.textContent.includes('ER_')), false);
    assert.equal(await page.evaluate(() => window.audit.imported), 0);
  });

  await test('conflito_exige_nova_analise', async () => {
    await previaPronta();
    await clicar('Aplicar');
    await falhar(1, { message: 'A prévia não corresponde mais ao hospital/convênio selecionado ou os itens foram alterados. Reanalise a planilha antes de aplicar.', data: { code: 'CONFLICT' } });
    assert.equal((await toasts()).at(-1)[0], 'error');
    assert.equal(await has('[data-testid=previa-decisoes]'), false, 'prévia descartada após conflito');
    assert.equal((await botao('Aplicar')).disabled, true);
  });

  await test('aplicacao_sem_confirmacao_total_nao_e_sucesso', async () => {
    await previaPronta();
    await clicar('Aplicar');
    await responder(1, { aplicados: 1 });
    assert.equal((await toasts()).some(([k]) => k === 'success'), false, 'sem toast de sucesso');
    assert.equal(await has('[role=dialog]'), true, 'modal permanece aberto');
    assert.equal((await botao('Aplicar')).disabled, true);
  });

  await test('sucesso_confirma_escopo_e_token_da_previa', async () => {
    await previaPronta();
    const token = (await calls()).length && await page.evaluate(() => window.audit.pending[0].input);
    await clicar('Aplicar');
    const ap = (await calls())[1];
    assert.deepEqual([ap.aplicar, ap.estabelecimentoId, ap.convenioId], [true, 1, 10]);
    assert.match(ap.tokenPrevia, /^1f10/);
    assert.deepEqual(ap.linhas, token.linhas);
    await responder(1);
    assert.equal((await toasts()).at(-1)[0], 'success');
    assert.equal(await has('[role=dialog]'), false);
    assert.equal(await page.evaluate(() => window.audit.imported), 1);
  });

  await browser.close();
  clearTimeout(timeout);
  fs.rmSync(temporary, { recursive: true, force: true });
  assert.deepEqual(requests, [], 'nenhuma requisição externa');
  assert.deepEqual(errors, [], 'sem erros de página: ' + errors.join(' | '));
  if (results.includes('FAIL')) { console.error('FAIL importarDecisoesGlosaDom'); process.exit(1); }
  console.log('PASS importarDecisoesGlosaDom');
})().catch((e) => {
  console.error('FAIL harness: ' + (e && e.stack || e));
  browser?.close().catch(() => {});
  process.exit(1);
});
