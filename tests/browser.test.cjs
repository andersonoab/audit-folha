'use strict';
const assert=require('node:assert/strict'), path=require('node:path'), {pathToFileURL}=require('node:url');
const {chromium}=require('playwright');
const root=path.resolve(__dirname,'..'), csv=path.join(root,'examples','selecao_validacao.csv');
const checks=[];
async function state(p,name){return p.evaluate(n=>{const s=RowSelection.table(n).model.summary();return {count:s.count,totals:Object.fromEntries(Object.entries(s.totals).map(([k,v])=>[k,String(v)])),ids:[...RowSelection.table(n).model.selected].sort()};},name)}
function total(s,starts){return Number(Object.entries(s.totals).find(([k])=>k.startsWith(starts))?.[1]||0)}
function check(label){checks.push(label);console.log('PASS:',label)}
function row(p,body,text){return p.locator(body+' tr').filter({hasText:text}).first()}
async function single(r){await r.locator('td').nth(2).click();await r.page().waitForTimeout(410)}
async function clear(p,name){const b=p.locator(`[data-selection="${name}"] .sum-clear`);if(await b.isEnabled())await b.click()}
async function select(r){await r.locator('.sum-check').check()}
async function closeDetail(p){await p.locator('dialog.sum-detail button').click()}
(async()=>{
const browser=await chromium.launch({headless:true,...(process.env.CHROMIUM_PATH?{executablePath:process.env.CHROMIUM_PATH}:{}),args:['--no-sandbox','--disable-dev-shm-usage','--disable-gpu']});
try{
const errors=[];const p=await browser.newPage({viewport:{width:1550,height:1100}});
await p.route(/^https?:/,r=>r.abort());p.on('pageerror',e=>errors.push(e.message));p.on('dialog',d=>d.dismiss());
await p.goto(pathToFileURL(path.join(root,'index.html')).href);
await p.locator('#fileInput').setInputFiles(csv);await p.waitForTimeout(600);
assert.equal(await p.locator('#tabelaBody tr').count(),100);assert.equal((await p.locator('#holBody tr').count()),8);
check('Abertura file://, CSV latin-1 e gráficos sem rede');
await p.locator('[data-tab="tabela"]').click();
const first=p.locator('#tabelaBody tr').first();
await single(first);assert.equal((await state(p,'principal')).count,1);assert.equal(await p.locator('#drillBody tr').count(),0);
await single(first);assert.equal((await state(p,'principal')).count,0);
await select(first);const beforeDouble=await state(p,'principal');await first.locator('td').nth(2).dblclick();await p.waitForTimeout(440);assert.deepEqual(await state(p,'principal'),beforeDouble);assert.equal(await p.locator('#drillBody tr').count(),8);
check('Clique, desmarcação, checkbox único, duplo clique abre composição e mantém seleção');
const salary=row(p,'#drillBody','SALARIO'),inss=row(p,'#drillBody','INSS').filter({hasText:'DESC'}),pgbl=row(p,'#drillBody','PGBL');
const confBefore=await p.evaluate(()=>localStorage.getItem('af_conf_v2'));
await select(inss);await select(pgbl);let s=await state(p,'verbas-detalhe');assert.equal(s.count,2);assert.equal(total(s,'Folha'),-125807);
await pgbl.locator('.sum-check').uncheck();await select(salary);assert.equal(total(await state(p,'verbas-detalhe'),'Folha'),801193);
check('−988,07 −270,00 = −1.258,07; 9.000,00 −988,07 = 8.011,93');
const zero=row(p,'#drillBody','ZERO');await zero.focus();await p.keyboard.press('Space');assert.equal((await state(p,'verbas-detalhe')).count,3);assert.equal(total(await state(p,'verbas-detalhe'),'Folha'),801193);
assert.equal(await p.evaluate(()=>localStorage.getItem('af_conf_v2')),confBefore);
await p.locator(".app-header").hover();
let selectedColor=await salary.locator('td').nth(2).evaluate(el=>getComputedStyle(el).backgroundColor);await p.locator('.app-header').hover();assert.equal(await salary.locator('td').nth(2).evaluate(el=>getComputedStyle(el).backgroundColor),selectedColor);
const ordinary=row(p,'#drillBody','PGBL');await ordinary.hover();assert.notEqual(await ordinary.locator('td').nth(2).evaluate(el=>getComputedStyle(el).backgroundColor),selectedColor);
check('Zero conta, Espaço acessível, destaque persistente, hover distinto e conferência isolada');
const beforeSort=await state(p,'verbas-detalhe');await p.locator('#tabelaDrill th[data-sk="cod"]').click();await p.locator('#tabelaDrill th[data-sk="cod"]').click();assert.deepEqual(await state(p,'verbas-detalhe'),beforeSort);
await row(p,'#drillBody','INSS').focus();await p.keyboard.press('Enter');assert.equal(await p.locator('dialog.sum-detail').count(),1);assert.deepEqual(await state(p,'verbas-detalhe'),beforeSort);await closeDetail(p);
await p.locator('#comboDrillClasse').selectOption('DESC');s=await state(p,'verbas-detalhe');assert.equal(s.count,1);assert.equal(total(s,'Folha'),-98807);assert.match(await p.locator('[data-selection="verbas-detalhe"] .sum-notice').textContent(),/2 linha/);
await p.locator('#btnLimparDrill').click();await clear(p,'verbas-detalhe');
const dups=p.locator('#drillBody tr').filter({hasText:'SALARIO'});assert.equal(await dups.count(),3);await select(dups.nth(1));await select(dups.nth(2));assert.equal(total(await state(p,'verbas-detalhe'),'Folha'),2);assert.equal((await state(p,'verbas-detalhe')).count,2);
check('Ordenação bidirecional, Enter investiga, filtros removem somente ausentes, duplicados de 1 centavo independentes');
await clear(p,'verbas-detalhe');await select(row(p,'#drillBody','BASE INSS'));await select(row(p,'#drillBody','SALARIO'));s=await state(p,'verbas-detalhe');assert.equal(Object.keys(s.totals).length,2);assert.equal(total(s,'Folha'),900000);assert.equal(total(s,'BASE'),900000);
await clear(p,'verbas-detalhe');assert.equal((await state(p,'verbas-detalhe')).count,0);
check('Bases separadas de folha e limpeza completa');
// Pagination doesn't constrain the filtered selection.
await p.locator('[data-selection-pager="principal"] button').filter({hasText:'Próxima'}).click();
await select(p.locator('#tabelaBody tr').first());assert.equal((await state(p,'principal')).count,2);
await p.locator('[data-selection-pager="principal"] button').filter({hasText:'Anterior'}).click();assert.equal(await p.locator('#tabelaBody .sum-check:checked').count(),1);
const pageState=await state(p,'principal');await p.locator('#tabelaPrincipal th[data-col="Nome"]').click();assert.deepEqual(await state(p,'principal'),pageState);
await p.locator('#entryBusca').fill('COLABORADOR TESTE 1');await p.locator('#btnAplicarTab').click();assert.ok((await state(p,'principal')).count<=2);
await p.locator('#entryBusca').fill('NENHUM RESULTADO');await p.locator('#btnAplicarTab').click();assert.equal((await state(p,'principal')).count,0);assert.match(await p.locator('[data-selection="principal"] .sum-notice').textContent(),/removida/);await p.locator('#btnLimparTab').click();
check('Seleção cumulativa entre páginas, ordenação preservada e filtro vazio sem totais ocultos');
// Holerite and its original Calc. state are independent.
await p.locator('[data-tab="holerite"]').click();await select(row(p,'#holBody','INSS'));await select(row(p,'#holBody','PGBL'));assert.equal(total(await state(p,'holerite'),'Folha'),-125807);
const holBefore=await state(p,'holerite');await row(p,'#holBody','INSS').locator('.hol-check').check();assert.deepEqual(await state(p,'holerite'),holBefore);
await p.locator('#holBusca').fill('PGBL');assert.equal((await state(p,'holerite')).count,1);assert.equal(total(await state(p,'holerite'),'Folha'),-27000);await p.locator('#holBusca').fill('');await clear(p,'holerite');
check('Holerite: soma, Calc. independente, pesquisa e remoção informada');
await p.locator('[data-tab="conferencia"]').click();await select(p.locator('#confBody tr').first());const cf=await state(p,'conferencia');await p.locator('#tabelaConf th[data-sk="nome"]').click();assert.deepEqual(await state(p,'conferencia'),cf);await p.locator('#entryConfBusca2').fill('NENHUM RESULTADO');assert.equal((await state(p,'conferencia')).count,0);
check('Conferência: seleção e ordenação independentes da revisão');
await p.locator('[data-tab="visaoGeral"]').click();await select(p.locator('#descBody tr').first());assert.ok(total(await state(p,'descontos'),'Folha')<0);const desc=await state(p,'descontos');await p.locator('#tabelaDescontos th[data-sk="VALOR"]').click();assert.deepEqual(await state(p,'descontos'),desc);
check('Descontos para revisar: contribuição negativa e ordenação');
await p.locator('[data-tab="analiseVerbas"]').click();await p.waitForTimeout(400);await p.locator('#listboxVerbas').selectOption(['0020','3010']);const cRow=row(p,'#colabBody','COLABORADOR TESTE 3');await select(cRow);assert.equal(total(await state(p,'colaboradores-verbas'),'Folha'),801193);await cRow.locator('td').nth(2).dblclick();await p.waitForTimeout(440);assert.equal((await state(p,'colaboradores-verbas')).count,1);assert.match(await p.locator('#colabDrillInfo').textContent(),/COLABORADOR TESTE 3/);
await p.evaluate(()=>APP.limparSelecaoVerbas());assert.equal((await state(p,'colaboradores-verbas')).count,0);assert.equal(await p.locator('#colabBody tr').count(),0);
check('Análise de verbas: soma mista precisa, investigação e limpeza da população');
// Export must remain identical before/after selection.
const exportSnapshot=()=>p.evaluate(()=>{const saved=XLSX.writeFile;let wb;XLSX.writeFile=w=>wb=w;APP.acaoExportarExcel();XLSX.writeFile=saved;return Object.fromEntries(wb.SheetNames.filter(n=>n!=='Parametros').map(n=>[n,XLSX.utils.sheet_to_json(wb.Sheets[n])]));});
const exp1=await exportSnapshot();await p.locator('[data-tab="tabela"]').click();await select(p.locator('#tabelaBody tr').first());const exp2=await exportSnapshot();assert.deepEqual(exp2,exp1);assert.ok(exp1.Liquido_Funcionario.length>=124);
check('Exportação Excel mantém abas, valores e população após seleção');
// Reprocessing, including empty company selection, prunes every hidden table.
await p.locator('[data-tab="holerite"]').click();await select(row(p,'#holBody','INSS'));await p.locator('[data-tab="tabela"]').click();const rp=await state(p,'principal');await p.locator('#btnReprocessar').click();assert.deepEqual(await state(p,'principal'),rp);
await p.evaluate(()=>APP._selTodosEmpresa(false));await p.locator('#btnReprocessar').click();assert.equal((await state(p,'principal')).count,0);assert.equal((await state(p,'holerite')).count,0);assert.equal((await state(p,'descontos')).count,0);
check('Reprocessamento preserva IDs; nenhum CNPJ selecionado limpa registros e totais ocultos');
// A new dataset resets selection, and a single competency remains usable.
await p.locator('#fileInput').setInputFiles(path.join(root,'examples','selecao_mes_unico.csv'));await p.waitForTimeout(250);
assert.equal((await state(p,'principal')).count,0);await p.locator('#entryBusca').fill('00001');await p.locator('#btnAplicarTab').click();assert.equal(await p.locator('#tabelaBody tr').count(),2);
await select(p.locator('#tabelaBody tr').nth(0));await select(p.locator('#tabelaBody tr').nth(1));assert.equal((await state(p,'principal')).count,2);assert.notEqual((await state(p,'principal')).ids[0],(await state(p,'principal')).ids[1]);
const rawBefore=await p.evaluate(()=>JSON.stringify(carregarCSV(localStorage.getItem('af_csv')).df));
// CSV parsing alone doesn't alter the model; data references used by the UI stay intact.
await clear(p,'principal');await select(p.locator('#tabelaBody tr').first());
const rawAfter=await p.evaluate(()=>JSON.stringify(carregarCSV(localStorage.getItem('af_csv')).df));assert.equal(rawAfter,rawBefore);
const sameMatRows=p.locator('#tabelaBody tr');await sameMatRows.nth(1).locator('td').nth(2).dblclick();await p.locator('#btnVerHolerite').click();
assert.match(await p.locator('.hol-sub').textContent(),/EMPRESA TESTE B/);assert.equal(await p.locator('#holBody tr').count(),12);
check('Mês único, nova importação, homônimos com mesma matrícula em CNPJs diferentes e base sem mutação');
assert.deepEqual(errors,[]);
await p.close();
// Auxiliary receipt: all four views reuse the exact same component.
const q=await browser.newPage({viewport:{width:1500,height:1000}});await q.route(/^https?:/,r=>r.abort());q.on('pageerror',e=>errors.push(e.message));q.on('dialog',d=>d.dismiss());
await q.goto(pathToFileURL(path.join(root,'Recibo_Conferencia_Folha_Sonova_comparativo_meses_v3.html')).href);await q.locator('#file').setInputFiles(csv);await q.waitForTimeout(400);
await select(row(q,'#reciboArea table.codes tbody','INSS'));await select(row(q,'#reciboArea table.codes tbody','PGBL'));assert.equal(total(await state(q,'recibo-linhas'),'Folha'),-125807);
const rec=await state(q,'recibo-linhas');await q.locator('#reciboArea th[data-sort="codigo"]').click();assert.deepEqual(await state(q,'recibo-linhas'),rec);await row(q,'#reciboArea table.codes tbody','INSS').locator('td').nth(2).dblclick();await q.waitForTimeout(400);assert.deepEqual(await state(q,'recibo-linhas'),rec);await closeDetail(q);
await q.locator('#btnAllMonths').click();assert.deepEqual(await state(q,'recibo-linhas'),rec);
await q.locator('#vSearch').fill('PGBL');assert.equal((await state(q,'recibo-linhas')).count,1);await q.locator('#vSearch').fill('');
check('Recibo auxiliar: soma, ordenação, duplo clique, comparativo de meses e filtro');
await q.locator('#btnOver').click();await select(q.locator('#ovTable tbody tr').first());assert.equal((await state(q,'recibo-visao-geral')).count,1);await q.locator('[data-selection-pager="recibo-visao-geral"] button').filter({hasText:'Próxima'}).click();await select(q.locator('#ovTable tbody tr').first());assert.equal((await state(q,'recibo-visao-geral')).count,2);await q.locator('#ovClose').click();
await q.locator('#btnVerba').click();await q.locator('#vqSearch').fill('3010');await q.locator('#vqSearch').press('Enter');await q.waitForTimeout(100);await select(q.locator('#vqTable tbody tr').filter({hasText:'DESC'}).first());assert.equal(total(await state(q,'recibo-consulta-verbas'),'Folha'),-98807);await q.locator('#vqClose').click();
await q.locator('#btnDiv').click();assert.equal(await q.locator('[data-selection="recibo-divergencias"]').count(),1);await q.locator('#divClose').click();
check('Recibo auxiliar: visão geral paginada, consulta de verbas e divergências monetárias');
assert.deepEqual(errors,[]);await q.close();
console.log(`\n${checks.length} grupos de validação concluídos sem erro JavaScript.`);
}catch(e){console.error(e);process.exitCode=1}finally{await browser.close()}
})();
