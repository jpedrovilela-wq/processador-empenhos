/* Processamento local no navegador. Não há upload de arquivos. */
const CORH = "ajustada CORH", SALDOS = "SALDOS DE EMPENHO", MARCADOR = "LINHA ACRESCENTADA +1 EMPENHO";
const LINHA_CABECALHO_CORH = 0, LINHA_INICIO_CORH = 1, LINHA_INICIO_SALDOS = 6;
const $ = (id) => document.getElementById(id);
let arquivo;

$("arquivo").addEventListener("change", (e) => {
  arquivo = e.target.files[0];
  $("nomeArquivo").textContent = arquivo ? arquivo.name : "Selecionar planilha Excel (.xlsx)";
  $("processar").disabled = !arquivo;
  $("status").textContent = arquivo ? "Arquivo selecionado. Escolha uma operação." : "Selecione uma planilha para começar.";
});
$("processar").onclick = executar;

function txt(v) { return v == null ? "" : String(v).trim(); }
function num(v) {
  if (typeof v === "number") return v;
  let s = txt(v).replace("R$", "").replace(/\s/g, "");
  if (s.includes(",") && s.includes(".")) s = s.replace(/\./g, "").replace(",", "."); else s = s.replace(",", ".");
  return Number(s) || 0;
}
function anoNota(nota) { const a = nota.length >= 15 ? Number(nota.slice(11, 15)) : 2020; return a >= 2000 && a <= 2100 ? a : 2020; }
function cell(ws, r, c) { return ws[XLSX.utils.encode_cell({r, c})]; }
function valor(ws, r, c) { return cell(ws, r, c)?.v; }
function set(ws, r, c, v, fmt) { const a = XLSX.utils.encode_cell({r, c}); ws[a] = { ...(ws[a] || {}), t: typeof v === "number" ? "n" : "s", v }; if (fmt) ws[a].z = fmt; }
function ref(ws) { return XLSX.utils.decode_range(ws["!ref"] || "A1:A1"); }
function expandir(ws, r, c) { const x = ref(ws); x.e.r = Math.max(x.e.r, r); x.e.c = Math.max(x.e.c, c); ws["!ref"] = XLSX.utils.encode_range(x); }
function clonar(obj) { return obj ? JSON.parse(JSON.stringify(obj)) : undefined; }
const BORDA_FINA = { style: "thin", color: { rgb: "000000" } };
function estilizarCabecalho(ws, linha) {
  const area = ref(ws);
  for (let c = 0; c <= area.e.c; c++) {
    const chave = XLSX.utils.encode_cell({ r: linha, c });
    if (!ws[chave]) ws[chave] = { t: "s", v: "" };
    ws[chave].s = {
      ...(ws[chave].s || {}),
      fill: { patternType: "solid", fgColor: { rgb: "C8C8C8" } },
      font: { ...((ws[chave].s || {}).font || {}), bold: true },
      alignment: { horizontal: "center", vertical: "center" },
      border: { top: BORDA_FINA, bottom: BORDA_FINA, left: BORDA_FINA, right: BORDA_FINA }
    };
  }
}
function bordasCompletas(ws) {
  const area = ref(ws);
  for (let r = area.s.r; r <= area.e.r; r++) for (let c = area.s.c; c <= area.e.c; c++) {
    const chave = XLSX.utils.encode_cell({ r, c });
    if (!ws[chave]) ws[chave] = { t: "s", v: "" };
    ws[chave].s = { ...(ws[chave].s || {}), border: { top: BORDA_FINA, bottom: BORDA_FINA, left: BORDA_FINA, right: BORDA_FINA } };
  }
}

function inserirLinha(ws, destino) {
  const area = ref(ws);
  for (let r = area.e.r; r >= destino; r--) for (let c = 0; c <= area.e.c; c++) {
    const de = XLSX.utils.encode_cell({r, c}), para = XLSX.utils.encode_cell({r: r + 1, c});
    if (ws[de]) ws[para] = clonar(ws[de]); else delete ws[para];
  }
  area.e.r++; ws["!ref"] = XLSX.utils.encode_range(area);
}
function copiarLinha(ws, origem, destino) {
  inserirLinha(ws, destino); const area = ref(ws);
  for (let c = 0; c <= area.e.c; c++) {
    const de = XLSX.utils.encode_cell({r: origem, c}), para = XLSX.utils.encode_cell({r: destino, c});
    if (ws[de]) ws[para] = clonar(ws[de]); else delete ws[para];
  }
}
function preencher(ws, r, ate, cor) { for (let c = 0; c <= ate; c++) { const x = cell(ws,r,c) || (ws[XLSX.utils.encode_cell({r,c})] = {t:"s",v:""}); x.s = {...(x.s||{}), fill:{patternType:"solid",fgColor:{rgb:cor}}}; } }

function processar(wb) {
  if (!wb.SheetNames.includes(CORH) || !wb.SheetNames.includes(SALDOS)) throw new Error("São necessárias as abas 'ajustada CORH' e 'SALDOS DE EMPENHO'.");
  const corh = wb.Sheets[CORH], saldos = wb.Sheets[SALDOS];
  [[14,"Saldo utilizado NE"],[15,"Saldo remanescente NE"],[16,"Controle Macro"]].forEach(([c,t]) => { if (!valor(corh,LINHA_CABECALHO_CORH,c)) set(corh,LINHA_CABECALHO_CORH,c,t); });
  const log = XLSX.utils.aoa_to_sheet([["Linha","Contrato","Município","Valor faltante","Motivo"]]); wb.Sheets.LOG = log; if (!wb.SheetNames.includes("LOG")) wb.SheetNames.push("LOG");
  const mapa = new Map(); const sr = ref(saldos);
  for (let r=LINHA_INICIO_SALDOS;r<=sr.e.r;r++) { const contrato=txt(valor(saldos,r,0)), nota=txt(valor(saldos,r,8)); if(!contrato||!nota) continue; const saldoQ=num(valor(saldos,r,16)), saldoO=num(valor(saldos,r,14)), saldo=saldoQ>0?saldoQ:saldoO, ano=anoNota(nota); if(saldo<=0 || ![2025,2026].includes(ano))continue; if(!mapa.has(contrato))mapa.set(contrato,[]); mapa.get(contrato).push({nota,fonte:txt(valor(saldos,r,6)),saldo,ano}); }
  mapa.forEach(x=>x.sort((a,b)=>a.ano-b.ano)); const restantes=new Map(); let inseridas=0, usados=0, nao=0;
  for(let r=ref(corh).e.r;r>=LINHA_INICIO_CORH;r--) { const contrato=txt(valor(corh,r,0)); if(!contrato||txt(valor(corh,r,7)))continue; const municipio=txt(valor(corh,r,2)), total=num(valor(corh,r,12)); if(total<=0)continue; const candidatos=mapa.get(contrato); if(!candidatos){preencher(corh,r,15,"FFFF00"); XLSX.utils.sheet_add_aoa(log,[[r+1,contrato,municipio,total,"Contrato não encontrado"]],{origin:-1});nao++;continue;} let restante=total, alocou=false, dest=r;
    for(const e of candidatos){if(restante<=.00001)break; const saldo=restantes.has(e.nota)?restantes.get(e.nota):e.saldo;if(saldo<=0)continue;const usar=Math.min(restante,saldo);if(alocou){dest++;copiarLinha(corh,r,dest);inseridas++;set(corh,r,16,"");set(corh,dest,16,MARCADOR);}set(corh,dest,7,e.nota.slice(11));set(corh,dest,9,e.fonte);set(corh,dest,11,e.ano>=2026?3:2);set(corh,dest,14,usar,"#,##0.00");set(corh,dest,15,saldo-usar,"#,##0.00");restantes.set(e.nota,saldo-usar);restante-=usar;usados++;alocou=true;}
    if(alocou&&restante>.01){preencher(corh,r,15,"FFC864");XLSX.utils.sheet_add_aoa(log,[[r+1,contrato,municipio,restante,"Saldo insuficiente"]],{origin:-1});}
  }
  estilizarCabecalho(corh, LINHA_CABECALHO_CORH);
  estilizarCabecalho(saldos, LINHA_INICIO_SALDOS - 1);
  estilizarCabecalho(log, 0);
  bordasCompletas(log);
  return {inseridas,usados,nao,inconsistencias: XLSX.utils.sheet_to_json(log,{header:1}).length - 1};
}
async function executar() {
  try { $("processar").disabled=true;$("status").textContent="Processando no seu navegador…";const dados=await arquivo.arrayBuffer(),wb=XLSX.read(dados,{type:"array",cellStyles:true});const resultado=processar(wb);const avisoLog=resultado.inconsistencias>0?` Foram encontradas ${resultado.inconsistencias} inconsistência(s); consulte a aba LOG do arquivo baixado.`:" Nenhuma inconsistência foi registrada na aba LOG.";$("status").textContent=`Concluído: ${resultado.inseridas} linha(s) inserida(s), ${resultado.usados} empenho(s) usado(s), ${resultado.nao} ocorrência(s) sem correspondência.${avisoLog} Selecione outro arquivo para iniciar um novo ciclo.`;const nome=arquivo.name.replace(/\.xlsx$/i,"")+"_processado.xlsx";XLSX.writeFile(wb,nome,{cellStyles:true});} catch(e) { console.error(e);$("status").textContent=`Erro: ${e.message}`;} finally { arquivo = undefined; $("arquivo").value = ""; $("nomeArquivo").textContent = "Selecionar planilha Excel (.xlsx)"; $("processar").disabled = true; }
}
