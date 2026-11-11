/* Processos APVS — gera a planilha no MESMO formato do DASHBOARD (modelo.xlsx),
   preenchendo só as células de dados. Abas, fórmulas, formatação e cores ficam intactas. */
(function (root) {
  "use strict";

  // Onde ficam os dados no modelo
  const BASE = { aba: "BASE DE DADOS", primeira: 7, colunas: ["B","C","D","E","F","J","K","L","M","N","P"] };
  const DIARIO = { aba: "APVS", primeira: 16, ultima: 36 };          // C=data, D=processos, F=redução, L=obs
  const DIVERG = { aba: "DIVERGÊNCIAS", primeira: 7, colunas: ["B","C","D","E","G"] };
  const CONFIG = { aba: "CONFIGURAÇÕES", col: "C", de: 5, ate: 60 };   // feriados

  const SEM_ERRO = "Sem Erro";

  /* ---------- datas no formato do Excel (número de série, hora local) ---------- */
  const EPOCH = Date.UTC(1899, 11, 30);
  function serialDe(y, m, d, h, mi, s) { return (Date.UTC(y, m - 1, d, h || 0, mi || 0, s || 0) - EPOCH) / 864e5; }
  function serialTexto(t) { // "2026-11-10T09:23" ou "2026-11-10T09:23:00"
    if (!t) return null;
    const m = String(t).match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/);
    return m ? serialDe(+m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)) : null;
  }
  function serialMs(ms) {
    if (ms == null) return null;
    const d = new Date(ms);
    return serialDe(d.getFullYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds());
  }
  const diaMs = ms => { const d = new Date(ms); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; };
  const r2 = n => Math.round(n * 100) / 100;

  /* ---------- XML ---------- */
  const xmlEsc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  const reEsc = s => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

  function attr(tag, nome) { const m = tag.match(new RegExp("\\s" + reEsc(nome) + '="([^"]*)"')); return m ? m[1] : null; }

  async function caminhosDasAbas(zip) {
    const wb = await zip.file("xl/workbook.xml").async("string");
    const rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    const alvo = {};
    (rels.match(/<Relationship\b[^>]*>/g) || []).forEach(t => { alvo[attr(t, "Id")] = attr(t, "Target"); });
    const mapa = {};
    (wb.match(/<sheet\b[^>]*>/g) || []).forEach(t => {
      const nome = attr(t, "name").replace(/&amp;/g, "&");
      let tg = alvo[attr(t, "r:id")] || "";
      tg = tg.replace(/^\//, "");
      mapa[nome] = tg.startsWith("xl/") ? tg : "xl/" + tg;
    });
    return { mapa, wb };
  }

  // Aplica valores em células existentes. valores: { "7": { B: "texto", C: 46265.5, J: null }, ... }
  function aplicar(xml, valores) {
    return xml.replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, row => {
      const rn = (row.match(/^<row\b[^>]*?\sr="(\d+)"/) || [])[1];
      const cels = rn && valores[rn];
      if (!cels) return row;
      for (const col in cels) {
        const ref = col + rn;
        const m = row.match(new RegExp('<c r="' + ref + '"([^>]*?)(?:/>|>[\\s\\S]*?</c>)'));
        if (!m) continue;
        const attrs = m[1].replace(/\s+t="[^"]*"/, "");
        const v = cels[col];
        let novo;
        if (v === null || v === undefined || v === "") novo = `<c r="${ref}"${attrs}/>`;
        else if (typeof v === "number") novo = `<c r="${ref}"${attrs}><v>${v}</v></c>`;
        else novo = `<c r="${ref}"${attrs} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
        row = row.replace(m[0], () => novo);
      }
      return row;
    });
  }

  // Tira os resultados guardados das fórmulas, para o Excel/Google recalcularem ao abrir
  function limparCacheFormulas(xml) {
    return xml
      .replace(/<c\b([^>]*?)>(<f\b[^>]*(?:\/>|>[\s\S]*?<\/f>))<v>[^<]*<\/v><\/c>/g, (m, a, f) => `<c${a.replace(/\s+t="[^"]*"/, "")}>${f}</c>`);
  }

  function ultimaLinha(xml) {
    let max = 0; xml.replace(/<row\b[^>]*?\sr="(\d+)"/g, (m, n) => { max = Math.max(max, +n); return m; });
    return max;
  }

  function feriados(xml) {
    const set = new Set();
    xml.replace(new RegExp('<c r="' + CONFIG.col + '(\\d+)"[^>]*?>(?:<f[\\s\\S]*?</f>)?<v>([\\d.]+)</v>', "g"), (m, rn, v) => {
      if (+rn >= CONFIG.de && +rn <= CONFIG.ate) set.add(Math.floor(+v));
      return m;
    });
    return set;
  }

  function diasUteis(ano, mes, fer) {
    const out = [];
    const n = new Date(ano, mes, 0).getDate();
    for (let d = 1; d <= n; d++) {
      const wd = new Date(ano, mes - 1, d).getDay();
      const s = Math.floor(serialDe(ano, mes, d));
      if (wd >= 1 && wd <= 5 && !fer.has(s)) out.push({ d, chave: `${ano}-${String(mes).padStart(2, "0")}-${String(d).padStart(2, "0")}`, serial: serialDe(ano, mes, d, 9, 0) });
    }
    return out;
  }

  /* ---------- regras de negócio ---------- */
  const auditoriaDe = p => !p.erro ? "" : (p.erro === SEM_ERRO ? "Aprovado" : "Reprovado");
  const obsDe = p => [p.erro2 ? "2º erro: " + p.erro2 : "", p.descricaoErro || ""].filter(Boolean).join(" · ");
  const revisado = p => (p.status === "comunicar" || p.status === "comunicado") && p.revisadoEm;
  // mês do processo na BASE: pela liberação do regulador; sem ela, pela revisão; sem as duas, pelo registro
  function mesDoProcesso(p) {
    if (p.regulacao) return String(p.regulacao).slice(0, 7);
    if (p.revisadoEm) return diaMs(p.revisadoEm).slice(0, 7);
    return p.criadoEm ? diaMs(p.criadoEm).slice(0, 7) : "";
  }

  function resumo(processos, mesISO) {
    const base = processos.filter(p => mesDoProcesso(p) === mesISO)
      .sort((a, b) => (serialTexto(a.regulacao) ?? serialMs(a.revisadoEm) ?? 1e9) - (serialTexto(b.regulacao) ?? serialMs(b.revisadoEm) ?? 1e9));
    const porDia = {};
    processos.filter(p => revisado(p) && diaMs(p.revisadoEm).slice(0, 7) === mesISO).forEach(p => {
      const k = diaMs(p.revisadoEm); const o = porDia[k] || (porDia[k] = { n: 0, red: 0 });
      o.n++; o.red += p.reducao || 0;
    });
    return { base, porDia };
  }

  async function gerar(JSZip, modelo, processos, mesISO, agora) {
    agora = agora || new Date();
    const [ano, mes] = mesISO.split("-").map(Number);
    const zip = await JSZip.loadAsync(modelo);
    const { mapa, wb } = await caminhosDasAbas(zip);
    for (const nome of [BASE.aba, DIARIO.aba]) if (!mapa[nome] || !zip.file(mapa[nome])) throw new Error(`O modelo não tem a aba "${nome}".`);
    const { base, porDia } = resumo(processos, mesISO);

    // BASE DE DADOS
    let xb = await zip.file(mapa[BASE.aba]).async("string");
    const fim = ultimaLinha(xb);
    const capacidade = fim - BASE.primeira + 1;
    if (base.length > capacidade) throw new Error(`O modelo tem espaço para ${capacidade} processos e o mês tem ${base.length}. Aumente as linhas da BASE DE DADOS no modelo.`);
    const vb = {};
    for (let r = BASE.primeira; r <= fim; r++) {
      const p = base[r - BASE.primeira];
      const linha = {}; BASE.colunas.forEach(c => linha[c] = null);
      if (p) Object.assign(linha, {
        B: p.protocolo, C: serialTexto(p.entrada), D: serialTexto(p.regulacao), E: (p.regulador || "").toUpperCase() || null, F: "APVS",
        J: auditoriaDe(p) || null, K: p.erro || null, L: p.porte || null, N: p.valorAprovado ?? null, P: obsDe(p) || null
      });
      vb[r] = linha;
    }
    xb = limparCacheFormulas(aplicar(xb, vb));
    zip.file(mapa[BASE.aba], xb);

    // APVS (diário)
    const fer = mapa[CONFIG.aba] && zip.file(mapa[CONFIG.aba]) ? feriados(await zip.file(mapa[CONFIG.aba]).async("string")) : new Set();
    const dias = diasUteis(ano, mes, fer);
    const hojeK = diaMs(agora.getTime());
    const slots = DIARIO.ultima - DIARIO.primeira + 1;
    // revisões feitas fora de dia útil (fim de semana/feriado) entram no dia útil anterior (ou no primeiro do mês)
    const tot = dias.map(() => ({ n: 0, red: 0 }));
    Object.keys(porDia).sort().forEach(k => {
      let i = dias.findIndex(x => x.chave === k);
      if (i < 0) { i = -1; dias.forEach((x, j) => { if (x.chave < k) i = j; }); if (i < 0) i = 0; }
      tot[i].n += porDia[k].n; tot[i].red += porDia[k].red;
    });
    const vd = {}; let nota = "";
    for (let s = 0; s < slots; s++) vd[DIARIO.primeira + s] = { C: null, D: null, F: null, L: null };
    dias.forEach((x, i) => {
      const s = Math.min(i, slots - 1), r = DIARIO.primeira + s;
      const passou = x.chave <= hojeK;
      if (i < slots) vd[r].C = x.serial;
      if (passou || tot[i].n) {
        vd[r].D = (vd[r].D || 0) + tot[i].n;
        vd[r].F = r2((vd[r].F || 0) + tot[i].red);
      }
      if (i >= slots) nota += (nota ? ", " : "Inclui ") + String(x.d).padStart(2, "0") + "/" + String(mes).padStart(2, "0");
    });
    if (nota) vd[DIARIO.ultima].L = nota;
    let xd = await zip.file(mapa[DIARIO.aba]).async("string");
    xd = limparCacheFormulas(aplicar(xd, vd));
    zip.file(mapa[DIARIO.aba], xd);

    // DIVERGÊNCIAS: o sistema não registra divergências; a aba sai vazia
    if (mapa[DIVERG.aba] && zip.file(mapa[DIVERG.aba])) {
      let xv = await zip.file(mapa[DIVERG.aba]).async("string");
      const fv = ultimaLinha(xv), vv = {};
      for (let r = DIVERG.primeira; r <= fv; r++) { vv[r] = {}; DIVERG.colunas.forEach(c => vv[r][c] = null); }
      zip.file(mapa[DIVERG.aba], limparCacheFormulas(aplicar(xv, vv)));
    }

    // Recalcular tudo ao abrir
    let wb2 = wb.replace(/<calcPr\b[^>]*\/>/, '<calcPr fullCalcOnLoad="1"/>');
    if (!/<calcPr/.test(wb2)) wb2 = wb2.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>');
    zip.file("xl/workbook.xml", wb2);

    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 },
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  root.PlanilhaAPVS = { gerar, resumo, diasUteis, mesDoProcesso };
})(typeof window !== "undefined" ? window : globalThis);
