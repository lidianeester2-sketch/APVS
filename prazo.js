/* Processos APVS — regra de prazo (mesma lógica do VOI)
   APVS: 3 dias úteis a partir da ENTRADA. Expediente: 09h–18h, segunda a sexta, sem feriados nacionais.
   Ex.: entrada segunda 10:00 → vence quinta 10:00. Entrada fora do expediente começa a contar no próximo horário útil. */
(function (root) {
  "use strict";
  const INICIO = 9, FIM = 18, DIAS_APVS = 3;
  // Feriados nacionais (mesma lista da aba CONFIGURAÇÕES da planilha). Acrescente os próximos anos aqui.
  const FERIADOS = new Set([
    "2025-01-01","2025-03-03","2025-03-04","2025-04-18","2025-04-21","2025-05-01","2025-06-19","2025-09-07","2025-10-12","2025-11-02","2025-11-15","2025-11-20","2025-12-25",
    "2026-01-01","2026-02-16","2026-02-17","2026-04-03","2026-04-21","2026-05-01","2026-06-04","2026-09-07","2026-10-12","2026-11-02","2026-11-15","2026-11-20","2026-12-25",
    "2027-01-01","2027-02-08","2027-02-09","2027-03-26","2027-04-21","2027-05-01","2027-05-27","2027-09-07","2027-10-12","2027-11-02","2027-11-15","2027-11-20","2027-12-25"
  ]);
  const z = n => String(n).padStart(2, "0");
  const chave = d => `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`;
  const diaUtil = d => { const w = d.getDay(); return w >= 1 && w <= 5 && !FERIADOS.has(chave(d)); };
  const proximoDia = d => { const n = new Date(d); n.setDate(n.getDate() + 1); n.setHours(INICIO, 0, 0, 0); return n; };

  // leva o horário para dentro do expediente
  function normalizar(d) {
    d = new Date(d);
    for (let i = 0; i < 400; i++) {
      if (!diaUtil(d)) { d = proximoDia(d); continue; }
      const h = d.getHours() + d.getMinutes() / 60;
      if (h < INICIO) { d.setHours(INICIO, 0, 0, 0); return d; }
      if (h >= FIM) { d = proximoDia(d); continue; }
      return d;
    }
    return d;
  }

  function somarHorasUteis(inicio, horas) {
    let d = normalizar(inicio), resta = horas * 60;
    for (let i = 0; i < 400 && resta > 0; i++) {
      const fimDia = new Date(d); fimDia.setHours(FIM, 0, 0, 0);
      const livre = (fimDia - d) / 6e4;
      if (resta <= livre) return new Date(d.getTime() + resta * 6e4);
      resta -= livre; d = normalizar(proximoDia(d));
    }
    return d;
  }

  function lerData(iso) { // "2026-11-11T09:16" (hora local)
    const m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/);
    return m ? new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) : null;
  }

  function vencimento(entradaIso) {
    const e = lerData(entradaIso);
    return e ? somarHorasUteis(e, DIAS_APVS * (FIM - INICIO)) : null;
  }

  const EM_ANDAMENTO = ["revisao", "devolvido", "comunicar"];
  // null = sem prazo (já comunicado ou sem data de entrada)
  function situacao(p, agora) {
    if (!p || !EM_ANDAMENTO.includes(p.status) || !p.entrada) return null;
    const v = vencimento(p.entrada); if (!v) return null;
    agora = agora || new Date();
    if (agora > v) return { nivel: "atrasado", vence: v };
    if (chave(v) === chave(agora)) return { nivel: "hoje", vence: v };
    return { nivel: "ok", vence: v };
  }

  // atrasados primeiro, depois os que vencem hoje, depois pelo vencimento mais próximo; sem data por último
  const PESO = { atrasado: 0, hoje: 1, ok: 2 };
  function comparar(a, b, agora) {
    const sa = situacao(a, agora), sb = situacao(b, agora);
    const pa = sa ? PESO[sa.nivel] : 3, pb = sb ? PESO[sb.nivel] : 3;
    if (pa !== pb) return pa - pb;
    if (sa && sb) return sa.vence - sb.vence;
    return 0;
  }

  root.PrazoAPVS = { vencimento, situacao, comparar, FERIADOS };
})(typeof window !== "undefined" ? window : globalThis);
