import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://dpkoudaggtvaujtzypeo.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRwa291ZGFnZ3R2YXVqdHp5cGVvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg0NTkyMTMsImV4cCI6MjA5NDAzNTIxM30.yKWJL-UyDbNriRVcLM0bjGShwtUiIW-eEO_gseqvg7A";
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

const COLORS = {
  bg: "#F0F4F8", card: "#FFFFFF", navy: "#0B2545", teal: "#1A6B72",
  accent: "#E87C2A", success: "#2D9C6A", warn: "#D4841A",
  danger: "#C0392B", muted: "#8899A8", border: "#DCE5EE", lightBg: "#F7FAFC",
};

const CVC_OPTS = ["VJID","VJIE","VSCD","VSCE","VFD","VFE"];
const ART_OPTS = ["ARD","ARE","AFD","AFE","ASCD","ASCE"];

const INITIAL_DISPOSITIVOS = {
  cvc: [], arterial: [],
  hd: false, svd: false, sne: false, sng: false, drenos: false,
  desinvadir: { cvc: {}, arterial: {}, hd: false, svd: false, sne: false, sng: false, drenos: false },
};

const INITIAL_ROUND = {
  contato: null, visitaFlex: null,
  rass: null, dor: null, delirium: null, contencao: null,
  dva: null, pam: "",
  suporteResp: null, vmProtetora: null, planoDesmame: [], preocResp: [], ims: "", progredirFuncional: null,
  viaAlimentar: null, aceitacao: null, metaCalorica: null, fono: null, glicemia: null, evacuacao: null,
  funcaoRenal: null, metaBH: null, pioraInfec: null, atb: null,
  tev: null, lamg: null, cornea: null, higieneOral: null, decubito: null,
  bundles: null, bundlesPendente: "", mudancaDecubito: null, lesaoPressao: null, avalEspecializada: null,
  dispositivos: INITIAL_DISPOSITIVOS,
  diretivas: [], pendenciaExame: null, descPendencia: "", previsaoAlta: null,
  diagnostico: "",
};

function hoje() { return new Date().toISOString().slice(0, 10); }
function calcIdade(d) {
  if (!d) return null;
  const n = new Date(d), h = new Date();
  let a = h.getFullYear() - n.getFullYear();
  if (h.getMonth() < n.getMonth() || (h.getMonth() === n.getMonth() && h.getDate() < n.getDate())) a--;
  return a;
}
function calcDias(d) {
  if (!d) return 0;
  const diff = Math.floor((new Date() - new Date(d)) / 86400000);
  return diff >= 0 ? diff : 0;
}
const gravBadge = (g) => ({ alta:[COLORS.danger,"⚠ Alta"], media:[COLORS.warn,"◈ Média"], baixa:[COLORS.success,"✓ Baixa"], livre:[COLORS.muted,"Livre"] }[g] || [COLORS.muted, g]);

function computeAlerts(r, p) {
  const a = [];
  if (!r || !p.nome) return a;
  if (r.tev === "Não") a.push("Sem profilaxia TEV");
  if (r.suporteResp === "VM invasiva" && !(r.planoDesmame?.length)) a.push("VM sem plano de desmame");
  if (r.dva === "Sim" && !r.pam) a.push("DVA sem meta PAM");
  if (r.pendenciaExame === "Sim" && r.descPendencia) a.push(r.descPendencia);
  if (r.lesaoPressao === "Sim") a.push("LPP presente");
  if (r.bundles === "Não" && r.bundlesPendente) a.push(`Bundle pendente: ${r.bundlesPendente}`);
  if (calcDias(p.dataAdm) > 7) a.push(`${calcDias(p.dataAdm)} dias internado`);
  return a;
}

// ── Gerador de relatório WhatsApp ─────────────────────────────────────────────
function gerarRelatorioGeral(patients, rounds) {
  const ocupados = patients.filter(p => p.nome);
  const vagos    = patients.filter(p => !p.nome);
  const altaHoje = ocupados.filter(p => rounds[p.id]?.previsaoAlta === "Hoje");
  const graves   = ocupados.filter(p => p.gravidade === "alta");
  const intub    = ocupados.filter(p => rounds[p.id]?.suporteResp === "VM invasiva");
  const dva      = ocupados.filter(p => rounds[p.id]?.dva === "Sim");
  const hd       = ocupados.filter(p => rounds[p.id]?.funcaoRenal === "Em HD");

  const alertas = [];
  ocupados.forEach(p => {
    computeAlerts(rounds[p.id], p).forEach(a => alertas.push(`⚠️ ${p.leito} — ${a}`));
  });

  const dt = new Date().toLocaleDateString("pt-BR") + " — " + new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });

  return `🏥 *ROUND UTI CLÍNICA — IMIP*
📅 *${dt}*

━━━━━━━━━━━━━━━━━
📊 *RESUMO GERAL*
━━━━━━━━━━━━━━━━━

🛏 *Leitos ocupados:* ${ocupados.length}/10
🔓 *Leitos vagos:* ${vagos.length > 0 ? vagos.map(p => p.leito).join(", ") : "Nenhum"}

🏠 *Previsão de alta hoje:* ${altaHoje.length > 0 ? altaHoje.map(p => p.leito).join(", ") : "Nenhum"}
⚠️ *Graves:* ${graves.length > 0 ? graves.map(p => p.leito).join(", ") : "Nenhum"}
🫁 *Intubados (VMI):* ${intub.length > 0 ? intub.map(p => p.leito).join(", ") : "Nenhum"}
💊 *Em DVA:* ${dva.length > 0 ? dva.map(p => p.leito).join(", ") : "Nenhum"}
🩸 *Em HD:* ${hd.length > 0 ? hd.map(p => p.leito).join(", ") : "Nenhum"}

━━━━━━━━━━━━━━━━━
🚨 *ALERTAS E PENDÊNCIAS*
━━━━━━━━━━━━━━━━━

${alertas.length > 0 ? alertas.join("\n") : "✅ Sem alertas no momento"}

━━━━━━━━━━━━━━━━━
📊 _Gerado pelo UTI Round — IMIP_`;
}

function gerarRelatorioEspecifico(patients, rounds) {
  const ocupados = patients.filter(p => p.nome);
  const dt = new Date().toLocaleDateString("pt-BR") + " — " + new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });

  const blocos = ocupados.map(p => {
    const r = rounds[p.id];
    const dev = r?.dispositivos || INITIAL_DISPOSITIVOS;
    const idade = calcIdade(p.dataNasc);
    const dias  = calcDias(p.dataAdm);
    const cvcList = (dev.cvc || []).join(", ");
    const artList = (dev.arterial || []).join(", ");
    const devList = [
      cvcList && `CVC (${cvcList})`,
      artList && `Art (${artList})`,
      dev.hd && "HD",
      dev.svd && "SVD",
      dev.sne && "SNE",
      dev.sng && "SNG",
      dev.drenos && "Drenos",
    ].filter(Boolean).join(" | ") || "Nenhum";

    return `🛏 *${p.leito} — ${p.nome}*
👤 ${idade ?? "?"}a | 🕐 ${dias}d internado
📋 *Diagnóstico:* ${r?.diagnostico || p.diagnostico || "—"}
⚠️ *Gravidade:* ${p.gravidade}

🧠 *Neuro:* RASS ${r?.rass || "—"} | Dor ${r?.dor === "Sim" ? "✅" : r?.dor === "Não" ? "❌" : "—"} | Delirium ${r?.delirium === "Sim" ? "✅" : r?.delirium === "Não" ? "❌" : "—"}
❤️ *Cardio:* DVA ${r?.dva === "Sim" ? "✅" : r?.dva === "Não" ? "❌" : "—"} | PAM: ${r?.pam || "—"} mmHg
🫁 *Resp:* ${r?.suporteResp || "—"} ${r?.planoDesmame?.length ? "| " + r.planoDesmame.join(", ") : ""}
🍽️ *Nutrição:* ${r?.viaAlimentar || "—"} | Meta calórica ${r?.metaCalorica === "Sim" ? "✅" : r?.metaCalorica === "Não" ? "❌" : "—"}
🫘 *Renal:* ${r?.funcaoRenal || "—"} | BH ${r?.metaBH || "—"}
🦠 *Infeccioso:* ATB ${r?.atb === "Sim" ? "✅" : r?.atb === "Não" ? "❌" : "—"} | Piora ${r?.pioraInfec === "Sim" ? "✅" : r?.pioraInfec === "Não" ? "❌" : "—"}
💉 *Profilaxias:* TEV ${r?.tev || "—"} | LAMG ${r?.lamg || "—"} | HO ${r?.higieneOral || "—"}
🔌 *Dispositivos:* ${devList}
${r?.pendenciaExame === "Sim" ? `📌 *Pendência:* ${r?.descPendencia || "—"}` : ""}
🏠 *Alta:* ${r?.previsaoAlta || "—"}`;
  });

  return `🏥 *ROUND UTI CLÍNICA — IMIP*
📅 *${dt}*

━━━━━━━━━━━━━━━━━
${blocos.join("\n\n━━━━━━━━━━━━━━━━━\n")}

━━━━━━━━━━━━━━━━━
📊 _Gerado pelo UTI Round — IMIP_`;
}

// ── Exportar Excel ────────────────────────────────────────────────────────────
function exportExcel(patients, rounds) {
  const rows = patients.filter(p => p.nome);
  if (!rows.length) { alert("Nenhum paciente para exportar."); return; }
  const data = rows.map(p => {
    const r = rounds[p.id] || {};
    const dev = r.dispositivos || INITIAL_DISPOSITIVOS;
    return {
      "Leito": p.leito, "Paciente": p.nome,
      "Data Nascimento": p.dataNasc || "", "Idade": calcIdade(p.dataNasc) ?? "",
      "Data Admissão": p.dataAdm || "", "Dias Internado": calcDias(p.dataAdm),
      "Gravidade": p.gravidade, "Diagnóstico": r.diagnostico || p.diagnostico || "",
      "Precaução Contato": r.contato || "", "Visita Flexibilizada": r.visitaFlex || "",
      "RASS": r.rass || "", "Dor": r.dor || "", "Delirium": r.delirium || "", "Contenção": r.contencao || "",
      "DVA": r.dva || "", "Meta PAM": r.pam || "",
      "Suporte Resp": r.suporteResp || "", "VM Protetora": r.vmProtetora || "",
      "Plano Desmame": (r.planoDesmame || []).join(" | "),
      "Preocupações Resp": (r.preocResp || []).join(" | "),
      "IMS": r.ims || "", "Progredir Funcional": r.progredirFuncional || "",
      "Via Alimentar": r.viaAlimentar || "", "Aceitação": r.aceitacao || "",
      "Meta Calórica": r.metaCalorica || "", "Fono": r.fono || "",
      "Glicemia": r.glicemia || "", "Evacuação": r.evacuacao || "",
      "Função Renal": r.funcaoRenal || "", "BH Meta": r.metaBH || "",
      "Piora Infecciosa": r.pioraInfec || "", "ATB": r.atb || "",
      "TEV": r.tev || "", "LAMG": r.lamg || "", "Córnea": r.cornea || "",
      "Higiene Oral": r.higieneOral || "", "Decúbito": r.decubito || "",
      "Bundles OK": r.bundles || "", "Bundle Pendente": r.bundlesPendente || "",
      "Mudança Decúbito": r.mudancaDecubito || "", "LPP": r.lesaoPressao || "",
      "Aval Especializada": r.avalEspecializada || "",
      "CVC": (dev.cvc || []).join(" | "), "CVC Desinvadir": Object.values(dev.desinvadir?.cvc || {}).some(v => v) ? "Sim" : "Não",
      "Cateter Arterial": (dev.arterial || []).join(" | "), "Art Desinvadir": Object.values(dev.desinvadir?.arterial || {}).some(v => v) ? "Sim" : "Não",
      "HD": dev.hd ? "Sim" : "Não", "HD Desinvadir": dev.desinvadir?.hd ? "Sim" : "Não",
      "SVD": dev.svd ? "Sim" : "Não", "SVD Desinvadir": dev.desinvadir?.svd ? "Sim" : "Não",
      "SNE": dev.sne ? "Sim" : "Não", "SNE Desinvadir": dev.desinvadir?.sne ? "Sim" : "Não",
      "SNG": dev.sng ? "Sim" : "Não", "SNG Desinvadir": dev.desinvadir?.sng ? "Sim" : "Não",
      "Drenos": dev.drenos ? "Sim" : "Não", "Drenos Desinvadir": dev.desinvadir?.drenos ? "Sim" : "Não",
      "Diretivas": (r.diretivas || []).join(" | "),
      "Pendência": r.pendenciaExame || "", "Desc Pendência": r.descPendencia || "",
      "Previsão Alta": r.previsaoAlta || "",
      "Alertas": computeAlerts(r, p).join(" | "),
    };
  });
  const ws = XLSX.utils.json_to_sheet(data);
  ws["!cols"] = Object.keys(data[0]).map(() => ({ wch: 20 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Round UTI");
  XLSX.writeFile(wb, `round-uti-${hoje()}.xlsx`);
}

// ── Atoms ─────────────────────────────────────────────────────────────────────
function Pill({ label, selected, onClick, color }) {
  return (
    <button onClick={onClick} style={{
      padding: "5px 14px", borderRadius: 20, cursor: "pointer", whiteSpace: "nowrap",
      border: `1.5px solid ${selected ? (color || COLORS.teal) : COLORS.border}`,
      background: selected ? (color || COLORS.teal) : "#fff",
      color: selected ? "#fff" : COLORS.navy,
      fontSize: 13, fontWeight: selected ? 600 : 400, transition: "all .15s",
    }}>{label}</button>
  );
}
function MultiPill({ label, checked, onChange, color }) {
  return (
    <button onClick={() => onChange(!checked)} style={{
      padding: "5px 14px", borderRadius: 20, cursor: "pointer",
      border: `1.5px solid ${checked ? (color || COLORS.accent) : COLORS.border}`,
      background: checked ? (color || COLORS.accent) : "#fff",
      color: checked ? "#fff" : COLORS.navy,
      fontSize: 13, fontWeight: checked ? 600 : 400, transition: "all .15s",
    }}>{label}</button>
  );
}
function SecHdr({ title, icon }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14, marginTop: 8 }}>
      <span style={{ fontSize: 17 }}>{icon}</span>
      <span style={{ fontWeight: 700, fontSize: 12, color: COLORS.navy, letterSpacing: ".6px", textTransform: "uppercase" }}>{title}</span>
      <div style={{ flex: 1, height: 1, background: COLORS.border }} />
    </div>
  );
}
function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 6, textTransform: "uppercase", letterSpacing: ".4px" }}>{label}</div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>{children}</div>
    </div>
  );
}
function TInput({ value, onChange, placeholder, w }) {
  return (
    <input value={value || ""} onChange={e => onChange(e.target.value)} placeholder={placeholder} style={{
      border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "6px 12px",
      fontSize: 13, color: COLORS.navy, outline: "none", background: "#fff",
      width: w || "100%", maxWidth: w || 240, boxSizing: "border-box",
    }} />
  );
}
function TArea({ value, onChange, placeholder, rows = 2 }) {
  return (
    <textarea value={value || ""} onChange={e => onChange(e.target.value)} placeholder={placeholder} rows={rows} style={{
      border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "8px 12px",
      fontSize: 13, color: COLORS.navy, outline: "none", width: "100%",
      resize: "vertical", background: "#fff", fontFamily: "inherit", boxSizing: "border-box",
    }} />
  );
}

// ── Seção Dispositivos ────────────────────────────────────────────────────────
function DispositivosSection({ dev, onChange }) {
  const d = dev || INITIAL_DISPOSITIVOS;

  const toggleCVC = (opt) => {
    const cur = d.cvc || [];
    const novo = cur.includes(opt) ? cur.filter(x => x !== opt) : [...cur, opt];
    const desCvc = { ...(d.desinvadir?.cvc || {}) };
    if (!novo.includes(opt)) delete desCvc[opt];
    onChange({ ...d, cvc: novo, desinvadir: { ...d.desinvadir, cvc: desCvc } });
  };
  const toggleArt = (opt) => {
    const cur = d.arterial || [];
    const novo = cur.includes(opt) ? cur.filter(x => x !== opt) : [...cur, opt];
    const desArt = { ...(d.desinvadir?.arterial || {}) };
    if (!novo.includes(opt)) delete desArt[opt];
    onChange({ ...d, arterial: novo, desinvadir: { ...d.desinvadir, arterial: desArt } });
  };
  const toggleSimples = (key) => onChange({ ...d, [key]: !d[key], desinvadir: { ...d.desinvadir, [key]: false } });
  const setDesinvCvc = (opt, val) => onChange({ ...d, desinvadir: { ...d.desinvadir, cvc: { ...(d.desinvadir?.cvc || {}), [opt]: val } } });
  const setDesinvArt = (opt, val) => onChange({ ...d, desinvadir: { ...d.desinvadir, arterial: { ...(d.desinvadir?.arterial || {}), [opt]: val } } });
  const setDesinvSimp = (key, val) => onChange({ ...d, desinvadir: { ...d.desinvadir, [key]: val } });

  const boxStyle = (active) => ({
    background: active ? COLORS.teal + "12" : COLORS.lightBg,
    borderRadius: 10, padding: "12px 14px",
    border: `1.5px solid ${active ? COLORS.teal : COLORS.border}`,
    marginBottom: 8,
  });

  return (
    <div>
      {/* CVC */}
      <div style={boxStyle((d.cvc || []).length > 0)}>
        <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.navy, marginBottom: 8 }}>🔵 Cateter Venoso Central</div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
          {CVC_OPTS.map(opt => (
            <button key={opt} onClick={() => toggleCVC(opt)} style={{
              padding: "4px 12px", borderRadius: 16, cursor: "pointer", fontSize: 12, fontWeight: 600,
              border: `1.5px solid ${(d.cvc || []).includes(opt) ? COLORS.teal : COLORS.border}`,
              background: (d.cvc || []).includes(opt) ? COLORS.teal : "#fff",
              color: (d.cvc || []).includes(opt) ? "#fff" : COLORS.navy,
            }}>{opt}</button>
          ))}
        </div>
        {(d.cvc || []).map(opt => (
          <div key={opt} style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 4 }}>
            <span style={{ fontSize: 12, color: COLORS.teal, fontWeight: 600, minWidth: 40 }}>{opt}</span>
            <span style={{ fontSize: 12, color: COLORS.muted }}>Desinvadir?</span>
            {["Sim","Não"].map(v => (
              <button key={v} onClick={() => setDesinvCvc(opt, v === "Sim")} style={{
                padding: "2px 10px", borderRadius: 12, cursor: "pointer", fontSize: 11, fontWeight: 600,
                border: `1.5px solid ${(d.desinvadir?.cvc?.[opt] ? "Sim" : "Não") === v ? (v === "Sim" ? COLORS.success : COLORS.danger) : COLORS.border}`,
                background: (d.desinvadir?.cvc?.[opt] ? "Sim" : "Não") === v ? (v === "Sim" ? COLORS.success : COLORS.danger) : "#fff",
                color: (d.desinvadir?.cvc?.[opt] ? "Sim" : "Não") === v ? "#fff" : COLORS.navy,
              }}>{v}</button>
            ))}
          </div>
        ))}
      </div>

      {/* Cateter Arterial */}
      <div style={boxStyle((d.arterial || []).length > 0)}>
        <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.navy, marginBottom: 8 }}>🔴 Cateter Arterial</div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
          {ART_OPTS.map(opt => (
            <button key={opt} onClick={() => toggleArt(opt)} style={{
              padding: "4px 12px", borderRadius: 16, cursor: "pointer", fontSize: 12, fontWeight: 600,
              border: `1.5px solid ${(d.arterial || []).includes(opt) ? COLORS.danger : COLORS.border}`,
              background: (d.arterial || []).includes(opt) ? COLORS.danger : "#fff",
              color: (d.arterial || []).includes(opt) ? "#fff" : COLORS.navy,
            }}>{opt}</button>
          ))}
        </div>
        {(d.arterial || []).map(opt => (
          <div key={opt} style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 4 }}>
            <span style={{ fontSize: 12, color: COLORS.danger, fontWeight: 600, minWidth: 40 }}>{opt}</span>
            <span style={{ fontSize: 12, color: COLORS.muted }}>Desinvadir?</span>
            {["Sim","Não"].map(v => (
              <button key={v} onClick={() => setDesinvArt(opt, v === "Sim")} style={{
                padding: "2px 10px", borderRadius: 12, cursor: "pointer", fontSize: 11, fontWeight: 600,
                border: `1.5px solid ${(d.desinvadir?.arterial?.[opt] ? "Sim" : "Não") === v ? (v === "Sim" ? COLORS.success : COLORS.danger) : COLORS.border}`,
                background: (d.desinvadir?.arterial?.[opt] ? "Sim" : "Não") === v ? (v === "Sim" ? COLORS.success : COLORS.danger) : "#fff",
                color: (d.desinvadir?.arterial?.[opt] ? "Sim" : "Não") === v ? "#fff" : COLORS.navy,
              }}>{v}</button>
            ))}
          </div>
        ))}
      </div>

      {/* Dispositivos simples */}
      {[["hd","🟣 Cateter de HD"],["svd","🟡 SVD"],["sne","🟠 SNE"],["sng","⚪ SNG"],["drenos","🟤 Drenos"]].map(([key, lbl]) => (
        <div key={key} style={boxStyle(d[key])}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <button onClick={() => toggleSimples(key)} style={{
              padding: "4px 14px", borderRadius: 16, cursor: "pointer", fontSize: 12, fontWeight: 700,
              border: `1.5px solid ${d[key] ? COLORS.teal : COLORS.border}`,
              background: d[key] ? COLORS.teal : "#fff",
              color: d[key] ? "#fff" : COLORS.navy,
            }}>{d[key] ? "✓ Presente" : "Ausente"}</button>
            <span style={{ fontSize: 12, fontWeight: 700, color: COLORS.navy }}>{lbl}</span>
            {d[key] && (
              <>
                <span style={{ fontSize: 12, color: COLORS.muted, marginLeft: 8 }}>Desinvadir?</span>
                {["Sim","Não"].map(v => (
                  <button key={v} onClick={() => setDesinvSimp(key, v === "Sim")} style={{
                    padding: "2px 10px", borderRadius: 12, cursor: "pointer", fontSize: 11, fontWeight: 600,
                    border: `1.5px solid ${(d.desinvadir?.[key] ? "Sim" : "Não") === v ? (v === "Sim" ? COLORS.success : COLORS.danger) : COLORS.border}`,
                    background: (d.desinvadir?.[key] ? "Sim" : "Não") === v ? (v === "Sim" ? COLORS.success : COLORS.danger) : "#fff",
                    color: (d.desinvadir?.[key] ? "Sim" : "Não") === v ? "#fff" : COLORS.navy,
                  }}>{v}</button>
                ))}
              </>
            )}
          </div>
        </div>
      ))}
    </div>
  );
}

// ── Modal relatório WhatsApp ───────────────────────────────────────────────────
function RelatorioModal({ patients, rounds, onClose, onSave }) {
  const [tipo, setTipo] = useState("geral");
  const texto = tipo === "geral"
    ? gerarRelatorioGeral(patients, rounds)
    : gerarRelatorioEspecifico(patients, rounds);

  const copiar = () => {
    navigator.clipboard.writeText(texto);
    alert("Copiado! Cole no WhatsApp.");
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,37,69,.55)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: "#fff", borderRadius: 16, padding: "24px", maxWidth: 640, width: "100%", maxHeight: "90vh", display: "flex", flexDirection: "column" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div style={{ fontWeight: 800, fontSize: 18, color: COLORS.navy }}>📋 Relatório</div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: COLORS.muted }}>×</button>
        </div>
        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          {[["geral","📊 Resumo Geral"],["especifico","🛏 Por Paciente"]].map(([v,l]) => (
            <button key={v} onClick={() => setTipo(v)} style={{ padding: "8px 18px", borderRadius: 10, border: `1.5px solid ${tipo === v ? COLORS.teal : COLORS.border}`, background: tipo === v ? COLORS.teal : "#fff", color: tipo === v ? "#fff" : COLORS.navy, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>{l}</button>
          ))}
        </div>
        <textarea readOnly value={texto} style={{ flex: 1, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 12, fontSize: 12, fontFamily: "monospace", resize: "none", background: COLORS.lightBg, color: COLORS.navy, minHeight: 300 }} />
        <div style={{ display: "flex", gap: 10, marginTop: 16, flexWrap: "wrap" }}>
          <button onClick={copiar} style={{ flex: 2, minWidth: 160, padding: "10px", borderRadius: 10, border: "none", background: "#25D366", color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>📲 Copiar para WhatsApp</button>
          <button onClick={() => exportExcel(patients, rounds)} style={{ flex: 1, minWidth: 120, padding: "10px", borderRadius: 10, border: "none", background: "#1A6B72", color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>📊 Exportar Excel</button>
          <button onClick={() => { onSave(tipo, texto); }} style={{ flex: 1, minWidth: 80, padding: "10px", borderRadius: 10, border: "none", background: COLORS.accent, color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>💾 Salvar</button>
          <button onClick={onClose} style={{ flex: 1, minWidth: 80, padding: "10px", borderRadius: 10, border: `1.5px solid ${COLORS.border}`, background: "#fff", color: COLORS.navy, fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Fechar</button>
        </div>
      </div>
    </div>
  );
}

// ── Modal histórico de relatórios ─────────────────────────────────────────────
function HistoricoModal({ relatorios, onClose }) {
  const [sel, setSel] = useState(null);
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,37,69,.55)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: "#fff", borderRadius: 16, padding: "24px", maxWidth: 700, width: "100%", maxHeight: "90vh", display: "flex", flexDirection: "column" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div style={{ fontWeight: 800, fontSize: 18, color: COLORS.navy }}>🗂 Histórico de Relatórios</div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: COLORS.muted }}>×</button>
        </div>
        {relatorios.length === 0 ? (
          <div style={{ textAlign: "center", color: COLORS.muted, padding: 40 }}>Nenhum relatório salvo.</div>
        ) : (
          <div style={{ display: "flex", gap: 12, flex: 1, overflow: "hidden" }}>
            <div style={{ width: 200, overflowY: "auto", borderRight: `1px solid ${COLORS.border}`, paddingRight: 12 }}>
              {relatorios.map((r, i) => (
                <div key={i} onClick={() => setSel(r)} style={{ padding: "8px 10px", borderRadius: 8, cursor: "pointer", marginBottom: 6, background: sel === r ? COLORS.teal + "18" : COLORS.lightBg, border: `1px solid ${sel === r ? COLORS.teal : COLORS.border}` }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.navy }}>{r.tipo === "geral" ? "📊 Geral" : "🛏 Por paciente"}</div>
                  <div style={{ fontSize: 11, color: COLORS.muted }}>{r.data}</div>
                </div>
              ))}
            </div>
            <div style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
              {sel ? (
                <>
                  <textarea readOnly value={sel.texto} style={{ flex: 1, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 12, fontSize: 12, fontFamily: "monospace", resize: "none", background: COLORS.lightBg, color: COLORS.navy }} />
                  <button onClick={() => { navigator.clipboard.writeText(sel.texto); alert("Copiado!"); }} style={{ marginTop: 10, padding: "9px", borderRadius: 10, border: "none", background: "#25D366", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>📲 Copiar para WhatsApp</button>
                </>
              ) : (
                <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100%", color: COLORS.muted }}>Selecione um relatório</div>
              )}
            </div>
          </div>
        )}
        <button onClick={onClose} style={{ marginTop: 16, padding: "10px", borderRadius: 10, border: `1.5px solid ${COLORS.border}`, background: "#fff", color: COLORS.navy, fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Fechar</button>
      </div>
    </div>
  );
}

// ── Modal editar paciente ─────────────────────────────────────────────────────
function EditPatientModal({ pat, onSave, onClear, onClose }) {
  const [nome,    setNome]    = useState(pat.nome || "");
  const [dataNasc,setNasc]   = useState(pat.dataNasc || "");
  const [dataAdm, setAdm]    = useState(pat.dataAdm || "");
  const [diag,    setDiag]   = useState(pat.diagnostico || "");
  const [grav,    setGrav]   = useState(pat.gravidade || "livre");
  const [confirm, setConfirm]= useState(false);
  const canSave = nome.trim().length > 0;

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,37,69,.6)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: "#fff", borderRadius: 16, padding: "28px 30px", width: "100%", maxWidth: 460, boxShadow: "0 20px 60px rgba(0,0,0,.25)", maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 22 }}>
          <div>
            <div style={{ fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: "uppercase" }}>{pat.leito}</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: COLORS.navy }}>{pat.nome ? "Editar Paciente" : "Admitir Paciente"}</div>
          </div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: COLORS.muted }}>×</button>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <div>
            <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase" }}>Nome completo *</div>
            <input value={nome} onChange={e => setNome(e.target.value)} placeholder="Nome do paciente"
              style={{ border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "9px 12px", fontSize: 14, color: COLORS.navy, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <div>
              <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase" }}>Data de nascimento</div>
              <input type="date" value={dataNasc} onChange={e => setNasc(e.target.value)}
                style={{ border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "9px 12px", fontSize: 14, color: COLORS.navy, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
              {calcIdade(dataNasc) !== null && <div style={{ fontSize: 12, color: COLORS.teal, marginTop: 4, fontWeight: 600 }}>→ {calcIdade(dataNasc)} anos</div>}
            </div>
            <div>
              <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase" }}>Data de admissão</div>
              <input type="date" value={dataAdm} onChange={e => setAdm(e.target.value)}
                style={{ border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "9px 12px", fontSize: 14, color: COLORS.navy, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
              {dataAdm && <div style={{ fontSize: 12, color: COLORS.teal, marginTop: 4, fontWeight: 600 }}>→ {calcDias(dataAdm)} dia(s)</div>}
            </div>
          </div>
          <div>
            <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase" }}>Diagnóstico</div>
            <input value={diag} onChange={e => setDiag(e.target.value)} placeholder="Ex: Sepse pulmonar"
              style={{ border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "9px 12px", fontSize: 14, color: COLORS.navy, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
          </div>
          <div>
            <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 8, textTransform: "uppercase" }}>Gravidade</div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {[["alta",COLORS.danger,"⚠ Alta"],["media",COLORS.warn,"◈ Média"],["baixa",COLORS.success,"✓ Baixa"],["livre",COLORS.muted,"Livre"]].map(([v,c,l]) => (
                <button key={v} onClick={() => setGrav(v)} style={{ padding: "6px 14px", borderRadius: 8, cursor: "pointer", border: `1.5px solid ${grav===v?c:COLORS.border}`, background: grav===v?c+"1A":"#fff", color: grav===v?c:COLORS.navy, fontSize: 13, fontWeight: 700 }}>{l}</button>
              ))}
            </div>
          </div>
        </div>
        {pat.nome && (
          <div style={{ marginTop: 20, paddingTop: 16, borderTop: `1px solid ${COLORS.border}` }}>
            {!confirm ? (
              <button onClick={() => setConfirm(true)} style={{ width: "100%", padding: "9px", borderRadius: 8, border: `1.5px solid ${COLORS.danger}`, background: "#fff", color: COLORS.danger, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>🗑 Limpar dados e desocupar leito</button>
            ) : (
              <div style={{ background: COLORS.danger+"12", borderRadius: 10, padding: 14 }}>
                <div style={{ fontSize: 13, color: COLORS.danger, fontWeight: 700, marginBottom: 10 }}>Tem certeza? Todos os dados serão apagados.</div>
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={() => { onClear(pat.id); onClose(); }} style={{ flex: 1, padding: "9px", borderRadius: 8, border: "none", background: COLORS.danger, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>Sim, limpar</button>
                  <button onClick={() => setConfirm(false)} style={{ flex: 1, padding: "9px", borderRadius: 8, border: `1.5px solid ${COLORS.border}`, background: "#fff", color: COLORS.navy, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Cancelar</button>
                </div>
              </div>
            )}
          </div>
        )}
        <div style={{ display: "flex", gap: 10, marginTop: 20 }}>
          <button onClick={onClose} style={{ flex: 1, padding: "10px", borderRadius: 10, border: `1.5px solid ${COLORS.border}`, background: "#fff", color: COLORS.navy, fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Cancelar</button>
          <button onClick={() => { if (canSave) { onSave(pat.id, { nome: nome.trim(), dataNasc, dataAdm, diagnostico: diag, gravidade: grav }); onClose(); }}}
            style={{ flex: 2, padding: "10px", borderRadius: 10, border: "none", background: canSave ? COLORS.teal : COLORS.border, color: "#fff", fontSize: 14, fontWeight: 700, cursor: canSave ? "pointer" : "not-allowed" }}>✓ Salvar</button>
        </div>
      </div>
    </div>
  );
}

// ── Modal Limpar Todos ────────────────────────────────────────────────────────
function ClearAllModal({ onConfirm, onClose }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,37,69,.6)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: "#fff", borderRadius: 16, padding: 32, width: "100%", maxWidth: 400, textAlign: "center" }}>
        <div style={{ fontSize: 48, marginBottom: 12 }}>🗑️</div>
        <div style={{ fontSize: 20, fontWeight: 800, color: COLORS.navy, marginBottom: 8 }}>Limpar todos os pacientes?</div>
        <div style={{ fontSize: 14, color: COLORS.muted, marginBottom: 24 }}>Todos os dados e rounds serão apagados permanentemente.</div>
        <div style={{ display: "flex", gap: 12 }}>
          <button onClick={onClose} style={{ flex: 1, padding: "12px", borderRadius: 10, border: `1.5px solid ${COLORS.border}`, background: "#fff", color: COLORS.navy, fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Cancelar</button>
          <button onClick={onConfirm} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: COLORS.danger, color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>Sim, limpar tudo</button>
        </div>
      </div>
    </div>
  );
}

// ── Patient Card ──────────────────────────────────────────────────────────────
function PatientCard({ pat, round, onSelect, onEdit }) {
  const [color, label] = gravBadge(pat.gravidade);
  const alerts = computeAlerts(round, pat);
  const isEmpty = !pat.nome;
  return (
    <div style={{ background: COLORS.card, borderRadius: 14, border: `1.5px solid ${isEmpty ? COLORS.border : color+"44"}`, padding: "14px 16px", position: "relative", overflow: "hidden", boxShadow: "0 2px 8px rgba(11,37,69,.06)" }}>
      {!isEmpty && <div style={{ position: "absolute", top: 0, left: 0, width: 4, height: "100%", background: color, borderRadius: "14px 0 0 14px" }} />}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginLeft: isEmpty ? 0 : 8 }}>
        <span style={{ fontSize: 11, fontWeight: 700, color: COLORS.muted }}>{pat.leito}</span>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          {!isEmpty && <span style={{ fontSize: 11, color, fontWeight: 600, background: color+"18", padding: "2px 8px", borderRadius: 10 }}>{label}</span>}
          <button onClick={e => { e.stopPropagation(); onEdit(pat.id); }} style={{ background: "none", border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "2px 7px", cursor: "pointer", fontSize: 12, color: COLORS.muted }}>{isEmpty ? "＋" : "✏️"}</button>
        </div>
      </div>
      {isEmpty ? (
        <div onClick={() => onEdit(pat.id)} style={{ marginTop: 10, cursor: "pointer" }}>
          <div style={{ color: COLORS.muted, fontSize: 13, marginBottom: 4 }}>Leito disponível</div>
          <div style={{ fontSize: 12, color: COLORS.teal, fontWeight: 700 }}>+ Admitir paciente</div>
        </div>
      ) : (
        <div onClick={() => onSelect(pat.id)} style={{ cursor: "pointer" }}>
          <div style={{ marginTop: 6, fontWeight: 700, fontSize: 15, color: COLORS.navy, marginLeft: 8 }}>{pat.nome}</div>
          <div style={{ fontSize: 12, color: COLORS.muted, marginTop: 2, marginLeft: 8 }}>{pat.diagnostico || "—"}</div>
          <div style={{ display: "flex", gap: 12, marginTop: 8, marginLeft: 8 }}>
            {pat.dataAdm && <span style={{ fontSize: 12, color: COLORS.muted }}>🕐 {calcDias(pat.dataAdm)}d</span>}
            {calcIdade(pat.dataNasc) !== null && <span style={{ fontSize: 12, color: COLORS.muted }}>👤 {calcIdade(pat.dataNasc)}a</span>}
          </div>
          {alerts.length > 0 && (
            <div style={{ marginTop: 8, marginLeft: 8, display: "flex", flexWrap: "wrap", gap: 4 }}>
              {alerts.map((a, i) => <span key={i} style={{ fontSize: 11, background: COLORS.danger+"15", color: COLORS.danger, padding: "2px 7px", borderRadius: 8, fontWeight: 600 }}>⚠ {a}</span>)}
            </div>
          )}
          <div style={{ marginTop: 8, marginLeft: 8 }}>
            {round ? <span style={{ fontSize: 11, background: COLORS.success+"20", color: COLORS.success, padding: "2px 8px", borderRadius: 8, fontWeight: 600 }}>✓ Round feito hoje</span>
                   : <span style={{ fontSize: 11, background: COLORS.warn+"18", color: COLORS.warn, padding: "2px 8px", borderRadius: 8, fontWeight: 600 }}>⏳ Pendente</span>}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Round Form ────────────────────────────────────────────────────────────────
function RoundForm({ pat, round, onChange, onBack }) {
  const r = round || { ...INITIAL_ROUND };
  const set = (k, v) => onChange({ ...r, [k]: v });
  const tog = (k, v) => { const a = r[k] || []; onChange({ ...r, [k]: a.includes(v) ? a.filter(x => x !== v) : [...a, v] }); };
  const s2 = ["Sim","Não"], s3 = ["Sim","Não","Sem indicação"];

  return (
    <div style={{ maxWidth: 820, margin: "0 auto", paddingBottom: 80 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 24 }}>
        <button onClick={onBack} style={{ background: "none", border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "6px 14px", cursor: "pointer", fontSize: 13, color: COLORS.navy, fontWeight: 600 }}>← Voltar</button>
        <div>
          <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, textTransform: "uppercase" }}>{pat.leito} · Round — {new Date().toLocaleDateString("pt-BR")}</div>
          <div style={{ fontSize: 20, fontWeight: 800, color: COLORS.navy }}>{pat.nome}</div>
          <div style={{ fontSize: 12, color: COLORS.muted, marginTop: 2 }}>
            {calcIdade(pat.dataNasc) !== null && `${calcIdade(pat.dataNasc)} anos`}
            {calcIdade(pat.dataNasc) !== null && pat.dataAdm && " · "}
            {pat.dataAdm && `${calcDias(pat.dataAdm)} dia(s) internado`}
          </div>
        </div>
      </div>

      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Diagnóstico" icon="🏥" />
        <TInput value={r.diagnostico} onChange={v => set("diagnostico", v)} placeholder="Ex: Sepse pulmonar" />
      </div>

      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Cuidados Gerais" icon="🛡️" />
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          <Field label="Precaução de contato?">{s2.map(o => <Pill key={o} label={o} selected={r.contato===o} onClick={() => set("contato",o)} />)}</Field>
          <Field label="Visita flexibilizada?">{s2.map(o => <Pill key={o} label={o} selected={r.visitaFlex===o} onClick={() => set("visitaFlex",o)} />)}</Field>
        </div>
      </div>

      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Neurológico" icon="🧠" />
        <Field label="Meta de sedação (RASS)">{["-5 a -4","-2 a 0","Não se aplica"].map(o => <Pill key={o} label={o} selected={r.rass===o} onClick={() => set("rass",o)} />)}</Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
          <Field label="Controle de dor?">{s2.map(o => <Pill key={o} label={o} selected={r.dor===o} onClick={() => set("dor",o)} />)}</Field>
          <Field label="Delirium?">{s2.map(o => <Pill key={o} label={o} selected={r.delirium===o} onClick={() => set("delirium",o)} />)}</Field>
          <Field label="Contenção mecânica?">{s2.map(o => <Pill key={o} label={o} selected={r.contencao===o} onClick={() => set("contencao",o)} />)}</Field>
        </div>
      </div>

      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Cardiovascular / Hemodinâmica" icon="❤️" />
        <div style={{ display: "flex", gap: 24, flexWrap: "wrap", alignItems: "flex-end" }}>
          <Field label="Suporte hemodinâmico / DVA?">{s2.map(o => <Pill key={o} label={o} selected={r.dva===o} onClick={() => set("dva",o)} />)}</Field>
          <Field label="Meta de PAM (mmHg)"><TInput value={r.pam} onChange={v => set("pam",v)} placeholder="Ex: 65" w="90px" /></Field>
        </div>
      </div>

      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Respiratório / Reabilitação" icon="🫁" />
        <Field label="Suporte respiratório">{["Sem suporte","O2 suplementar","VNI","VM invasiva"].map(o => <Pill key={o} label={o} selected={r.suporteResp===o} onClick={() => set("suporteResp",o)} />)}</Field>
        {(r.suporteResp==="VM invasiva"||r.suporteResp==="VNI") && <Field label="VM protetora?">{s2.map(o => <Pill key={o} label={o} selected={r.vmProtetora===o} onClick={() => set("vmProtetora",o)} />)}</Field>}
        <Field label="Plano de desmame">{["Redução de parâmetros","TRE hoje","Ex-TOT","Sem proposta"].map(o => <MultiPill key={o} label={o} checked={(r.planoDesmame||[]).includes(o)} onChange={() => tog("planoDesmame",o)} />)}</Field>
        <Field label="Preocupações respiratórias">{["Piora respiratória","Piora de secreção","Risco de broncoaspiração"].map(o => <MultiPill key={o} label={o} checked={(r.preocResp||[]).includes(o)} onChange={() => tog("preocResp",o)} color={COLORS.danger} />)}</Field>
        <div style={{ display: "flex", gap: 24, flexWrap: "wrap" }}>
          <Field label="Escala IMS"><TInput value={r.ims} onChange={v => set("ims",v)} placeholder="0-10" w="80px" /></Field>
          <Field label="Progredir nível funcional?">{s2.map(o => <Pill key={o} label={o} selected={r.progredirFuncional===o} onClick={() => set("progredirFuncional",o)} />)}</Field>
        </div>
      </div>

      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Gastrointestinal / Nutrição" icon="🍽️" />
        <Field label="Via alimentar">{["VO","SNE/GTT","NPT","Zero"].map(o => <Pill key={o} label={o} selected={r.viaAlimentar===o} onClick={() => set("viaAlimentar",o)} />)}</Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
          <Field label="Aceitação">{["Normal","Baixa","Intolerância"].map(o => <Pill key={o} label={o} selected={r.aceitacao===o} onClick={() => set("aceitacao",o)} />)}</Field>
          <Field label="Meta calórica?">{s2.map(o => <Pill key={o} label={o} selected={r.metaCalorica===o} onClick={() => set("metaCalorica",o)} />)}</Field>
          <Field label="Avaliação fono?">{s2.map(o => <Pill key={o} label={o} selected={r.fono===o} onClick={() => set("fono",o)} />)}</Field>
          <Field label="Glicemia adequada?">{s2.map(o => <Pill key={o} label={o} selected={r.glicemia===o} onClick={() => set("glicemia",o)} />)}</Field>
          <Field label="Evacuação < 3 dias?">{s2.map(o => <Pill key={o} label={o} selected={r.evacuacao===o} onClick={() => set("evacuacao",o)} />)}</Field>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 14 }}>
        <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", border: `1px solid ${COLORS.border}` }}>
          <SecHdr title="Renal" icon="🫘" />
          <Field label="Função renal em piora?">{["Sim","Não","Em HD"].map(o => <Pill key={o} label={o} selected={r.funcaoRenal===o} onClick={() => set("funcaoRenal",o)} />)}</Field>
          <Field label="Meta de balanço hídrico">{["Positivo","Negativo","Neutro"].map(o => <Pill key={o} label={o} selected={r.metaBH===o} onClick={() => set("metaBH",o)} />)}</Field>
        </div>
        <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", border: `1px solid ${COLORS.border}` }}>
          <SecHdr title="Infeccioso" icon="🦠" />
          <Field label="Piora infecciosa?">{s2.map(o => <Pill key={o} label={o} selected={r.pioraInfec===o} onClick={() => set("pioraInfec",o)} />)}</Field>
          <Field label="Em uso de ATB?">{s2.map(o => <Pill key={o} label={o} selected={r.atb===o} onClick={() => set("atb",o)} />)}</Field>
        </div>
      </div>

      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Profilaxias" icon="💉" />
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
          <Field label="Profilaxia TEV">{s3.map(o => <Pill key={o} label={o} selected={r.tev===o} onClick={() => set("tev",o)} color={o==="Não"?COLORS.danger:COLORS.teal} />)}</Field>
          <Field label="Profilaxia LAMG">{s3.map(o => <Pill key={o} label={o} selected={r.lamg===o} onClick={() => set("lamg",o)} />)}</Field>
          <Field label="Úlcera de córnea">{s3.map(o => <Pill key={o} label={o} selected={r.cornea===o} onClick={() => set("cornea",o)} />)}</Field>
          <Field label="Higiene oral">{s3.map(o => <Pill key={o} label={o} selected={r.higieneOral===o} onClick={() => set("higieneOral",o)} />)}</Field>
          <Field label="Decúbito elevado">{s3.map(o => <Pill key={o} label={o} selected={r.decubito===o} onClick={() => set("decubito",o)} />)}</Field>
          <Field label="Bundles OK?">{s2.map(o => <Pill key={o} label={o} selected={r.bundles===o} onClick={() => set("bundles",o)} color={o==="Não"?COLORS.danger:COLORS.teal} />)}</Field>
        </div>
        {r.bundles==="Não" && <div style={{ marginTop: 8 }}><Field label="Bundle pendente"><TInput value={r.bundlesPendente} onChange={v => set("bundlesPendente",v)} placeholder="Qual bundle?" /></Field></div>}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14, marginTop: 4 }}>
          <Field label="Mudança de decúbito?">{s2.map(o => <Pill key={o} label={o} selected={r.mudancaDecubito===o} onClick={() => set("mudancaDecubito",o)} />)}</Field>
          <div>
            <Field label="Lesão por pressão?">{s2.map(o => <Pill key={o} label={o} selected={r.lesaoPressao===o} onClick={() => set("lesaoPressao",o)} color={o==="Sim"?COLORS.danger:COLORS.teal} />)}</Field>
            {r.lesaoPressao==="Sim" && (
              <Field label="Avaliação especializada?">{["Comissão curativo","Cirurgia plástica","Não necessário"].map(o => <Pill key={o} label={o} selected={r.avalEspecializada===o} onClick={() => set("avalEspecializada",o)} />)}</Field>
            )}
          </div>
        </div>
      </div>

      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Dispositivos Invasivos" icon="🔌" />
        <DispositivosSection dev={r.dispositivos} onChange={v => set("dispositivos", v)} />
      </div>

      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Objetivos de Cuidado e Planejamento" icon="📋" />
        <Field label="Diretivas de cuidado">{["Não RCP","Não HD","Não IOT","Não DVA","Não coletar exames","Sem diretivas"].map(o => <MultiPill key={o} label={o} checked={(r.diretivas||[]).includes(o)} onChange={() => tog("diretivas",o)} color={o==="Sem diretivas"?COLORS.success:COLORS.danger} />)}</Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          <div>
            <Field label="Pendência de exame/procedimento?">{s2.map(o => <Pill key={o} label={o} selected={r.pendenciaExame===o} onClick={() => set("pendenciaExame",o)} />)}</Field>
            {r.pendenciaExame==="Sim" && <TArea value={r.descPendencia} onChange={v => set("descPendencia",v)} placeholder="Descreva a pendência..." />}
          </div>
          <Field label="Previsão de alta">{["Hoje","24–48h","> 48h"].map(o => <Pill key={o} label={o} selected={r.previsaoAlta===o} onClick={() => set("previsaoAlta",o)} />)}</Field>
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 12 }}>
        <button onClick={onBack} style={{ padding: "10px 24px", borderRadius: 10, border: `1.5px solid ${COLORS.border}`, background: "#fff", color: COLORS.navy, fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Cancelar</button>
        <button onClick={onBack} style={{ padding: "10px 28px", borderRadius: 10, border: "none", background: COLORS.teal, color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>✓ Salvar Round</button>
      </div>
    </div>
  );
}

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [patients,      setPatients]      = useState([]);
  const [rounds,        setRounds]        = useState({});
  const [selected,      setSelected]      = useState(null);
  const [editingId,     setEditingId]     = useState(null);
  const [showRelatorio, setShowRelatorio] = useState(false);
  const [showHistorico, setShowHistorico] = useState(false);
  const [showClearAll,  setShowClearAll]  = useState(false);
  const [search,        setSearch]        = useState("");
  const [filter,        setFilter]        = useState("todos");
  const [loading,       setLoading]       = useState(true);
  const [saving,        setSaving]        = useState(false);
  const [error,         setError]         = useState(null);
  const [relatorios,    setRelatorios]    = useState([]);

  const dataHoje = hoje();

  // ── Carregar dados ──────────────────────────────────────────────────────────
  useEffect(() => {
    async function load() {
      setLoading(true); setError(null);
      try {
        const { data: pats, error: e1 } = await supabase.from("patients").select("*").order("id");
        if (e1) throw e1;
        const { data: rds, error: e2 } = await supabase.from("rounds").select("*").eq("data", dataHoje);
        if (e2) throw e2;

        setPatients((pats || []).map(p => ({
          ...p, nome: p.nome||"", diagnostico: p.diagnostico||"",
          gravidade: p.gravidade||"livre", dataNasc: p.dataNasc||null, dataAdm: p.dataAdm||null,
        })));

        const map = {};
        (rds || []).forEach(r => {
          const rd = r.round_data || {};
          if (!rd.dispositivos) rd.dispositivos = INITIAL_DISPOSITIVOS;
          else {
            if (!rd.dispositivos.cvc) rd.dispositivos.cvc = [];
            if (!rd.dispositivos.arterial) rd.dispositivos.arterial = [];
            if (!rd.dispositivos.desinvadir) rd.dispositivos.desinvadir = INITIAL_DISPOSITIVOS.desinvadir;
          }
          map[r.patient_id] = rd;
        });
        setRounds(map);

        // Carregar relatórios salvos (últimos 7 dias)
        const seteDiasAtras = new Date(); seteDiasAtras.setDate(seteDiasAtras.getDate() - 7);
        const { data: rels } = await supabase.from("relatorios").select("*").gte("created_at", seteDiasAtras.toISOString()).order("created_at", { ascending: false });
        if (rels) setRelatorios(rels.map(r => ({ tipo: r.tipo, texto: r.texto, data: new Date(r.created_at).toLocaleString("pt-BR") })));
      } catch (err) {
        setError("Erro ao conectar ao banco de dados.");
        console.error(err);
      }
      setLoading(false);
    }
    load();
  }, []);

  // ── Salvar round ────────────────────────────────────────────────────────────
  const handleChange = useCallback(async (r) => {
    setRounds(prev => ({ ...prev, [selected]: r }));
    setSaving(true);
    await supabase.from("rounds").upsert({ patient_id: selected, data: dataHoje, round_data: r, updated_at: new Date().toISOString() }, { onConflict: "patient_id,data" });
    setSaving(false);
  }, [selected]);

  const savePatient = async (id, data) => {
    setPatients(prev => prev.map(p => p.id===id ? {...p,...data} : p));
    await supabase.from("patients").update({ nome: data.nome, dataNasc: data.dataNasc||null, dataAdm: data.dataAdm||null, diagnostico: data.diagnostico, gravidade: data.gravidade, updated_at: new Date().toISOString() }).eq("id", id);
  };

  const clearPatient = async (id) => {
    setPatients(prev => prev.map(p => p.id===id ? {...p, nome:"", dataNasc:null, dataAdm:null, diagnostico:"", gravidade:"livre"} : p));
    setRounds(prev => { const n={...prev}; delete n[id]; return n; });
    await supabase.from("patients").update({ nome:"", dataNasc:null, dataAdm:null, diagnostico:"", gravidade:"livre" }).eq("id", id);
    await supabase.from("rounds").delete().eq("patient_id", id);
  };

  const clearAll = async () => {
    setPatients(prev => prev.map(p => ({...p, nome:"", dataNasc:null, dataAdm:null, diagnostico:"", gravidade:"livre"})));
    setRounds({}); setShowClearAll(false);
    await supabase.from("patients").update({ nome:"", dataNasc:null, dataAdm:null, diagnostico:"", gravidade:"livre" }).gte("id", 1);
    await supabase.from("rounds").delete().gte("patient_id", 1);
  };

  const salvarRelatorio = async (tipo, texto) => {
    const novoRel = { tipo, texto, data: new Date().toLocaleString("pt-BR") };
    setRelatorios(prev => [novoRel, ...prev]);
    await supabase.from("relatorios").insert({ tipo, texto, created_at: new Date().toISOString() });
  };

  const selPat  = patients.find(p => p.id === selected);
  const editPat = patients.find(p => p.id === editingId);
  const date = new Date().toLocaleDateString("pt-BR", { weekday:"long", day:"2-digit", month:"long", year:"numeric" });
  const withPats   = patients.filter(p => p.nome);
  const roundsDone = withPats.filter(p => rounds[p.id]).length;
  const total      = withPats.length;

  const filtered = patients.filter(p => {
    if (search) { const q=search.toLowerCase(); if (!p.nome?.toLowerCase().includes(q) && !p.leito?.toLowerCase().includes(q)) return false; }
    if (filter==="pendente" && (!p.nome || rounds[p.id])) return false;
    if (filter==="alta" && p.gravidade!=="alta") return false;
    return true;
  });

  if (loading) return (
    <div style={{ background:COLORS.bg, minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", fontFamily:"'DM Sans','Segoe UI',sans-serif" }}>
      <div style={{ textAlign:"center" }}>
        <div style={{ fontSize:48, marginBottom:16 }}>🏥</div>
        <div style={{ fontSize:18, fontWeight:700, color:COLORS.navy }}>Carregando dados da UTI...</div>
        <div style={{ fontSize:13, color:COLORS.muted, marginTop:8 }}>Conectando ao banco de dados</div>
      </div>
    </div>
  );

  if (error) return (
    <div style={{ background:COLORS.bg, minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", fontFamily:"'DM Sans','Segoe UI',sans-serif" }}>
      <div style={{ textAlign:"center", maxWidth:400 }}>
        <div style={{ fontSize:48, marginBottom:16 }}>⚠️</div>
        <div style={{ fontSize:18, fontWeight:700, color:COLORS.danger, marginBottom:8 }}>{error}</div>
        <button onClick={() => window.location.reload()} style={{ padding:"10px 24px", borderRadius:10, border:"none", background:COLORS.teal, color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer" }}>Tentar novamente</button>
      </div>
    </div>
  );

  if (selected && selPat) return (
    <div style={{ background:COLORS.bg, minHeight:"100vh", fontFamily:"'DM Sans','Segoe UI',sans-serif" }}>
      <div style={{ background:COLORS.navy, padding:"12px 24px", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
        <div style={{ color:"#fff", fontWeight:800, fontSize:16 }}>🏥 UTI Clínica — IMIP</div>
        <div style={{ display:"flex", alignItems:"center", gap:12 }}>
          {saving && <span style={{ color:"#8BBBD9", fontSize:12 }}>💾 Salvando...</span>}
          <div style={{ color:"#8BBBD9", fontSize:13 }}>{date}</div>
        </div>
      </div>
      <div style={{ padding:"24px" }}>
        <RoundForm pat={selPat} round={rounds[selected]||{...INITIAL_ROUND, diagnostico:selPat.diagnostico||""}} onChange={handleChange} onBack={() => setSelected(null)} />
      </div>
    </div>
  );

  return (
    <div style={{ background:COLORS.bg, minHeight:"100vh", fontFamily:"'DM Sans','Segoe UI',sans-serif" }}>
      <div style={{ background:COLORS.navy, padding:"13px 28px", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
        <div style={{ display:"flex", alignItems:"center", gap:12 }}>
          <span style={{ fontSize:22 }}>🏥</span>
          <div>
            <div style={{ color:"#fff", fontWeight:800, fontSize:17 }}>UTI Clínica — IMIP</div>
            <div style={{ color:"#8BBBD9", fontSize:12 }}>Round Multidisciplinar · 10 leitos</div>
          </div>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
          <div style={{ color:"#8BBBD9", fontSize:13 }}>{date}</div>
          <button onClick={() => setShowHistorico(true)} style={{ padding:"8px 14px", borderRadius:10, border:`1.5px solid #ffffff33`, background:"transparent", color:"#fff", fontSize:13, fontWeight:600, cursor:"pointer" }}>🗂 Histórico</button>
          <button onClick={() => setShowClearAll(true)} style={{ padding:"8px 14px", borderRadius:10, border:`1.5px solid ${COLORS.danger}55`, background:"transparent", color:COLORS.danger, fontSize:13, fontWeight:700, cursor:"pointer" }}>🗑 Limpar</button>
          <button onClick={() => setShowRelatorio(true)} style={{ padding:"8px 18px", borderRadius:10, border:"none", background:COLORS.accent, color:"#fff", fontSize:13, fontWeight:700, cursor:"pointer" }}>📋 Relatório</button>
        </div>
      </div>

      <div style={{ padding:"24px 28px" }}>
        <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:20 }}>
          {[["Leitos ocupados",total,COLORS.navy,"🛏"],["Rounds feitos",roundsDone,COLORS.success,"✅"],["Pendentes",total-roundsDone,COLORS.warn,"⏳"],["Alertas",patients.filter(p=>computeAlerts(rounds[p.id],p).length>0).length,COLORS.danger,"⚠️"]].map(([lbl,val,color,icon]) => (
            <div key={lbl} style={{ background:"#fff", borderRadius:12, padding:"16px 20px", border:`1px solid ${COLORS.border}` }}>
              <div style={{ fontSize:22, marginBottom:4 }}>{icon}</div>
              <div style={{ fontSize:26, fontWeight:800, color }}>{val}</div>
              <div style={{ fontSize:12, color:COLORS.muted, fontWeight:600 }}>{lbl}</div>
            </div>
          ))}
        </div>

        <div style={{ background:"#fff", borderRadius:12, padding:"14px 20px", border:`1px solid ${COLORS.border}`, marginBottom:20 }}>
          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}>
            <span style={{ fontSize:13, fontWeight:700, color:COLORS.navy }}>Progresso do round</span>
            <span style={{ fontSize:13, color:COLORS.muted }}>{roundsDone}/{total} leitos</span>
          </div>
          <div style={{ background:COLORS.border, borderRadius:8, height:8 }}>
            <div style={{ width:`${total?(roundsDone/total)*100:0}%`, height:"100%", background:COLORS.teal, borderRadius:8, transition:"width .4s" }} />
          </div>
        </div>

        <div style={{ display:"flex", gap:12, marginBottom:20, flexWrap:"wrap" }}>
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="🔍 Buscar paciente ou leito..."
            style={{ flex:1, minWidth:200, border:`1.5px solid ${COLORS.border}`, borderRadius:10, padding:"9px 16px", fontSize:14, color:COLORS.navy, outline:"none", background:"#fff" }} />
          {[["todos","Todos"],["pendente","Pendentes"],["alta","Alta gravidade"]].map(([v,l]) => (
            <button key={v} onClick={() => setFilter(v)} style={{ padding:"9px 18px", borderRadius:10, border:`1.5px solid ${filter===v?COLORS.teal:COLORS.border}`, background:filter===v?COLORS.teal:"#fff", color:filter===v?"#fff":COLORS.navy, fontSize:13, fontWeight:600, cursor:"pointer" }}>{l}</button>
          ))}
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))", gap:14 }}>
          {filtered.map(p => <PatientCard key={p.id} pat={p} round={rounds[p.id]} onSelect={setSelected} onEdit={setEditingId} />)}
        </div>
      </div>

      {editingId && editPat && <EditPatientModal pat={editPat} onSave={savePatient} onClear={clearPatient} onClose={() => setEditingId(null)} />}
      {showRelatorio && <RelatorioModal patients={patients} rounds={rounds} onClose={() => setShowRelatorio(false)} onSave={salvarRelatorio} />}
      {showHistorico && <HistoricoModal relatorios={relatorios} onClose={() => setShowHistorico(false)} />}
      {showClearAll  && <ClearAllModal onConfirm={clearAll} onClose={() => setShowClearAll(false)} />}
    </div>
  );
}
