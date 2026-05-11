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

const INITIAL_ROUND = {
  contato: null, visitaFlex: null,
  rass: null, dor: null, delirium: null, contencao: null,
  dva: null, pam: "",
  suporteResp: null, vmProtetora: null, planoDesmame: [], preocResp: [], ims: "", progredirFuncional: null,
  viaAlimentar: null, aceitacao: null, metaCalorica: null, fono: null, glicemia: null, evacuacao: null,
  funcaoRenal: null, metaBH: null, pioraInfec: null, atb: null,
  tev: null, lamg: null, cornea: null, higieneOral: null, decubito: null,
  bundles: null, bundlesPendente: "", mudancaDecubito: null, lesaoPressao: null, avalEspecializada: null,
  dispositivos: {
    cvc: { data: "", desinvadir: false }, hd: { data: "", desinvadir: false },
    svd: { data: "", desinvadir: false }, arterial: { data: "", desinvadir: false },
    sondaEnteral: { data: "", desinvadir: false }, drenos: { data: "", desinvadir: false },
  },
  diretivas: [], pendenciaExame: null, descPendencia: "", previsaoAlta: null,
  diagnostico: "",
};

// ── Helpers de data ───────────────────────────────────────────────────────────
function calcIdade(dataNasc) {
  if (!dataNasc) return null;
  const hoje = new Date();
  const nasc = new Date(dataNasc);
  let idade = hoje.getFullYear() - nasc.getFullYear();
  const m = hoje.getMonth() - nasc.getMonth();
  if (m < 0 || (m === 0 && hoje.getDate() < nasc.getDate())) idade--;
  return idade;
}
function calcDias(dataAdm) {
  if (!dataAdm) return 0;
  const hoje = new Date();
  const adm = new Date(dataAdm);
  const diff = Math.floor((hoje - adm) / (1000 * 60 * 60 * 24));
  return diff >= 0 ? diff : 0;
}
function hoje() {
  return new Date().toISOString().slice(0, 10);
}

const gravBadge = (g) => ({
  alta:  [COLORS.danger,  "⚠ Alta"],
  media: [COLORS.warn,    "◈ Média"],
  baixa: [COLORS.success, "✓ Baixa"],
  livre: [COLORS.muted,   "Livre"],
}[g] || [COLORS.muted, g]);

function computeAlerts(round, pat) {
  const a = [];
  if (!round || !pat.nome) return a;
  if (round.tev === "Não") a.push("Sem profilaxia TEV");
  if (round.suporteResp === "VM invasiva" && !(round.planoDesmame?.length)) a.push("VM sem plano de desmame");
  if (round.dva === "Sim" && !round.pam) a.push("DVA sem meta PAM");
  if (round.pendenciaExame === "Sim" && !round.descPendencia) a.push("Pendência sem descrição");
  const dias = calcDias(pat.dataAdm);
  if (dias > 7) a.push(`${dias} dias internado`);
  return a;
}

// ── Exportar Excel ────────────────────────────────────────────────────────────
function exportExcel(patients, rounds) {
  const rows = patients.filter(p => p.nome);
  if (!rows.length) { alert("Nenhum paciente para exportar."); return; }
  const data = rows.map(p => {
    const r   = rounds[p.id] || {};
    const dev = r.dispositivos || INITIAL_ROUND.dispositivos;
    return {
      "Leito": p.leito,
      "Paciente": p.nome,
      "Data de Nascimento": p.dataNasc || "",
      "Idade (anos)": calcIdade(p.dataNasc) ?? "",
      "Data de Admissão": p.dataAdm || "",
      "Dias Internado": calcDias(p.dataAdm),
      "Gravidade": p.gravidade,
      "Diagnóstico": r.diagnostico || p.diagnostico || "",
      "Precaução de Contato": r.contato || "",
      "Visita Flexibilizada": r.visitaFlex || "",
      "Meta Sedação RASS": r.rass || "",
      "Controle de Dor": r.dor || "",
      "Delirium": r.delirium || "",
      "Contenção Mecânica": r.contencao || "",
      "DVA": r.dva || "",
      "Meta PAM (mmHg)": r.pam || "",
      "Suporte Respiratório": r.suporteResp || "",
      "VM Protetora": r.vmProtetora || "",
      "Plano de Desmame": (r.planoDesmame || []).join(" | "),
      "Preocupações Resp.": (r.preocResp || []).join(" | "),
      "Escala IMS": r.ims || "",
      "Progredir Funcional": r.progredirFuncional || "",
      "Via Alimentar": r.viaAlimentar || "",
      "Aceitação": r.aceitacao || "",
      "Meta Calórica": r.metaCalorica || "",
      "Avaliação Fono": r.fono || "",
      "Controle Glicêmico": r.glicemia || "",
      "Evacuação < 3 dias": r.evacuacao || "",
      "Função Renal em Piora": r.funcaoRenal || "",
      "Meta Balanço Hídrico": r.metaBH || "",
      "Piora Infecciosa": r.pioraInfec || "",
      "ATB": r.atb || "",
      "Profilaxia TEV": r.tev || "",
      "Profilaxia LAMG": r.lamg || "",
      "Úlcera de Córnea": r.cornea || "",
      "Higiene Oral": r.higieneOral || "",
      "Decúbito Elevado": r.decubito || "",
      "Bundles OK": r.bundles || "",
      "Bundle Pendente": r.bundlesPendente || "",
      "Mudança Decúbito": r.mudancaDecubito || "",
      "Lesão por Pressão": r.lesaoPressao || "",
      "Avaliação Especializada": r.avalEspecializada || "",
      "CVC Data": dev.cvc?.data || "",           "CVC Desinvadir": dev.cvc?.desinvadir ? "Sim" : "Não",
      "HD Data": dev.hd?.data || "",             "HD Desinvadir": dev.hd?.desinvadir ? "Sim" : "Não",
      "SVD Data": dev.svd?.data || "",           "SVD Desinvadir": dev.svd?.desinvadir ? "Sim" : "Não",
      "Art. Data": dev.arterial?.data || "",     "Art. Desinvadir": dev.arterial?.desinvadir ? "Sim" : "Não",
      "SNE Data": dev.sondaEnteral?.data || "",  "SNE Desinvadir": dev.sondaEnteral?.desinvadir ? "Sim" : "Não",
      "Dreno Data": dev.drenos?.data || "",      "Dreno Desinvadir": dev.drenos?.desinvadir ? "Sim" : "Não",
      "Diretivas": (r.diretivas || []).join(" | "),
      "Pendência Exame": r.pendenciaExame || "",
      "Descrição Pendência": r.descPendencia || "",
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

// ── Modal editar paciente ─────────────────────────────────────────────────────
function EditPatientModal({ pat, onSave, onClear, onClose }) {
  const [nome,    setNome]    = useState(pat.nome || "");
  const [dataNasc,setNasc]   = useState(pat.dataNasc || "");
  const [dataAdm, setAdm]    = useState(pat.dataAdm || "");
  const [diag,    setDiag]   = useState(pat.diagnostico || "");
  const [grav,    setGrav]   = useState(pat.gravidade || "livre");
  const [confirm, setConfirm]= useState(false);
  const canSave = nome.trim().length > 0;
  const idadeCalc = calcIdade(dataNasc);
  const diasCalc  = calcDias(dataAdm);

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,37,69,.6)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: "#fff", borderRadius: 16, padding: "28px 30px", width: "100%", maxWidth: 460, boxShadow: "0 20px 60px rgba(0,0,0,.25)", maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 22 }}>
          <div>
            <div style={{ fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: "uppercase", letterSpacing: ".5px" }}>{pat.leito}</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: COLORS.navy }}>{pat.nome ? "Editar Paciente" : "Admitir Paciente"}</div>
          </div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: COLORS.muted }}>×</button>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <div>
            <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase", letterSpacing: ".4px" }}>Nome completo *</div>
            <input value={nome} onChange={e => setNome(e.target.value)} placeholder="Nome do paciente"
              style={{ border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "9px 12px", fontSize: 14, color: COLORS.navy, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <div>
              <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase", letterSpacing: ".4px" }}>Data de nascimento</div>
              <input type="date" value={dataNasc} onChange={e => setNasc(e.target.value)}
                style={{ border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "9px 12px", fontSize: 14, color: COLORS.navy, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
              {idadeCalc !== null && <div style={{ fontSize: 12, color: COLORS.teal, marginTop: 4, fontWeight: 600 }}>→ {idadeCalc} anos</div>}
            </div>
            <div>
              <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase", letterSpacing: ".4px" }}>Data de admissão na UTI</div>
              <input type="date" value={dataAdm} onChange={e => setAdm(e.target.value)}
                style={{ border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "9px 12px", fontSize: 14, color: COLORS.navy, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
              {dataAdm && <div style={{ fontSize: 12, color: COLORS.teal, marginTop: 4, fontWeight: 600 }}>→ {diasCalc} dia(s) internado</div>}
            </div>
          </div>
          <div>
            <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 5, textTransform: "uppercase", letterSpacing: ".4px" }}>Diagnóstico principal</div>
            <input value={diag} onChange={e => setDiag(e.target.value)} placeholder="Ex: Sepse pulmonar"
              style={{ border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "9px 12px", fontSize: 14, color: COLORS.navy, outline: "none", width: "100%", background: "#fff", boxSizing: "border-box" }} />
          </div>
          <div>
            <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, marginBottom: 8, textTransform: "uppercase", letterSpacing: ".4px" }}>Gravidade</div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {[["alta", COLORS.danger, "⚠ Alta"], ["media", COLORS.warn, "◈ Média"], ["baixa", COLORS.success, "✓ Baixa"], ["livre", COLORS.muted, "Livre"]].map(([v, c, l]) => (
                <button key={v} onClick={() => setGrav(v)} style={{ padding: "6px 14px", borderRadius: 8, cursor: "pointer", border: `1.5px solid ${grav === v ? c : COLORS.border}`, background: grav === v ? c + "1A" : "#fff", color: grav === v ? c : COLORS.navy, fontSize: 13, fontWeight: 700 }}>{l}</button>
              ))}
            </div>
          </div>
        </div>

        {pat.nome && (
          <div style={{ marginTop: 20, paddingTop: 16, borderTop: `1px solid ${COLORS.border}` }}>
            {!confirm ? (
              <button onClick={() => setConfirm(true)} style={{ width: "100%", padding: "9px", borderRadius: 8, border: `1.5px solid ${COLORS.danger}`, background: "#fff", color: COLORS.danger, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                🗑 Limpar dados e desocupar leito
              </button>
            ) : (
              <div style={{ background: COLORS.danger + "12", borderRadius: 10, padding: "14px" }}>
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
          <button
            onClick={() => { if (canSave) { onSave(pat.id, { nome: nome.trim(), dataNasc, dataAdm, diagnostico: diag, gravidade: grav }); onClose(); } }}
            style={{ flex: 2, padding: "10px", borderRadius: 10, border: "none", background: canSave ? COLORS.teal : COLORS.border, color: "#fff", fontSize: 14, fontWeight: 700, cursor: canSave ? "pointer" : "not-allowed" }}>
            ✓ Salvar
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Modal Limpar Todos ────────────────────────────────────────────────────────
function ClearAllModal({ onConfirm, onClose }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,37,69,.6)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: "#fff", borderRadius: 16, padding: "32px", width: "100%", maxWidth: 400, boxShadow: "0 20px 60px rgba(0,0,0,.25)", textAlign: "center" }}>
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
  const idade = calcIdade(pat.dataNasc);
  const dias  = calcDias(pat.dataAdm);

  return (
    <div style={{
      background: COLORS.card, borderRadius: 14,
      border: `1.5px solid ${isEmpty ? COLORS.border : color + "44"}`,
      padding: "14px 16px", position: "relative", overflow: "hidden",
      boxShadow: "0 2px 8px rgba(11,37,69,.06)",
    }}>
      {!isEmpty && <div style={{ position: "absolute", top: 0, left: 0, width: 4, height: "100%", background: color, borderRadius: "14px 0 0 14px" }} />}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginLeft: isEmpty ? 0 : 8 }}>
        <span style={{ fontSize: 11, fontWeight: 700, color: COLORS.muted, letterSpacing: ".5px" }}>{pat.leito}</span>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          {!isEmpty && <span style={{ fontSize: 11, color, fontWeight: 600, background: color + "18", padding: "2px 8px", borderRadius: 10 }}>{label}</span>}
          <button
            onClick={e => { e.stopPropagation(); onEdit(pat.id); }}
            title={isEmpty ? "Admitir paciente" : "Editar paciente"}
            style={{ background: "none", border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "2px 7px", cursor: "pointer", fontSize: 12, color: COLORS.muted }}>
            {isEmpty ? "＋" : "✏️"}
          </button>
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
            {pat.dataAdm && <span style={{ fontSize: 12, color: COLORS.muted }}>🕐 {dias}d internado</span>}
            {idade !== null && <span style={{ fontSize: 12, color: COLORS.muted }}>👤 {idade}a</span>}
          </div>
          {alerts.length > 0 && (
            <div style={{ marginTop: 8, marginLeft: 8, display: "flex", flexWrap: "wrap", gap: 4 }}>
              {alerts.map((a, i) => <span key={i} style={{ fontSize: 11, background: COLORS.danger + "15", color: COLORS.danger, padding: "2px 7px", borderRadius: 8, fontWeight: 600 }}>⚠ {a}</span>)}
            </div>
          )}
          <div style={{ marginTop: 8, marginLeft: 8 }}>
            {round
              ? <span style={{ fontSize: 11, background: COLORS.success + "20", color: COLORS.success, padding: "2px 8px", borderRadius: 8, fontWeight: 600 }}>✓ Round feito hoje</span>
              : <span style={{ fontSize: 11, background: COLORS.warn + "18", color: COLORS.warn, padding: "2px 8px", borderRadius: 8, fontWeight: 600 }}>⏳ Pendente</span>}
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
  const setDev = (dev, f, v) => onChange({ ...r, dispositivos: { ...r.dispositivos, [dev]: { ...r.dispositivos[dev], [f]: v } } });
  const s2 = ["Sim", "Não"], s3 = ["Sim", "Não", "Sem indicação"];
  const idade = calcIdade(pat.dataNasc);
  const dias  = calcDias(pat.dataAdm);

  return (
    <div style={{ maxWidth: 820, margin: "0 auto", paddingBottom: 80 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 24 }}>
        <button onClick={onBack} style={{ background: "none", border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "6px 14px", cursor: "pointer", fontSize: 13, color: COLORS.navy, fontWeight: 600 }}>← Voltar</button>
        <div>
          <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 700, letterSpacing: ".5px", textTransform: "uppercase" }}>{pat.leito} · Round — {new Date().toLocaleDateString("pt-BR")}</div>
          <div style={{ fontSize: 20, fontWeight: 800, color: COLORS.navy }}>{pat.nome}</div>
          <div style={{ fontSize: 12, color: COLORS.muted, marginTop: 2 }}>
            {idade !== null && `${idade} anos`}{idade !== null && pat.dataAdm && " · "}{pat.dataAdm && `${dias} dia(s) internado`}
          </div>
        </div>
      </div>

      {/* Diagnóstico */}
      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Diagnóstico" icon="🏥" />
        <TInput value={r.diagnostico} onChange={v => set("diagnostico", v)} placeholder="Ex: Sepse pulmonar" />
      </div>

      {/* Cuidados Gerais */}
      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Cuidados Gerais" icon="🛡️" />
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          <Field label="Precaução de contato?">{s2.map(o => <Pill key={o} label={o} selected={r.contato === o} onClick={() => set("contato", o)} />)}</Field>
          <Field label="Visita flexibilizada?">{s2.map(o => <Pill key={o} label={o} selected={r.visitaFlex === o} onClick={() => set("visitaFlex", o)} />)}</Field>
        </div>
      </div>

      {/* Neurológico */}
      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Neurológico" icon="🧠" />
        <Field label="Meta de sedação (RASS)">{["-5 a -4", "-2 a 0", "Não se aplica"].map(o => <Pill key={o} label={o} selected={r.rass === o} onClick={() => set("rass", o)} />)}</Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
          <Field label="Controle de dor?">{s2.map(o => <Pill key={o} label={o} selected={r.dor === o} onClick={() => set("dor", o)} />)}</Field>
          <Field label="Delirium?">{s2.map(o => <Pill key={o} label={o} selected={r.delirium === o} onClick={() => set("delirium", o)} />)}</Field>
          <Field label="Contenção mecânica?">{s2.map(o => <Pill key={o} label={o} selected={r.contencao === o} onClick={() => set("contencao", o)} />)}</Field>
        </div>
      </div>

      {/* Cardiovascular */}
      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Cardiovascular / Hemodinâmica" icon="❤️" />
        <div style={{ display: "flex", gap: 24, flexWrap: "wrap", alignItems: "flex-end" }}>
          <Field label="Suporte hemodinâmico / DVA?">{s2.map(o => <Pill key={o} label={o} selected={r.dva === o} onClick={() => set("dva", o)} />)}</Field>
          <Field label="Meta de PAM (mmHg)"><TInput value={r.pam} onChange={v => set("pam", v)} placeholder="Ex: 65" w="90px" /></Field>
        </div>
      </div>

      {/* Respiratório */}
      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Respiratório / Reabilitação" icon="🫁" />
        <Field label="Suporte respiratório">{["Sem suporte", "O2 suplementar", "VNI", "VM invasiva"].map(o => <Pill key={o} label={o} selected={r.suporteResp === o} onClick={() => set("suporteResp", o)} />)}</Field>
        {(r.suporteResp === "VM invasiva" || r.suporteResp === "VNI") && <Field label="VM protetora?">{s2.map(o => <Pill key={o} label={o} selected={r.vmProtetora === o} onClick={() => set("vmProtetora", o)} />)}</Field>}
        <Field label="Plano de desmame">{["Redução de parâmetros", "TRE hoje", "Ex-TOT", "Sem proposta de desmame"].map(o => <MultiPill key={o} label={o} checked={(r.planoDesmame || []).includes(o)} onChange={() => tog("planoDesmame", o)} />)}</Field>
        <Field label="Preocupações respiratórias">{["Piora respiratória", "Piora de secreção", "Risco de broncoaspiração"].map(o => <MultiPill key={o} label={o} checked={(r.preocResp || []).includes(o)} onChange={() => tog("preocResp", o)} color={COLORS.danger} />)}</Field>
        <div style={{ display: "flex", gap: 24, flexWrap: "wrap" }}>
          <Field label="Escala IMS"><TInput value={r.ims} onChange={v => set("ims", v)} placeholder="0-10" w="80px" /></Field>
          <Field label="Progredir nível funcional?">{s2.map(o => <Pill key={o} label={o} selected={r.progredirFuncional === o} onClick={() => set("progredirFuncional", o)} />)}</Field>
        </div>
      </div>

      {/* GI */}
      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Gastrointestinal / Nutrição" icon="🍽️" />
        <Field label="Via alimentar">{["VO", "SNE/GTT", "NPT", "Zero"].map(o => <Pill key={o} label={o} selected={r.viaAlimentar === o} onClick={() => set("viaAlimentar", o)} />)}</Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
          <Field label="Aceitação">{["Normal", "Baixa", "Intolerância"].map(o => <Pill key={o} label={o} selected={r.aceitacao === o} onClick={() => set("aceitacao", o)} />)}</Field>
          <Field label="Meta calórica?">{s2.map(o => <Pill key={o} label={o} selected={r.metaCalorica === o} onClick={() => set("metaCalorica", o)} />)}</Field>
          <Field label="Avaliação fono?">{s2.map(o => <Pill key={o} label={o} selected={r.fono === o} onClick={() => set("fono", o)} />)}</Field>
          <Field label="Glicemia adequada?">{s2.map(o => <Pill key={o} label={o} selected={r.glicemia === o} onClick={() => set("glicemia", o)} />)}</Field>
          <Field label="Evacuação < 3 dias?">{s2.map(o => <Pill key={o} label={o} selected={r.evacuacao === o} onClick={() => set("evacuacao", o)} />)}</Field>
        </div>
      </div>

      {/* Renal + Infeccioso */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 14 }}>
        <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", border: `1px solid ${COLORS.border}` }}>
          <SecHdr title="Renal" icon="🫘" />
          <Field label="Função renal em piora?">{["Sim", "Não", "Em HD"].map(o => <Pill key={o} label={o} selected={r.funcaoRenal === o} onClick={() => set("funcaoRenal", o)} />)}</Field>
          <Field label="Meta de balanço hídrico">{["Positivo", "Negativo", "Neutro"].map(o => <Pill key={o} label={o} selected={r.metaBH === o} onClick={() => set("metaBH", o)} />)}</Field>
        </div>
        <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", border: `1px solid ${COLORS.border}` }}>
          <SecHdr title="Infeccioso" icon="🦠" />
          <Field label="Piora infecciosa?">{s2.map(o => <Pill key={o} label={o} selected={r.pioraInfec === o} onClick={() => set("pioraInfec", o)} />)}</Field>
          <Field label="Em uso de ATB?">{s2.map(o => <Pill key={o} label={o} selected={r.atb === o} onClick={() => set("atb", o)} />)}</Field>
        </div>
      </div>

      {/* Profilaxias */}
      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Profilaxias" icon="💉" />
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
          <Field label="Profilaxia TEV">{s3.map(o => <Pill key={o} label={o} selected={r.tev === o} onClick={() => set("tev", o)} color={o === "Não" ? COLORS.danger : COLORS.teal} />)}</Field>
          <Field label="Profilaxia LAMG">{s3.map(o => <Pill key={o} label={o} selected={r.lamg === o} onClick={() => set("lamg", o)} />)}</Field>
          <Field label="Úlcera de córnea">{s3.map(o => <Pill key={o} label={o} selected={r.cornea === o} onClick={() => set("cornea", o)} />)}</Field>
          <Field label="Higiene oral">{s3.map(o => <Pill key={o} label={o} selected={r.higieneOral === o} onClick={() => set("higieneOral", o)} />)}</Field>
          <Field label="Decúbito elevado">{s3.map(o => <Pill key={o} label={o} selected={r.decubito === o} onClick={() => set("decubito", o)} />)}</Field>
          <Field label="Bundles OK?">{s2.map(o => <Pill key={o} label={o} selected={r.bundles === o} onClick={() => set("bundles", o)} color={o === "Não" ? COLORS.danger : COLORS.teal} />)}</Field>
        </div>
        {r.bundles === "Não" && <div style={{ marginTop: 8 }}><Field label="Bundle pendente"><TInput value={r.bundlesPendente} onChange={v => set("bundlesPendente", v)} placeholder="Qual bundle?" /></Field></div>}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14, marginTop: 4 }}>
          <Field label="Mudança de decúbito?">{s2.map(o => <Pill key={o} label={o} selected={r.mudancaDecubito === o} onClick={() => set("mudancaDecubito", o)} />)}</Field>
          <Field label="Lesão por pressão?">{s2.map(o => <Pill key={o} label={o} selected={r.lesaoPressao === o} onClick={() => set("lesaoPressao", o)} />)}</Field>
          <Field label="Avaliação especializada?">{["Sim", "Não", "Comissão curativo", "Cirurgia plástica", "N/A"].map(o => <Pill key={o} label={o} selected={r.avalEspecializada === o} onClick={() => set("avalEspecializada", o)} />)}</Field>
        </div>
      </div>

      {/* Dispositivos */}
      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Dispositivos Invasivos" icon="🔌" />
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          {[["cvc", "Cateter Venoso Central"], ["hd", "Cateter de HD"], ["svd", "SVD"], ["arterial", "Cateter Arterial"], ["sondaEnteral", "Sonda Enteral"], ["drenos", "Drenos"]].map(([key, lbl]) => (
            <div key={key} style={{ background: COLORS.lightBg, borderRadius: 10, padding: "12px 14px", border: `1px solid ${COLORS.border}` }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.navy, marginBottom: 8 }}>{lbl}</div>
              <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                <input type="date" value={r.dispositivos[key]?.data || ""} onChange={e => setDev(key, "data", e.target.value)}
                  style={{ border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "4px 8px", fontSize: 12, color: COLORS.navy, background: "#fff" }} />
                <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: COLORS.teal, fontWeight: 600, cursor: "pointer" }}>
                  <input type="checkbox" checked={r.dispositivos[key]?.desinvadir || false} onChange={e => setDev(key, "desinvadir", e.target.checked)} />
                  Desinvadir?
                </label>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Objetivos */}
      <div style={{ background: COLORS.card, borderRadius: 14, padding: "18px 20px", marginBottom: 14, border: `1px solid ${COLORS.border}` }}>
        <SecHdr title="Objetivos de Cuidado e Planejamento" icon="📋" />
        <Field label="Diretivas de cuidado">{["Não RCP", "Não HD", "Não IOT", "Não DVA", "Não coletar exames", "Sem diretivas"].map(o => <MultiPill key={o} label={o} checked={(r.diretivas || []).includes(o)} onChange={() => tog("diretivas", o)} color={o === "Sem diretivas" ? COLORS.success : COLORS.danger} />)}</Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          <div>
            <Field label="Pendência de exame/procedimento?">{s2.map(o => <Pill key={o} label={o} selected={r.pendenciaExame === o} onClick={() => set("pendenciaExame", o)} />)}</Field>
            {r.pendenciaExame === "Sim" && <TArea value={r.descPendencia} onChange={v => set("descPendencia", v)} placeholder="Descreva a pendência..." />}
          </div>
          <Field label="Previsão de alta">{["Hoje", "24–48h", "> 48h"].map(o => <Pill key={o} label={o} selected={r.previsaoAlta === o} onClick={() => set("previsaoAlta", o)} />)}</Field>
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 12 }}>
        <button onClick={onBack} style={{ padding: "10px 24px", borderRadius: 10, border: `1.5px solid ${COLORS.border}`, background: "#fff", color: COLORS.navy, fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Cancelar</button>
        <button onClick={onBack} style={{ padding: "10px 28px", borderRadius: 10, border: "none", background: COLORS.teal, color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>✓ Salvar Round</button>
      </div>
    </div>
  );
}

// ── Report Modal ──────────────────────────────────────────────────────────────
function ReportModal({ patients, rounds, onClose }) {
  const rows = patients.filter(p => p.nome);
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,37,69,.55)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: "#fff", borderRadius: 16, padding: "28px 32px", maxWidth: 960, width: "100%", maxHeight: "85vh", overflow: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <div style={{ fontWeight: 800, fontSize: 20, color: COLORS.navy }}>📊 Resumo Geral da UTI</div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: COLORS.muted }}>×</button>
        </div>
        {rows.length === 0 ? (
          <div style={{ textAlign: "center", padding: "40px", color: COLORS.muted }}>Nenhum paciente admitido.</div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead><tr style={{ background: COLORS.navy, color: "#fff" }}>
                {["Leito", "Paciente", "Idade", "Dias", "Diagnóstico", "VM", "DVA", "ATB", "Alta", "Pendências", "Alertas"].map(h => (
                  <th key={h} style={{ padding: "10px 12px", textAlign: "left", fontWeight: 700, whiteSpace: "nowrap" }}>{h}</th>
                ))}
              </tr></thead>
              <tbody>{rows.map((p, i) => {
                const r = rounds[p.id]; const alerts = computeAlerts(r, p);
                return <tr key={p.id} style={{ background: i % 2 === 0 ? "#fff" : COLORS.lightBg }}>
                  <td style={{ padding: "9px 12px", fontWeight: 700, color: COLORS.navy }}>{p.leito}</td>
                  <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{p.nome}</td>
                  <td style={{ padding: "9px 12px", textAlign: "center" }}>{calcIdade(p.dataNasc) ?? "-"}</td>
                  <td style={{ padding: "9px 12px", textAlign: "center" }}>{calcDias(p.dataAdm)}d</td>
                  <td style={{ padding: "9px 12px", color: COLORS.muted }}>{r?.diagnostico || p.diagnostico}</td>
                  <td style={{ padding: "9px 12px", textAlign: "center" }}><span style={{ color: r?.suporteResp === "VM invasiva" ? COLORS.danger : COLORS.success, fontWeight: 700 }}>{r?.suporteResp === "VM invasiva" ? "Sim" : "Não"}</span></td>
                  <td style={{ padding: "9px 12px", textAlign: "center" }}><span style={{ color: r?.dva === "Sim" ? COLORS.danger : COLORS.success, fontWeight: 700 }}>{r?.dva || "-"}</span></td>
                  <td style={{ padding: "9px 12px", textAlign: "center" }}>{r?.atb || "-"}</td>
                  <td style={{ padding: "9px 12px" }}>{r?.previsaoAlta || "-"}</td>
                  <td style={{ padding: "9px 12px", color: COLORS.warn, maxWidth: 160 }}>{r?.descPendencia || "-"}</td>
                  <td style={{ padding: "9px 12px" }}>{alerts.map((a, j) => <div key={j} style={{ color: COLORS.danger, fontSize: 11, fontWeight: 600 }}>⚠ {a}</div>)}</td>
                </tr>;
              })}</tbody>
            </table>
          </div>
        )}
        <div style={{ display: "flex", gap: 12, marginTop: 20, justifyContent: "flex-end" }}>
          <button onClick={() => exportExcel(patients, rounds)} style={{ padding: "9px 22px", borderRadius: 10, border: "none", background: COLORS.teal, color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>⬇ Exportar Excel (.xlsx)</button>
          <button onClick={onClose} style={{ padding: "9px 22px", borderRadius: 10, border: `1.5px solid ${COLORS.border}`, background: "#fff", color: COLORS.navy, fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Fechar</button>
        </div>
      </div>
    </div>
  );
}

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [patients,     setPatients]     = useState([]);
  const [rounds,       setRounds]       = useState({});
  const [selected,     setSelected]     = useState(null);
  const [editingId,    setEditingId]    = useState(null);
  const [showReport,   setShowReport]   = useState(false);
  const [showClearAll, setShowClearAll] = useState(false);
  const [search,       setSearch]       = useState("");
  const [filter,       setFilter]       = useState("todos");
  const [loading,      setLoading]      = useState(true);
  const [saving,       setSaving]       = useState(false);
  const [error,        setError]        = useState(null);

  const dataHoje = hoje();

  // ── Carregar dados ──────────────────────────────────────────────────────────
  useEffect(() => {
    async function load() {
      setLoading(true);
      setError(null);
      try {
        const { data: pats, error: e1 } = await supabase
          .from("patients").select("*").order("id");
        if (e1) throw e1;

        const { data: rds, error: e2 } = await supabase
          .from("rounds").select("*").eq("data", dataHoje);
        if (e2) throw e2;

        // Garantir estrutura mínima em cada paciente
        const patsNorm = (pats || []).map(p => ({
          ...p,
          nome:        p.nome        || "",
          diagnostico: p.diagnostico || "",
          gravidade:   p.gravidade   || "livre",
          dataNasc:    p.dataNasc    || null,
          dataAdm:     p.dataAdm     || null,
        }));

        setPatients(patsNorm);

        const map = {};
        (rds || []).forEach(r => {
          // Garantir que dispositivos existe no round carregado
          const rd = r.round_data || {};
          if (!rd.dispositivos) rd.dispositivos = INITIAL_ROUND.dispositivos;
          map[r.patient_id] = rd;
        });
        setRounds(map);
      } catch (err) {
        setError("Erro ao conectar ao banco de dados. Verifique sua conexão.");
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
    await supabase.from("rounds").upsert({
      patient_id: selected,
      data: dataHoje,
      round_data: r,
      updated_at: new Date().toISOString(),
    }, { onConflict: "patient_id,data" });
    setSaving(false);
  }, [selected]);

  // ── Salvar paciente ─────────────────────────────────────────────────────────
  const savePatient = async (id, data) => {
    setPatients(prev => prev.map(p => p.id === id ? { ...p, ...data } : p));
    await supabase.from("patients").update({
      nome:        data.nome,
      dataNasc:    data.dataNasc || null,
      dataAdm:     data.dataAdm  || null,
      diagnostico: data.diagnostico,
      gravidade:   data.gravidade,
      updated_at:  new Date().toISOString(),
    }).eq("id", id);
  };

  // ── Limpar paciente ─────────────────────────────────────────────────────────
  const clearPatient = async (id) => {
    setPatients(prev => prev.map(p => p.id === id
      ? { ...p, nome: "", dataNasc: null, dataAdm: null, diagnostico: "", gravidade: "livre" } : p));
    setRounds(prev => { const n = { ...prev }; delete n[id]; return n; });
    await supabase.from("patients").update({ nome: "", dataNasc: null, dataAdm: null, diagnostico: "", gravidade: "livre" }).eq("id", id);
    await supabase.from("rounds").delete().eq("patient_id", id);
  };

  // ── Limpar todos ────────────────────────────────────────────────────────────
  const clearAll = async () => {
    setPatients(prev => prev.map(p => ({ ...p, nome: "", dataNasc: null, dataAdm: null, diagnostico: "", gravidade: "livre" })));
    setRounds({});
    setShowClearAll(false);
    await supabase.from("patients").update({ nome: "", dataNasc: null, dataAdm: null, diagnostico: "", gravidade: "livre" }).gte("id", 1);
    await supabase.from("rounds").delete().gte("patient_id", 1);
  };

  const selPat  = patients.find(p => p.id === selected);
  const editPat = patients.find(p => p.id === editingId);
  const date    = new Date().toLocaleDateString("pt-BR", { weekday: "long", day: "2-digit", month: "long", year: "numeric" });

  // Só conta como ocupado quem tem nome
  const withPats   = patients.filter(p => p.nome);
  const roundsDone = withPats.filter(p => rounds[p.id]).length;
  const total      = withPats.length;

  // Filtros — leitos vazios sempre aparecem em "Todos"
  const filtered = patients.filter(p => {
    if (search) {
      const q = search.toLowerCase();
      if (!p.nome?.toLowerCase().includes(q) && !p.leito?.toLowerCase().includes(q)) return false;
    }
    if (filter === "pendente" && (!p.nome || rounds[p.id])) return false;
    if (filter === "alta" && p.gravidade !== "alta") return false;
    return true;
  });

  // ── Loading ─────────────────────────────────────────────────────────────────
  if (loading) return (
    <div style={{ background: COLORS.bg, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'DM Sans','Segoe UI',sans-serif" }}>
      <div style={{ textAlign: "center" }}>
        <div style={{ fontSize: 48, marginBottom: 16 }}>🏥</div>
        <div style={{ fontSize: 18, fontWeight: 700, color: COLORS.navy }}>Carregando dados da UTI...</div>
        <div style={{ fontSize: 13, color: COLORS.muted, marginTop: 8 }}>Conectando ao banco de dados</div>
      </div>
    </div>
  );

  if (error) return (
    <div style={{ background: COLORS.bg, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'DM Sans','Segoe UI',sans-serif" }}>
      <div style={{ textAlign: "center", maxWidth: 400 }}>
        <div style={{ fontSize: 48, marginBottom: 16 }}>⚠️</div>
        <div style={{ fontSize: 18, fontWeight: 700, color: COLORS.danger, marginBottom: 8 }}>{error}</div>
        <button onClick={() => window.location.reload()} style={{ padding: "10px 24px", borderRadius: 10, border: "none", background: COLORS.teal, color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>Tentar novamente</button>
      </div>
    </div>
  );

  // ── Tela de round ───────────────────────────────────────────────────────────
  if (selected && selPat) return (
    <div style={{ background: COLORS.bg, minHeight: "100vh", fontFamily: "'DM Sans','Segoe UI',sans-serif" }}>
      <div style={{ background: COLORS.navy, padding: "12px 24px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ color: "#fff", fontWeight: 800, fontSize: 16 }}>🏥 UTI Clínica — IMIP</div>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          {saving && <span style={{ color: "#8BBBD9", fontSize: 12 }}>💾 Salvando...</span>}
          <div style={{ color: "#8BBBD9", fontSize: 13 }}>{date}</div>
        </div>
      </div>
      <div style={{ padding: "24px" }}>
        <RoundForm
          pat={selPat}
          round={rounds[selected] || { ...INITIAL_ROUND, diagnostico: selPat.diagnostico || "" }}
          onChange={handleChange}
          onBack={() => setSelected(null)}
        />
      </div>
    </div>
  );

  // ── Dashboard ───────────────────────────────────────────────────────────────
  return (
    <div style={{ background: COLORS.bg, minHeight: "100vh", fontFamily: "'DM Sans','Segoe UI',sans-serif" }}>
      <div style={{ background: COLORS.navy, padding: "13px 28px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <span style={{ fontSize: 22 }}>🏥</span>
          <div>
            <div style={{ color: "#fff", fontWeight: 800, fontSize: 17 }}>UTI Clínica — IMIP</div>
            <div style={{ color: "#8BBBD9", fontSize: 12 }}>Round Multidisciplinar · 10 leitos</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ color: "#8BBBD9", fontSize: 13 }}>{date}</div>
          <button onClick={() => setShowClearAll(true)} style={{ padding: "8px 16px", borderRadius: 10, border: `1.5px solid ${COLORS.danger}55`, background: "transparent", color: COLORS.danger, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>🗑 Limpar todos</button>
          <button onClick={() => setShowReport(true)} style={{ padding: "8px 18px", borderRadius: 10, border: "none", background: COLORS.accent, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>📊 Relatório</button>
        </div>
      </div>

      <div style={{ padding: "24px 28px" }}>
        {/* Stats */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14, marginBottom: 20 }}>
          {[
            ["Leitos ocupados", total, COLORS.navy, "🛏"],
            ["Rounds feitos", roundsDone, COLORS.success, "✅"],
            ["Pendentes", total - roundsDone, COLORS.warn, "⏳"],
            ["Alertas", patients.filter(p => computeAlerts(rounds[p.id], p).length > 0).length, COLORS.danger, "⚠️"],
          ].map(([lbl, val, color, icon]) => (
            <div key={lbl} style={{ background: "#fff", borderRadius: 12, padding: "16px 20px", border: `1px solid ${COLORS.border}` }}>
              <div style={{ fontSize: 22, marginBottom: 4 }}>{icon}</div>
              <div style={{ fontSize: 26, fontWeight: 800, color }}>{val}</div>
              <div style={{ fontSize: 12, color: COLORS.muted, fontWeight: 600 }}>{lbl}</div>
            </div>
          ))}
        </div>

        {/* Progress */}
        <div style={{ background: "#fff", borderRadius: 12, padding: "14px 20px", border: `1px solid ${COLORS.border}`, marginBottom: 20 }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
            <span style={{ fontSize: 13, fontWeight: 700, color: COLORS.navy }}>Progresso do round</span>
            <span style={{ fontSize: 13, color: COLORS.muted }}>{roundsDone}/{total} leitos</span>
          </div>
          <div style={{ background: COLORS.border, borderRadius: 8, height: 8 }}>
            <div style={{ width: `${total ? (roundsDone / total) * 100 : 0}%`, height: "100%", background: COLORS.teal, borderRadius: 8, transition: "width .4s" }} />
          </div>
        </div>

        {/* Filtros */}
        <div style={{ display: "flex", gap: 12, marginBottom: 20, flexWrap: "wrap" }}>
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="🔍 Buscar paciente ou leito..."
            style={{ flex: 1, minWidth: 200, border: `1.5px solid ${COLORS.border}`, borderRadius: 10, padding: "9px 16px", fontSize: 14, color: COLORS.navy, outline: "none", background: "#fff" }} />
          {[["todos", "Todos"], ["pendente", "Pendentes"], ["alta", "Alta gravidade"]].map(([v, l]) => (
            <button key={v} onClick={() => setFilter(v)} style={{ padding: "9px 18px", borderRadius: 10, border: `1.5px solid ${filter === v ? COLORS.teal : COLORS.border}`, background: filter === v ? COLORS.teal : "#fff", color: filter === v ? "#fff" : COLORS.navy, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>{l}</button>
          ))}
        </div>

        {/* Grid leitos */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(220px,1fr))", gap: 14 }}>
          {filtered.map(p => (
            <PatientCard key={p.id} pat={p} round={rounds[p.id]} onSelect={setSelected} onEdit={setEditingId} />
          ))}
        </div>
      </div>

      {editingId && editPat && <EditPatientModal pat={editPat} onSave={savePatient} onClear={clearPatient} onClose={() => setEditingId(null)} />}
      {showReport    && <ReportModal patients={patients} rounds={rounds} onClose={() => setShowReport(false)} />}
      {showClearAll  && <ClearAllModal onConfirm={clearAll} onClose={() => setShowClearAll(false)} />}
    </div>
  );
}
