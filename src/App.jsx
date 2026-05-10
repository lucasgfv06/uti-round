import { useState, useCallback } from "react";

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
  resumoClinico: "", planoDia: "", diagnostico: "",
};

const INITIAL_PATIENTS = [
  { id: 1,  leito: "L01", nome: "Maria Silva",   idade: 67, dias: 5,  diagnostico: "Sepse pulmonar",    gravidade: "alta"  },
  { id: 2,  leito: "L02", nome: "João Ferreira", idade: 54, dias: 2,  diagnostico: "IAM com choque",    gravidade: "alta"  },
  { id: 3,  leito: "L03", nome: "Ana Costa",     idade: 72, dias: 8,  diagnostico: "TCE grave",         gravidade: "media" },
  { id: 4,  leito: "L04", nome: "Carlos Mendes", idade: 45, dias: 1,  diagnostico: "Pós-op cardíaco",   gravidade: "baixa" },
  { id: 5,  leito: "L05", nome: "Rosa Lima",     idade: 81, dias: 12, diagnostico: "DPOC exacerbado",   gravidade: "media" },
  { id: 6,  leito: "L06", nome: "Pedro Alves",   idade: 38, dias: 3,  diagnostico: "Pancreatite grave", gravidade: "alta"  },
  { id: 7,  leito: "L07", nome: "", idade: null, dias: 0, diagnostico: "", gravidade: "livre" },
  { id: 8,  leito: "L08", nome: "", idade: null, dias: 0, diagnostico: "", gravidade: "livre" },
  { id: 9,  leito: "L09", nome: "", idade: null, dias: 0, diagnostico: "", gravidade: "livre" },
  { id: 10, leito: "L10", nome: "", idade: null, dias: 0, diagnostico: "", gravidade: "livre" },
];

const gravBadge = (g) => ({
  alta:  [COLORS.danger,  "⚠ Alta"],
  media: [COLORS.warn,    "◈ Média"],
  baixa: [COLORS.success, "✓ Baixa"],
  livre: [COLORS.muted,   "Livre"],
}[g] || [COLORS.muted, g]);

function computeAlerts(round, pat) {
  const a = [];
  if (!round) return a;
  if (round.tev === "Não") a.push("Sem profilaxia TEV");
  if (round.suporteResp === "VM invasiva" && !(round.planoDesmame?.length)) a.push("VM sem plano de desmame");
  if (round.dva === "Sim" && !round.pam) a.push("DVA sem meta PAM");
  if (round.pendenciaExame === "Sim" && !round.descPendencia) a.push("Pendência sem descrição");
  if (pat.dias > 7) a.push(`${pat.dias} dias internado`);
  return a;
}

// ── Atoms ────────────────────────────────────────────────────────────────────
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
    <input value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} style={{
      border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "6px 12px",
      fontSize: 13, color: COLORS.navy, outline: "none", background: "#fff",
      width: w || "100%", maxWidth: w || 240, boxSizing: "border-box",
    }} />
  );
}
function TArea({ value, onChange, placeholder, rows = 2 }) {
  return (
    <textarea value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} rows={rows} style={{
      border: `1.5px solid ${COLORS.border}`, borderRadius: 8, padding: "8px 12px",
      fontSize: 13, color: COLORS.navy, outline: "none", width: "100%",
      resize: "vertical", background: "#fff", fontFamily: "inherit", boxSizing: "border-box",
    }} />
  );
}

// ── Modal editar/admitir paciente ─────────────────────────────────────────────
function EditPatientModal({ pat, onSave, onClear, onClose }) {
  const [nome,       setNome]       = useState(pat.nome || "");
  const [idade,      setIdade]      = useState(pat.idade ? String(pat.idade) : "");
  const [dias,       setDias]       = useState(pat.dias  ? String(pat.dias)  : "");
  const [diag,       setDiag]       = useState(pat.diagnostico || "");
  const [grav,       setGrav]       = useState(pat.gravidade || "livre");
  const [confirmClear, setConfirm]  = useState(false);

  const canSave = nome.trim().length > 0;

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(11,37,69,.6)", zIndex:200, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div style={{ background:"#fff", borderRadius:16, padding:"28px 30px", width:"100%", maxWidth:440, boxShadow:"0 20px 60px rgba(0,0,0,.25)" }}>

        {/* Header */}
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:22 }}>
          <div>
            <div style={{ fontSize:11, fontWeight:700, color:COLORS.muted, textTransform:"uppercase", letterSpacing:".5px" }}>{pat.leito}</div>
            <div style={{ fontSize:18, fontWeight:800, color:COLORS.navy }}>{pat.nome ? "Editar Paciente" : "Admitir Paciente"}</div>
          </div>
          <button onClick={onClose} style={{ background:"none", border:"none", fontSize:22, cursor:"pointer", color:COLORS.muted }}>×</button>
        </div>

        <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
          {/* Nome */}
          <div>
            <div style={{ fontSize:11, color:COLORS.muted, fontWeight:700, marginBottom:5, textTransform:"uppercase", letterSpacing:".4px" }}>Nome completo *</div>
            <input value={nome} onChange={e => setNome(e.target.value)} placeholder="Nome do paciente"
              style={{ border:`1.5px solid ${canSave || !nome ? COLORS.border : COLORS.danger}`, borderRadius:8, padding:"9px 12px", fontSize:14, color:COLORS.navy, outline:"none", width:"100%", background:"#fff", boxSizing:"border-box" }} />
          </div>

          {/* Idade + Dias */}
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
            <div>
              <div style={{ fontSize:11, color:COLORS.muted, fontWeight:700, marginBottom:5, textTransform:"uppercase", letterSpacing:".4px" }}>Idade</div>
              <input type="number" value={idade} onChange={e => setIdade(e.target.value)} placeholder="Anos"
                style={{ border:`1.5px solid ${COLORS.border}`, borderRadius:8, padding:"9px 12px", fontSize:14, color:COLORS.navy, outline:"none", width:"100%", background:"#fff", boxSizing:"border-box" }} />
            </div>
            <div>
              <div style={{ fontSize:11, color:COLORS.muted, fontWeight:700, marginBottom:5, textTransform:"uppercase", letterSpacing:".4px" }}>Dias internado</div>
              <input type="number" value={dias} onChange={e => setDias(e.target.value)} placeholder="Dias"
                style={{ border:`1.5px solid ${COLORS.border}`, borderRadius:8, padding:"9px 12px", fontSize:14, color:COLORS.navy, outline:"none", width:"100%", background:"#fff", boxSizing:"border-box" }} />
            </div>
          </div>

          {/* Diagnóstico */}
          <div>
            <div style={{ fontSize:11, color:COLORS.muted, fontWeight:700, marginBottom:5, textTransform:"uppercase", letterSpacing:".4px" }}>Diagnóstico principal</div>
            <input value={diag} onChange={e => setDiag(e.target.value)} placeholder="Ex: Sepse pulmonar"
              style={{ border:`1.5px solid ${COLORS.border}`, borderRadius:8, padding:"9px 12px", fontSize:14, color:COLORS.navy, outline:"none", width:"100%", background:"#fff", boxSizing:"border-box" }} />
          </div>

          {/* Gravidade */}
          <div>
            <div style={{ fontSize:11, color:COLORS.muted, fontWeight:700, marginBottom:8, textTransform:"uppercase", letterSpacing:".4px" }}>Gravidade</div>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
              {[["alta", COLORS.danger, "⚠ Alta"], ["media", COLORS.warn, "◈ Média"], ["baixa", COLORS.success, "✓ Baixa"], ["livre", COLORS.muted, "Livre"]].map(([v, c, l]) => (
                <button key={v} onClick={() => setGrav(v)} style={{
                  padding:"6px 14px", borderRadius:8, cursor:"pointer",
                  border:`1.5px solid ${grav === v ? c : COLORS.border}`,
                  background: grav === v ? c + "1A" : "#fff",
                  color: grav === v ? c : COLORS.navy, fontSize:13, fontWeight:700,
                }}>{l}</button>
              ))}
            </div>
          </div>
        </div>

        {/* Limpar dados */}
        {pat.nome && (
          <div style={{ marginTop:20, paddingTop:16, borderTop:`1px solid ${COLORS.border}` }}>
            {!confirmClear ? (
              <button onClick={() => setConfirm(true)} style={{ width:"100%", padding:"9px", borderRadius:8, border:`1.5px solid ${COLORS.danger}`, background:"#fff", color:COLORS.danger, fontSize:13, fontWeight:700, cursor:"pointer" }}>
                🗑 Limpar dados e desocupar leito
              </button>
            ) : (
              <div style={{ background:COLORS.danger + "12", borderRadius:10, padding:"14px" }}>
                <div style={{ fontSize:13, color:COLORS.danger, fontWeight:700, marginBottom:10 }}>Tem certeza? Todos os dados e o round serão apagados.</div>
                <div style={{ display:"flex", gap:8 }}>
                  <button onClick={() => { onClear(pat.id); onClose(); }} style={{ flex:1, padding:"9px", borderRadius:8, border:"none", background:COLORS.danger, color:"#fff", fontSize:13, fontWeight:700, cursor:"pointer" }}>Sim, limpar</button>
                  <button onClick={() => setConfirm(false)} style={{ flex:1, padding:"9px", borderRadius:8, border:`1.5px solid ${COLORS.border}`, background:"#fff", color:COLORS.navy, fontSize:13, fontWeight:600, cursor:"pointer" }}>Cancelar</button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Ações */}
        <div style={{ display:"flex", gap:10, marginTop:20 }}>
          <button onClick={onClose} style={{ flex:1, padding:"10px", borderRadius:10, border:`1.5px solid ${COLORS.border}`, background:"#fff", color:COLORS.navy, fontSize:14, fontWeight:600, cursor:"pointer" }}>Cancelar</button>
          <button
            onClick={() => { if (canSave) { onSave(pat.id, { nome: nome.trim(), idade: idade ? parseInt(idade) : null, dias: dias ? parseInt(dias) : 0, diagnostico: diag, gravidade: grav }); onClose(); } }}
            style={{ flex:2, padding:"10px", borderRadius:10, border:"none", background: canSave ? COLORS.teal : COLORS.border, color:"#fff", fontSize:14, fontWeight:700, cursor: canSave ? "pointer" : "not-allowed" }}
          >✓ Salvar</button>
        </div>
      </div>
    </div>
  );
}

// ── Patient Card ─────────────────────────────────────────────────────────────
function PatientCard({ pat, round, onSelect, onEdit }) {
  const [color, label] = gravBadge(pat.gravidade);
  const alerts = computeAlerts(round, pat);
  const isEmpty = !pat.nome;

  return (
    <div style={{ background:COLORS.card, borderRadius:14, border:`1.5px solid ${isEmpty ? COLORS.border : color+"44"}`, padding:"14px 16px", position:"relative", overflow:"hidden", boxShadow:"0 2px 8px rgba(11,37,69,.06)" }}>
      {!isEmpty && <div style={{ position:"absolute", top:0, left:0, width:4, height:"100%", background:color, borderRadius:"14px 0 0 14px" }} />}

      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginLeft: isEmpty ? 0 : 8 }}>
        <span style={{ fontSize:11, fontWeight:700, color:COLORS.muted, letterSpacing:".5px" }}>{pat.leito}</span>
        <div style={{ display:"flex", gap:6, alignItems:"center" }}>
          {!isEmpty && <span style={{ fontSize:11, color, fontWeight:600, background:color+"18", padding:"2px 8px", borderRadius:10 }}>{label}</span>}
          <button onClick={e => { e.stopPropagation(); onEdit(pat.id); }}
            title={isEmpty ? "Admitir paciente" : "Editar paciente"}
            style={{ background:"none", border:`1px solid ${COLORS.border}`, borderRadius:6, padding:"2px 7px", cursor:"pointer", fontSize:12, color:COLORS.muted }}>
            {isEmpty ? "＋" : "✏️"}
          </button>
        </div>
      </div>

      {isEmpty ? (
        <div onClick={() => onEdit(pat.id)} style={{ marginTop:10, cursor:"pointer" }}>
          <div style={{ color:COLORS.muted, fontSize:13, marginBottom:4 }}>Leito disponível</div>
          <div style={{ fontSize:12, color:COLORS.teal, fontWeight:700 }}>+ Admitir paciente</div>
        </div>
      ) : (
        <div onClick={() => onSelect(pat.id)} style={{ cursor:"pointer" }}>
          <div style={{ marginTop:6, fontWeight:700, fontSize:15, color:COLORS.navy, marginLeft:8 }}>{pat.nome}</div>
          <div style={{ fontSize:12, color:COLORS.muted, marginTop:2, marginLeft:8 }}>{pat.diagnostico || "—"}</div>
          <div style={{ display:"flex", gap:12, marginTop:8, marginLeft:8 }}>
            <span style={{ fontSize:12, color:COLORS.muted }}>🕐 {pat.dias}d</span>
            {pat.idade && <span style={{ fontSize:12, color:COLORS.muted }}>👤 {pat.idade}a</span>}
          </div>
          {alerts.length > 0 && (
            <div style={{ marginTop:8, marginLeft:8, display:"flex", flexWrap:"wrap", gap:4 }}>
              {alerts.map((a,i) => <span key={i} style={{ fontSize:11, background:COLORS.danger+"15", color:COLORS.danger, padding:"2px 7px", borderRadius:8, fontWeight:600 }}>⚠ {a}</span>)}
            </div>
          )}
          <div style={{ marginTop:8, marginLeft:8 }}>
            {round
              ? <span style={{ fontSize:11, background:COLORS.success+"20", color:COLORS.success, padding:"2px 8px", borderRadius:8, fontWeight:600 }}>✓ Round feito</span>
              : <span style={{ fontSize:11, background:COLORS.warn+"18", color:COLORS.warn, padding:"2px 8px", borderRadius:8, fontWeight:600 }}>⏳ Pendente</span>}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Round Form ───────────────────────────────────────────────────────────────
function RoundForm({ pat, round, onChange, onBack }) {
  const r = round || { ...INITIAL_ROUND };
  const set = (k, v) => onChange({ ...r, [k]: v });
  const tog = (k, v) => { const a = r[k]||[]; onChange({ ...r, [k]: a.includes(v) ? a.filter(x=>x!==v) : [...a,v] }); };
  const setDev = (dev, f, v) => onChange({ ...r, dispositivos: { ...r.dispositivos, [dev]: { ...r.dispositivos[dev], [f]: v } } });
  const s2 = ["Sim","Não"], s3 = ["Sim","Não","Sem indicação"];

  return (
    <div style={{ maxWidth:820, margin:"0 auto", paddingBottom:80 }}>
      <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:24 }}>
        <button onClick={onBack} style={{ background:"none", border:`1.5px solid ${COLORS.border}`, borderRadius:8, padding:"6px 14px", cursor:"pointer", fontSize:13, color:COLORS.navy, fontWeight:600 }}>← Voltar</button>
        <div>
          <div style={{ fontSize:11, color:COLORS.muted, fontWeight:700, letterSpacing:".5px", textTransform:"uppercase" }}>{pat.leito} · Round Multidisciplinar</div>
          <div style={{ fontSize:20, fontWeight:800, color:COLORS.navy }}>{pat.nome}</div>
        </div>
      </div>

      {/* Identificação */}
      <div style={{ background:COLORS.card, borderRadius:14, padding:"18px 20px", marginBottom:14, border:`1px solid ${COLORS.border}` }}>
        <SecHdr title="Identificação" icon="🏥" />
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
          <div><div style={{ fontSize:11, color:COLORS.muted, fontWeight:700, marginBottom:6, textTransform:"uppercase", letterSpacing:".4px" }}>Diagnóstico</div><TInput value={r.diagnostico} onChange={v=>set("diagnostico",v)} placeholder="Ex: Sepse pulmonar" /></div>
          <div><div style={{ fontSize:11, color:COLORS.muted, fontWeight:700, marginBottom:6, textTransform:"uppercase", letterSpacing:".4px" }}>Resumo clínico</div><TInput value={r.resumoClinico} onChange={v=>set("resumoClinico",v)} placeholder="Situação atual..." /></div>
        </div>
        <div style={{ marginTop:14 }}><div style={{ fontSize:11, color:COLORS.muted, fontWeight:700, marginBottom:6, textTransform:"uppercase", letterSpacing:".4px" }}>Plano do dia</div><TArea value={r.planoDia} onChange={v=>set("planoDia",v)} placeholder="Plano de cuidados para hoje..." rows={2} /></div>
      </div>

      {/* Cuidados Gerais */}
      <div style={{ background:COLORS.card, borderRadius:14, padding:"18px 20px", marginBottom:14, border:`1px solid ${COLORS.border}` }}>
        <SecHdr title="Cuidados Gerais" icon="🛡️" />
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
          <Field label="Precaução de contato?">{s2.map(o=><Pill key={o} label={o} selected={r.contato===o} onClick={()=>set("contato",o)}/>)}</Field>
          <Field label="Visita flexibilizada?">{s2.map(o=><Pill key={o} label={o} selected={r.visitaFlex===o} onClick={()=>set("visitaFlex",o)}/>)}</Field>
        </div>
      </div>

      {/* Neurológico */}
      <div style={{ background:COLORS.card, borderRadius:14, padding:"18px 20px", marginBottom:14, border:`1px solid ${COLORS.border}` }}>
        <SecHdr title="Neurológico" icon="🧠" />
        <Field label="Meta de sedação (RASS)">{["-5 a -4","-2 a 0","Não se aplica"].map(o=><Pill key={o} label={o} selected={r.rass===o} onClick={()=>set("rass",o)}/>)}</Field>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:14 }}>
          <Field label="Controle de dor?">{s2.map(o=><Pill key={o} label={o} selected={r.dor===o} onClick={()=>set("dor",o)}/>)}</Field>
          <Field label="Delirium?">{s2.map(o=><Pill key={o} label={o} selected={r.delirium===o} onClick={()=>set("delirium",o)}/>)}</Field>
          <Field label="Contenção mecânica?">{s2.map(o=><Pill key={o} label={o} selected={r.contencao===o} onClick={()=>set("contencao",o)}/>)}</Field>
        </div>
      </div>

      {/* Cardiovascular */}
      <div style={{ background:COLORS.card, borderRadius:14, padding:"18px 20px", marginBottom:14, border:`1px solid ${COLORS.border}` }}>
        <SecHdr title="Cardiovascular / Hemodinâmica" icon="❤️" />
        <div style={{ display:"flex", gap:24, flexWrap:"wrap", alignItems:"flex-end" }}>
          <Field label="Suporte hemodinâmico / DVA?">{s2.map(o=><Pill key={o} label={o} selected={r.dva===o} onClick={()=>set("dva",o)}/>)}</Field>
          <Field label="Meta de PAM (mmHg)"><TInput value={r.pam} onChange={v=>set("pam",v)} placeholder="Ex: 65" w="90px"/></Field>
        </div>
      </div>

      {/* Respiratório */}
      <div style={{ background:COLORS.card, borderRadius:14, padding:"18px 20px", marginBottom:14, border:`1px solid ${COLORS.border}` }}>
        <SecHdr title="Respiratório / Reabilitação" icon="🫁" />
        <Field label="Suporte respiratório">{["Sem suporte","O2 suplementar","VNI","VM invasiva"].map(o=><Pill key={o} label={o} selected={r.suporteResp===o} onClick={()=>set("suporteResp",o)}/>)}</Field>
        {(r.suporteResp==="VM invasiva"||r.suporteResp==="VNI")&&<Field label="VM protetora?">{s2.map(o=><Pill key={o} label={o} selected={r.vmProtetora===o} onClick={()=>set("vmProtetora",o)}/>)}</Field>}
        <Field label="Plano de desmame">{["Redução de parâmetros","TRE hoje","Ex-TOT","Sem proposta de desmame"].map(o=><MultiPill key={o} label={o} checked={(r.planoDesmame||[]).includes(o)} onChange={()=>tog("planoDesmame",o)}/>)}</Field>
        <Field label="Preocupações respiratórias">{["Piora respiratória","Piora de secreção","Risco de broncoaspiração"].map(o=><MultiPill key={o} label={o} checked={(r.preocResp||[]).includes(o)} onChange={()=>tog("preocResp",o)} color={COLORS.danger}/>)}</Field>
        <div style={{ display:"flex", gap:24, flexWrap:"wrap" }}>
          <Field label="Escala IMS"><TInput value={r.ims} onChange={v=>set("ims",v)} placeholder="0-10" w="80px"/></Field>
          <Field label="Progredir nível funcional?">{s2.map(o=><Pill key={o} label={o} selected={r.progredirFuncional===o} onClick={()=>set("progredirFuncional",o)}/>)}</Field>
        </div>
      </div>

      {/* GI */}
      <div style={{ background:COLORS.card, borderRadius:14, padding:"18px 20px", marginBottom:14, border:`1px solid ${COLORS.border}` }}>
        <SecHdr title="Gastrointestinal / Nutrição" icon="🍽️" />
        <Field label="Via alimentar">{["VO","SNE/GTT","NPT","Zero"].map(o=><Pill key={o} label={o} selected={r.viaAlimentar===o} onClick={()=>set("viaAlimentar",o)}/>)}</Field>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:14 }}>
          <Field label="Aceitação">{["Normal","Baixa","Intolerância"].map(o=><Pill key={o} label={o} selected={r.aceitacao===o} onClick={()=>set("aceitacao",o)}/>)}</Field>
          <Field label="Meta calórica?">{s2.map(o=><Pill key={o} label={o} selected={r.metaCalorica===o} onClick={()=>set("metaCalorica",o)}/>)}</Field>
          <Field label="Avaliação fono?">{s2.map(o=><Pill key={o} label={o} selected={r.fono===o} onClick={()=>set("fono",o)}/>)}</Field>
          <Field label="Glicemia adequada?">{s2.map(o=><Pill key={o} label={o} selected={r.glicemia===o} onClick={()=>set("glicemia",o)}/>)}</Field>
          <Field label="Evacuação < 3 dias?">{s2.map(o=><Pill key={o} label={o} selected={r.evacuacao===o} onClick={()=>set("evacuacao",o)}/>)}</Field>
        </div>
      </div>

      {/* Renal + Infeccioso */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:14 }}>
        <div style={{ background:COLORS.card, borderRadius:14, padding:"18px 20px", border:`1px solid ${COLORS.border}` }}>
          <SecHdr title="Renal" icon="🫘" />
          <Field label="Função renal em piora?">{["Sim","Não","Em HD"].map(o=><Pill key={o} label={o} selected={r.funcaoRenal===o} onClick={()=>set("funcaoRenal",o)}/>)}</Field>
          <Field label="Meta de balanço hídrico">{["Positivo","Negativo","Neutro"].map(o=><Pill key={o} label={o} selected={r.metaBH===o} onClick={()=>set("metaBH",o)}/>)}</Field>
        </div>
        <div style={{ background:COLORS.card, borderRadius:14, padding:"18px 20px", border:`1px solid ${COLORS.border}` }}>
          <SecHdr title="Infeccioso" icon="🦠" />
          <Field label="Piora infecciosa?">{s2.map(o=><Pill key={o} label={o} selected={r.pioraInfec===o} onClick={()=>set("pioraInfec",o)}/>)}</Field>
          <Field label="Em uso de ATB?">{s2.map(o=><Pill key={o} label={o} selected={r.atb===o} onClick={()=>set("atb",o)}/>)}</Field>
        </div>
      </div>

      {/* Profilaxias */}
      <div style={{ background:COLORS.card, borderRadius:14, padding:"18px 20px", marginBottom:14, border:`1px solid ${COLORS.border}` }}>
        <SecHdr title="Profilaxias" icon="💉" />
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:14 }}>
          <Field label="Profilaxia TEV">{s3.map(o=><Pill key={o} label={o} selected={r.tev===o} onClick={()=>set("tev",o)} color={o==="Não"?COLORS.danger:COLORS.teal}/>)}</Field>
          <Field label="Profilaxia LAMG">{s3.map(o=><Pill key={o} label={o} selected={r.lamg===o} onClick={()=>set("lamg",o)}/>)}</Field>
          <Field label="Úlcera de córnea">{s3.map(o=><Pill key={o} label={o} selected={r.cornea===o} onClick={()=>set("cornea",o)}/>)}</Field>
          <Field label="Higiene oral">{s3.map(o=><Pill key={o} label={o} selected={r.higieneOral===o} onClick={()=>set("higieneOral",o)}/>)}</Field>
          <Field label="Decúbito elevado">{s3.map(o=><Pill key={o} label={o} selected={r.decubito===o} onClick={()=>set("decubito",o)}/>)}</Field>
          <Field label="Bundles OK?">{s2.map(o=><Pill key={o} label={o} selected={r.bundles===o} onClick={()=>set("bundles",o)} color={o==="Não"?COLORS.danger:COLORS.teal}/>)}</Field>
        </div>
        {r.bundles==="Não"&&<div style={{ marginTop:8 }}><Field label="Bundle pendente"><TInput value={r.bundlesPendente} onChange={v=>set("bundlesPendente",v)} placeholder="Qual bundle?"/></Field></div>}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:14, marginTop:4 }}>
          <Field label="Mudança de decúbito?">{s2.map(o=><Pill key={o} label={o} selected={r.mudancaDecubito===o} onClick={()=>set("mudancaDecubito",o)}/>)}</Field>
          <Field label="Lesão por pressão?">{s2.map(o=><Pill key={o} label={o} selected={r.lesaoPressao===o} onClick={()=>set("lesaoPressao",o)}/>)}</Field>
          <Field label="Avaliação especializada?">{["Sim","Não","Comissão curativo","Cirurgia plástica","N/A"].map(o=><Pill key={o} label={o} selected={r.avalEspecializada===o} onClick={()=>set("avalEspecializada",o)}/>)}</Field>
        </div>
      </div>

      {/* Dispositivos */}
      <div style={{ background:COLORS.card, borderRadius:14, padding:"18px 20px", marginBottom:14, border:`1px solid ${COLORS.border}` }}>
        <SecHdr title="Dispositivos Invasivos" icon="🔌" />
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
          {[["cvc","Cateter Venoso Central"],["hd","Cateter de HD"],["svd","SVD"],["arterial","Cateter Arterial"],["sondaEnteral","Sonda Enteral"],["drenos","Drenos"]].map(([key,lbl])=>(
            <div key={key} style={{ background:COLORS.lightBg, borderRadius:10, padding:"12px 14px", border:`1px solid ${COLORS.border}` }}>
              <div style={{ fontSize:12, fontWeight:700, color:COLORS.navy, marginBottom:8 }}>{lbl}</div>
              <div style={{ display:"flex", gap:10, alignItems:"center", flexWrap:"wrap" }}>
                <input type="date" value={r.dispositivos[key].data} onChange={e=>setDev(key,"data",e.target.value)}
                  style={{ border:`1px solid ${COLORS.border}`, borderRadius:6, padding:"4px 8px", fontSize:12, color:COLORS.navy, background:"#fff" }}/>
                <label style={{ display:"flex", alignItems:"center", gap:5, fontSize:12, color:COLORS.teal, fontWeight:600, cursor:"pointer" }}>
                  <input type="checkbox" checked={r.dispositivos[key].desinvadir} onChange={e=>setDev(key,"desinvadir",e.target.checked)}/>
                  Desinvadir?
                </label>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Objetivos */}
      <div style={{ background:COLORS.card, borderRadius:14, padding:"18px 20px", marginBottom:14, border:`1px solid ${COLORS.border}` }}>
        <SecHdr title="Objetivos de Cuidado e Planejamento" icon="📋" />
        <Field label="Diretivas de cuidado">{["Não RCP","Não HD","Não IOT","Não DVA","Não coletar exames","Sem diretivas"].map(o=><MultiPill key={o} label={o} checked={(r.diretivas||[]).includes(o)} onChange={()=>tog("diretivas",o)} color={o==="Sem diretivas"?COLORS.success:COLORS.danger}/>)}</Field>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
          <div>
            <Field label="Pendência de exame/procedimento?">{s2.map(o=><Pill key={o} label={o} selected={r.pendenciaExame===o} onClick={()=>set("pendenciaExame",o)}/>)}</Field>
            {r.pendenciaExame==="Sim"&&<TArea value={r.descPendencia} onChange={v=>set("descPendencia",v)} placeholder="Descreva a pendência..."/>}
          </div>
          <Field label="Previsão de alta">{["Hoje","24–48h","> 48h"].map(o=><Pill key={o} label={o} selected={r.previsaoAlta===o} onClick={()=>set("previsaoAlta",o)}/>)}</Field>
        </div>
      </div>

      <div style={{ display:"flex", justifyContent:"flex-end", gap:12 }}>
        <button onClick={onBack} style={{ padding:"10px 24px", borderRadius:10, border:`1.5px solid ${COLORS.border}`, background:"#fff", color:COLORS.navy, fontSize:14, fontWeight:600, cursor:"pointer" }}>Cancelar</button>
        <button onClick={onBack} style={{ padding:"10px 28px", borderRadius:10, border:"none", background:COLORS.teal, color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer" }}>✓ Salvar Round</button>
      </div>
    </div>
  );
}

// ── Report Modal ──────────────────────────────────────────────────────────────
function ReportModal({ patients, rounds, onClose }) {
  const rows = patients.filter(p => p.nome);
  const csv = [
    ["Leito","Nome","Diagnóstico","Dias","VM","DVA","ATB","HD","Alta","Pendência","Alertas"].join(";"),
    ...rows.map(p => {
      const r = rounds[p.id];
      return [p.leito, p.nome, r?.diagnostico||p.diagnostico, p.dias, r?.suporteResp==="VM invasiva"?"Sim":"Não", r?.dva||"-", r?.atb||"-", r?.funcaoRenal==="Em HD"?"Sim":"Não", r?.previsaoAlta||"-", r?.descPendencia||"-", computeAlerts(r,p).join(" | ")].join(";");
    })
  ].join("\n");

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(11,37,69,.55)", zIndex:100, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div style={{ background:"#fff", borderRadius:16, padding:"28px 32px", maxWidth:960, width:"100%", maxHeight:"85vh", overflow:"auto" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:20 }}>
          <div style={{ fontWeight:800, fontSize:20, color:COLORS.navy }}>📊 Resumo Geral da UTI</div>
          <button onClick={onClose} style={{ background:"none", border:"none", fontSize:22, cursor:"pointer", color:COLORS.muted }}>×</button>
        </div>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
            <thead><tr style={{ background:COLORS.navy, color:"#fff" }}>
              {["Leito","Paciente","Diagnóstico","Dias","VM","DVA","ATB","Alta","Pendências","Alertas"].map(h=>(
                <th key={h} style={{ padding:"10px 12px", textAlign:"left", fontWeight:700, whiteSpace:"nowrap" }}>{h}</th>
              ))}
            </tr></thead>
            <tbody>{rows.map((p,i)=>{
              const r=rounds[p.id]; const alerts=computeAlerts(r,p);
              return <tr key={p.id} style={{ background:i%2===0?"#fff":COLORS.lightBg }}>
                <td style={{ padding:"9px 12px", fontWeight:700, color:COLORS.navy }}>{p.leito}</td>
                <td style={{ padding:"9px 12px", whiteSpace:"nowrap" }}>{p.nome}</td>
                <td style={{ padding:"9px 12px", color:COLORS.muted }}>{r?.diagnostico||p.diagnostico}</td>
                <td style={{ padding:"9px 12px", textAlign:"center" }}>{p.dias}d</td>
                <td style={{ padding:"9px 12px", textAlign:"center" }}><span style={{ color:r?.suporteResp==="VM invasiva"?COLORS.danger:COLORS.success, fontWeight:700 }}>{r?.suporteResp==="VM invasiva"?"Sim":"Não"}</span></td>
                <td style={{ padding:"9px 12px", textAlign:"center" }}><span style={{ color:r?.dva==="Sim"?COLORS.danger:COLORS.success, fontWeight:700 }}>{r?.dva||"-"}</span></td>
                <td style={{ padding:"9px 12px", textAlign:"center" }}>{r?.atb||"-"}</td>
                <td style={{ padding:"9px 12px" }}>{r?.previsaoAlta||"-"}</td>
                <td style={{ padding:"9px 12px", color:COLORS.warn, maxWidth:160 }}>{r?.descPendencia||"-"}</td>
                <td style={{ padding:"9px 12px" }}>{alerts.map((a,j)=><div key={j} style={{ color:COLORS.danger, fontSize:11, fontWeight:600 }}>⚠ {a}</div>)}</td>
              </tr>;
            })}</tbody>
          </table>
        </div>
        <div style={{ display:"flex", gap:12, marginTop:20, justifyContent:"flex-end" }}>
          <button onClick={()=>{ const blob=new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"}); const u=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=u; a.download="round-uti.csv"; a.click(); }} style={{ padding:"9px 22px", borderRadius:10, border:"none", background:COLORS.teal, color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer" }}>⬇ Exportar CSV</button>
          <button onClick={onClose} style={{ padding:"9px 22px", borderRadius:10, border:`1.5px solid ${COLORS.border}`, background:"#fff", color:COLORS.navy, fontSize:14, fontWeight:600, cursor:"pointer" }}>Fechar</button>
        </div>
      </div>
    </div>
  );
}

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [patients, setPatients] = useState(INITIAL_PATIENTS);
  const [rounds,   setRounds]   = useState({});
  const [selected, setSelected] = useState(null);
  const [editingId,setEditingId]= useState(null);
  const [showReport,setShowReport]=useState(false);
  const [search,   setSearch]   = useState("");
  const [filter,   setFilter]   = useState("todos");

  const selPat  = patients.find(p=>p.id===selected);
  const editPat = patients.find(p=>p.id===editingId);
  const date = new Date().toLocaleDateString("pt-BR",{weekday:"long",day:"2-digit",month:"long",year:"numeric"});

  const handleChange = useCallback(r=>setRounds(prev=>({...prev,[selected]:r})),[selected]);

  const savePatient  = (id,data) => setPatients(prev=>prev.map(p=>p.id===id?{...p,...data}:p));
  const clearPatient = (id)      => { setPatients(prev=>prev.map(p=>p.id===id?{...p,nome:"",idade:null,dias:0,diagnostico:"",gravidade:"livre"}:p)); setRounds(prev=>{const n={...prev};delete n[id];return n;}); };

  const withPats  = patients.filter(p=>p.nome);
  const roundsDone= withPats.filter(p=>rounds[p.id]).length;
  const total     = withPats.length;

  const filtered = patients.filter(p=>{
    if(search && !p.nome?.toLowerCase().includes(search.toLowerCase()) && !p.leito.toLowerCase().includes(search.toLowerCase())) return false;
    if(filter==="pendente" && (rounds[p.id]||!p.nome)) return false;
    if(filter==="alta" && p.gravidade!=="alta") return false;
    return true;
  });

  // ── Tela de round ──
  if(selected && selPat) return (
    <div style={{ background:COLORS.bg, minHeight:"100vh", fontFamily:"'DM Sans','Segoe UI',sans-serif" }}>
      <div style={{ background:COLORS.navy, padding:"12px 24px", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
        <div style={{ color:"#fff", fontWeight:800, fontSize:16 }}>🏥 UTI Clínica — IMIP</div>
        <div style={{ color:"#8BBBD9", fontSize:13 }}>{date}</div>
      </div>
      <div style={{ padding:"24px" }}>
        <RoundForm pat={selPat} round={rounds[selected]||{...INITIAL_ROUND,diagnostico:selPat.diagnostico}} onChange={handleChange} onBack={()=>setSelected(null)}/>
      </div>
    </div>
  );

  // ── Dashboard ──
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
        <div style={{ display:"flex", alignItems:"center", gap:14 }}>
          <div style={{ color:"#8BBBD9", fontSize:13 }}>{date}</div>
          <button onClick={()=>setShowReport(true)} style={{ padding:"8px 18px", borderRadius:10, border:"none", background:COLORS.accent, color:"#fff", fontSize:13, fontWeight:700, cursor:"pointer" }}>📊 Relatório</button>
        </div>
      </div>

      <div style={{ padding:"24px 28px" }}>
        {/* Stats */}
        <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:20 }}>
          {[["Leitos ocupados",total,COLORS.navy,"🛏"],["Rounds feitos",roundsDone,COLORS.success,"✅"],["Pendentes",total-roundsDone,COLORS.warn,"⏳"],["Alertas",patients.filter(p=>computeAlerts(rounds[p.id],p).length>0).length,COLORS.danger,"⚠️"]].map(([lbl,val,color,icon])=>(
            <div key={lbl} style={{ background:"#fff", borderRadius:12, padding:"16px 20px", border:`1px solid ${COLORS.border}` }}>
              <div style={{ fontSize:22, marginBottom:4 }}>{icon}</div>
              <div style={{ fontSize:26, fontWeight:800, color }}>{val}</div>
              <div style={{ fontSize:12, color:COLORS.muted, fontWeight:600 }}>{lbl}</div>
            </div>
          ))}
        </div>

        {/* Progress */}
        <div style={{ background:"#fff", borderRadius:12, padding:"14px 20px", border:`1px solid ${COLORS.border}`, marginBottom:20 }}>
          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}>
            <span style={{ fontSize:13, fontWeight:700, color:COLORS.navy }}>Progresso do round</span>
            <span style={{ fontSize:13, color:COLORS.muted }}>{roundsDone}/{total} leitos</span>
          </div>
          <div style={{ background:COLORS.border, borderRadius:8, height:8 }}>
            <div style={{ width:`${total?(roundsDone/total)*100:0}%`, height:"100%", background:COLORS.teal, borderRadius:8, transition:"width .4s" }}/>
          </div>
        </div>

        {/* Filtros */}
        <div style={{ display:"flex", gap:12, marginBottom:20, flexWrap:"wrap" }}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar paciente ou leito..."
            style={{ flex:1, minWidth:200, border:`1.5px solid ${COLORS.border}`, borderRadius:10, padding:"9px 16px", fontSize:14, color:COLORS.navy, outline:"none", background:"#fff" }}/>
          {[["todos","Todos"],["pendente","Pendentes"],["alta","Alta gravidade"]].map(([v,l])=>(
            <button key={v} onClick={()=>setFilter(v)} style={{ padding:"9px 18px", borderRadius:10, border:`1.5px solid ${filter===v?COLORS.teal:COLORS.border}`, background:filter===v?COLORS.teal:"#fff", color:filter===v?"#fff":COLORS.navy, fontSize:13, fontWeight:600, cursor:"pointer" }}>{l}</button>
          ))}
        </div>

        {/* Grid 10 leitos */}
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))", gap:14 }}>
          {filtered.map(p=><PatientCard key={p.id} pat={p} round={rounds[p.id]} onSelect={setSelected} onEdit={setEditingId}/>)}
        </div>
      </div>

      {editingId && editPat && <EditPatientModal pat={editPat} onSave={savePatient} onClear={clearPatient} onClose={()=>setEditingId(null)}/>}
      {showReport && <ReportModal patients={patients} rounds={rounds} onClose={()=>setShowReport(false)}/>}
    </div>
  );
}
