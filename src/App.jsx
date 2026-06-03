import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabase.js";
import Auth from "./Auth.jsx";

// ─── Config (salvo no localStorage) ───────────────────────────────────────
const DEFAULT_CONFIG = {
  debito:        { label:"Débito",        prazo:1,  taxa:1.69, cor:"#60A5FA", bg:"#0F2942" },
  credito_vista: { label:"Créd. à Vista", prazo:30, taxa:2.69, cor:"#34D399", bg:"#052E1A" },
  alelo:         { label:"Alelo",         prazo:1,  taxa:1.80, cor:"#FCD34D", bg:"#2A1F06" },
  ticket:        { label:"Ticket",        prazo:1,  taxa:1.80, cor:"#F97316", bg:"#2A1006" },
  vr:            { label:"VR",            prazo:1,  taxa:1.80, cor:"#A78BFA", bg:"#1E1244" },
};
const loadConfig = () => { try { return JSON.parse(localStorage.getItem("rede_config")) || DEFAULT_CONFIG; } catch { return DEFAULT_CONFIG; } };

// ─── Helpers ───────────────────────────────────────────────────────────────
const addDays = (s, n) => { const d = new Date(s + "T12:00:00"); d.setDate(d.getDate() + n); return d.toISOString().split("T")[0]; };
const R  = (v) => new Intl.NumberFormat("pt-BR", { style:"currency", currency:"BRL" }).format(v ?? 0);
const D  = (s) => s ? s.split("-").reverse().join("/") : "—";
const Dt = (s) => s ? new Date(s).toLocaleDateString("pt-BR") : "—";

function calcForSave(rawTxs, config) {
  return rawTxs.map(tx => {
    const key = config[tx.type] ? tx.type : "debito";
    const { prazo, taxa } = config[key];
    return {
      date: tx.date, description: tx.description || null,
      type: tx.type, gross_amount: Number(tx.gross_amount),
      net_amount: Number(tx.gross_amount) * (1 - taxa / 100),
      taxa_pct: taxa, prazo_dias: prazo,
      settlement_date: addDays(tx.date, prazo),
      card_brand: tx.card_brand || null, nsu: tx.nsu || null,
    };
  });
}

// ─── File readers ──────────────────────────────────────────────────────────
const toBase64 = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result.split(",")[1]); r.onerror = rej; r.readAsDataURL(f); });
const toBuffer = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsArrayBuffer(f); });
const toText   = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsText(f, "UTF-8"); });

const PARSE_PROMPT = `Você é um especialista em extratos de maquininha de cartão e vouchers brasileiros (Rede/Redecard).
Analise o extrato e extraia TODAS as transações.
Responda APENAS com JSON válido — sem markdown, sem texto extra, sem explicações.
Formato:
{"transactions":[{"date":"YYYY-MM-DD","description":"string","type":"debito|credito_vista|alelo|ticket|vr","gross_amount":número,"card_brand":"string ou null","nsu":"string ou null"}]}
Regras: debito=cartão débito, credito_vista=crédito à vista, alelo/ticket/vr=vouchers respectivos. Estornos=gross_amount negativo. Ignore cabeçalhos e totais.`;

// ─── App ───────────────────────────────────────────────────────────────────
export default function App() {
  const [session, setSession]     = useState(undefined); // undefined = carregando
  const [config, setConfig]       = useState(loadConfig);
  const [settlements, setSettlements] = useState([]);
  const [imports, setImports]     = useState([]);
  const [recon, setRecon]         = useState({});
  const [loading, setLoading]     = useState(true);
  const [importing, setImporting] = useState(false);
  const [importMsg, setImportMsg] = useState("");
  const [error, setError]         = useState(null);
  const [tab, setTab]             = useState("painel");
  const [showCfg, setShowCfg]     = useState(false);
  const [dragging, setDragging]   = useState(false);
  const [showAll, setShowAll]     = useState(false);
  const fileRef     = useRef(null);
  const reconTimers = useRef({});
  const today = new Date().toISOString().split("T")[0];

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, session) => setSession(session));
    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => { if (session) loadAllData(); }, [session]);

  async function loadAllData() {
    setLoading(true);
    try {
      const [{ data: txs, error: e1 }, { data: recons }, { data: imps }] = await Promise.all([
        supabase.from("transactions").select("*, imports(filename, imported_at)").order("settlement_date"),
        supabase.from("reconciliations").select("*"),
        supabase.from("imports").select("*").order("imported_at", { ascending: false }),
      ]);
      if (e1) throw e1;
      setSettlements(txs || []);
      setRecon((recons || []).reduce((acc, r) => ({ ...acc, [r.settlement_date]: r.actual_amount != null ? String(r.actual_amount) : "" }), {}));
      setImports(imps || []);
    } catch (e) {
      setError("Erro Supabase: " + e.message + " — verifique suas variáveis de ambiente");
    } finally { setLoading(false); }
  }

  const parseFile = useCallback(async (file) => {
    setImporting(true); setError(null);
    const ext = file.name.split(".").pop().toLowerCase();
    try {
      let messages;
      setImportMsg("Lendo arquivo…");
      if (ext === "pdf") {
        setImportMsg("Enviando PDF para IA…");
        const b64 = await toBase64(file);
        messages = [{ role:"user", content:[{ type:"document", source:{ type:"base64", media_type:"application/pdf", data:b64 }},{ type:"text", text:PARSE_PROMPT }]}];
      } else if (["xlsx","xls"].includes(ext)) {
        setImportMsg("Convertendo planilha…");
        const ab = await toBuffer(file);
        const wb = XLSX.read(ab, { type:"array" });
        const csv = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
        messages = [{ role:"user", content:`${PARSE_PROMPT}\n\nConteúdo:\n${csv}` }];
      } else {
        const txt = await toText(file);
        messages = [{ role:"user", content:`${PARSE_PROMPT}\n\nConteúdo:\n${txt}` }];
      }
      setImportMsg("IA classificando transações…");
      const res = await fetch("/api/parse", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ model:"claude-sonnet-4-20250514", max_tokens:4000, messages }) });
      const data = await res.json();
      if (data.error) throw new Error(data.error.message);
      const raw = data.content.map(b => b.text || "").join("").replace(/```json|```/gi,"").trim();
      const parsed = JSON.parse(raw);
      const rawTxs = (parsed.transactions || []).filter(t => t.date && t.gross_amount !== undefined);
      if (rawTxs.length === 0) throw new Error("Nenhuma transação encontrada no arquivo");
      const calculated = calcForSave(rawTxs, config);
      const grossTotal = calculated.reduce((a,t) => a + t.gross_amount, 0);
      const netTotal   = calculated.reduce((a,t) => a + t.net_amount,   0);
      setImportMsg("Salvando no banco de dados…");
      const { data: imp, error: ie } = await supabase.from("imports")
        .insert({ filename: file.name, transaction_count: calculated.length, gross_total: grossTotal, net_total: netTotal })
        .select().single();
      if (ie) throw ie;
      setImportMsg("Arquivando documento…");
      const month = today.slice(0, 7);
      const storagePath = `${month}/${imp.id}/${file.name}`;
      const { error: ue } = await supabase.storage.from("extratos").upload(storagePath, file, { upsert: true });
      if (!ue) await supabase.from("imports").update({ storage_path: storagePath }).eq("id", imp.id);
      const { error: te } = await supabase.from("transactions").insert(calculated.map(t => ({ ...t, import_id: imp.id })));
      if (te) throw te;
      setImportMsg("Concluído!");
      await loadAllData();
      setTab("previsao");
    } catch (e) { setError("Erro ao importar: " + e.message); }
    finally { setImporting(false); setImportMsg(""); }
  }, [config, today]);

  const onDrop = (e) => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if (f) parseFile(f); };
  const onPick = (e) => { const f = e.target.files[0]; if (f) parseFile(f); e.target.value = ""; };

  const updateRecon = (day, value) => {
    setRecon(p => ({ ...p, [day]: value }));
    clearTimeout(reconTimers.current[day]);
    reconTimers.current[day] = setTimeout(async () => {
      const val = parseFloat(value);
      await supabase.from("reconciliations").upsert({ settlement_date: day, actual_amount: isNaN(val) ? null : val, updated_at: new Date().toISOString() }, { onConflict: "settlement_date" });
    }, 800);
  };

  const updateCfg = (k, field, v) => {
    const nc = { ...config, [k]: { ...config[k], [field]: parseFloat(v) || 0 } };
    setConfig(nc); localStorage.setItem("rede_config", JSON.stringify(nc));
  };

  async function deleteImport(id, storagePath) {
    if (!window.confirm("Remover esta importação e todas as suas transações?")) return;
    if (storagePath) await supabase.storage.from("extratos").remove([storagePath]);
    await supabase.from("imports").delete().eq("id", id);
    await loadAllData();
  }

  async function downloadFile(storagePath, filename) {
    const { data } = await supabase.storage.from("extratos").download(storagePath);
    if (!data) return;
    const url = URL.createObjectURL(data);
    const a = document.createElement("a"); a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
  }

  // ─── Derived data ──────────────────────────────────────────────────────
  const grouped = settlements.reduce((acc, s) => { (acc[s.settlement_date] = acc[s.settlement_date] || []).push(s); return acc; }, {});
  const allDays     = Object.keys(grouped).sort();
  const futureDays  = allDays.filter(d => d >= today);
  const pastDays    = allDays.filter(d => d <  today).reverse();
  const previsaoDays = showAll ? allDays : futureDays;
  const conferenciaDays = [...pastDays.slice(0, 90), ...futureDays.slice(0, 14)].sort();

  const totGross = settlements.reduce((a,s) => a + s.gross_amount, 0);
  const totNet   = settlements.reduce((a,s) => a + s.net_amount,   0);
  const totTax   = totGross - totNet;

  const byType = settlements.reduce((acc, s) => {
    const k = config[s.type] ? s.type : "debito";
    if (!acc[k]) acc[k] = { gross:0, net:0, count:0 };
    acc[k].gross += s.gross_amount; acc[k].net += s.net_amount; acc[k].count++;
    return acc;
  }, {});

  const pendingRecon = pastDays.filter(d => grouped[d]?.length > 0 && (!recon[d] || recon[d] === "")).length;

  const TypeChip = ({ type, amount }) => {
    const neg = amount < 0;
    const cfg = config[type] || config.debito;
    if (neg) return <span style={{background:"#2A0A0A",color:"#F87171",borderRadius:100,padding:"2px 8px",fontSize:10,fontWeight:700,fontFamily:"var(--mono)"}}>EST</span>;
    return <span style={{background:cfg.bg,color:cfg.cor,borderRadius:100,padding:"2px 8px",fontSize:10,fontWeight:700,fontFamily:"var(--mono)",whiteSpace:"nowrap"}}>{cfg.label}</span>;
  };

  const TABS = [
    { id:"painel",      label:"Painel" },
    { id:"importar",    label:"Importar",    badge: imports.length || null },
    { id:"previsao",    label:"Previsão",    badge: futureDays.length || null },
    { id:"conferencia", label:"Conferência", badge: pendingRecon || null, warn: pendingRecon > 0 },
    { id:"transacoes",  label:"Transações",  badge: settlements.length || null },
  ];

  const css = `
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:opsz,wght@9..40,400;9..40,500;9..40,700&family=JetBrains+Mono:wght@400;600&display=swap');
    :root{--mono:'JetBrains Mono',monospace;}
    *{box-sizing:border-box;margin:0;padding:0;}
    ::-webkit-scrollbar{width:5px;height:5px;}
    ::-webkit-scrollbar-track{background:#0D1320;}
    ::-webkit-scrollbar-thumb{background:#1E2D45;border-radius:4px;}
    .tab{padding:10px 16px;border:none;background:none;cursor:pointer;font-family:inherit;font-size:13px;font-weight:500;color:#64748B;border-bottom:2px solid transparent;transition:.15s;white-space:nowrap;}
    .tab:hover{color:#94A3B8;} .tab.on{color:#10B981;border-bottom-color:#10B981;}
    .bdg{display:inline-flex;align-items:center;justify-content:center;background:#1E2D45;color:#64748B;border-radius:10px;font-size:10px;padding:1px 6px;margin-left:5px;font-family:var(--mono);}
    .tab.on .bdg{background:#042F21;color:#10B981;} .bdg.w{background:#2A1F06;color:#F59E0B;}
    .tr:hover{background:#0D1C2E;}
    .inp{background:#050A12;border:1px solid #1E2D45;border-radius:6px;padding:7px 11px;color:#E2E8F0;font-family:var(--mono);font-size:12px;outline:none;transition:border-color .15s;}
    .inp:focus{border-color:#10B981;}
    .btn-g{background:#10B981;color:#001A0E;border:none;border-radius:8px;padding:9px 18px;font-family:inherit;font-weight:700;font-size:13px;cursor:pointer;transition:.15s;}
    .btn-g:hover{background:#059669;}
    .btn-o{background:transparent;color:#64748B;border:1px solid #1E2D45;border-radius:8px;padding:8px 14px;font-family:inherit;font-size:12px;cursor:pointer;transition:.15s;}
    .btn-o:hover{border-color:#334155;color:#CBD5E1;}
    .btn-r{background:transparent;color:#EF4444;border:1px solid #3B1111;border-radius:6px;padding:5px 10px;font-family:inherit;font-size:11px;cursor:pointer;transition:.15s;}
    .btn-r:hover{background:#3B1111;}
    .zone{border:2px dashed #1E2D45;border-radius:14px;padding:40px 36px;text-align:center;cursor:pointer;transition:.25s;}
    .zone:hover,.zone.drag{border-color:#10B981;background:#021911;}
    .card{background:#0C1520;border:1px solid #1A2840;border-radius:12px;}
    .stat{background:#0C1520;border:1px solid #1A2840;border-radius:10px;padding:16px 20px;}
    .mono{font-family:var(--mono);}
    @keyframes fd{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
    .fd{animation:fd .25s ease;}
    @keyframes sp{to{transform:rotate(360deg)}} .spin{animation:sp 1s linear infinite;display:inline-block;}
    .recon-in{background:#050A12;border:1px solid #1E2D45;border-radius:6px;padding:6px 10px;color:#E2E8F0;font-family:var(--mono);font-size:12px;width:140px;text-align:right;outline:none;transition:border-color .15s;}
    .recon-in:focus{border-color:#10B981;}
    .imp-row{display:flex;align-items:center;gap:12px;padding:11px 16px;border-bottom:1px solid #111C2C;}
    .imp-row:last-child{border-bottom:none;} .imp-row:hover{background:#0D1C2E;}
    .day-hd{display:flex;align-items:center;justify-content:space-between;padding:11px 16px;border-bottom:1px solid #1A2840;}
    .tx-row{display:flex;align-items:center;gap:12px;padding:7px 16px;cursor:default;}
    .tx-row:not(:last-child){border-bottom:1px solid #111C2C;}
  `;

  if (session === undefined) return null;
  if (!session) return <Auth />;

  return (
    <div style={{fontFamily:"'DM Sans','Helvetica Neue',sans-serif",background:"#070C16",minHeight:"100vh",color:"#E2E8F0"}}>
      <style>{css}</style>

      {/* ── HEADER ── */}
      <div style={{background:"#0C1520",borderBottom:"1px solid #1A2840",height:54,display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 24px",position:"sticky",top:0,zIndex:100}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div style={{width:32,height:32,borderRadius:8,background:"linear-gradient(135deg,#10B981,#059669)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15}}>💳</div>
          <div>
            <div style={{fontSize:14,fontWeight:700,color:"#F1F5F9",lineHeight:1.1}}>Conferência Rede</div>
            <div style={{fontSize:10,color:"#64748B",marginTop:1}}>{imports.length} importações · {settlements.length} transações · {R(totNet)} acumulado</div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <button className="btn-o" onClick={()=>setShowCfg(p=>!p)}>⚙ Taxas & Prazos</button>
          <button className="btn-g" onClick={()=>{setTab("importar");setTimeout(()=>fileRef.current?.click(),100)}}>+ Importar Extrato</button>
          <button className="btn-o" style={{fontSize:12}} onClick={()=>supabase.auth.signOut()} title="Sair">Sair</button>
          <input ref={fileRef} type="file" accept=".pdf,.xlsx,.xls,.csv,.txt" style={{display:"none"}} onChange={onPick}/>
        </div>
      </div>

      {/* ── CONFIG ── */}
      {showCfg && (
        <div className="fd" style={{background:"#0C1520",borderBottom:"1px solid #1A2840",padding:"14px 24px"}}>
          <div style={{display:"flex",gap:12,alignItems:"flex-end",flexWrap:"wrap"}}>
            <span style={{fontSize:11,color:"#64748B",fontWeight:700,letterSpacing:"0.06em",alignSelf:"center"}}>TAXAS E PRAZOS</span>
            {Object.entries(config).map(([k,c])=>(
              <div key={k} style={{background:"#0D1A2A",border:`1px solid ${c.bg}`,borderRadius:10,padding:"10px 14px"}}>
                <div style={{fontSize:11,color:c.cor,fontWeight:700,marginBottom:8}}>{c.label}</div>
                <div style={{display:"flex",gap:10}}>
                  <div><div style={{fontSize:10,color:"#475569",marginBottom:3}}>Prazo (dias)</div><input className="inp" style={{width:74}} type="number" value={c.prazo} onChange={e=>updateCfg(k,"prazo",e.target.value)}/></div>
                  <div><div style={{fontSize:10,color:"#475569",marginBottom:3}}>Taxa (%)</div><input className="inp" style={{width:74}} type="number" step="0.01" value={c.taxa} onChange={e=>updateCfg(k,"taxa",e.target.value)}/></div>
                </div>
              </div>
            ))}
          </div>
          <div style={{marginTop:10,fontSize:11,color:"#475569"}}>⚠ Taxas se aplicam a novas importações. Importações anteriores mantêm os valores originais salvos.</div>
        </div>
      )}

      {/* ── TABS ── */}
      <div style={{background:"#0C1520",borderBottom:"1px solid #1A2840",padding:"0 24px",display:"flex",gap:2}}>
        {TABS.map(t=>(
          <button key={t.id} className={`tab ${tab===t.id?"on":""}`} onClick={()=>setTab(t.id)}>
            {t.label}
            {t.badge>0&&<span className={`bdg${t.warn?" w":""}`}>{t.badge}</span>}
          </button>
        ))}
      </div>

      {/* ── CONTENT ── */}
      <div style={{padding:24,maxWidth:1200,margin:"0 auto"}}>

        {loading && (
          <div style={{textAlign:"center",padding:"80px 0"}}>
            <div className="spin" style={{fontSize:32,marginBottom:14}}>⚙</div>
            <div style={{color:"#10B981",fontWeight:600,marginBottom:6}}>Carregando histórico do Supabase…</div>
            <div style={{color:"#475569",fontSize:13}}>Buscando todas as transações e conferências salvas</div>
          </div>
        )}

        {!loading&&error&&(
          <div style={{background:"#1C0808",border:"1px solid #7F1D1D",borderRadius:10,padding:14,color:"#EF4444",marginBottom:16,fontSize:13}}>⚠ {error}</div>
        )}

        {!loading&&(
          <>
            {/* ══════════ PAINEL ══════════ */}
            {tab==="painel"&&(
              <div className="fd">
                <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
                  {[
                    {l:"Líquido Acumulado", v:R(totNet),      c:"#10B981"},
                    {l:"Bruto Total",       v:R(totGross),    c:"#F1F5F9"},
                    {l:"Taxas Pagas",       v:R(totTax),      c:"#EF4444"},
                    {l:"Extratos",          v:imports.length, c:"#F1F5F9"},
                  ].map(s=>(
                    <div key={s.l} className="stat">
                      <div style={{fontSize:10,color:"#64748B",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:7}}>{s.l}</div>
                      <div className="mono" style={{fontSize:20,fontWeight:700,color:s.c}}>{s.v}</div>
                    </div>
                  ))}
                </div>

                <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:22}}>
                  {Object.entries(byType).map(([k,v])=>{
                    const c=config[k]||config.debito;
                    return (
                      <div key={k} style={{background:"#0C1520",border:`1px solid ${c.bg}`,borderRadius:10,padding:"10px 16px",minWidth:130}}>
                        <div style={{fontSize:10,color:c.cor,fontWeight:700,marginBottom:5}}>{c.label}</div>
                        <div className="mono" style={{fontSize:15,fontWeight:700,color:"#F1F5F9"}}>{R(v.net)}</div>
                        <div style={{fontSize:11,color:"#475569",marginTop:2}}>{v.count} transações</div>
                      </div>
                    );
                  })}
                </div>

                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
                  <div className="card">
                    <div style={{padding:"12px 16px",borderBottom:"1px solid #1A2840"}}>
                      <span style={{fontSize:12,fontWeight:700,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.06em"}}>Próximos Recebimentos</span>
                    </div>
                    {futureDays.length===0
                      ? <div style={{padding:24,color:"#475569",fontSize:13,textAlign:"center"}}>Nenhum futuro cadastrado</div>
                      : futureDays.slice(0,7).map(day=>{
                          const net=(grouped[day]||[]).reduce((a,s)=>a+s.net_amount,0);
                          const n=Math.round((new Date(day)-new Date(today))/86400000);
                          return (
                            <div key={day} className="tr" style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 16px",borderBottom:"1px solid #0D1520"}}>
                              <div>
                                <div className="mono" style={{fontSize:13,color:"#F1F5F9",fontWeight:600}}>{D(day)}</div>
                                <div style={{fontSize:11,color:"#475569"}}>{(grouped[day]||[]).length} lançamento{(grouped[day]||[]).length!==1?"s":""} · em {n} dia{n!==1?"s":""}</div>
                              </div>
                              <div className="mono" style={{fontSize:14,fontWeight:700,color:"#10B981"}}>{R(net)}</div>
                            </div>
                          );
                        })
                    }
                  </div>

                  <div className="card">
                    <div style={{padding:"12px 16px",borderBottom:"1px solid #1A2840"}}>
                      <span style={{fontSize:12,fontWeight:700,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.06em"}}>Últimas Importações</span>
                    </div>
                    {imports.length===0
                      ? <div style={{padding:24,color:"#475569",fontSize:13,textAlign:"center"}}>Nenhum extrato importado ainda</div>
                      : imports.slice(0,5).map(imp=>(
                          <div key={imp.id} className="tr" style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 16px",borderBottom:"1px solid #0D1520"}}>
                            <div style={{minWidth:0}}>
                              <div style={{fontSize:12,color:"#CBD5E1",fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:200}}>{imp.filename}</div>
                              <div style={{fontSize:11,color:"#475569"}}>{Dt(imp.imported_at)} · {imp.transaction_count} transações</div>
                            </div>
                            <div className="mono" style={{fontSize:13,fontWeight:600,color:"#10B981",flexShrink:0,marginLeft:8}}>{R(imp.net_total)}</div>
                          </div>
                        ))
                    }
                    {imports.length>5&&(
                      <div style={{padding:"10px 16px"}}>
                        <button className="btn-o" style={{width:"100%",fontSize:11}} onClick={()=>setTab("importar")}>Ver todas ({imports.length}) →</button>
                      </div>
                    )}
                  </div>
                </div>

                {pendingRecon>0&&(
                  <div style={{marginTop:16,background:"#2A1F06",border:"1px solid #92400E",borderRadius:10,padding:"12px 16px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                    <div style={{fontSize:13,color:"#F59E0B"}}>⚠ {pendingRecon} dia{pendingRecon!==1?"s":""} com recebimento pendente de confirmação</div>
                    <button className="btn-o" style={{fontSize:12}} onClick={()=>setTab("conferencia")}>Conferir →</button>
                  </div>
                )}
              </div>
            )}

            {/* ══════════ IMPORTAR ══════════ */}
            {tab==="importar"&&(
              <div className="fd">
                {importing
                  ? <div style={{textAlign:"center",padding:"40px",background:"#0C1520",borderRadius:14,border:"1px solid #1A2840",marginBottom:20}}>
                      <div className="spin" style={{fontSize:32,marginBottom:14}}>⚙</div>
                      <div style={{color:"#10B981",fontWeight:600,marginBottom:6}}>{importMsg}</div>
                      <div style={{color:"#475569",fontSize:13}}>As transações serão somadas ao histórico acumulado</div>
                    </div>
                  : <div className={`zone${dragging?" drag":""}`} style={{marginBottom:20}}
                      onClick={()=>fileRef.current.click()}
                      onDragOver={e=>{e.preventDefault();setDragging(true)}} onDragLeave={()=>setDragging(false)} onDrop={onDrop}>
                      <div style={{fontSize:40,marginBottom:12}}>📂</div>
                      <div style={{fontSize:16,fontWeight:700,color:"#F1F5F9",marginBottom:6}}>Arraste ou clique para importar novo extrato</div>
                      <div style={{color:"#64748B",fontSize:13,marginBottom:14}}>As transações são somadas ao histórico — arquivo salvo no Supabase Storage</div>
                      <div style={{display:"inline-flex",gap:8}}>
                        {["PDF","XLSX","XLS","CSV"].map(f=><span key={f} style={{background:"#1A2840",color:"#64748B",borderRadius:6,padding:"3px 10px",fontSize:11,fontFamily:"var(--mono)",fontWeight:700}}>{f}</span>)}
                      </div>
                    </div>
                }

                {imports.length>0&&(
                  <div className="card" style={{overflow:"hidden"}}>
                    <div style={{padding:"12px 16px",borderBottom:"1px solid #1A2840",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                      <div style={{fontSize:12,fontWeight:700,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.06em"}}>Histórico de Importações</div>
                      <div style={{fontSize:11,color:"#475569"}}>{imports.length} arquivos · {R(totGross)} bruto acumulado</div>
                    </div>
                    {imports.map(imp=>(
                      <div key={imp.id} className="imp-row">
                        <div style={{width:36,height:36,borderRadius:8,background:"#0A1322",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>
                          {imp.filename?.toLowerCase().endsWith(".pdf")?"📄":"📊"}
                        </div>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{fontSize:13,color:"#CBD5E1",fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{imp.filename}</div>
                          <div style={{fontSize:11,color:"#475569"}}>{Dt(imp.imported_at)} · {imp.transaction_count} transações · {R(imp.gross_total)} bruto</div>
                        </div>
                        <div style={{textAlign:"right",marginRight:12,flexShrink:0}}>
                          <div className="mono" style={{fontSize:13,fontWeight:600,color:"#10B981"}}>{R(imp.net_total)}</div>
                          <div style={{fontSize:10,color:"#475569"}}>líquido</div>
                        </div>
                        <div style={{display:"flex",gap:6,flexShrink:0}}>
                          {imp.storage_path&&<button className="btn-o" style={{fontSize:11,padding:"5px 10px"}} onClick={()=>downloadFile(imp.storage_path,imp.filename)} title="Baixar arquivo">↓</button>}
                          <button className="btn-r" onClick={()=>deleteImport(imp.id,imp.storage_path)} title="Remover importação">✕</button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {/* ══════════ PREVISÃO ══════════ */}
            {tab==="previsao"&&(
              <div className="fd">
                {settlements.length===0
                  ? <div style={{textAlign:"center",padding:60,color:"#475569"}}>Importe um extrato para ver a previsão</div>
                  : <>
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:18}}>
                        <div style={{fontSize:13,color:"#64748B"}}>
                          Mostrando <span style={{color:"#94A3B8"}}>{showAll?"todo o histórico acumulado":"apenas recebimentos futuros"}</span>
                          {" · "}{previsaoDays.length} dias
                        </div>
                        <button className="btn-o" style={{fontSize:12}} onClick={()=>setShowAll(p=>!p)}>
                          {showAll?"Só futuros":"Ver histórico completo"}
                        </button>
                      </div>

                      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
                        {(()=>{
                          const pDays=previsaoDays;
                          const pG=pDays.reduce((a,d)=>(grouped[d]||[]).reduce((s,i)=>s+i.gross_amount,0)+a,0);
                          const pN=pDays.reduce((a,d)=>(grouped[d]||[]).reduce((s,i)=>s+i.net_amount,0)+a,0);
                          return [
                            {l:"Dias",   v:pDays.length, c:"#F1F5F9"},
                            {l:"Bruto",  v:R(pG),        c:"#F1F5F9"},
                            {l:"Taxas",  v:R(pG-pN),     c:"#EF4444"},
                            {l:"Líquido",v:R(pN),        c:"#10B981"},
                          ].map(s=>(
                            <div key={s.l} className="stat">
                              <div style={{fontSize:10,color:"#64748B",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:7}}>{s.l}</div>
                              <div className="mono" style={{fontSize:20,fontWeight:700,color:s.c}}>{s.v}</div>
                            </div>
                          ));
                        })()}
                      </div>

                      {previsaoDays.length===0
                        ? <div style={{textAlign:"center",padding:40,color:"#475569",fontSize:13}}>Nenhum recebimento futuro. Clique em "Ver histórico completo".</div>
                        : <div style={{display:"flex",flexDirection:"column",gap:10}}>
                            {previsaoDays.map(day=>{
                              const items=grouped[day]||[];
                              const dNet=items.reduce((a,i)=>a+i.net_amount,0);
                              const dGross=items.reduce((a,i)=>a+i.gross_amount,0);
                              const isToday=day===today, isPast=day<today;
                              return (
                                <div key={day} className="card" style={{overflow:"hidden",...(isToday?{borderColor:"#10B981"}:{})}}>
                                  <div className="day-hd" style={{background:isToday?"#041F12":isPast?"#080D14":"#0C1520"}}>
                                    <div style={{display:"flex",alignItems:"center",gap:10}}>
                                      <span className="mono" style={{fontSize:14,fontWeight:700,color:isToday?"#10B981":isPast?"#475569":"#F1F5F9"}}>{D(day)}</span>
                                      {isToday&&<span style={{background:"#10B981",color:"#001A0E",borderRadius:100,padding:"1px 7px",fontSize:9,fontWeight:800}}>HOJE</span>}
                                      {isPast&&<span style={{color:"#475569",fontSize:11}}>passado</span>}
                                      <span style={{color:"#334155",fontSize:12}}>{items.length} lançamento{items.length!==1?"s":""}</span>
                                    </div>
                                    <div style={{display:"flex",gap:22}}>
                                      <div style={{textAlign:"right"}}>
                                        <div style={{fontSize:9,color:"#475569",marginBottom:1,fontWeight:700}}>BRUTO</div>
                                        <div className="mono" style={{fontSize:12,color:"#64748B"}}>{R(dGross)}</div>
                                      </div>
                                      <div style={{textAlign:"right"}}>
                                        <div style={{fontSize:9,color:"#475569",marginBottom:1,fontWeight:700}}>LÍQUIDO</div>
                                        <div className="mono" style={{fontSize:16,fontWeight:700,color:"#10B981"}}>{R(dNet)}</div>
                                      </div>
                                    </div>
                                  </div>
                                  {items.map((it,idx)=>(
                                    <div key={idx} className="tx-row tr" style={{...( idx<items.length-1?{borderBottom:"1px solid #111C2C"}:{})}}>
                                      <TypeChip type={it.type} amount={it.gross_amount}/>
                                      <div style={{flex:1,fontSize:12,color:"#94A3B8",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                                        <span style={{color:"#CBD5E1"}}>{it.description||"—"}</span>
                                        {it.card_brand&&<span style={{color:"#475569",marginLeft:8}}>· {it.card_brand}</span>}
                                        {it.nsu&&<span className="mono" style={{color:"#334155",marginLeft:8,fontSize:10}}>NSU {it.nsu}</span>}
                                      </div>
                                      {it.imports?.filename&&(
                                        <span style={{fontSize:10,color:"#1E3A5F",background:"#0F2942",borderRadius:4,padding:"1px 5px",flexShrink:0,maxWidth:110,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={it.imports.filename}>
                                          {it.imports.filename.replace(/\.[^.]+$/,"").slice(0,14)}
                                        </span>
                                      )}
                                      <div className="mono" style={{fontSize:12,color:"#475569",textAlign:"right",minWidth:86}}>{R(it.gross_amount)}</div>
                                      <div style={{textAlign:"right",minWidth:100}}>
                                        <div className="mono" style={{fontSize:13,fontWeight:700,color:it.net_amount<0?"#EF4444":"#10B981"}}>{R(it.net_amount)}</div>
                                        <div style={{fontSize:10,color:"#475569"}}>taxa {it.taxa_pct?.toFixed(2)}%</div>
                                      </div>
                                    </div>
                                  ))}
                                </div>
                              );
                            })}
                          </div>
                      }
                    </>
                }
              </div>
            )}

            {/* ══════════ CONFERÊNCIA ══════════ */}
            {tab==="conferencia"&&(
              <div className="fd">
                {settlements.length===0
                  ? <div style={{textAlign:"center",padding:60,color:"#475569"}}>Importe um extrato para conferir</div>
                  : <>
                      {pendingRecon>0&&(
                        <div style={{background:"#2A1F06",border:"1px solid #92400E",borderRadius:10,padding:"10px 16px",marginBottom:14,fontSize:13,color:"#F59E0B"}}>
                          ⚠ {pendingRecon} dia{pendingRecon!==1?"s":""} com depósito esperado ainda não confirmado
                        </div>
                      )}
                      <div style={{fontSize:13,color:"#64748B",marginBottom:14}}>
                        Digite o valor recebido em cada data. Salvo automaticamente no Supabase. 💾
                      </div>
                      <div className="card" style={{overflow:"hidden"}}>
                        <table style={{width:"100%",borderCollapse:"collapse"}}>
                          <thead>
                            <tr style={{background:"#0A1322"}}>
                              {["Data","Lançtos","Modalidades","Previsto","Recebido","Diferença","Status"].map(h=>(
                                <th key={h} style={{padding:"10px 12px",fontSize:10,fontWeight:700,color:"#475569",textAlign:"left",textTransform:"uppercase",letterSpacing:"0.06em",borderBottom:"1px solid #1A2840",whiteSpace:"nowrap"}}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {conferenciaDays.map(day=>{
                              const items=grouped[day]||[];
                              if(!items.length) return null;
                              const expected=items.reduce((a,i)=>a+i.net_amount,0);
                              const rawVal=recon[day];
                              const actual=rawVal!==undefined&&rawVal!==""?parseFloat(rawVal)||0:null;
                              const diff=actual!==null?actual-expected:null;
                              const isPast=day<today, isToday=day===today;
                              const dayTypes=[...new Set(items.map(i=>i.type))];
                              let s={txt:"—",c:"#475569"};
                              if(actual!==null){ if(Math.abs(diff)<0.02) s={txt:"✓ OK",c:"#10B981"}; else if(diff>0) s={txt:"↑ A mais",c:"#F59E0B"}; else s={txt:"↓ A menos",c:"#EF4444"}; }
                              else if(isPast) s={txt:"⏳ Pendente",c:"#F59E0B"};
                              else s={txt:"◷ Futuro",c:"#334155"};
                              return (
                                <tr key={day} className="tr" style={{borderBottom:"1px solid #0D1520"}}>
                                  <td style={{padding:"11px 12px"}}>
                                    <span className="mono" style={{fontSize:13,fontWeight:600,color:isToday?"#10B981":isPast?"#9CA3AF":"#F1F5F9"}}>
                                      {D(day)}{isToday&&<span style={{marginLeft:6,fontSize:9,background:"#10B981",color:"#001A0E",borderRadius:100,padding:"1px 6px",fontWeight:800}}>HOJE</span>}
                                    </span>
                                  </td>
                                  <td style={{padding:"11px 12px"}}><span style={{fontSize:12,color:"#94A3B8"}}>{items.length}</span></td>
                                  <td style={{padding:"11px 12px"}}>
                                    <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                                      {dayTypes.map(t=>{const c=config[t]||config.debito;return<span key={t} style={{background:c.bg,color:c.cor,borderRadius:100,padding:"1px 6px",fontSize:9,fontWeight:700,fontFamily:"var(--mono)"}}>{c.label}</span>;})}
                                    </div>
                                  </td>
                                  <td style={{padding:"11px 12px"}}><span className="mono" style={{fontSize:13,fontWeight:700,color:"#10B981"}}>{R(expected)}</span></td>
                                  <td style={{padding:"11px 12px"}}>
                                    <input className="recon-in" type="number" step="0.01" placeholder="0,00" value={recon[day]??""} onChange={e=>updateRecon(day,e.target.value)}/>
                                  </td>
                                  <td style={{padding:"11px 12px"}}>
                                    <span className="mono" style={{fontSize:13,fontWeight:600,color:diff===null?"#334155":Math.abs(diff)<0.02?"#10B981":diff>0?"#F59E0B":"#EF4444"}}>
                                      {diff!==null?(diff>=0?"+":"")+R(diff):"—"}
                                    </span>
                                  </td>
                                  <td style={{padding:"11px 12px"}}><span style={{fontSize:12,fontWeight:600,color:s.c}}>{s.txt}</span></td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </>
                }
              </div>
            )}

            {/* ══════════ TRANSAÇÕES ══════════ */}
            {tab==="transacoes"&&(
              <div className="fd">
                {settlements.length===0
                  ? <div style={{textAlign:"center",padding:60,color:"#475569"}}>Importe um extrato para ver as transações</div>
                  : <div className="card" style={{overflow:"hidden"}}>
                      <div style={{padding:"12px 16px",borderBottom:"1px solid #1A2840",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                        <div style={{fontSize:12,fontWeight:700,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.06em"}}>Todas as Transações</div>
                        <div style={{fontSize:11,color:"#475569"}}>{settlements.length} registros · {R(totNet)} líquido acumulado</div>
                      </div>
                      <div style={{overflowX:"auto"}}>
                        <table style={{width:"100%",borderCollapse:"collapse",minWidth:900}}>
                          <thead>
                            <tr style={{background:"#0A1322"}}>
                              {["Arquivo","Data","Tipo","Descrição","NSU","D+","Liquidação","Bruto","Taxa","Líquido"].map(h=>(
                                <th key={h} style={{padding:"9px 12px",fontSize:10,fontWeight:700,color:"#475569",textAlign:"left",textTransform:"uppercase",letterSpacing:"0.06em",borderBottom:"1px solid #1A2840",whiteSpace:"nowrap"}}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {settlements.map((s,i)=>(
                              <tr key={i} className="tr" style={{borderBottom:"1px solid #0D1520"}}>
                                <td style={{padding:"8px 12px",maxWidth:120}}><span style={{fontSize:10,color:"#334155",display:"block",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={s.imports?.filename}>{s.imports?.filename?.replace(/\.[^.]+$/,"").slice(0,15)||"—"}</span></td>
                                <td style={{padding:"8px 12px"}}><span className="mono" style={{fontSize:11,color:"#94A3B8"}}>{D(s.date)}</span></td>
                                <td style={{padding:"8px 12px"}}><TypeChip type={s.type} amount={s.gross_amount}/></td>
                                <td style={{padding:"8px 12px",maxWidth:160}}><span style={{fontSize:11,color:"#CBD5E1"}}>{s.description||"—"}</span></td>
                                <td style={{padding:"8px 12px"}}><span className="mono" style={{fontSize:10,color:"#334155"}}>{s.nsu||"—"}</span></td>
                                <td style={{padding:"8px 12px"}}><span className="mono" style={{fontSize:11,color:"#64748B"}}>D+{s.prazo_dias}</span></td>
                                <td style={{padding:"8px 12px"}}><span className="mono" style={{fontSize:11,color:s.settlement_date<today?"#475569":s.settlement_date===today?"#10B981":"#F1F5F9"}}>{D(s.settlement_date)}</span></td>
                                <td style={{padding:"8px 12px",textAlign:"right"}}><span className="mono" style={{fontSize:11,color:"#64748B"}}>{R(s.gross_amount)}</span></td>
                                <td style={{padding:"8px 12px",textAlign:"right"}}><span className="mono" style={{fontSize:11,color:"#EF4444"}}>{s.taxa_pct?.toFixed(2)}%</span></td>
                                <td style={{padding:"8px 12px",textAlign:"right"}}><span className="mono" style={{fontSize:12,fontWeight:700,color:s.net_amount<0?"#EF4444":"#10B981"}}>{R(s.net_amount)}</span></td>
                              </tr>
                            ))}
                          </tbody>
                          <tfoot>
                            <tr style={{background:"#0A1322",borderTop:"2px solid #1A2840"}}>
                              <td colSpan={7} style={{padding:"10px 12px",fontSize:12,fontWeight:700,color:"#64748B"}}>TOTAL ACUMULADO</td>
                              <td style={{padding:"10px 12px",textAlign:"right"}}><span className="mono" style={{fontSize:12,fontWeight:700,color:"#F1F5F9"}}>{R(totGross)}</span></td>
                              <td style={{padding:"10px 12px",textAlign:"right"}}><span className="mono" style={{fontSize:12,fontWeight:700,color:"#EF4444"}}>{R(totTax)}</span></td>
                              <td style={{padding:"10px 12px",textAlign:"right"}}><span className="mono" style={{fontSize:12,fontWeight:700,color:"#10B981"}}>{R(totNet)}</span></td>
                            </tr>
                          </tfoot>
                        </table>
                      </div>
                    </div>
                }
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
