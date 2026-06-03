import { useState } from "react";
import { supabase } from "./supabase.js";

export default function Auth() {
  const [email, setEmail]     = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError]     = useState(null);

  async function handleSubmit(e) {
    e.preventDefault();
    setLoading(true); setError(null);
    try {
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) throw error;
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  const css = `
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:opsz,wght@9..40,400;9..40,500;9..40,700&family=JetBrains+Mono:wght@400;600&display=swap');
    *{box-sizing:border-box;margin:0;padding:0;}
    .auth-inp{display:block;width:100%;background:#050A12;border:1px solid #1E2D45;border-radius:8px;padding:11px 14px;color:#E2E8F0;font-family:'DM Sans',sans-serif;font-size:14px;outline:none;transition:border-color .15s;margin-bottom:12px;}
    .auth-inp:focus{border-color:#10B981;}
    .auth-btn{width:100%;background:#10B981;color:#001A0E;border:none;border-radius:8px;padding:12px;font-family:'DM Sans',sans-serif;font-weight:700;font-size:14px;cursor:pointer;transition:.15s;margin-top:4px;}
    .auth-btn:hover{background:#059669;} .auth-btn:disabled{opacity:.5;cursor:default;}
    .auth-link{background:none;border:none;color:#10B981;font-family:'DM Sans',sans-serif;font-size:13px;cursor:pointer;text-decoration:underline;}
  `;

  return (
    <div style={{fontFamily:"'DM Sans','Helvetica Neue',sans-serif",background:"#070C16",minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",color:"#E2E8F0"}}>
      <style>{css}</style>
      <div style={{width:"100%",maxWidth:380,padding:"0 20px"}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{width:52,height:52,borderRadius:12,background:"linear-gradient(135deg,#10B981,#059669)",display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:24,marginBottom:16}}>💳</div>
          <div style={{fontSize:22,fontWeight:700,color:"#F1F5F9"}}>Conferência Rede</div>
          <div style={{fontSize:13,color:"#64748B",marginTop:4}}>Entre na sua conta</div>
        </div>

        <div style={{background:"#0C1520",border:"1px solid #1A2840",borderRadius:14,padding:28}}>
          {error && (
            <div style={{background:"#1C0808",border:"1px solid #7F1D1D",borderRadius:8,padding:"10px 14px",color:"#EF4444",fontSize:13,marginBottom:16}}>
              {error}
            </div>
          )}

          <form onSubmit={handleSubmit}>
            <label style={{fontSize:12,color:"#64748B",fontWeight:600,display:"block",marginBottom:5}}>E-MAIL</label>
            <input className="auth-inp" type="email" placeholder="seu@email.com" value={email} onChange={e=>setEmail(e.target.value)} required autoFocus/>

            <label style={{fontSize:12,color:"#64748B",fontWeight:600,display:"block",marginBottom:5}}>SENHA</label>
            <input className="auth-inp" type="password" placeholder="••••••••" value={password} onChange={e=>setPassword(e.target.value)} required minLength={6}/>

            <button className="auth-btn" type="submit" disabled={loading}>
              {loading ? "Aguarde…" : "Entrar"}
            </button>
          </form>
        </div>
      </div>
    </div>
  );
}
