# 💳 Conferência Rede

Sistema de conferência de extratos da Rede — Débito, Crédito, Alelo, Ticket e VR.

---

## 🚀 Deploy no Netlify (recomendado)

### 1. Coloque o projeto no GitHub
Crie um repositório no GitHub e suba os arquivos.

### 2. Conecte ao Netlify
- Entre em netlify.com → "Add new site" → "Import an existing project"
- Escolha o repositório do GitHub

### 3. Configure o build (já vem preenchido pelo netlify.toml)
- Build command: `npm run build`
- Publish directory: `dist`

### 4. Adicione a variável de ambiente
- No Netlify: Site configuration → Environment variables → Add a variable
- Key: `ANTHROPIC_API_KEY`
- Value: sua chave (em console.anthropic.com)

### 5. Deploy!
Clique em "Deploy site". Em ~1 minuto o site estará no ar.

---

## 💻 Rodar localmente (desenvolvimento)

```bash
npm install
cp .env.example .env   # edite e coloque sua ANTHROPIC_API_KEY
npm run dev            # abre em http://localhost:3000
```

---

## 📁 Estrutura

```
conferencia-rede/
├── netlify/
│   └── functions/
│       └── parse.js      ← API serverless (chave segura no servidor)
├── netlify.toml           ← Config do Netlify (build + redirects)
├── server.js              ← Express para dev local
├── src/
│   ├── main.jsx
│   └── App.jsx
└── package.json
```
