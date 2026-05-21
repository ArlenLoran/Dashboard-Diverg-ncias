import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import dotenv from "dotenv";
import cors from "cors";

import { GoogleGenAI } from "@google/genai";

dotenv.config();

// Initialize Gemini API
const genAI = new GoogleGenAI({ 
  apiKey: process.env.GEMINI_API_KEY || "",
  httpOptions: {
    headers: {
      'User-Agent': 'aistudio-build',
    }
  }
});

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Enable CORS for SharePoint integration
  app.use(cors());
  app.use(express.json());
  
  // Request logging
  app.use((req, res, next) => {
    console.log(`${req.method} ${req.url}`);
    next();
  });

  // AI Analysis Route
  app.post("/api/ai/analyze", async (req, res) => {
    try {
      const { metricTitle, objective, rules, data, history } = req.body;
      
      if (!metricTitle) {
        return res.status(400).json({ error: "Missing metric title" });
      }

      const prompt = `
        Você é um analista de dados especialista em logística e qualidade operacional.
        Analise a seguinte métrica e forneça insights acionáveis.

        MÉTRICA: ${metricTitle}
        OBJETIVO: ${objective}
        REGRAS DE NEGÓCIO: ${rules?.join(', ') || 'Não informadas'}
        
        SITUAÇÃO ATUAL:
        - Quantidade de divergências atual: ${data?.length || 0}
        - Histórico das últimas 10 atualizações (quantidade de linhas): ${history?.join(', ') || 'Sem histórico'}

        DADOS RECENTES (AMOSTRA):
        ${JSON.stringify(data?.slice(0, 5) || [], null, 2)}

        QUESTÕES PARA ANALISAR:
        1. Qual a tendência atual baseada no histórico? (Melhorando, Piorando ou Estável)
        2. Qual a provável causa raiz baseada nas regras e nos dados?
        3. Quais são as 3 recomendações imediatas para o time operacional?
        4. Com base no SLA atual (${history?.[0] || 0} erros), qual o nível de risco?

        Responda de forma concisa, profissional e formatada em Markdown. Use emojis para destacar pontos importantes.
      `;

      const response = await genAI.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: prompt
      });

      res.json({ analysis: response.text });
    } catch (error: any) {
      console.error("AI analysis error:", error);
      res.status(500).json({ error: error.message });
    }
  });
  
  // API Route to proxy Power Automate
  app.post("/api/query", async (req, res) => {
    console.log("Received query request:", req.body);
    try {
      const { query, id_score } = req.body || {};
      if (!query) {
        return res.status(400).json({ error: "Missing query in request body" });
      }
      
      let API_URL = process.env.POWER_AUTOMATE_URL;
      if (!API_URL || !API_URL.startsWith("https://") || API_URL.includes("POWER_AUTOMATE_URL")) {
        API_URL = "https://51a805d34213e248a3506f5db8fe28.55.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/655aac37bdea49b1b1221a2f37198754/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-2l0x4h5cwmpZ20RCIbMrzaR0860ka4aB8_dDOVQQHQ";
      }

      console.log("Proxying query through Express server to:", API_URL);

      const response = await fetch(API_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({ query, id_score })
      });

      if (!response.ok) {
        const errorText = await response.text();
        console.error(`Power Automate error (${response.status}):`, errorText);
        return res.status(response.status).json({ error: `Power Automate error: ${response.status}`, details: errorText });
      }

      const data = await response.json();
      res.json(data);
    } catch (error: any) {
      console.error("Query proxy error:", error);
      res.status(500).json({ error: error.message });
    }
  });

  // API Route to proxy email sending via Power Automate
  app.post("/api/send-email", async (req, res) => {
    console.log("Received send-email request:", req.body);
    try {
      const { emails, Title, BodyEmail, Attachments } = req.body || {};
      
      let EMAIL_URL = process.env.POWER_AUTOMATE_EMAIL_URL;
      if (!EMAIL_URL || !EMAIL_URL.startsWith("https://") || EMAIL_URL.includes("POWER_AUTOMATE_EMAIL_URL")) {
        EMAIL_URL = "https://51a805d34213e248a3506f5db8fe28.55.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/3e998fbce06445cdae41e91bfa5547de/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=6ujIHRbfYrL4_kZ9ZWcuu_LkL9TZyUWMs-VJes1U-As";
      }

      console.log("Proxying email through Express server to:", EMAIL_URL);

      const response = await fetch(EMAIL_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          emails: emails || "",
          Title: Title || "Notificação do Painel Operacional",
          BodyEmail: BodyEmail || "",
          Attachments: Attachments || []
        })
      });

      if (!response.ok) {
        const errorText = await response.text();
        console.error(`Power Automate Email error (${response.status}):`, errorText);
        return res.status(response.status).json({ error: `Power Automate email error: ${response.status}`, details: errorText });
      }

      let data = {};
      const resText = await response.text();
      if (resText) {
        try {
          data = JSON.parse(resText);
        } catch {
          data = { message: resText };
        }
      }
      res.json(data);
    } catch (error: any) {
      console.error("Email proxy query error:", error);
      res.status(500).json({ error: error.message });
    }
  });

  // Health check
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('/admin', (req, res) => {
      res.sendFile(path.join(distPath, 'admin.html'));
    });
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
