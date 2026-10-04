import express from "express";
import { createServer } from "http";
import { Server } from "socket.io";
import { createServer as createViteServer } from "vite";
import path from "path";
import { fileURLToPath } from "url";
import { GoogleGenAI, Type } from "@google/genai";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  app.use(express.json({ limit: "25mb" }));
  const httpServer = createServer(app);
  const io = new Server(httpServer, {
    cors: {
      origin: "*",
    },
  });

  const PORT = 3000;

  app.post("/api/ocr-id-document", async (req, res) => {
    try {
      const { images } = req.body;
      if (!Array.isArray(images) || images.length === 0) {
        return res.status(400).json({ error: "Nessuna immagine fornita." });
      }

      if (!process.env.GEMINI_API_KEY) {
        return res.status(500).json({ error: "Chiave API Gemini non configurata sul server." });
      }

      const ai = new GoogleGenAI({
        apiKey: process.env.GEMINI_API_KEY,
        httpOptions: {
          headers: {
            "User-Agent": "aistudio-build",
          },
        },
      });

      const parts: any[] = images.map((img: { data: string; mimeType?: string }) => {
        const cleanBase64 = img.data.includes(",") ? img.data.split(",")[1] : img.data;
        return {
          inlineData: {
            mimeType: img.mimeType || "image/jpeg",
            data: cleanBase64,
          },
        };
      });

      parts.push({
        text:
          "Analizza questo documento d'identità italiano (es. Carta d'Identità elettronica o cartacea, Patente di guida, Passaporto). " +
          "Estrai il Nome e Cognome completi del titolare (in formato 'NOME COGNOME' tutto maiuscolo), la Data di Scadenza del documento (in formato YYYY-MM-DD), " +
          "il Numero del Documento e il Tipo di Documento. Se un dato non è visibile restituisci stringa vuota.",
      });

      const response = await ai.models.generateContent({
        model: "gemini-3.8-flash",
        contents: { parts },
        config: {
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              fullName: {
                type: Type.STRING,
                description: "Nome e Cognome del titolare in maiuscolo (es. MARIO ROSSI)",
              },
              expiryDate: {
                type: Type.STRING,
                description: "Data di scadenza del documento in formato YYYY-MM-DD (es. 2030-05-14)",
              },
              documentNumber: {
                type: Type.STRING,
                description: "Numero del documento d'identità (es. CA12345AA)",
              },
              documentType: {
                type: Type.STRING,
                description: "Tipo di documento (es. Carta d'Identità, Patente, Passaporto)",
              },
            },
            required: ["fullName", "expiryDate"],
          },
        },
      });

      const rawText = response.text || "{}";
      const parsed = JSON.parse(rawText);
      return res.json(parsed);
    } catch (error: any) {
      console.error("Errore OCR documento:", error);
      return res.status(500).json({ error: error?.message || "Errore durante il riconoscimento del documento." });
    }
  });
  
  // In-memory state for messages
  let messages: any[] = [];

  io.on("connection", (socket) => {
    console.log("A user connected:", socket.id);

    // Send initial state
    socket.emit("init_messages", messages);

    socket.on("send_message", (msg) => {
      const newMessage = {
        ...msg,
        id: Math.random().toString(36).substr(2, 9),
        timestamp: new Date().toISOString(),
        isRead: false,
      };
      messages.push(newMessage);
      io.emit("new_message", newMessage);

      if (msg.asEmail) {
        console.log(`[EMAIL NOTIFICATION] Sending email for message ${newMessage.id} to target ${msg.targetId}`);
        // Qui andrebbe la logica di invio email reale (es. nodemailer)
      }
    });

    socket.on("mark_as_read", (id) => {
      messages = messages.map(m => m.id === id ? { ...m, isRead: true } : m);
      io.emit("message_read", id);
    });

    socket.on("mark_target_as_read", ({ targetId, userId }) => {
      // Mark all messages for this target as read if they weren't sent by this user
      messages = messages.map(m => 
        (m.targetId === targetId && m.senderId !== userId) ? { ...m, isRead: true } : m
      );
      io.emit("target_read", { targetId, userId });
    });

    socket.on("disconnect", () => {
      console.log("User disconnected:", socket.id);
    });
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.join(__dirname, "dist")));
    app.get("*all", (req, res) => {
      res.sendFile(path.join(__dirname, "dist", "index.html"));
    });
  }

  httpServer.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
