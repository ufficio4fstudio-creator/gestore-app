import express from "express";
import { createServer } from "http";
import { Server } from "socket.io";
import { createServer as createViteServer } from "vite";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const httpServer = createServer(app);
  const io = new Server(httpServer, {
    cors: {
      origin: "*",
    },
  });

  const PORT = 3000;
  
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
    app.get("*", (req, res) => {
      res.sendFile(path.join(__dirname, "dist", "index.html"));
    });
  }

  httpServer.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
