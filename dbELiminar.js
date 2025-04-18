const express = require("express");
const app = express();
const PORT = 3000;

const { connectDB, getDB } = require("./db");

connectDB()
  .then(() => {
    const db = getDB();
    const usuarios = db.collection("usuarios");

    // Todas tus rutas que usan la DB van acá
    app.get("/verUsuarios", async (req, res) => {
      const lista = await usuarios.find().toArray();
      res.json(lista);
    });

    app.get("/agregar", async (req, res) => {
      res.json({ mensaje: "Usuario agregado", id: resultado.insertedId });
    });

    // Solo después de conectar, se puede arrancar el servidor
    app.listen(PORT, () => {
      console.log(`Servidor corriendo en http://localhost:${PORT}`);
    });
  })
  .catch((err) => {
    console.error("Error conectando a MongoDB:", err);
  });
