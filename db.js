// db.js
const { MongoClient } = require("mongodb");

const uri = process.env.URI_MONGODB; // Cambia por tu URI si es remota
const client = new MongoClient(uri);

let db;

async function connectDB() {
  try {
    await client.connect();
    db = client.db("creadorDePacMongoDB"); // Reemplaza con el nombre de tu DB
    console.log("✅ Conectado a creadorDePacMongoDB");
  } catch (err) {
    console.error("❌ Error al conectar a creadorDePacMongoDB:", err);
  }
}

function getDB() {
  if (!db) {
    throw new Error(
      "❌ La base de datos no está conectada, creadorDePacMongoDB"
    );
  }
  return db;
}

module.exports = {
  connectDB,
  getDB,
};
