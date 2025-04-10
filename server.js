console.log("Iniciando servidor...");

const express = require("express");
const axios = require("axios");
const cors = require("cors");
const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json()); // Middleware para parsear JSON

app.get("/gmailPag.html", async (req, res) => {
  console.warn("hola web gmail");
});

// Ruta para manejar la redirección después de la autenticación de Google
app.get("/oauth2callback", async (req, res) => {
  const code = req.query.code; // El código de autorización que Google envía como query string

  if (!code) {
    return res
      .status(400)
      .send("Error: No se recibió el código de autorización");
  }

  // Asegúrate de que la URI de redirección sea exactamente la misma que configuraste en Google Cloud Console
  const redirectUri = `http://localhost:${PORT}/oauth2callback`; // Esta URI debe coincidir con lo que configuraste en Google Cloud

  try {
    // Usamos URLSearchParams para enviar los parámetros en formato URL encoded
    const params = new URLSearchParams();
    params.append("code", code);
    params.append(
      "client_id",
      "45594330364-68qsjfc7lo95iq95fvam08hb55oktu4c.apps.googleusercontent.com"
    ); // Tu client_id

    const SCOPES = [
      "https://www.googleapis.com/auth/gmail.readonly", // Solo lectura de correos de Gmail
    ];

    params.append("client_secret", "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"); // Tu client_secret
    params.append("redirect_uri", redirectUri); // La misma URI de redirección
    params.append("grant_type", "authorization_code");
    params.append("scope", SCOPES); // Alcance de solo lectura de Gmail

    // Realizamos la solicitud POST a Google para obtener el access token
    const response = await axios.post(
      "https://oauth2.googleapis.com/token",
      params, // Pasamos los parámetros como datos del cuerpo
      {
        headers: {
          "Content-Type": "application/x-www-form-urlencoded", // Indicamos que los datos se envían como formulario URL encoded
        },
      }
    );

    const accessToken = response.data.access_token;
    console.log("Access Token:", accessToken);

    res.redirect(`http://127.0.0.1:5500/gmailPag.html?tok=${accessToken}`);

    let respuestaEnviar = await obtenerEmailsConAsuntoDesignacion(accessToken);

    res.json(respuestaEnviar);
    // Devuelve el access token al frontend
    //res.json({ access_token: accessToken });
  } catch (error) {
    console.error(error.response ? error.response.data : error);
    res.status(500).send("Error al obtener el token");
  }
});

async function obtenerEmailsConAsuntoDesignacion(token) {
  console.log(token);
  if (!token) {
    console.log("No se pudo obtener un token.");
    return;
  }

  // URL para realizar la consulta de correos con el asunto "Designación APD"
  const url =
    "https://www.googleapis.com/gmail/v1/users/me/messages?q=subject:Designación%20APD";

  try {
    // Realizar la solicitud GET con el token de autorización
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });
    const data = response.data;

    if (data.messages && data.messages.length > 0) {
      // Obtener detalles completos de cada mensaje
      const messages = data.messages;
      const messageDetailsPromises = messages.map(async (message) => {
        const messageResponse = await axios.get(
          `https://www.googleapis.com/gmail/v1/users/me/messages/${message.id}`,
          {
            headers: {
              Authorization: `Bearer ${token}`,
            },
          }
        );
        return messageResponse.data;
      });

      // Esperar a que todas las promesas se resuelvan y devolver los correos completos
      const messageDetails = await Promise.all(messageDetailsPromises);
      console.log(messageDetails);
      return messageDetails;
    }
  } catch (error) {
    console.error(
      "Error al obtener los correos:",
      error.response ? error.response.data : error.message
    );
  }
}

app.post("/obtenerMails", async (req, res) => {
  const token = req.body.token;
  console.log(token);
  let resEnviar = await obtenerEmailsConAsuntoDesignacion(token);
  res.json(resEnviar);
});

// Middleware para manejar el acceso al token
app.use(express.json()); // Asegura que se pueda recibir el JSON en el cuerpo de la solicitud

// Ruta para leer los correos electrónicos
app.post("/getEmails", async (req, res) => {
  const { access_token } = req.body; // El Access Token debe ser enviado desde el frontend

  if (!access_token) {
    return res.status(400).send("Error: No se recibió el Access Token.");
  }

  try {
    // Solicitar los correos utilizando el Access Token
    const response = await axios.get(
      "https://gmail.googleapis.com/gmail/v1/users/me/messages",
      {
        headers: {
          Authorization: `Bearer ${access_token}`,
        },
      }
    );

    // Si la solicitud es exitosa, se devuelve la lista de mensajes
    console.log("Emails:", response.data);
    res.json(response.data); // Devolver la lista de correos al frontend
  } catch (error) {
    console.error("Error al obtener los correos:", error);
    res.status(500).send("Error al obtener los correos.");
  }
});

// Iniciar el servidor
app.listen(PORT, () => {
  console.log(`Servidor escuchando en http://localhost:${PORT}`);
});
