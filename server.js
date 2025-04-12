console.log("Iniciando servidor...");

const express = require("express");
const axios = require("axios");
const cors = require("cors");
const app = express();
const PORT = process.env.PORT || 10000;
const maxMensajes = 10;

//excel

const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

app.use(cors());
app.use(express.json()); // Middleware para parsear JSON

app.get("/gmailPag.html", async (req, res) => {
  console.warn("redireccionando a gmailPag");
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
  //  const redirectUri = `http://localhost:${PORT}/oauth2callback`; // Esta URI debe coincidir con lo que configuraste en Google Cloud
  const redirectUri = `https://creador-de-pac-backend.onrender.com/oauth2callback`;

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

    res.redirect(
      `https://adrianbenitezdev.github.io/CREADOR-DE-PAC/gmailPag.html?tok=${accessToken}`
    );

    //  let respuestaEnviar = await obtenerEmailsConAsuntoDesignacion(accessToken);

    // res.json(respuestaEnviar);
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
  } else {
    console.log("Obteniendo mensajes únicos con pausas de 200ms...");
  }

  const url =
    "https://www.googleapis.com/gmail/v1/users/me/messages?q=subject:Designación%20APD";

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    const data = response.data;

    if (data.messages && data.messages.length > 0) {
      const threadIdsUnicos = new Set();
      const messageDetails = [];

      for (const message of data.messages) {
        if (messageDetails.length >= maxMensajes) break;

        try {
          const messageResponse = await axios.get(
            `https://www.googleapis.com/gmail/v1/users/me/messages/${message.id}`,
            {
              headers: {
                Authorization: `Bearer ${token}`,
              },
            }
          );

          const detalle = messageResponse.data;

          // Solo agregar si el threadId es nuevo
          if (!threadIdsUnicos.has(detalle.threadId)) {
            threadIdsUnicos.add(detalle.threadId);
            messageDetails.push(detalle);
          }

          // Esperar 200ms para evitar sobrecarga
          await new Promise((resolve) => setTimeout(resolve, 200));
        } catch (error) {
          console.error(
            `Error al obtener el mensaje con ID ${message.id}:`,
            error.message
          );
        }
      }

      console.log(messageDetails);
      return messageDetails;
    } else {
      console.log("No se encontraron mensajes con ese asunto.");
      return [];
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

//PARTE PARA REALIZAR UN ARCHIVO EXCEL Y ENVIARLO AL CLIENTE

// Ruta para modificar el archivo y servirlo
app.get("/descargar", (req, res) => {
  console.log("escuchando -descargar-");

  const rutaArchivo = path.join(__dirname, "plantilla_pac.xlsx");

  // Leer el archivo original
  const workbook = XLSX.readFile(rutaArchivo);
  const hoja = workbook.Sheets[workbook.SheetNames[0]];

  // Modificar la celda C19
  hoja["C19"] = { t: "s", v: "Dato nuevo desde el servidor" };

  // Actualizar el rango si es necesario
  const rango = XLSX.utils.decode_range(hoja["!ref"]);
  rango.e.r = Math.max(rango.e.r, 18); // fila 19 (0 indexado)
  rango.e.c = Math.max(rango.e.c, 2); // columna C (0 indexado)
  hoja["!ref"] = XLSX.utils.encode_range(rango);

  // Escribir en un archivo temporal
  const archivoTemporal = path.join(__dirname, "plantilla-modificada.xlsx");
  XLSX.writeFile(workbook, archivoTemporal);

  // Enviar el archivo al cliente
  res.download(archivoTemporal, "plantilla-modificada.xlsx", (err) => {
    if (err) {
      console.error("Error al enviar el archivo:", err);
      res.status(500).send("Error al enviar el archivo");
    } else {
      // Opcional: eliminar archivo temporal si no lo necesitas
      fs.unlink(archivoTemporal, () => {});
    }
  });
});

// Iniciar el servidor
app.listen(PORT, () => {
  console.log(`Servidor escuchando en puerto ${PORT}`);
  console.log("--versión con excel!");
});
