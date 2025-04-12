console.log("Iniciando servidor...");

const express = require("express");
const axios = require("axios");
const cors = require("cors");
const app = express();
const PORT = process.env.PORT || 10000;
const maxMensajes = 10;

//excel

const ExcelJS = require("exceljs");
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
app.get("/descargar", async (req, res) => {
  console.log("escuchando -descargar-");

  const rutaArchivo = path.join(__dirname, "plantilla_pac.xlsx");

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(rutaArchivo);

  const worksheet = workbook.getWorksheet(1); // Primera hoja (también puedes usar nombre)

  worksheet.getCell("C19").value = "new date";

  // Preparar para enviar el archivo directamente como descarga
  res.setHeader(
    "Content-Disposition",
    "attachment; filename=plantilla-formateada.xlsx"
  );
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );

  await workbook.xlsx.write(res);
  res.end();
});

//visualización previa antes de descargar

app.get("/ver", async (req, res) => {
  const rutaArchivo = path.join(__dirname, "plantilla_pac.xlsx");

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(rutaArchivo);

  const worksheet = workbook.getWorksheet(1);
  worksheet.getCell("C19").value = "Dato previo a la descarga";

  const mergeMap = {};
  worksheet._merges.forEach((merge) => {
    const topLeft = merge.tl;
    mergeMap[topLeft] = {
      colspan: merge.br.col - merge.tl.col + 1,
      rowspan: merge.br.row - merge.tl.row + 1,
    };
  });

  let html = '<table border="1" style="border-collapse: collapse;">';

  worksheet.eachRow((row, rowNum) => {
    html += "<tr>";
    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      const cellId = `${cell.address}`;
      const merge = mergeMap[cellId];

      // Evitar celdas que están dentro de una combinación pero no son la celda superior izquierda
      if (
        Object.values(mergeMap).some((m) => {
          const start = worksheet.getCell(cellId)._mergeStart;
          return start && start !== cellId;
        })
      ) {
        return; // Saltamos celdas combinadas duplicadas
      }

      let style = "";
      const font = cell.style?.font || {};
      const fill = cell.style?.fill || {};
      const alignment = cell.style?.alignment || {};

      if (font.bold) style += "font-weight:bold;";
      if (alignment.horizontal) style += `text-align:${alignment.horizontal};`;

      if (fill.fgColor?.argb) {
        const bg = `#${fill.fgColor.argb.slice(2)}`;
        style += `background-color:${bg};`;
      }

      const colspan = merge?.colspan ? `colspan="${merge.colspan}"` : "";
      const rowspan = merge?.rowspan ? `rowspan="${merge.rowspan}"` : "";

      html += `<td ${colspan} ${rowspan} style="${style}">${
        cell.value ?? ""
      }</td>`;
    });
    html += "</tr>";
  });

  html += "</table>";

  res.send(`
    <html>
      <head>
        <title>Vista previa del Excel</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          td { padding: 6px 12px; min-width: 80px; }
        </style>
      </head>
      <body>
        <h2>Vista previa del archivo Excel</h2>
        ${html}
        <br><br>
        <a href="/descargar">📥 Descargar Excel con formato</a>
      </body>
    </html>
  `);
});

// Iniciar el servidor

app.listen(PORT, () => {
  console.log(`Servidor escuchando en puerto ${PORT}`);
  console.log("--versión con excel 2!");
});
