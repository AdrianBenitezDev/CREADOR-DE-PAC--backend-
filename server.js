console.log("Iniciando servidor...");

let accessToken = null;
let refreshToken = null;
let sub = null;

const express = require("express");
const axios = require("axios");
const cors = require("cors");
const app = express();
const PORT = process.env.PORT || 10000;
const maxMensajes = 10;
const { connectDB, getDB } = require("./db");

app.use(express.static("plantilla_pac/resources"));

//excel

const ExcelJS = require("exceljs");
const path = require("path");

app.use(cors());
app.use(express.json()); // Middleware para parsear JSON

app.post("/obtenerMails", async (req, res) => {
  //const token = req.body.token;
  const maxFilaReq = req.body.maxFila;

  let maxFila = 10;
  if (maxFilaReq == 10 || 20 || 30) {
    maxFila = maxFilaReq;
  } else {
    maxFila = 10;
  }

  //console.log(token);
  let resEnviar = await obtenerEmailsConAsuntoDesignacion(maxFila);
  res.json(resEnviar);
});

//PARTE PARA REALIZAR UN ARCHIVO EXCEL Y ENVIARLO AL CLIENTE

// Ruta para modificar el archivo y servirlo
app.post("/generarPac", async (req, res) => {
  const datosPac = req.body.objeto;
  const headerPac = JSON.parse(req.body.headerPac);

  const rutaArchivo = path.join(__dirname, `plantilla_pac_${maxMensajes}.xlsx`);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(rutaArchivo);

  const worksheet = workbook.getWorksheet(1); // Primera hoja (también puedes usar nombre)

  //aleramos celdas puntuales
  worksheet.getCell("A5").value = "Domicilio: " + headerPac.domicilio;
  worksheet.getCell("A6").value = "Telefono: " + headerPac.telefono;
  worksheet.getCell("A10").value = "Categoria: " + headerPac.categoria;
  worksheet.getCell("A11").value = "Turno: " + headerPac.turno;
  worksheet.getCell("A12").value =
    "Desfavorabilidad: " + headerPac.desfavorabilidad;
  worksheet.getCell("K6").value = headerPac.titlePac;
  worksheet.getCell("AI5").value = headerPac.numDistrito;
  worksheet.getCell("AM5").value = headerPac.tipoOrganizacion;
  worksheet.getCell("AQ5").value = headerPac.escuela;

  datosPac.forEach((fila, numeroFila) => {
    worksheet.getCell(`A${19 + numeroFila}`).value = fila.cupof; //cupof a19
    worksheet.getCell(`D${19 + numeroFila}`).value = fila.dni; //dni
    worksheet.getCell(`G${19 + numeroFila}`).value = fila.name; //name
    worksheet.getCell(`H${19 + numeroFila}`).value = fila.revista; //resvista
    worksheet.getCell(`J${19 + numeroFila}`).value = fila.pid; //resvista
    worksheet.getCell(`K${19 + numeroFila}`).value = fila.mod; //mod
    worksheet.getCell(`M${19 + numeroFila}`).value = fila.year; //año
    worksheet.getCell(`N${19 + numeroFila}`).value = fila.seccion; //seccion
    worksheet.getCell(`O${19 + numeroFila}`).value = fila.turno; //turno
  });

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
const fs = require("fs");
const { callbackify } = require("util");
//const { getHeapCodeStatistics } = require("v8");
app.post("/ver", async (req, res) => {
  const datosPac = req.body.objeto;
  const headerPac = JSON.parse(req.body.headerPac);

  console.log("--ver");

  const htmlPath = path.join(__dirname, "/plantilla_pac/ANVERSO.html");
  //const htmlPathReverso = path.join(__dirname, "/plantilla_pac/REVERSO.html");

  let html1 = fs.readFileSync(htmlPath, "utf8");
  //let html2 = fs.readFileSync(htmlPathReverso, "utf8");

  let agregarHTML = "";

  datosPac.forEach((fila) => {
    agregarHTML += `<tr style="height: 53px">
      <th id="412696113R18" style="height: 53px;" class="row-headers-background">
        <div class="row-header-wrapper" style="line-height: 53px"></div>
      </th>
      <td class="s75" dir="ltr">${fila.cupof}</td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75">${fila.dni}</td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s76">${fila.name}</td>
      <td class="s75">${fila.revista}</td>
      <td class="s75"></td>
      <td class="s75">${fila.pid}</td>
      <td class="s75">${fila.mod}</td>
      <td class="s75"></td>
      <td class="s75">${fila.year}</td>
      <td class="s75">${fila.seccion}</td>
      <td class="s75">${fila.turno}</td>
      <td class="s77"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s77"></td>
      <td class="s77"></td>
      <td class="s77"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s75"></td>
      <td class="s77" dir="ltr"></td>
    </tr>`;
  });
  html1 = html1.replace("{{domicilio}}", headerPac.domicilio);
  html1 = html1.replace("{{telefono}}", headerPac.telefono);
  html1 = html1.replace("{{email}}", headerPac.email);

  html1 = html1.replace("{{categoria}}", headerPac.categoria);
  html1 = html1.replace("{{turno}}", headerPac.turno);
  html1 = html1.replace("{{desfavorabilidad}}", headerPac.desfavorabilidad);

  html1 = html1.replace("{{PAC}}", headerPac.titlePac || ""); // solo si existe esa propiedad
  html1 = html1.replace("{{distrito}}", headerPac.numDistrito); // según el nombre que usás
  html1 = html1.replace("{{organizacion}}", headerPac.tipoOrganizacion);
  html1 = html1.replace("{{escuela}}", headerPac.escuela);

  html1 = html1.replace("<tr><td>inyectorAnverso</td></tr>", agregarHTML);
  //html2 = html2.replace("{{inyectorReverso}}", htmlReverso || "");

  const htmlC = html1;

  res.send(htmlC);
});

async function obtenerEmailsConAsuntoDesignacion(maxFila) {
  const url =
    "https://www.googleapis.com/gmail/v1/users/me/messages?q=subject:Designación%20APD";

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    const data = response.data;

    if (data.messages && data.messages.length > 0) {
      const threadIdsUnicos = new Set();
      const messageDetails = [];

      for (const message of data.messages) {
        if (messageDetails.length >= maxFila) break;

        try {
          const messageResponse = await axios.get(
            `https://www.googleapis.com/gmail/v1/users/me/messages/${message.id}`,
            {
              headers: {
                Authorization: `Bearer ${accessToken}`,
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

      //console.log(messageDetails);
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
    refrescarAccessToken(obtenerEmailsConAsuntoDesignacion(maxFila));
  }
}
//obtener mensajes con palabras personalizadas
async function obtenerEmailsConAsuntoDesignacionPersonalizado(
  maxFila,
  datosConsulta
) {
  //preparamos los datos para concatenarlos en la URL
  let datosConsultaPreparado = encodeURIComponent(datosConsulta);

  const url =
    "https://www.googleapis.com/gmail/v1/users/me/messages?q=Designación%20APD%20" +
    datosConsultaPreparado;

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    const data = response.data;

    if (data.messages && data.messages.length > 0) {
      const threadIdsUnicos = new Set();
      const messageDetails = [];

      for (const message of data.messages) {
        if (messageDetails.length >= maxFila) break;

        try {
          const messageResponse = await axios.get(
            `https://www.googleapis.com/gmail/v1/users/me/messages/${message.id}`,
            {
              headers: {
                Authorization: `Bearer ${accessToken}`,
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

      //console.log(messageDetails);
      return messageDetails;
    } else {
      console.log(
        "No se encontraron coincidencias con los parametros ingresados"
      );
      console.log(url);
      return [];
    }
  } catch (error) {
    console.error(
      "Error al obtener los correos:",
      error.response ? error.response.data : error.message
    );
    refrescarAccessToken(
      obtenerEmailsConAsuntoDesignacionPersonalizado(maxFila, datosConsulta)
    );
  }
}

app.post("/obtenerMailsPersonalizado", async (req, res) => {
  const maxFilaReq = req.body.maxFila;
  let datosConsultaSinRevisar = req.body.datosConsulta;
  let datosConsulta = "";

  let maxFila = 10;
  if (maxFilaReq == 10 || 20 || 30) {
    maxFila = maxFilaReq;
  } else {
    maxFila = 10;
  }

  if (/[^a-zA-Z0-9\s]/.test(datosConsultaSinRevisar)) {
    // Contiene caracteres especiales
    console.error("error:consultaPersonalizada tiene caracteres especiales");
    return;
  } else {
    // Solo tiene letras, números o espacios
    datosConsulta = datosConsultaSinRevisar;
  }

  //console.log(token);
  let resEnviar = await obtenerEmailsConAsuntoDesignacionPersonalizado(
    maxFila,
    datosConsulta
  );
  res.json(resEnviar);
});

// CON BASE DE DATOS COLOCAMOS ADENTRO LAS CONSULTAS QUE UTILIZAN LA BASE DE DATOS
connectDB().then(() => {
  const db = getDB();
  const usuarios = db.collection("usuarios");

  //se autentica una sola vez!!

  app.get("/oauth2callback", async (req, res) => {
    const code = req.query.code;
    if (!code) return res.status(400).send("Falta el código");

    const redirectUri =
      "https://creador-de-pac-backend.onrender.com/oauth2callback";

    try {
      const params = new URLSearchParams();
      params.append("code", code);
      params.append(
        "client_id",
        "45594330364-68qsjfc7lo95iq95fvam08hb55oktu4c.apps.googleusercontent.com"
      );

      params.append("client_secret", process.env.MY_CLIENT_SECRET);
      params.append("redirect_uri", redirectUri);
      params.append("grant_type", "authorization_code");

      const tokenRes = await axios.post(
        "https://oauth2.googleapis.com/token",
        params,
        {
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
        }
      );

      //obtenemos el refresh token y el token, las declaramos para ser utilizadas en la app
      accessToken = tokenRes.data.access_token;
      refreshToken = tokenRes.data.refresh_token;

      //obtenemos la info del usuario de googles
      const profileRes = await axios.get(
        "https://www.googleapis.com/oauth2/v3/userinfo",
        {
          headers: { Authorization: `Bearer ${accessToken}` },
        }
      );
      const profile = profileRes.data;
      sub = profile.sub;

      //guardamos el usuario con el refresh token VERIFICANDO QUE NO EXISTA EL USUARIO PREVIAMENTE

      agregarAndActualizarUsuariosDb(
        usuarios,
        profile.email,
        profile.sub,
        profile.name,
        profile.picture,
        refreshToken,
        accessToken
      );

      // Página HTML con postMessage
      res.send(`
        <html>
          <body>
            <script>
              window.opener.postMessage({
                profile: ${JSON.stringify(profile)}
              }, "https://adrianbenitezdev.github.io");
              window.close();
            </script>
          </body>
        </html>
      `);
    } catch (err) {
      console.error(err);
      res.status(500).send("Error en la autenticación" + err);
    }
  });

  app.get("/all", async (req, res) => {
    const usuariosLista = await usuarios.find().toArray();
    res.json(usuariosLista);
  });

  app.post("/obtenerVariablesGlobales", async (req, res) => {
    const user_id = req.body.user_google_id;
    //realizamos la consulta para traer los datos de mongoDb
    let resp = await leerUsuarios(usuarios, user_id);

    //actualizamos las variable globales de los tokens
    accessToken = resp.access_Token;
    refreshToken = resp.refresh_Token;
    sub = user_id;

    if (resp) {
      res.json({
        google_id: resp.google_id, // <- este es el `sub`, tu identificador clave
        nombre: resp.nombre,
        foto: resp.foto,
        email: resp.email,
      });
    } else {
      res.json({ mensaje: "no hay usuario guardado para el id:" + user_id });
    }
  });
});

app.listen(PORT, () => {
  console.log(`Servidor escuchando en puerto ${PORT}`);
  console.log("--versión con excel donDB!");
});

async function agregarAndActualizarUsuariosDb(
  usuarios,
  email,
  sub,
  names,
  foto,
  refresh_Token,
  access_Token
) {
  const filtro = { google_id: sub };

  const nuevo = {
    google_id: sub,
    email: email,
    nombre: names,
    foto: foto,
    accessToken: access_Token,
    refresh_token: refresh_Token,
  };

  const actualizacion = { $set: nuevo };

  const resultado = await usuarios.updateOne(filtro, actualizacion, {
    upsert: true,
  });

  console.log(resultado);
}

async function leerUsuarios(usuarios, sub) {
  console.log("sub: " + sub);
  const usuariosLista = await usuarios.find().toArray();
  const usuarioEncontrado = usuariosLista.find((ele) => ele.google_id === sub);

  return usuarioEncontrado || false;
}

async function refrescarAccessToken(callback) {
  const params = new URLSearchParams();
  params.append(
    "client_id",
    "45594330364-68qsjfc7lo95iq95fvam08hb55oktu4c.apps.googleusercontent.com"
  );
  params.append("client_secret", process.env.MY_CLIENT_SECRET);
  params.append("refresh_token", refreshToken);
  params.append("grant_type", "refresh_token");

  try {
    const tokenRes = await axios.post(
      "https://oauth2.googleapis.com/token",
      params,
      {
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
      }
    );

    const accessToken = tokenRes.data.access_token;

    // Actualiza el access token en la base de datos
    await actualizarTokenEnBD(accessToken);

    if (callback) callback(null, accessToken);
  } catch (error) {
    console.error(
      "Error al refrescar el token:",
      error.response?.data || error
    );
    if (callback) callback(error);
  }
}

async function actualizarTokenEnBD(nuevoToken) {
  await db
    .collection("usuarios")
    .updateOne({ google_id: sub }, { $set: { access_token: nuevoToken } });
}
