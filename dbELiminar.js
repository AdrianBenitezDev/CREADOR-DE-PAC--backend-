let accessToken = null;
let refreshToken = null;

async function inicializarUsuariosTokens() {
  const usuario = await db.collection("usuarios").findOne({ sub: "loquesea" });
  accessToken = usuario.access_token;
  refreshToken = usuario.refresh_token;
}

async function hacerConsultaAApiExterna() {
  try {
    const response = await axios.get(
      "https://api.externaservice.com/endpoint",
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );
    return response.data;
  } catch (error) {
    if (error.response && error.response.status === 401) {
      // Token expiró, pedimos uno nuevo
      const nuevoToken = await refrescarAccessToken();
      accessToken = nuevoToken;
      await actualizarTokenEnBD(nuevoToken);
      // Reintentamos la consulta
      return hacerConsultaAApiExterna();
    } else {
      throw error;
    }
  }
}

async function refrescarAccessToken() {
  const response = await axios.post("https://api.externaservice.com/token", {
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    client_id: "tu_client_id",
    client_secret: "tu_secret",
  });

  return response.data.access_token;
}

async function actualizarTokenEnBD(nuevoToken) {
  await db
    .collection("usuarios")
    .updateOne({ sub: "loquesea" }, { $set: { access_token: nuevoToken } });
}
