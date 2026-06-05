require("dotenv").config();
const express = require("express");
const bodyParser = require("body-parser"); // Leer y entender los datos del formulario
const db = require("./config/db"); // Ahora es el pool
const textdb = require("./config/db"); // Ahora es el pool
const XLSX = require("xlsx");
const app = express(); //Crear una aplicación express
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const sql = require("mssql");
const nodemailer = require("nodemailer");
const session = require("express-session");
const puppeteer = require("puppeteer");
//////////////////////////////////////////////////////////////////
const dayjs = require("dayjs");
const utc = require("dayjs/plugin/utc");
const timezone = require("dayjs/plugin/timezone");
const localizedFormat = require("dayjs/plugin/localizedFormat");
require("dayjs/locale/es");

dayjs.extend(utc);
dayjs.extend(timezone);
dayjs.extend(localizedFormat);
dayjs.locale("es");

const BASE_URL = process.env.BASE_URL || "http://localhost:3000";
const CV_EXTERNAL_PATH =
  process.env.CV_EXTERNAL_PATH || "/Recursos/FrmRep.aspx";

function construirUrlCvExterno(noTrabajador, tipoTrabajador) {
  const baseNormalizada = String(BASE_URL || "").replace(/\/+$/, "");
  const rutaNormalizada = String(CV_EXTERNAL_PATH || "").startsWith("/")
    ? CV_EXTERNAL_PATH
    : `/${CV_EXTERNAL_PATH}`;
  const w = encodeURIComponent(noTrabajador);
  const t = encodeURIComponent(tipoTrabajador);
  return `${baseNormalizada}${rutaNormalizada}?W=${w}&T=${t}`;
}

app.locals.formatearFecha = function (fecha) {
  if (!fecha) return "N/A";

  let year, month, day;

  if (fecha instanceof Date) {
    year = fecha.getUTCFullYear();
    month = fecha.getUTCMonth();
    day = fecha.getUTCDate();
  } else if (typeof fecha === "string") {
    const fechaLimpia = fecha.split("T")[0].split(" ")[0];
    const partes = fechaLimpia.split("-");
    if (partes.length === 3 && partes[0].length === 4) {
      year = parseInt(partes[0], 10);
      month = parseInt(partes[1], 10) - 1;
      day = parseInt(partes[2], 10);
    } else {
      return "N/A";
    }
  } else {
    return "N/A";
  }

  if (Number.isNaN(year) || Number.isNaN(month) || Number.isNaN(day)) {
    return "N/A";
  }

  const nombresMeses = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
  ];

  return `${day} de ${nombresMeses[month]} de ${year}`;
};

const multer = require("multer");
const uploadPath = path.join(__dirname, "uploads");

// Crear carpeta si no existe
if (!fs.existsSync(uploadPath)) {
  fs.mkdirSync(uploadPath);
}

// Configurar multer
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadPath);
  },
  filename: function (req, file, cb) {
    const ext = path.extname(file.originalname);
    const noTrabajador = req.session?.user?.noTrabajador || "desconocido";
    const nombreTrabajador =
      req.session?.user?.nombreTrabajador || "desconocido";
    const nombreLimpio = nombreTrabajador + "_" + file.fieldname + ext;
    cb(null, nombreLimpio);
  },
});

const upload = multer({ storage });
app.use(express.static("public")); // Es donde esta guardado mi CSS, imágenes, etc.
app.use("/uploads", express.static(path.join(__dirname, "uploads"))); // Acceso a documentos cargados
app.use("/reportes", express.static(path.join(__dirname, "reportes"))); // Acceso a reportes generados
app.use(bodyParser.urlencoded({ extended: false })); //decodificar la url de como llegan los datos
app.use(bodyParser.json()); //permitir decodificar archivos json

app.set("view engine", "ejs");

app.use(
  session({
    secret: process.env.SESSION_SECRET || "change_this_secret", // Llave unica de la sesión
    resave: false,
    saveUninitialized: false,
  }),
);

async function obtenerDatosCurriculum(noTrabajador, tipoTrabajador) {
  const [rows] = await db.query(
    "SELECT * FROM trabajador WHERE noTrabajador = ? AND tipoTrabajador = ?",
    [noTrabajador, tipoTrabajador],
  );

  if (!rows || rows.length === 0) {
    return null;
  }

  const trabajador = rows[0];
  const idReal = trabajador.idTrabajador;

  const [escolaridad] = await db.query(
    "SELECT * FROM escolaridad WHERE idTrabajador = ?",
    [idReal],
  );

  const [experienciaPJ] = await db.query(
    `SELECT
        institucion,
        adscripcion,
        cargo,
        campoExperiencia,
        periodoInicio,
        periodoFin
     FROM experienciapj
     WHERE idTrabajador = ?`,
    [idReal],
  );

  const [actualizacion] = await db.query(
    "SELECT * FROM actualizacionprofesional WHERE idTrabajador = ?",
    [idReal],
  );

  return {
    trabajador,
    escolaridad,
    experienciaPJ,
    actualizacion,
  };
}

app.get("/acuse/:noTrabajador/:tipoTrabajador", async (req, res) => {
  const { noTrabajador, tipoTrabajador } = req.params;

  const renderErrorAccesoAcuse = (status, titulo, mensaje, detalle = null) => {
    return res.status(status).render("error-acceso-acuse", {
      titulo,
      mensaje,
      detalle,
      volverUrl: "/index",
      loginUrl: "/login",
      dashboardUrl: "/admin/dashboard",
      esAdmin: !!req.session.admin,
      usuarioSesion: req.session.user || null,
    });
  };

  try {
    // Seguridad: solo el propio trabajador logueado o un administrador pueden ver el acuse
    const esAdmin = !!req.session.admin;
    const usuarioSesion = req.session.user;

    if (!esAdmin) {
      // Si no es admin, debe estar logueado como trabajador
      if (!usuarioSesion) {
        return renderErrorAccesoAcuse(
          401,
          "Sesión expirada",
          "Tu sesión terminó y por eso no se pudo abrir el acuse.",
          "Vuelve a iniciar sesión para continuar.",
        );
      }

      // El trabajador solo puede ver su propio acuse (normalizamos a string para evitar problemas de tipo)
      const noTrabajadorSesion = String(
        usuarioSesion.noTrabajador || "",
      ).trim();
      const noTrabajadorParam = String(noTrabajador || "").trim();
      const tipoSesion = String(usuarioSesion.tipoTrabajador || "")
        .trim()
        .toUpperCase();
      const tipoParam = String(tipoTrabajador || "")
        .trim()
        .toUpperCase();

      if (
        noTrabajadorSesion !== noTrabajadorParam ||
        tipoSesion !== tipoParam
      ) {
        return renderErrorAccesoAcuse(
          403,
          "No tienes permiso para este acuse",
          "El enlace que abriste no coincide con tu cuenta de trabajador.",
          `Estás en sesión como No. ${noTrabajadorSesion || "N/A"} Tipo ${tipoSesion || "N/A"}, pero intentaste abrir No. ${noTrabajadorParam || "N/A"} Tipo ${tipoParam || "N/A"}.`,
        );
      }
    }
    const [rows] = await db.query(
      "SELECT * FROM trabajador WHERE noTrabajador = ? AND tipoTrabajador = ?",
      [noTrabajador, tipoTrabajador],
    );

    if (!rows || rows.length === 0) {
      return renderErrorAccesoAcuse(
        404,
        "No se encontro informacion del acuse",
        "No localizamos un trabajador con los datos del enlace.",
        "Verifica el numero de trabajador y el tipo (C o H).",
      );
    }
    const trabajador = rows[0];
    const idReal = trabajador.idTrabajador;

    const [hijos] = await db.query(
      `SELECT 
          inicialesHijo, 
          sexoHijo, 
          fechaNacimientoHijo, 
          (CONVERT(int,CONVERT(char(8),GETDATE(),112))-CONVERT(int,CONVERT(char(8),fechaNacimientoHijo,112)))/10000 AS edadHijo 
       FROM hijo 
       WHERE idTrabajador = ?`,
      [idReal],
    );

    // C) Escolaridad
    const [escolaridad] = await db.query(
      "SELECT * FROM escolaridad WHERE idTrabajador = ?",
      [idReal],
    );

    // D) Experiencia (Corrección: La columna en BD se llama 'cargo')
    const [experienciaPJ] = await db.query(
      `SELECT 
          institucion, 
          adscripcion, 
          cargo, 
          campoExperiencia, 
          periodoInicio, 
          periodoFin 
       FROM experienciapj 
       WHERE idTrabajador = ?`,
      [idReal],
    );

    // E) Actualización
    const [actualizacion] = await db.query(
      "SELECT * FROM actualizacionprofesional WHERE idTrabajador = ?",
      [idReal],
    );

    // Renderizar la vista (formatearFecha ahora viene de app.locals)
    res.render("acuse", {
      trabajador,
      hijos,
      escolaridad,
      experienciaPJ,
      actualizacion,
      noTrabajador,
    });
  } catch (error) {
    console.error("Error en /acuse:", error);
    return renderErrorAccesoAcuse(
      500,
      "Error al generar el acuse",
      "Ocurrio un problema interno al abrir el acuse.",
      error.message,
    );
  }
});

async function renderCurriculumPrivado(req, res, noTrabajador, tipoTrabajador) {
  const noTrabajadorSafe = String(noTrabajador || "").trim();
  const tipoTrabajadorSafe = String(tipoTrabajador || "")
    .trim()
    .toUpperCase();

  if (!noTrabajadorSafe || !tipoTrabajadorSafe) {
    return res.status(400).send("Consulta invalida: W y T son requeridos.");
  }

  try {
    // Seguridad: solo el propio trabajador logueado o un administrador pueden ver el curriculum
    const esAdmin = !!req.session.admin;
    const usuarioSesion = req.session.user;

    if (!esAdmin) {
      if (!usuarioSesion) {
        return res.redirect("/login");
      }

      const noTrabajadorSesion = String(
        usuarioSesion.noTrabajador || "",
      ).trim();
      const noTrabajadorParam = noTrabajadorSafe;
      const tipoSesion = String(usuarioSesion.tipoTrabajador || "")
        .trim()
        .toUpperCase();
      const tipoParam = tipoTrabajadorSafe;

      if (
        noTrabajadorSesion !== noTrabajadorParam ||
        tipoSesion !== tipoParam
      ) {
        return res
          .status(403)
          .send("No tienes permiso para ver este curriculum.");
      }
    }

    const datos = await obtenerDatosCurriculum(
      noTrabajadorSafe,
      tipoTrabajadorSafe,
    );
    if (!datos) {
      return res.status(404).send("Trabajador no encontrado");
    }

    return res.render("curriculum", {
      trabajador: datos.trabajador,
      escolaridad: datos.escolaridad,
      experienciaPJ: datos.experienciaPJ,
      actualizacion: datos.actualizacion,
    });
  } catch (error) {
    console.error("Error en /curriculum:", error);
    return res
      .status(500)
      .send("Error interno al generar el curriculum: " + error.message);
  }
}

app.get("/curriculum", async (req, res) => {
  const noTrabajador = req.query.W;
  const tipoTrabajador = req.query.T;
  return renderCurriculumPrivado(req, res, noTrabajador, tipoTrabajador);
});

app.get("/curriculum/:noTrabajador/:tipoTrabajador", async (req, res) => {
  const { noTrabajador, tipoTrabajador } = req.params;
  return renderCurriculumPrivado(req, res, noTrabajador, tipoTrabajador);
});

// Enlace publico para visualizar curriculum sin autenticacion
app.get(
  "/curriculum_transparencia/:noTrabajador/:tipoTrabajador",
  async (req, res) => {
    const { noTrabajador, tipoTrabajador } = req.params;

    try {
      const datos = await obtenerDatosCurriculum(noTrabajador, tipoTrabajador);
      if (!datos) {
        return res.status(404).send("Trabajador no encontrado");
      }

      return res.render("curriculum", {
        trabajador: datos.trabajador,
        escolaridad: datos.escolaridad,
        experienciaPJ: datos.experienciaPJ,
        actualizacion: datos.actualizacion,
      });
    } catch (error) {
      console.error("Error en /curriculum_transparencia:", error);
      return res
        .status(500)
        .send("Error interno al generar el curriculum: " + error.message);
    }
  },
);
app.get("/login", (req, res) => {
  if (req.session.user) {
    return res.redirect("/index");
  }
  res.render("login", { error: null });
});

app.get("/index", async (req, res) => {
  if (!req.session.user) {
    return res.redirect("/login");
  }

  // Refresca datos del trabajador desde BD para evitar usar una sesión desactualizada.
  const [trabajadorActualRows] = await db.query(
    "SELECT * FROM trabajador WHERE idTrabajador = ?",
    [req.session.user.idTrabajador],
  );

  if (!trabajadorActualRows || trabajadorActualRows.length === 0) {
    req.session.destroy(() => {});
    return res.redirect("/login");
  }

  const trabajadorActual = trabajadorActualRows[0];
  req.session.user = trabajadorActual;

  const nombreTrabajador = trabajadorActual.nombreTrabajador;
  const uploadDir = path.join(__dirname, "uploads");

  // Función helper para verificar si existe el PDF
  const checkFile = (suffix) =>
    fs.existsSync(path.join(uploadDir, `${nombreTrabajador}_${suffix}.pdf`));

  // Consultar experiencia laboral existente
  const [experienciaPJ] = await db.query(
    "SELECT * FROM experienciapj WHERE idTrabajador = ?",
    [trabajadorActual.idTrabajador],
  );

  // Consultar hijos existentes
  const [hijos] = await db.query("SELECT * FROM hijo WHERE idTrabajador = ?", [
    trabajadorActual.idTrabajador,
  ]);

  // Consultar escolaridad (primaria, secundaria, técnica, bachillerato)
  const [escolaridad] = await db.query(
    "SELECT * FROM escolaridad WHERE idTrabajador = ?",
    [trabajadorActual.idTrabajador],
  );

  // Consultar actualizaciones profesionales existentes
  const [actualizacion] = await db.query(
    "SELECT tema, fecha, institucion, documento, idTrabajador FROM actualizacionprofesional WHERE idTrabajador = ? ORDER BY fecha DESC",
    [trabajadorActual.idTrabajador],
  );

  // Filtrar por nivel académico
  const tecnica =
    escolaridad.filter(
      (e) =>
        e.nivelAcademico && e.nivelAcademico.toUpperCase().includes("TECNICA"),
    )[0] || null;

  const bachillerato =
    escolaridad.filter(
      (e) =>
        e.nivelAcademico && e.nivelAcademico.toUpperCase() === "BACHILLERATO",
    )[0] || null;

  // Detectar si ya tiene registrada primaria y secundaria en la base de datos
  const tienePrimaria = escolaridad.some(
    (e) =>
      e.nivelAcademico && e.nivelAcademico.toUpperCase().includes("PRIMARIA"),
  );

  const tieneSecundaria = escolaridad.some(
    (e) =>
      e.nivelAcademico && e.nivelAcademico.toUpperCase().includes("SECUNDARIA"),
  );

  // Obtener datos de enfermedad crónica del usuario logueado
  const tieneEnfermedadCronica =
    trabajadorActual.tieneEnfermedadCronica === "Si" ? "1" : "0";
  const tipoCronica = trabajadorActual.tipoCronica || "";

  res.render("index", {
    user: trabajadorActual,
    experienciaPJ,
    hijos,
    escolaridad,
    actualizacion,
    tecnica,
    bachillerato,
    tienePrimaria,
    tieneSecundaria,
    tieneEnfermedadCronica,
    tipoCronica,
    existeIne: checkFile("INE"),
    existeActa: checkFile("ACTA-NACIMIENTO"),
    existeDomicilio: checkFile("COMPROBANTE-DOMICILIO"),
    existeRfc: checkFile("RFC"),
    existeGrado: checkFile("GRADO-ESTUDIOS"),
  });
});

app.get("/success", (req, res) => {
  res.render("success");
});

app.get("/buscador", (req, res) => {
  res.render("buscador"); // si usas EJS
});

app.get("/existe/:noTrabajador/:tipoTrabajador", async (req, res) => {
  const { noTrabajador, tipoTrabajador } = req.params;

  try {
    const [result] = await db.query(
      "SELECT COUNT(*) AS total FROM trabajador WHERE noTrabajador = ? AND tipoTrabajador = ?",
      [noTrabajador, tipoTrabajador],
    );
    const existe = result[0].total > 0;
    res.json({ existe });
  } catch (error) {
    console.error("Error al verificar trabajador:", error);
    res.status(500).json({ error: "Error interno" });
  }
});

app.get("/registrar-password", (req, res) => {
  res.render("registrarPassword", { error: null });
});

app.get("/insert", (req, res) => {
  res.render("insert", { error: null });
});

app.get("/", (req, res) => {
  res.redirect("/login");
});

app.get("/logout", (req, res) => {
  if (req.session) {
    req.session.destroy((err) => {
      if (err) {
        console.error("Error al destruir sesión:", err);
      }
      res.redirect("/login");
    });
  } else {
    res.redirect("/login");
  }
});

app.post("/eliminar-pdf/:tipo", (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: "No autorizado" });
  }
  const tiposValidos = ["INE", "ACTA-NACIMIENTO", "COMPROBANTE-DOMICILIO", "RFC", "GRADO-ESTUDIOS"];
  const tipo = req.params.tipo;
  if (!tiposValidos.includes(tipo)) {
    return res.status(400).json({ success: false, message: "Tipo de documento no válido" });
  }
  const nombreTrabajador = req.session.user.nombreTrabajador;
  const filePath = path.join(__dirname, "uploads", `${nombreTrabajador}_${tipo}.pdf`);
  try {
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, message: "Error al eliminar el archivo" });
  }
});
function calcularEdad(fecha) {
  return fecha ? dayjs().diff(fecha, "year") : 0;
}
function convertirFecha(fechaDDMMYYYY) {
  // 1. Validar si está vacío, es undefined, null o son ceros
  if (
    !fechaDDMMYYYY ||
    fechaDDMMYYYY.trim() === "" ||
    fechaDDMMYYYY === "00-00-0000"
  ) {
    return null; // <--- ESTO ES LA CLAVE: Enviar NULL a SQL Server
  }

  const partes = fechaDDMMYYYY.split("-");

  // 2. Validación extra por si el formato viene mal
  if (partes.length !== 3) return null;

  const dia = partes[0];
  const mes = partes[1];
  const anio = partes[2];

  return `${anio}-${mes}-${dia}`; // Formato correcto YYYY-MM-DD
}

function formatearFecha(fecha) {
  if (!fecha || fecha === "0000-00-00") {
    return "N/A";
  }

  const d = new Date(fecha);
  if (Number.isNaN(d.getTime())) {
    return "N/A";
  }

  const opciones = { day: "2-digit", month: "long", year: "numeric", timeZone: "UTC" };
  const fechaFormateada = d.toLocaleDateString("es-MX", opciones);

  return fechaFormateada.charAt(0).toUpperCase() + fechaFormateada.slice(1);
}

function convertirFechaTexto(fechaInput) {
  if (!fechaInput) return null;

  let anio, mes;
  const partes = fechaInput.split("-");

  // Caso 1: Formato YYYY-MM (El que envía tu formulario nuevo)
  if (partes[0].length === 4) {
    anio = partes[0];
    mes = partes[1];
  }
  // Caso 2: Formato MM-YYYY (Por si acaso llega al revés)
  else {
    mes = partes[0];
    anio = partes[1];
  }

  const nombresMeses = [
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE",
  ];

  const mesIndex = parseInt(mes, 10) - 1;

  if (mesIndex >= 0 && mesIndex < 12) {
    return `${nombresMeses[mesIndex]} ${anio}`;
  }

  return null;
}

function mesAnioAFechaOrdenable(fechaMMYYYY) {
  if (!fechaMMYYYY) return null;
  const [mes, anio] = fechaMMYYYY.split("-");
  return new Date(`${anio}-${mes}-01`); // YYYY-MM-DD
}
// Ruta para guardar datos

function validarFechas(fecha) {
  // 1️⃣ Si está vacía, nula o indefinida
  if (!fecha || fecha.trim() === "") return false;

  // 2️⃣ Intentar crear un objeto Date (maneja YYYY-MM-DD o DD/MM/YYYY)
  let partes;

  if (fecha.includes("-")) {
    partes = fecha.split("-");
  } else if (fecha.includes("/")) {
    partes = fecha.split("/");
  } else {
    return false;
  }

  let anio, mes, dia;

  // Detectar formato (si es YYYY-MM-DD o DD-MM-YYYY)
  if (partes[0].length === 4) {
    // YYYY-MM-DD
    [anio, mes, dia] = partes;
  } else {
    // DD-MM-YYYY o DD/MM/YYYY
    [dia, mes, anio] = partes;
  }

  // 3️⃣ Validar números y rangos
  anio = parseInt(anio, 10);
  mes = parseInt(mes, 10);
  dia = parseInt(dia, 10);

  if (!anio || !mes || !dia || mes < 1 || mes > 12 || dia < 1 || dia > 31)
    return false;

  // 4️⃣ Crear objeto Date y validar que coincida
  const fechaObj = new Date(anio, mes - 1, dia);
  if (isNaN(fechaObj.getTime())) return false;

  // 5️⃣ (Opcional) No permitir fechas futuras
  const hoy = new Date();
  if (fechaObj > hoy) return false;

  return true;
}

app.post(
  "/guardar",
  upload.fields([
    { name: "INE", maxCount: 1 },
    { name: "ACTA-NACIMIENTO", maxCount: 1 },
    { name: "COMPROBANTE-DOMICILIO", maxCount: 1 },
    { name: "RFC", maxCount: 1 },
    { name: "GRADO-ESTUDIOS", maxCount: 1 },
  ]),
  async (req, res) => {
    if (!req.session.user) {
      return res.redirect("/login");
    }

    const ineFile = req.files["ine_pdf"]?.[0];
    const actaFile = req.files["acta_pdf"]?.[0];
    const domicilioFile = req.files["comprobante_domicilio_pdf"]?.[0];
    const fiscalFile = req.files["comprobante_fiscal_pdf"]?.[0];

    const {
      telefono,
      fechaIngreso,
      adscripcionActual,
      lenguaSenias,
      sexo,
      fechaNacimiento,
      municipioNacimiento,
      estadoNacimiento,
      codigopostal,
      correo,
      papa,
      embarazo,
      cantidadHijos,
      cronica,
      civil,
      sangre,
      telefonocasa,
      telefonofamiliar,
      parentesco,
      primaria,
      secundaria_institucion,
    } = req.body;
    const { idTrabajador, noTrabajador } = req.session.user;

    const lugarNacimiento = `${municipioNacimiento.trim().toUpperCase()}, ${estadoNacimiento.trim().toUpperCase()}`;

    let cargoActual = req.body.cargoActual;

    if (cargoActual === "Otro") {
      cargoActual = req.body.otroCargo?.trim().toUpperCase() || "OTRO";
    }

    const callenumero = req.body.callenumero.toUpperCase();
    const colonia = req.body.colonia.toUpperCase();
    const municipio = req.body.municipio.toUpperCase();
    const estado = req.body.estado.toUpperCase();

    let actualCallenumero = req.body.actual_callenumero || callenumero;
    let actualColonia = req.body.actual_colonia || colonia;
    let actualMunicipio = req.body.actual_municipio || municipio;
    let actualEstado = req.body.actual_estado || estado;
    let actualCP = req.body.actual_codigopostal || codigopostal;

    // console.log("Papa: " + papa + " y " + sexo); // Log de apoyo para validar bandera de padre/madre
    let cronicaTexto = req.body.cronicaTexto;
    let nombreConyuge = req.body.nombreConyuge;
    let fechaConyuge = req.body.fechaConyuge;
    let sexoConyuge = req.body.sexoConyuge;

    let tieneCronica;
    let padre;
    let embarazada;

    let comunidadIndigena =
      req.body.indigena === "1"
        ? req.body.comunidadIndigena?.trim().toUpperCase() || "N/A"
        : "N/A";

    if (comunidadIndigena === "OTRO") {
      // Manejo seguro por si llega como array (duplicado) o undefined
      const otraComunidadVal = Array.isArray(req.body.otraComunidad)
        ? req.body.otraComunidad[0]
        : req.body.otraComunidad;
      comunidadIndigena =
        String(otraComunidadVal || "")
          .trim()
          .toUpperCase() || "OTRO";
    }

    let familiaLinguistica =
      req.body.hablaLenguaIndigena === "1"
        ? req.body.familiaLinguistica?.trim().toUpperCase() || "N/A"
        : "N/A";

    if (familiaLinguistica === "OTRA") {
      // Manejo seguro por si llega como array
      const otraLenguaVal = Array.isArray(req.body.otraLengua)
        ? req.body.otraLengua[0]
        : req.body.otraLengua;
      familiaLinguistica =
        String(otraLenguaVal || "")
          .trim()
          .toUpperCase() || "OTRO";
    }

    let tipoDiscapacidad =
      req.body.discapacidad === "1"
        ? req.body.tipoDiscapacidad?.trim().toUpperCase() || "NO"
        : "No";

    let discapacidadEspecificada = req.body.discapacidadTexto;

    if (req.body.discapacidad === "0") {
      discapacidadEspecificada = "N/A";
    }
    const vigenciaINE = req.body.vigenciaINE;
    const tieneTecnica = req.body.tieneTecnica === "1";
    const tecnicaTitulo = (
      req.body.tecnica_titulo ||
      req.body.tecnica_institucion ||
      ""
    )
      .trim()
      .toUpperCase();
    const tecnicaInstitucion = (req.body.tecnica_institucion || "")
      .trim()
      .toUpperCase();
    const tecnicaFecha =
      tieneTecnica && req.body.tecnica_fecha
        ? convertirFechaTexto(req.body.tecnica_fecha)
        : null;
    const tecnicaDocumento = tieneTecnica
      ? req.body.tecnica_documento || null
      : null;
    const tecnicaEstatus = tieneTecnica
      ? req.body.tecnica_estatus || null
      : null;
    const tecnicaCedula = tieneTecnica ? req.body.tecnica_cedula || null : null;
    const carreraTecnicaComercial =
      tieneTecnica && tecnicaTitulo ? tecnicaTitulo : "No";

    let cantidadHijosSegura = parseInt(cantidadHijos, 10);
    if (Number.isNaN(cantidadHijosSegura) || cantidadHijosSegura < 0) {
      cantidadHijosSegura = 0;
    }
    cantidadHijosSegura = Math.min(10, cantidadHijosSegura);

    const telefonoCasaLimpio = (telefonocasa || "").trim();

    const Bachillerato = req.body.tieneBachillerato === "1" ? "Si" : "No";
    const Licenciatura = req.body.tieneLicenciatura === "1" ? "Si" : "No";
    const Maestria = req.body.tieneMaestria === "1" ? "Si" : "No";
    const Doctorado = req.body.tieneDoctorado === "1" ? "Si" : "No";
    const Posdoctorado = req.body.tienePosdoctorado === "1" ? "Si" : "No";
    const Especialidad = req.body.tieneEspecialidad === "1" ? "Si" : "No";
    
    const prepaTitulo = Bachillerato === "Si"
      ? req.body.prepa_titulo.trim().toUpperCase()
      : null;
    const prepaFecha = Bachillerato === "Si"
      ? convertirFechaTexto(req.body.prepa_fecha)
      : null;
    const prepaInst = Bachillerato === "Si"
      ? req.body.prepa_institucion.trim().toUpperCase()
      : null;
    const prepaDocumento = Bachillerato === "Si" ? req.body.prepa_documento : null;
    const prepaEstatus = Bachillerato === "Si" ? req.body.prepa_estatus : null;

    let dominioLSM = req.body.nivelLSM;

    if (lenguaSenias === "No") {
      dominioLSM = "N/A";
    }

    const query =
      "UPDATE trabajador SET fechaIngreso=?, adscripcionActual=?,cargoActual=?,sexoTrabajador=?,fechaNacimiento=?,lugarNacimiento=?,estadoCivil=?,nombreConyuge=?,fechaNacimientoConyuge=?,sexoConyuge=?,tipoSangre=?,vigenciaINE=?,calleNumero=?,colonia=?,municipio=?,estado=?,codigoPostal=?,actualCalleNumero=?,actualColonia=?,actualMunicipio=?,actualEstado=?,actualCP=?,noTelefono=?,telefonoCasa=?,telefonoFamiliar=?,parentescoFamiliar=?,correoElectronico=?,comunidadIndigena=?,familiaLinguistica=?,tipoDiscapacidad=?,discapacidadEspecificada=?,lenguajeSenias=?,dominioLSM=?,tieneEnfermedadCronica=?,tipoCronica=?,es_padre_madre=?,embarazada=?,cantidadHijos=?,primaria=?,secundaria=?,carreraTecnicaComercial=?,Licenciatura=?,Maestria=?,Doctorado=?,Posdoctorado=?,Especialidad=?,Bachillerato=? WHERE idTrabajador = ?";

    if (papa == 1 && sexo == "Femenino") {
      padre = "Madre";
    } else if (papa == 1 && sexo == "Masculino") {
      padre = "Padre";
    } else {
      padre = "No";
    }

    if (cronica == 1) {
      tieneCronica = "Si";
    } else {
      tieneCronica = "No";
    }

    if (cronica == 0) {
      cronicaTexto = "N/A";
    }

    if (embarazo == 1 && sexo == "Femenino") {
      embarazada = "Embarazada";
    } else if (sexo == "Masculino") {
      embarazada = "No aplica";
    } else {
      embarazada = "No";
    }

    if (civil != "Casado/a") {
      nombreConyuge = "N/A";
      fechaConyuge = "00-00-0000";
      sexoConyuge = "N/A";
    }

    try {
      const fechaFormateada = convertirFecha(fechaConyuge);
      const fechaIngresoFormateada = convertirFecha(fechaIngreso);
      const fechaNacimientoFormateada = convertirFecha(fechaNacimiento);
      // Logs de depuración de fechas (mantener comentados para pruebas puntuales)
      // console.log("nacimiento Conyuge" + fechaFormateada);
      // console.log("Fecha de ingreso" + fechaIngresoFormateada);
      // console.log("Fecha de nacimiento" + fechaNacimientoFormateada);

      const [result] = await db.query(query, [
        fechaIngresoFormateada,
        adscripcionActual,
        cargoActual,
        sexo,
        fechaNacimientoFormateada,
        lugarNacimiento,
        civil,
        nombreConyuge,
        fechaFormateada,
        sexoConyuge,
        sangre,
        vigenciaINE,
        callenumero,
        colonia,
        municipio,
        estado,
        codigopostal,
        actualCallenumero,
        actualColonia,
        actualMunicipio,
        actualEstado,
        actualCP,
        telefono,
        telefonoCasaLimpio || null,
        telefonofamiliar,
        parentesco,
        correo,
        comunidadIndigena,
        familiaLinguistica,
        tipoDiscapacidad,
        discapacidadEspecificada,
        lenguaSenias,
        dominioLSM,
        tieneCronica,
        cronicaTexto,
        padre,
        embarazada,
        cantidadHijosSegura,
        primaria,
        secundaria_institucion,
        carreraTecnicaComercial,
        Licenciatura,
        Maestria,
        Doctorado,
        Posdoctorado,
        Especialidad,
        Bachillerato,
        idTrabajador,
      ]);
      // console.log("✅ Actualización exitosa:", result); // Log de seguimiento general de actualización

      const deleteQuery = "DELETE FROM hijo WHERE idTrabajador = ?";
      await db.query(deleteQuery, [idTrabajador]);

      // ... después de borrar los hijos anteriores ...

      if (cantidadHijosSegura > 0) {
        for (let i = 0; i < cantidadHijosSegura; i++) {
          let iniciales = req.body[`inicialesHijo${i}`];
          const sexo = req.body[`selectHijo${i}`];
          const fechaNacimiento = req.body[`fechaNacimientoHijo${i}`];

          if (iniciales && fechaNacimiento) {
            // Pequeña validación extra
            iniciales = iniciales.toUpperCase();

            // 1. Convertir la fecha para SQL
            const fechaFormateada = convertirFecha(fechaNacimiento);

            // 2. CALCULAR LA EDAD (¡Esta es la parte clave!)
            // Usamos la función optimizada 'calcularEdad' que definimos antes
            const edad = calcularEdad(fechaFormateada);

            // 3. Insertar en la BD incluyendo la columna 'edadHijo'
            // Asegúrate de agregar 'edadHijo' en el INSERT y un '?' extra en VALUES
            const queryChild =
              "INSERT INTO hijo (inicialesHijo, sexoHijo, fechaNacimientoHijo, edadHijo, idTrabajador) VALUES (?, ?, ?, ?, ?)";

            await db.query(queryChild, [
              iniciales,
              sexo,
              fechaFormateada,
              edad, // <--- Aquí enviamos la edad calculada
              idTrabajador,
            ]);
            // console.log(`✅ Hijo ${i + 1} guardado con edad: ${edad}`); // Útil solo en pruebas de registro de hijos
          }
        }
      }
      const deleteQuery2 = "DELETE FROM escolaridad WHERE idTrabajador = ?";
      await db.query(deleteQuery2, [idTrabajador]);

      await db.query("DELETE FROM escolaridad WHERE idTrabajador = ?", [
        idTrabajador,
      ]);

      // NOTA: El INSERT de escolaridad se realiza después con el array escolaridades
      // para evitar duplicados. Ver línea ~900 donde se procesa el array completo.

      // console.log(`Escolaridad eliminada para trabajador ${idTrabajador}`); // Log informativo, mantener comentado

      const escolaridades = [
        // CARRERA TÉCNICA O COMERCIAL
        {
          nivel: "TECNICA",
          titulo: tieneTecnica ? tecnicaTitulo || null : null,
          fecha: tecnicaFecha,
          institucion: tieneTecnica ? tecnicaInstitucion || null : null,
          documento: tecnicaDocumento,
          estatus: tecnicaEstatus,
          cedula: tecnicaCedula,
        },

        // BACHILLERATO
        {
          nivel: "BACHILLERATO",
          titulo: req.body.prepa_titulo || null,
          fecha: req.body.prepa_fecha
            ? convertirFechaTexto(req.body.prepa_fecha)
            : null,
          institucion: req.body.prepa_institucion || null,
          documento: req.body.prepa_documento || null,
          estatus: req.body.prepa_estatus || null,
          cedula: null,
        },

        // LICENCIATURAS (hasta 3)
        {
          nivel: "LICENCIATURA",
          titulo: req.body.licenciatura_titulo1 || null,
          fecha: req.body.licenciatura_fecha1
            ? convertirFechaTexto(req.body.licenciatura_fecha1)
            : null,
          institucion: req.body.licenciatura_institucion1 || null,
          documento: req.body.licenciatura_documento1 || null,
          estatus: req.body.licenciatura_estatus1 || null,
          cedula: req.body.licenciatura_cedula1 || null,
        },
        {
          nivel: "LICENCIATURA",
          titulo: req.body.licenciatura_titulo2 || null,
          fecha: req.body.licenciatura_fecha2
            ? convertirFechaTexto(req.body.licenciatura_fecha2)
            : null,
          institucion: req.body.licenciatura_institucion2 || null,
          documento: req.body.licenciatura_documento2 || null,
          estatus: req.body.licenciatura_estatus2 || null,
          cedula: req.body.licenciatura_cedula2 || null,
        },
        {
          nivel: "LICENCIATURA",
          titulo: req.body.licenciatura_titulo3 || null,
          fecha: req.body.licenciatura_fecha3
            ? convertirFechaTexto(req.body.licenciatura_fecha3)
            : null,
          institucion: req.body.licenciatura_institucion3 || null,
          documento: req.body.licenciatura_documento3 || null,
          estatus: req.body.licenciatura_estatus3 || null,
          cedula: req.body.licenciatura_cedula3 || null,
        },

        // MAESTRÍAS
        {
          nivel: "MAESTRIA",
          titulo: req.body.maestria_titulo1 || null,
          fecha: req.body.maestria_fecha1
            ? convertirFechaTexto(req.body.maestria_fecha1)
            : null,
          institucion: req.body.maestria_institucion1 || null,
          documento: req.body.maestria_documento1 || null,
          estatus: req.body.maestria_estatus1 || null,
          cedula: req.body.maestria_cedula1 || null,
        },
        {
          nivel: "MAESTRIA",
          titulo: req.body.maestria_titulo2 || null,
          fecha: req.body.maestria_fecha2
            ? convertirFechaTexto(req.body.maestria_fecha2)
            : null,
          institucion: req.body.maestria_institucion2 || null,
          documento: req.body.maestria_documento2 || null,
          estatus: req.body.maestria_estatus2 || null,
          cedula: req.body.maestria_cedula2 || null,
        },
        {
          nivel: "MAESTRIA",
          titulo: req.body.maestria_titulo3 || null,
          fecha: req.body.maestria_fecha3
            ? convertirFechaTexto(req.body.maestria_fecha3)
            : null,
          institucion: req.body.maestria_institucion3 || null,
          documento: req.body.maestria_documento3 || null,
          estatus: req.body.maestria_estatus3 || null,
          cedula: req.body.maestria_cedula3 || null,
        },

        // ESPECIALIDADES
        {
          nivel: "ESPECIALIDAD",
          titulo: req.body.especialidad_titulo1 || null,
          fecha: req.body.especialidad_fecha1
            ? convertirFechaTexto(req.body.especialidad_fecha1)
            : null,
          institucion: req.body.especialidad_institucion1 || null,
          documento: req.body.especialidad_documento1 || null,
          estatus: req.body.especialidad_estatus1 || null,
          cedula: req.body.especialidad_cedula1 || null,
        },
        {
          nivel: "ESPECIALIDAD",
          titulo: req.body.especialidad_titulo2 || null,
          fecha: req.body.especialidad_fecha2
            ? convertirFechaTexto(req.body.especialidad_fecha2)
            : null,
          institucion: req.body.especialidad_institucion2 || null,
          documento: req.body.especialidad_documento2 || null,
          estatus: req.body.especialidad_estatus2 || null,
          cedula: req.body.especialidad_cedula2 || null,
        },
        {
          nivel: "ESPECIALIDAD",
          titulo: req.body.especialidad_titulo3 || null,
          fecha: req.body.especialidad_fecha3
            ? convertirFechaTexto(req.body.especialidad_fecha3)
            : null,
          institucion: req.body.especialidad_institucion3 || null,
          documento: req.body.especialidad_documento3 || null,
          estatus: req.body.especialidad_estatus3 || null,
          cedula: req.body.especialidad_cedula3 || null,
        },

        // DOCTORADOS
        {
          nivel: "DOCTORADO",
          titulo: req.body.doctorado_titulo1 || null,
          fecha: req.body.doctorado_fecha1
            ? convertirFechaTexto(req.body.doctorado_fecha1)
            : null,
          institucion: req.body.doctorado_institucion1 || null,
          documento: req.body.doctorado_documento1 || null,
          estatus: req.body.doctorado_estatus1 || null,
          cedula: req.body.doctorado_cedula1 || null,
        },
        {
          nivel: "DOCTORADO",
          titulo: req.body.doctorado_titulo2 || null,
          fecha: req.body.doctorado_fecha2
            ? convertirFechaTexto(req.body.doctorado_fecha2)
            : null,
          institucion: req.body.doctorado_institucion2 || null,
          documento: req.body.doctorado_documento2 || null,
          estatus: req.body.doctorado_estatus2 || null,
          cedula: req.body.doctorado_cedula2 || null,
        },

        // POSDOCTORADOS
        {
          nivel: "POSDOCTORADO",
          titulo: req.body.posdoctorado_titulo1 || null,
          fecha: req.body.posdoctorado_fecha1
            ? convertirFechaTexto(req.body.posdoctorado_fecha1)
            : null,
          institucion: req.body.posdoctorado_institucion1 || null,
          documento: req.body.posdoctorado_documento1 || null,
          estatus: req.body.posdoctorado_estatus1 || null,
          cedula: req.body.posdoctorado_cedula1 || null,
        },
        {
          nivel: "POSDOCTORADO",
          titulo: req.body.posdoctorado_titulo2 || null,
          fecha: req.body.posdoctorado_fecha2
            ? convertirFechaTexto(req.body.posdoctorado_fecha2)
            : null,
          institucion: req.body.posdoctorado_institucion2 || null,
          documento: req.body.posdoctorado_documento2 || null,
          estatus: req.body.posdoctorado_estatus2 || null,
          cedula: req.body.posdoctorado_cedula2 || null,
        },
      ];

      // Recorremos cada escolaridad y guardamos solo si tienen algún dato real
      for (let escolaridad of escolaridades) {
        // console.log(escolaridad.fecha); // Depuración de fechas de escolaridad
        const hayDatos =
          escolaridad.titulo || escolaridad.fecha || escolaridad.institucion;

        if (hayDatos) {
          const insertQuery = `
            INSERT INTO escolaridad (
                nivelAcademico, nombreTitulo, fechaObtencion,
                institucion, documentoAdquirido, estatus, cedulaProfesional,idTrabajador
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            `;

          await db.query(insertQuery, [
            // Asegúrate de tener el idTrabajador después de insertar en trabajador
            escolaridad.nivel,
            escolaridad.titulo,
            escolaridad.fecha,
            escolaridad.institucion,
            escolaridad.documento,
            escolaridad.estatus,
            escolaridad.cedula,
            idTrabajador,
          ]);

          // console.log(`✅ Escolaridad ${escolaridad.nivel} guardada`); // Log de seguimiento por registro
        }
      }

      const deleteQuery4 =
        "DELETE FROM actualizacionprofesional WHERE idTrabajador = ?";
      await db.query(deleteQuery4, [idTrabajador]);

      let actualizaciones = [];
      for (let i = 0; i < 11; i++) {
        const tema = req.body[`actualizacion_tema${i}`];
        const fecha = req.body[`actualizacion_fecha${i}`];
        const institucion = req.body[`actualizacion_institucion${i}`];
        const documento = req.body[`actualizacion_documento${i}`];

        const hayDatos = tema || fecha || institucion;

        if (hayDatos) {
          actualizaciones.push({ tema, fecha, institucion, documento });
        }
      }

      // Paso 2: ordenar por fecha descendente
      actualizaciones.sort((a, b) => {
        const fa = new Date(convertirFecha(a.fecha));
        const fb = new Date(convertirFecha(b.fecha));
        return fb - fa; // descendente
      });

      // Paso 3: insertar ya ordenado
      for (let act of actualizaciones) {
        const fechaConvertida = act.fecha ? convertirFecha(act.fecha) : null;

        await db.query(
          `INSERT INTO actualizacionprofesional (tema, fecha, institucion, documento, idTrabajador)
            VALUES (?, ?, ?, ?, ?)`,
          [
            act.tema,
            fechaConvertida,
            act.institucion,
            act.documento,
            idTrabajador,
          ],
        );
      }

      const deleteQuery3 = "DELETE FROM experienciapj WHERE idTrabajador = ?";
      await db.query(deleteQuery3, [idTrabajador]);

      cantidadExpPJ = req.body["experienciaTotal"];

      if (cantidadExpPJ && cantidadExpPJ > 0) {
        let experiencias = [];

        for (let i = 0; i < cantidadExpPJ; i++) {
          const rawInicio = req.body[`experiencia_inicio${i}`]; // ej. "05-2023"
          const rawFin = req.body[`experiencia_fin${i}`]; // ej. "08-2024"
          let institucion = req.body[`experiencia_institucion${i}`];
          let adscripcion = req.body[`experiencia_adscripcion${i}`];
          const puesto = req.body[`experiencia_puesto${i}`];
          const campo = req.body[`experiencia_campo${i}`];

          const periodoInicio = convertirFechaTexto(rawInicio);
          const periodoFin = rawFin
            ? convertirFechaTexto(rawFin)
            : "A LA FECHA";

          if (institucion === "OTRA") {
            adscripcion = "";
            institucion = req.body[`experiencia_otrainst${i}`];
          } else {
            institucion = "PODER JUDICIAL DEL ESTADO DE HIDALGO";
          }

          // Convertimos fechas para ordenar correctamente
          const fechaOrdenable = (mesY) => {
            if (!mesY || typeof mesY !== "string") return new Date(0); // valor mínimo
            const [mes, anio] = mesY.split("-");
            return new Date(parseInt(anio), parseInt(mes) - 1);
          };

          const finOrden = rawFin
            ? fechaOrdenable(rawFin)
            : new Date(9999, 11, 31);

          experiencias.push({
            periodoInicio,
            periodoFin,
            inicioOrden: fechaOrdenable(rawInicio),
            finOrden,
            institucion,
            adscripcion,
            puesto,
            campo,
          });
        }

        // Ordenar por fecha de fin descendente
        experiencias.sort((a, b) => b.finOrden - a.finOrden);

        for (let exp of experiencias) {
          await db.query(
            `INSERT INTO experienciapj (periodoInicio, periodoFin, institucion, adscripcion, cargo, campoExperiencia, idTrabajador)
                    VALUES (?, ?, ?, ?, ?, ?, ?)`,
            [
              exp.periodoInicio,
              exp.periodoFin,
              exp.institucion,
              exp.adscripcion,
              exp.puesto,
              exp.campo,
              idTrabajador,
            ],
          );
          // console.log(`✅ Experiencia insertada: ${exp.puesto}`); // Solo necesario en pruebas de carga de experiencia
        }
      }

      const idTrab = idTrabajador;
      req.session.idTrabajador = idTrab;
      const tipoTrabajadorSesion = req.session.user?.tipoTrabajador || "C";
      res.redirect(`/acuse/${noTrabajador}/${tipoTrabajadorSesion}`);
    } catch (err) {
      console.error("❌ Error al guardar:", err);
      res.send("Error al guardar");
    }
  },
);

app.post("/login", async (req, res) => {
  const curp = req.body.curp?.toUpperCase().trim();
  const rfc = req.body.rfc?.toUpperCase().trim();
  const password = req.body.password;

  // Validar campos vacíos
  if (!curp && !rfc) {
    return res.status(400).render("login", {
      error: "Debes ingresar tu CURP o RFC para continuar.",
      curp: "",
      rfc: "",
    });
  }

  if (!curp) {
    return res.status(400).render("login", {
      error: "El campo CURP está vacío. Ingrésalo para continuar.",
      curp: "",
      rfc,
    });
  }

  if (!rfc) {
    return res.status(400).render("login", {
      error: "El campo RFC está vacío. Ingrésalo para continuar.",
      curp,
      rfc: "",
    });
  }

  if (!password || password.trim() === "") {
    return res.status(400).render("login", {
      error: "Ingresa tu contraseña para continuar.",
      curp,
      rfc,
    });
  }

  try {
    // Buscar por CURP
    const [curpResults] = await db.query(
      "SELECT * FROM trabajador WHERE curp = ?",
      [curp]
    );

    // Buscar por RFC
    const [rfcResults] = await db.query(
      "SELECT * FROM trabajador WHERE rfc = ?",
      [rfc]
    );

    // Validar si CURP no existe
    if (curpResults.length === 0 && rfcResults.length === 0) {
      return res.status(401).render("login", {
        error: "CURP y RFC no encontrados. Verifica que ambos datos sean correctos.",
        curp,
        rfc,
      });
    }

    if (curpResults.length === 0) {
      return res.status(401).render("login", {
        error: "CURP no encontrado. Verifica que esté escrito correctamente.",
        curp,
        rfc,
      });
    }

    if (rfcResults.length === 0) {
      return res.status(401).render("login", {
        error: "RFC no encontrado. Verifica que esté escrito correctamente.",
        curp,
        rfc,
      });
    }

    // Buscar coincidencia exacta de CURP y RFC
    const [results] = await db.query(
      "SELECT * FROM trabajador WHERE curp = ? AND rfc = ?",
      [curp, rfc]
    );

    if (results.length === 0) {
      return res.status(401).render("login", {
        error: "CURP y RFC no coinciden. Verifica que ambos datos pertenezcan a la misma persona.",
        curp,
        rfc,
      });
    }

    const usuario = results[0];

    // Validar si el usuario tiene contraseña registrada
    if (!usuario.password || usuario.password === "") {
      return res.status(401).render("login", {
        error: "Este usuario no tiene contraseña registrada. Contacta al administrador.",
        curp,
        rfc,
      });
    }

    if (usuario.password !== password) {
      return res.status(401).render("login", {
        error: "Contraseña incorrecta. Verifica e intenta nuevamente.",
        curp,
        rfc,
      });
    }

    // 🔹 Destruir sesión previa antes de asignar nueva
    req.session.regenerate((err) => {
      if (err) {
        console.error("❌ Error regenerando sesión:", err);
        return res.status(500).render("login", {
          error: "Error al iniciar sesión. Intenta nuevamente.",
          curp,
          rfc,
        });
      }

      // Asignar usuario a la nueva sesión
      req.session.user = usuario;

      // Redirigir al index
      return res.redirect("/index");
    });
  } catch (err) {
    console.error("❌ Error en el query:", err);
    return res.status(500).render("login", {
      error: "Ocurrió un error al consultar la base de datos. Intenta nuevamente más tarde.",
      curp,
      rfc,
    });
  }
});

app.get("/documentos/:idTrabajador", async (req, res) => {
  try {
    const { idTrabajador } = req.params;
    const [trabajadorRows] = await db.query(
      "SELECT * FROM trabajador WHERE idTrabajador = ?",
      [idTrabajador]
    );

    if (!trabajadorRows || trabajadorRows.length === 0) {
      return res.status(404).send("Trabajador no encontrado");
    }

    const trabajador = trabajadorRows[0];
    const uploadsPath = path.join(__dirname, "uploads");
    const reportesPath = path.join(__dirname, "reportes");

    const normalizarNombre = (valor) =>
      String(valor || "")
        .trim()
        .toUpperCase()
        .replace(/\s+/g, " ")
        .split(" ")
        .filter(Boolean)
        .sort()
        .join(" ");

    const nombreTrabajadorNormalizado = normalizarNombre(trabajador.nombreTrabajador);

    const documentos = [];

    if (fs.existsSync(reportesPath)) {
      const archivos = fs.readdirSync(reportesPath);
      archivos.forEach((archivo) => {
        if (archivo.includes(trabajador.nombreTrabajador)) {
          const ruta = path.join(reportesPath, archivo);
          const stat = fs.statSync(ruta);
          documentos.push({
            nombre: archivo,
            tipo: archivo.endsWith(".pdf") ? "pdf" : "excel",
            ruta: `/reportes/${archivo}`,
            tamano: (stat.size / 1024).toFixed(2) + " KB",
            fecha: stat.mtime.toLocaleDateString("es-MX")
          });
        }
      });
    }

    if (fs.existsSync(uploadsPath)) {
      const archivos = fs.readdirSync(uploadsPath);
      archivos.forEach((archivo) => {
        const ruta = path.join(uploadsPath, archivo);
        const stat = fs.statSync(ruta);
        if (!stat.isFile()) return;

        const baseSinExt = path.parse(archivo).name;
        const idxUltimoGuionBajo = baseSinExt.lastIndexOf("_");
        if (idxUltimoGuionBajo <= 0) return;

        const nombreEnArchivo = baseSinExt.slice(0, idxUltimoGuionBajo);
        const nombreArchivoNormalizado = normalizarNombre(nombreEnArchivo);
        
        if (nombreArchivoNormalizado === nombreTrabajadorNormalizado) {
          const tipoDoc = baseSinExt.slice(idxUltimoGuionBajo + 1);
          documentos.push({
            nombre: `${tipoDoc}.pdf`,
            tipo: "anexo",
            ruta: `/uploads/${archivo}`,
            tamano: (stat.size / 1024).toFixed(2) + " KB",
            fecha: stat.mtime.toLocaleDateString("es-MX")
          });
        }
      });
    }

    res.render("documentos", { trabajador, documentos });
  } catch (error) {
    console.error("Error al listar documentos:", error);
    res.status(500).send("Error al cargar documentos");
  }
});
app.get("/admin/login", (req, res) => {
  if (req.session.admin) {
    return res.redirect("/admin/dashboard");
  }
  res.render("admin/login", { error: null });
});

// POST - Procesar login de administrador
app.post("/admin/login", async (req, res) => {
  const { usuario, password } = req.body;

  if (!usuario || !password) {
    return res.render("admin/login", {
      error: "Usuario y contraseña requeridos",
    });
  }

  try {
    const [rows] = await db.query(
      "SELECT * FROM administrador WHERE usuario = ? AND password = ? AND activo = 1",
      [usuario, password],
    );

    if (rows && rows.length > 0) {
      // Guardar admin en sesión
      req.session.admin = rows[0];
      return res.redirect("/admin/dashboard");
    } else {
      return res.render("admin/login", {
        error: "Usuario o contraseña incorrectos",
      });
    }
  } catch (error) {
    console.error("Error en login admin:", error);
    return res.render("admin/login", { error: "Error interno del servidor" });
  }
});

// Middleware para verificar si es admin
const verificarAdmin = (req, res, next) => {
  if (!req.session.admin) {
    return res.redirect("/admin/login");
  }
  next();
};

function normalizarSexo(valor) {
  const sexo = String(valor || "")
    .trim()
    .toUpperCase();

  if (!sexo) return "";

  // Casos habituales completos
  if (sexo === "MASCULINO" || sexo === "HOMBRE" || sexo === "M") {
    return "Hombre";
  }
  if (sexo === "FEMENINO" || sexo === "MUJER" || sexo === "F") {
    return "Mujer";
  }

  // Fallback por inicial (ej. "Masculino ", "Femenino ", abreviaturas)
  if (sexo.startsWith("M")) return "Hombre";
  if (sexo.startsWith("F")) return "Mujer";

  return "";
}

// GET - Dashboard del administrador (Unificado: crear + buscar + últimos)
app.get("/admin/dashboard", verificarAdmin, async (req, res) => {
  let ultimosTrabajadores = [];

  try {
    // Obtener los 5 trabajadores más recientes (por ID más alto)
    const [resultados] = await db.query(
      "SELECT TOP 5 idTrabajador, nombreTrabajador, curp, noTrabajador FROM trabajador ORDER BY idTrabajador DESC",
      [],
    );
    ultimosTrabajadores = resultados || [];
  } catch (error) {
    console.error("Error al obtener últimos trabajadores:", error);
    ultimosTrabajadores = [];
  }

  return res.render("admin/dashboard", {
    admin: req.session.admin,
    ultimosTrabajadores: ultimosTrabajadores,
    usuariosBuscados: [],
    error: null,
    mensaje: null,
    busqueda: "",
  });
});

// POST - Crear nuevo trabajador (AJAX)
app.post("/admin/crear-usuario", verificarAdmin, async (req, res) => {
  const {
    nombre,
    rfc,
    curp,
    claveTrabjador,
    esAdmin,
    usuarioAdmin,
    correoAdmin,
    passwordAdmin,
    adminPasswordActual,
  } = req.body;

  const perfilAdmin =
    esAdmin === true || esAdmin === "true" || esAdmin === "on";

  try {
    if (perfilAdmin) {
      // Crear solo administrador (no trabajador)
      if (
        !usuarioAdmin ||
        !correoAdmin ||
        !passwordAdmin ||
        !adminPasswordActual
      ) {
        return res.status(400).json({
          success: false,
          error:
            "Para crear un administrador debes llenar usuario, correo, contraseña y confirmar tu contraseña actual.",
        });
      }

      const idAdminActual = req.session.admin?.idAdmin;

      const [adminValido] = await db.query(
        "SELECT * FROM administrador WHERE idAdmin = ? AND password = ? AND activo = 1",
        [idAdminActual, adminPasswordActual],
      );

      if (!adminValido || adminValido.length === 0) {
        return res.status(400).json({
          success: false,
          error: "Tu contraseña de administrador no es correcta.",
        });
      }

      const [adminExistente] = await db.query(
        "SELECT * FROM administrador WHERE usuario = ?",
        [usuarioAdmin],
      );

      if (adminExistente && adminExistente.length > 0) {
        return res.status(400).json({
          success: false,
          error: "Ya existe un administrador con ese usuario.",
        });
      }
      await db.query(
        "INSERT INTO administrador (usuario, password, correo, activo) VALUES (?, ?, ?, 1)",
        [usuarioAdmin, passwordAdmin, correoAdmin],
      );
      return res.json({
        success: true,
        mensaje: `Administrador ${usuarioAdmin} creado exitosamente`,
      });
    }

    // Perfil de trabajador: validar datos básicos
    const RFC_MAX = 13;

    const nombreLimpio = (nombre || "").trim().toUpperCase();
    const rfcLimpio = (rfc || "").trim().toUpperCase();
    const curpLimpio = (curp || "").trim().toUpperCase();
    const claveLimpia = (claveTrabjador || "").toString().trim();
    const tipoLimpio = (req.body.tipoTrabajador || "").trim().toUpperCase();

    if (
      !nombreLimpio ||
      !rfcLimpio ||
      !curpLimpio ||
      !claveLimpia ||
      !tipoLimpio
    ) {
      return res.status(400).json({
        success: false,
        error:
          "Todos los campos son obligatorios, incluyendo el tipo de trabajador (C u H)",
      });
    }

    if (tipoLimpio !== "C" && tipoLimpio !== "H") {
      return res.status(400).json({
        success: false,
        error: "El tipo de trabajador debe ser C (Confianza) o H (Honorario)",
      });
    }

    if (rfcLimpio.length > RFC_MAX) {
      return res.status(400).json({
        success: false,
        error: `El RFC no puede tener más de ${RFC_MAX} caracteres`,
      });
    }

    // Verificar que CURP y RFC no existan para el trabajador
    const [existe] = await db.query(
      "SELECT * FROM trabajador WHERE curp = ? OR rfc = ?",
      [curpLimpio, rfcLimpio],
    );

    if (existe && existe.length > 0) {
      return res.status(400).json({
        success: false,
        error: "El CURP o RFC ya existe en el sistema",
      });
    }

    await db.query(
      "INSERT INTO trabajador (nombreTrabajador, rfc, curp, noTrabajador, tipoTrabajador) VALUES (?, ?, ?, ?, ?)",
      [nombreLimpio, rfcLimpio, curpLimpio, claveLimpia, tipoLimpio],
    );

    return res.json({
      success: true,
      mensaje: `Trabajador ${nombreLimpio} creado exitosamente`,
    });
  } catch (error) {
    console.error("Error al crear trabajador:", error);
    return res.status(500).json({
      success: false,
      error: "Error al crear el trabajador. Intenta de nuevo.",
    });
  }
});

// POST - Buscar trabajadores (AJAX)
app.post("/admin/buscar-trabajadores", verificarAdmin, async (req, res) => {
  const { busqueda } = req.body;

  if (!busqueda || busqueda.trim() === "") {
    return res.json({
      success: true,
      usuarios: [],
      error: "Ingresa un nombre o CURP para buscar",
    });
  }

  try {
    const [usuarios] = await db.query(
      "SELECT idTrabajador, noTrabajador, nombreTrabajador, curp, rfc, correoElectronico, cargoActual FROM trabajador WHERE nombreTrabajador LIKE ? OR curp LIKE ? OR CAST(noTrabajador AS VARCHAR(50)) LIKE ? ORDER BY nombreTrabajador",
      [`%${busqueda}%`, `%${busqueda}%`, `%${busqueda}%`],
    );

    return res.json({
      success: true,
      usuarios: usuarios || [],
      error:
        usuarios && usuarios.length === 0
          ? "No se encontraron resultados"
          : null,
    });
  } catch (error) {
    console.error("Error al buscar trabajadores:", error);
    return res.status(500).json({
      success: false,
      error: "Error al buscar trabajadores",
      usuarios: [],
    });
  }
});

// GET - Ver detalles del trabajador
app.get("/admin/usuario/:idTrabajador", verificarAdmin, async (req, res) => {
  const { idTrabajador } = req.params;

  try {
    const [trabajador] = await db.query(
      "SELECT * FROM trabajador WHERE idTrabajador = ?",
      [idTrabajador],
    );

    if (!trabajador || trabajador.length === 0) {
      return res.status(404).render("admin/usuario-detalle", {
        trabajador: null,
        error: "Trabajador no encontrado",
      });
    }

    const [hijos] = await db.query(
      "SELECT inicialesHijo, sexoHijo, fechaNacimientoHijo FROM hijo WHERE idTrabajador = ?",
      [idTrabajador],
    );

    const [escolaridad] = await db.query(
      "SELECT * FROM escolaridad WHERE idTrabajador = ?",
      [idTrabajador],
    );

    const [experienciaPJ] = await db.query(
      "SELECT * FROM experienciapj WHERE idTrabajador = ?",
      [idTrabajador],
    );

    return res.render("admin/usuario-detalle", {
      trabajador: trabajador[0],
      hijos,
      escolaridad,
      experienciaPJ,
      error: null,
    });
  } catch (error) {
    console.error("Error al obtener detalles del trabajador:", error);
    return res.status(500).render("admin/usuario-detalle", {
      trabajador: null,
      error: "Error al cargar los detalles del trabajador",
    });
  }
});

// POST - Generar acuse desde admin (redirige al acuse HTML por noTrabajador)
app.post("/admin/generar-acuse", verificarAdmin, async (req, res) => {
  const { idTrabajador } = req.body;

  if (!idTrabajador) {
    return res.status(400).json({ error: "ID de trabajador requerido" });
  }

  try {
    const [trabajador] = await db.query(
      "SELECT noTrabajador, tipoTrabajador FROM trabajador WHERE idTrabajador = ?",
      [idTrabajador],
    );

    if (!trabajador || trabajador.length === 0) {
      return res.status(404).json({ error: "Trabajador no encontrado" });
    }

    const { noTrabajador, tipoTrabajador } = trabajador[0];
    return res.redirect(`/acuse/${noTrabajador}/${tipoTrabajador}`);
  } catch (error) {
    console.error("Error al redirigir al acuse desde admin:", error);
    return res.status(500).json({ error: "Error al generar el acuse" });
  }
});

// POST - Editar datos críticos del trabajador (noTrabajador, tipoTrabajador)
app.post("/admin/editar-datos-criticos", verificarAdmin, async (req, res) => {
  const { idTrabajador, nuevoNoTrabajador, nuevoTipo, password } = req.body;

  if (!idTrabajador || !nuevoNoTrabajador || !nuevoTipo || !password) {
    return res
      .status(400)
      .json({ success: false, error: "Todos los campos son obligatorios" });
  }

  if (nuevoTipo !== "C" && nuevoTipo !== "H") {
    return res
      .status(400)
      .json({ success: false, error: "Tipo de trabajador inválido" });
  }

  try {
    // Validar contraseña del administrador en sesión
    const idAdmin = req.session.admin?.idAdmin;
    const [adminRows] = await db.query(
      "SELECT * FROM administrador WHERE idAdmin = ? AND password = ? AND activo = 1",
      [idAdmin, password],
    );
    if (!adminRows || adminRows.length === 0) {
      return res.status(403).json({
        success: false,
        error: "Contraseña de administrador incorrecta",
      });
    }

    // Verificar que el nuevo número+tipo no exista en otro trabajador
    const [duplicado] = await db.query(
      "SELECT idTrabajador FROM trabajador WHERE noTrabajador = ? AND tipoTrabajador = ? AND idTrabajador <> ?",
      [nuevoNoTrabajador, nuevoTipo, idTrabajador],
    );
    if (duplicado && duplicado.length > 0) {
      return res.status(400).json({
        success: false,
        error: `Ya existe un trabajador con el número ${nuevoNoTrabajador} y tipo ${nuevoTipo === "C" ? "Confianza" : "Honorario"}`,
      });
    }

    await db.query(
      "UPDATE trabajador SET noTrabajador = ?, tipoTrabajador = ? WHERE idTrabajador = ?",
      [nuevoNoTrabajador, nuevoTipo, idTrabajador],
    );

    return res.json({
      success: true,
      mensaje: "Datos actualizados correctamente",
    });
  } catch (error) {
    console.error("Error al editar datos críticos:", error);
    return res
      .status(500)
      .json({ success: false, error: "Error interno al actualizar los datos" });
  }
});

// POST - Eliminar trabajador (con confirmación de contraseña)
app.post("/admin/eliminar-trabajador", verificarAdmin, async (req, res) => {
  const { idTrabajador, password } = req.body;

  if (!idTrabajador || !password) {
    return res
      .status(400)
      .json({ success: false, error: "ID y contraseña son requeridos" });
  }

  try {
    // Validar contraseña del administrador en sesión
    const idAdmin = req.session.admin?.idAdmin;
    const [adminRows] = await db.query(
      "SELECT * FROM administrador WHERE idAdmin = ? AND password = ? AND activo = 1",
      [idAdmin, password],
    );
    if (!adminRows || adminRows.length === 0) {
      return res.status(403).json({
        success: false,
        error: "Contraseña de administrador incorrecta",
      });
    }

    // Eliminar registros relacionados antes del trabajador principal
    await db.query("DELETE FROM hijo WHERE idTrabajador = ?", [idTrabajador]);
    await db.query("DELETE FROM escolaridad WHERE idTrabajador = ?", [
      idTrabajador,
    ]);
    await db.query("DELETE FROM experienciapj WHERE idTrabajador = ?", [
      idTrabajador,
    ]);
    await db.query(
      "DELETE FROM actualizacionprofesional WHERE idTrabajador = ?",
      [idTrabajador],
    );
    await db.query("DELETE FROM trabajador WHERE idTrabajador = ?", [
      idTrabajador,
    ]);

    return res.json({
      success: true,
      mensaje: "Trabajador eliminado correctamente",
    });
  } catch (error) {
    console.error("Error al eliminar trabajador:", error);
    return res.status(500).json({
      success: false,
      error: "Error interno al eliminar el trabajador",
    });
  }
});

// GET - Generar reporte de transparencia en Excel
app.get("/admin/reporte-transparencia", verificarAdmin, async (req, res) => {
  const { fechaInicio, fechaFin } = req.query;

  if (!fechaInicio || !fechaFin) {
    return res.status(400).json({ error: "Debes enviar fechaInicio y fechaFin" });
  }

  const inicio = dayjs(fechaInicio);
  const fin = dayjs(fechaFin);
  const hoy = dayjs().startOf("day");

  if (!inicio.isValid() || !fin.isValid()) {
    return res.status(400).json({ error: "Las fechas enviadas no son válidas" });
  }

  if (inicio.isAfter(fin, "day")) {
    return res
      .status(400)
      .json({ error: "La fecha de inicio no puede ser mayor a la fecha de término" });
  }

  if (fin.isAfter(hoy, "day")) {
    return res
      .status(400)
      .json({ error: "La fecha de término no puede ser posterior al día de hoy" });
  }

  try {
    const [trabajadores] = await db.query(
      `
      SELECT
        t.idTrabajador,
        t.noTrabajador,
        t.tipoTrabajador,
        t.nombreTrabajador,
        t.cargoActual,
        t.sexoTrabajador,
        t.adscripcionActual,
        ISNULL(exp.numExperiencia, 0) AS numExperiencia,
        esc.nivelAcademico,
        esc.nombreTitulo
      FROM trabajador t
      OUTER APPLY (
        SELECT COUNT(1) AS numExperiencia
        FROM experienciapj e
        WHERE e.idTrabajador = t.idTrabajador
      ) exp
      OUTER APPLY (
        SELECT TOP 1 s.nivelAcademico, s.nombreTitulo
        FROM escolaridad s
        WHERE s.idTrabajador = t.idTrabajador
        ORDER BY CASE UPPER(ISNULL(s.nivelAcademico, ''))
          WHEN 'DOCTORADO' THEN 1
          WHEN 'MAESTRIA' THEN 2
          WHEN 'LICENCIATURA' THEN 3
          WHEN 'BACHILLERATO' THEN 4
          WHEN 'TECNICA' THEN 5
          WHEN 'SECUNDARIA' THEN 6
          WHEN 'PRIMARIA' THEN 7
          ELSE 8
        END
      ) esc
      ORDER BY t.nombreTrabajador ASC
      `,
      [],
    );

    // Nueva hoja: Tabla_350631 con la vista de experiencia
    const [vistaExperiencia] = await db.query(
      "SELECT * FROM VistaExperiencia",
      [],
    );

    const fechaInicioTexto = inicio.format("YYYY-MM-DD");
    const fechaFinTexto = fin.format("YYYY-MM-DD");
    const fechaActualizacion = dayjs().format("YYYY-MM-DD");
    const ejercicio = inicio.year();
    const aplicarSexo = !inicio.isBefore(dayjs("2023-04-01"), "day");

    // Usar la variable BASE_URL para el hipervínculo de trayectoria
    const baseUrlTrayectoria = `${BASE_URL}/curriculum_transparencia/`;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Reporte Transparencia");

    const columnas = [
      "Ejercicio",
      "Fecha de inicio del periodo que se informa",
      "Fecha de término del periodo que se informa",
      "Denominación de puesto (Redactados con perspectiva de género)",
      "Nombre",
      "ESTE CRITERIO APLICA A PARTIR DEL 01/04/2023 -> Sexo (Hombre Mujer)",
      "Área de adscripción",
      "Nivel máximo de estudios concluido y comprobable (catálogo)",
      "Carrera genérica, en su caso",
      "Experiencia laboral Tabla_350631",
      "Hipervínculo al documento que contenga la trayectoria",
      "Sanciones Administrativas definitivas aplicadas por la autoridad competente (catálogo)",
      "Hipervínculo a la resolución donde se observe la aprobación de la sanción",
      "Área(s) responsable(s) que genera(n), posee(n), publica(n) y actualizan la información",
      "Fecha de actualización",
      "Notas",
    ];

    worksheet.addRow(columnas);

    for (const t of trabajadores || []) {
      const fila = [
        ejercicio,
        fechaInicioTexto,
        fechaFinTexto,
        t.cargoActual || "",
        t.nombreTrabajador || "",
      ];

      fila.push(aplicarSexo ? normalizarSexo(t.sexoTrabajador) : "");

      const noTrabajador = t.noTrabajador || "";
      const tipoTrabajadorRep = t.tipoTrabajador || "";
      const hipervinculo =
        noTrabajador && tipoTrabajadorRep
          ? `${baseUrlTrayectoria}${noTrabajador}/${tipoTrabajadorRep}`
          : "";

      fila.push(
        t.adscripcionActual || "",
        t.nivelAcademico || "",
        t.nombreTitulo || "",
        noTrabajador,
        hipervinculo,
        "", // Llenado manualmente por rh
        "", // llenado manualmente por rh
        "Recursos Humanos",
        fechaActualizacion,
        "", // Notas: pendiente por indicación
      );

      worksheet.addRow(fila);
    }

    const header = worksheet.getRow(1);
    header.font = { bold: true };
    header.alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };
    header.height = 30;

    worksheet.columns = columnas.map((col) => ({
      width: Math.min(Math.max(col.length + 4, 18), 60),
    }));

    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    worksheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: columnas.length },
    };

    // Hoja 2: exportar VistaExperiencia en Tabla_350631
    const wsExperiencia = workbook.addWorksheet("Tabla_350631");

    if (vistaExperiencia && vistaExperiencia.length > 0) {
      const columnasVista = Object.keys(vistaExperiencia[0]);
      wsExperiencia.addRow(columnasVista);

      for (const row of vistaExperiencia) {
        wsExperiencia.addRow(columnasVista.map((col) => row[col]));
      }

      const headerExp = wsExperiencia.getRow(1);
      headerExp.font = { bold: true };
      headerExp.alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: true,
      };

      wsExperiencia.columns = columnasVista.map((col) => ({
        width: Math.min(Math.max(String(col).length + 4, 15), 50),
      }));
    }

    const nombreArchivo = `REPORTE_TRANSPARENCIA_${fechaInicioTexto}_A_${fechaFinTexto}.xlsx`;
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${nombreArchivo}"`,
    );

    await workbook.xlsx.write(res);
    return res.end();
  } catch (error) {
    console.error("Error al generar reporte de transparencia:", error);
    return res.status(500).json({ error: "Error al generar el reporte de transparencia" });
  }
});

// GET - Gestionar administradores (lista todos los admins)
app.get("/admin/gestionar-admins", verificarAdmin, async (req, res) => {
  try {
    const [admins] = await db.query(
      "SELECT idAdmin, usuario, correo, activo FROM administrador ORDER BY idAdmin ASC",
      [],
    );
    return res.render("admin/gestionar-admins", {
      admin: req.session.admin,
      administradores: admins || [],
      error: null,
      mensaje: null,
    });
  } catch (error) {
    console.error("Error al listar administradores:", error);
    return res.render("admin/gestionar-admins", {
      admin: req.session.admin,
      administradores: [],
      error: "Error al obtener la lista de administradores",
      mensaje: null,
    });
  }
});

// POST - Desactivar administrador (borrado lógico)
app.post("/admin/desactivar-admin", verificarAdmin, async (req, res) => {
  const { idAdminObjetivo, password } = req.body;
  const idAdminActual = req.session.admin?.idAdmin;

  if (!idAdminObjetivo || !password) {
    return res.status(400).json({
      success: false,
      error: "ID del administrador y contraseña son requeridos",
    });
  }

  // No puede desactivarse a sí mismo
  if (parseInt(idAdminObjetivo) === idAdminActual) {
    return res
      .status(400)
      .json({ success: false, error: "No puedes desactivar tu propia cuenta" });
  }

  try {
    // Validar contraseña del admin en sesión
    const [adminValido] = await db.query(
      "SELECT * FROM administrador WHERE idAdmin = ? AND password = ? AND activo = 1",
      [idAdminActual, password],
    );
    if (!adminValido || adminValido.length === 0) {
      return res.status(403).json({
        success: false,
        error: "Contraseña de administrador incorrecta",
      });
    }

    // Verificar que quede al menos un admin activo después de desactivar
    const [conteo] = await db.query(
      "SELECT COUNT(*) AS total FROM administrador WHERE activo = 1 AND idAdmin <> ?",
      [idAdminObjetivo],
    );
    if (!conteo || conteo[0].total < 1) {
      return res.status(400).json({
        success: false,
        error:
          "No se puede desactivar: debe quedar al menos un administrador activo",
      });
    }

    await db.query("UPDATE administrador SET activo = 0 WHERE idAdmin = ?", [
      idAdminObjetivo,
    ]);

    return res.json({
      success: true,
      mensaje: "Administrador desactivado correctamente",
    });
  } catch (error) {
    console.error("Error al desactivar administrador:", error);
    return res.status(500).json({
      success: false,
      error: "Error interno al desactivar el administrador",
    });
  }
});

// POST - Activar administrador
app.post("/admin/activar-admin", verificarAdmin, async (req, res) => {
  const { idAdminObjetivo, password } = req.body;
  const idAdminActual = req.session.admin?.idAdmin;

  if (!idAdminObjetivo || !password) {
    return res.status(400).json({
      success: false,
      error: "ID del administrador y contraseña son requeridos",
    });
  }

  try {
    const [adminValido] = await db.query(
      "SELECT * FROM administrador WHERE idAdmin = ? AND password = ? AND activo = 1",
      [idAdminActual, password],
    );
    if (!adminValido || adminValido.length === 0) {
      return res.status(403).json({
        success: false,
        error: "Contraseña de administrador incorrecta",
      });
    }

    await db.query("UPDATE administrador SET activo = 1 WHERE idAdmin = ?", [
      idAdminObjetivo,
    ]);

    return res.json({
      success: true,
      mensaje: "Administrador activado correctamente",
    });
  } catch (error) {
    console.error("Error al activar administrador:", error);
    return res.status(500).json({
      success: false,
      error: "Error interno al activar el administrador",
    });
  }
});

// POST - Eliminar administrador (borrado físico)
app.post("/admin/eliminar-admin", verificarAdmin, async (req, res) => {
  const { idAdminObjetivo, password } = req.body;
  const idAdminActual = req.session.admin?.idAdmin;

  if (!idAdminObjetivo || !password) {
    return res.status(400).json({
      success: false,
      error: "ID del administrador y contraseña son requeridos",
    });
  }

  // No puede eliminarse a sí mismo
  if (parseInt(idAdminObjetivo) === idAdminActual) {
    return res
      .status(400)
      .json({ success: false, error: "No puedes eliminar tu propia cuenta" });
  }

  try {
    // Validar contraseña del admin en sesión
    const [adminValido] = await db.query(
      "SELECT * FROM administrador WHERE idAdmin = ? AND password = ? AND activo = 1",
      [idAdminActual, password],
    );
    if (!adminValido || adminValido.length === 0) {
      return res.status(403).json({
        success: false,
        error: "Contraseña de administrador incorrecta",
      });
    }

    // Verificar que quede al menos un admin activo después de eliminar
    const [conteo] = await db.query(
      "SELECT COUNT(*) AS total FROM administrador WHERE activo = 1 AND idAdmin <> ?",
      [idAdminObjetivo],
    );
    if (!conteo || conteo[0].total < 1) {
      return res.status(400).json({
        success: false,
        error:
          "No se puede eliminar: debe quedar al menos un administrador activo",
      });
    }

    await db.query("DELETE FROM administrador WHERE idAdmin = ?", [
      idAdminObjetivo,
    ]);

    return res.json({
      success: true,
      mensaje: "Administrador eliminado permanentemente",
    });
  } catch (error) {
    console.error("Error al eliminar administrador:", error);
    return res.status(500).json({
      success: false,
      error: "Error interno al eliminar el administrador",
    });
  }
});

// GET - Logout de administrador
app.get("/admin/logout", (req, res) => {
  req.session.admin = null;
  res.redirect("/admin/login");
});

// Puerto
app.listen(3001, () => {
  console.log("Servidor corriendo en http://localhost:3001/login");
});

async function convertExcelToPdf(excelPath, pdfPath) {
  // 1. Leer el archivo Excel
  const workbook = XLSX.readFile(excelPath);
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];

  // 2. Convertir a HTML (automático, sin especificar campos)
  const html = XLSX.utils.sheet_to_html(worksheet);

  // 3. Crear archivo HTML temporal
  const htmlPath = excelPath.replace(".xlsx", ".html");
  fs.writeFileSync(
    htmlPath,
    `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body { font-family: Arial, sans-serif; }
                table { border-collapse: collapse; width: 100%; }
                th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                th { background-color: #f2f2f2; }
                tr:nth-child(even) { background-color: #f9f9f9; }
            </style>
        </head>
        <body>
            <h1>Reporte Generado</h1>
            ${html}
        </body>
        </html>
    `,
  );

  // 4. Convertir HTML a PDF usando Puppeteer
  const browser = await puppeteer.launch({
    headless: "new",
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });
  const page = await browser.newPage();

  // Usar file:/// con forward slashes para compatibilidad Windows
  const fileUrlHtmlLegacy = `file:///${path.resolve(htmlPath).replace(/\\/g, "/")}`;
  await page.goto(fileUrlHtmlLegacy, {
    waitUntil: "networkidle0",
  });

  await page.pdf({
    path: pdfPath,
    format: "letter",
    printBackground: true,
    margin: {
      top: "20mm",
      bottom: "20mm",
      left: "10mm",
      right: "10mm",
    },
  });

  await browser.close();

  // 5. Eliminar archivo HTML temporal
  fs.unlinkSync(htmlPath);
}

function getSafeCellValue(cell) {
  try {
    if (cell.value === null) return "";
    if (cell.value === undefined) return "";
    if (typeof cell.value === "object") return JSON.stringify(cell.value);
    return String(cell.value);
  } catch (error) {
    console.error(`Error en celda ${cell.address}: ${error.message}`);
    return "";
  }
}
async function excelToHtmlWithStyles(excelPath, htmlPath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelPath);

  // Decodifica “A1” → {row:1, col:1}
  function decodeAddress(address) {
    const match = address.match(/^([A-Z]+)(\d+)$/);
    if (!match) throw new Error(`Dirección inválida: ${address}`);
    const [, letters, number] = match;
    let col = 0;
    for (let i = 0; i < letters.length; i++) {
      col = col * 26 + (letters.charCodeAt(i) - 64);
    }
    return { row: parseInt(number, 10), col };
  }

  const logoPath = path.join(__dirname, "public/images", "pj-logo-verde.png");
  const logoExt = path.extname(logoPath).slice(1); // "png" o "jpg"
  const logoData = fs.readFileSync(logoPath).toString("base64");
  const logoSrcLeft = `data:image/${logoExt};base64,${logoData}`;

  const logoPath2 = path.join(
    __dirname,
    "public/images",
    "consejo-logo-verde.png",
  );
  const logoExt2 = path.extname(logoPath2).slice(1); // "png" o "jpg"
  const logoData2 = fs.readFileSync(logoPath2).toString("base64");
  const logoSrcRight = `data:image/${logoExt2};base64,${logoData2}`;

  let html = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><title>Reporte Excel</title>
  <style>
    body { font-family: Arial, sans-serif; }
    table { border-collapse: collapse; width: 100%; }
    td, th { border: 1px solid #ddd; padding: 8px; }
    .bold { font-weight: bold; }
    .italic { font-style: italic; }
    .underline { text-decoration: underline; }
    .center { text-align: center; }
    .right { text-align: right; }
    .left { text-align: left; }
  </style>
</head>
<body>
<div style="
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
">
  <!-- Logo izquierda -->
  <img src="${logoSrcLeft}" alt="Logo Izquierda" style="height:150px;">

  <!-- Logo derecha -->
  <img src="${logoSrcRight}" alt="Logo Derecha" style="height:80px;">
</div>
`;

  workbook.eachSheet((worksheet) => {
    html += `<h2></h2><table>`;

    // 1) Merges
    const merges = worksheet.model.merges || [];
    const mergedCells = new Map();
    merges.forEach((range) => {
      const [start, end] = range.split(":");
      const s = decodeAddress(start);
      const e = decodeAddress(end);
      mergedCells.set(`${s.row},${s.col}`, {
        top: s.row,
        left: s.col,
        bottom: e.row,
        right: e.col,
      });
      for (let r = s.row; r <= e.row; r++) {
        for (let c = s.col; c <= e.col; c++) {
          if (r !== s.row || c !== s.col)
            mergedCells.set(`${r},${c}`, { merged: true });
        }
      }
    });

    // 2) Filas con límite de 10 vacías seguidas
    const rowCount = worksheet.rowCount || 0;
    const MAX_COLUMNS = 5; // <-- tu límite
    const colCount = Math.min(worksheet.columnCount || 0, MAX_COLUMNS);
    let emptyStreak = 0;
    const MAX_EMPTY = 10;

    for (let rowNum = 1; rowNum <= rowCount; rowNum++) {
      const row = worksheet.getRow(rowNum);
      if (row.hidden) continue;

      // Detectar si fila vacía
      let isEmpty = true;
      for (let c = 1; c <= colCount; c++) {
        if ((row.getCell(c).text || "").toString().trim() !== "") {
          isEmpty = false;
          break;
        }
      }
      if (isEmpty) {
        emptyStreak++;
        if (emptyStreak >= MAX_EMPTY) break; // salta todo el resto
        continue; // ignora esta fila
      }
      emptyStreak = 0; // reinicia racha
      html += "<tr>";

      // Columnas
      for (let colNum = 1; colNum <= colCount; colNum++) {
        const key = `${rowNum},${colNum}`;
        if (mergedCells.has(key) && mergedCells.get(key).merged) continue;

        const cell = row.getCell(colNum);
        if (!cell) {
          html += "<td></td>";
          continue;
        }
        const text = getSafeCellValue(cell);
        const cellStyles = [];
        const style = cell.style || {};

        // rowspan/colspan
        let cellAttrs = "";
        if (mergedCells.has(key) && !mergedCells.get(key).merged) {
          const rng = mergedCells.get(key);
          const rs = rng.bottom - rng.top + 1;
          const cs = rng.right - rng.left + 1;
          cellAttrs = ` rowspan="${rs}" colspan="${cs}"`;
        }

        // clases de fuente y alineación
        const cls = [];
        if (style.font) {
          if (style.font.bold) cls.push("bold");
          if (style.font.italic) cls.push("italic");
          if (style.font.underline) cls.push("underline");
        }
        if (style.alignment) {
          if (style.alignment.horizontal === "center") cls.push("center");
          if (style.alignment.horizontal === "right") cls.push("right");
        }
        const classAttr = cls.length ? ` class="${cls.join(" ")}"` : "";

        // fondo
        let bg = "";
        if (
          style.fill &&
          style.fill.type === "pattern" &&
          style.fill.fgColor?.argb
        ) {
          bg = `background-color: #${style.fill.fgColor.argb.slice(2)};`;
        }
        // bordes
        let border = "";
        ["top", "left", "bottom", "right"].forEach((side) => {
          const b = style.border?.[side];
          if (b && b.style) {
            const color = b.color?.argb ? `#${b.color.argb.slice(2)}` : "#000";
            border += `border-${side}: ${getBorderWidth(b.style)} ${color};`;
          }
        });
        const styleAttr = (bg + border).trim();
        const styleString = styleAttr ? ` style="${styleAttr}"` : "";

        html += `<td${cellAttrs}${classAttr}${styleString}>${text}</td>`;
      }

      html += "</tr>";
    }

    html += "</table>";
  });

  html += "</body></html>";
  fs.writeFileSync(htmlPath, html);
}

function getBorderWidth(s) {
  return (
    {
      thin: "1px solid",
      medium: "2px solid",
      thick: "3px solid",
      dashed: "1px dashed",
      dotted: "1px dotted",
      double: "3px double",
      hair: "1px solid",
    }[s] || "1px solid"
  );
}

app.post("/registrar-password", async (req, res) => {
  const { curp, rfc, password, confirmar } = req.body;

  if (!curp || !rfc || !password || !confirmar) {
    return res.render("registrarPassword", {
      error: "Todos los campos son obligatorios",
    });
  }

  if (password !== confirmar) {
    return res.render("registrarPassword", {
      error: "Las contraseñas no coinciden",
    });
  }

  try {
    const [results] = await db.query(
      "SELECT * FROM trabajador WHERE curp = ? AND rfc = ?",
      [curp.toUpperCase(), rfc.toUpperCase()],
    );

    if (results.length === 0) {
      return res.render("registrarPassword", {
        error: "CURP o RFC no encontrados",
      });
    }

    const trabajador = results[0];

    if (trabajador.password) {
      return res.render("registrarPassword", {
        error:
          "Ya tienes una contraseña registrada. Utiliza el medio de ingreso normal.",
      });
    }

    // Guardar contraseña sin cifrar (¡peligroso!)
    // await db.query('UPDATE trabajador SET password = ? WHERE curp = ?', [password, curp.toUpperCase()]);

    // Versión segura: usar bcrypt
    //const bcrypt = require('bcrypt');
    //const hash = await bcrypt.hash(password, 10);
    await db.query(
      "UPDATE trabajador SET password = ? WHERE curp = ? AND rfc = ?",
      [password, curp.toUpperCase(), rfc.toUpperCase()],
    );

    return res.send(`
        <script>
          alert('Contraseña registrada exitosamente. Ya puedes iniciar sesión.');
          window.location.href = '/login';
        </script>
      `);
  } catch (err) {
    console.error("Error registrando contraseña:", err);
    return res.render("registrarPassword", {
      error: "Error interno. Intenta más tarde.",
    });
  }
});

// Configuración de Nodemailer (SMTP)
const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST || "smtp.office365.com",
  port: process.env.SMTP_PORT ? parseInt(process.env.SMTP_PORT, 10) : 587,
  secure: false,
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  },
});

// GET mostrar formulario
app.get("/recuperar-password", (req, res) => {
  res.render("recuperar-password", { error: null, mensaje: null });
});

app.post("/recuperar-password", async (req, res) => {
  const { correoe, curp, action } = req.body;

  try {
    // Recuperar CLAVE por CORREO directamente desde trabajador
    if (action === "recuperar-clave") {
      if (!correoe) {
        return res.render("recuperar-password", {
          error: "Debe ingresar un correo electrónico",
          mensaje: null,
        });
      }

      const correoTrim = correoe.trim();

      const [rows] = await db.query(
        "SELECT TOP 1 password, correoElectronico FROM trabajador WHERE LTRIM(RTRIM(correoElectronico)) = ?",
        [correoTrim],
      );

      if (!rows || rows.length === 0) {
        return res.render("recuperar-password", {
          error: "No se encontró trabajador con ese correo",
          mensaje: null,
        });
      }

      const trabajador = rows[0];

      if (!trabajador.correoElectronico) {
        return res.render("recuperar-password", {
          error: "El trabajador existe pero no tiene correo registrado",
          mensaje: null,
        });
      }

      await transporter.sendMail({
        from: '"Soporte Recursos Humanos" <desarrollo.sistemas@pjhidalgo.gob.mx>',
        to: correoTrim,
        subject: "Recuperación de clave",
        html: `<p>Su clave de acceso al módulo de actualización de datos personales de los Trabajadores del Poder Judicial de Hidalgo es: <b>${trabajador.password}</b></p>`,
      });

      return res.render("recuperar-password", {
        error: null,
        mensaje: `La clave fue enviada al correo ${correoTrim}`,
      });
    }

    // Recuperar CORREO por CURP directamente desde trabajador
    else if (action === "recuperar-correo") {
      if (!curp) {
        return res.render("recuperar-password", {
          error: "Debe ingresar la CURP",
          mensaje: null,
        });
      }

      const curpTrim = curp.trim().toUpperCase();

      const [rows] = await db.query(
        "SELECT TOP 1 correoElectronico FROM trabajador WHERE LTRIM(RTRIM(curp)) = ?",
        [curpTrim],
      );

      if (!rows || rows.length === 0) {
        return res.render("recuperar-password", {
          error: "No se encontró usuario con esa CURP",
          mensaje: null,
        });
      }

      const trabajador = rows[0];

      if (!trabajador.correoElectronico) {
        return res.render("recuperar-password", {
          error: "El trabajador existe pero no tiene correo registrado",
          mensaje: null,
        });
      }

      return res.render("recuperar-password", {
        error: null,
        mensaje: `El correo registrado para la CURP ${curpTrim} es ${trabajador.correoElectronico}`,
      });
    }

    // Acción desconocida
    else {
      return res.render("recuperar-password", {
        error: "Acción no reconocida",
        mensaje: null,
      });
    }
  } catch (err) {
    console.error("Error en recuperación:", err);
    res.status(500).render("recuperar-password", {
      error: "Ocurrió un error en el servidor",
      mensaje: null,
    });
  }
});
