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
          "Sesion expirada",
          "Tu sesion termino y por eso no se pudo abrir el acuse.",
          "Vuelve a iniciar sesion para continuar.",
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
          `Estas en sesion como No. ${noTrabajadorSesion || "N/A"} Tipo ${tipoSesion || "N/A"}, pero intentaste abrir No. ${noTrabajadorParam || "N/A"} Tipo ${tipoParam || "N/A"}.`,
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
  req.session.destroy(() => {
    res.status(200).send("Sesión cerrada");
  });
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
  const curp = req.body.curp?.toUpperCase();
  const rfc = req.body.rfc?.toUpperCase();
  const password = req.body.password;

  try {
    const [results] = await db.query(
      "SELECT * FROM trabajador WHERE curp = ? AND rfc = ?",
      [curp, rfc],
    );

    if (results.length === 0) {
      return res.status(401).render("login", {
        error: "CURP o RFC no encontrado. Verifica los datos.",
        curp,
        rfc,
      });
    }

    const usuario = results[0];

    if (usuario.password !== password) {
      return res.status(401).render("login", {
        error: "Contraseña incorrecta. Intenta nuevamente.",
        curp,
        rfc,
      });
    }

    // 🔹 Destruir sesión previa antes de asignar nueva
    req.session.regenerate((err) => {
      if (err) {
        console.error("❌ Error regenerando sesión:", err);
        return res.send("Error al iniciar sesión");
      }

      // Asignar usuario a la nueva sesión
      req.session.user = usuario;

      // Redirigir al index
      return res.redirect("/index");
    });
  } catch (err) {
    console.error("❌ Error en el query:", err);
    return res.status(500).render("login", {
      error:
        "Ocurrió un error al consultar la base de datos. Intenta nuevamente más tarde.",
      curp,
      rfc,
    });
  }
});

app.get("/reporte/:noTrabajador/:tipoTrabajador", async (req, res) => {
  const { noTrabajador, tipoTrabajador } = req.params;
  const previewOnly = req.query.preview === "1";

  const rutaPlantilla = "./plantillas/curriculum.xlsx";

  try {
    const [trabajadorRows] = await db.query(
      "SELECT * FROM trabajador WHERE noTrabajador = ? AND tipoTrabajador = ?",
      [noTrabajador, tipoTrabajador],
    );
    const trabajador = trabajadorRows[0];

    if (!trabajador) {
      return res.status(404).send("No se encontró el trabajador");
    }

    const idTrabajador = trabajador.idTrabajador;

    const [hijos] = await db.query(
      "SELECT * FROM hijo WHERE idTrabajador = ?",
      [idTrabajador],
    );

    const [escolaridades] = await db.query(
      "SELECT * FROM escolaridad WHERE idTrabajador = ?",
      [idTrabajador],
    );

    const [actualizaciones] = await db.query(
      "SELECT * FROM actualizacionprofesional WHERE idTrabajador = ?",
      [idTrabajador],
    );

    const [experiencias] = await db.query(
      "SELECT * FROM experienciapj WHERE idTrabajador = ?",
      [idTrabajador],
    );

    // console.log("Experiencias obtenidas:", experiencias.length); // Log opcional para validar conteo

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(rutaPlantilla);
    const worksheet = workbook.worksheets[0];
    worksheet.getCell("B7").value = trabajador.nombreTrabajador;
    worksheet.getCell("B11").value = trabajador.primaria;
    worksheet.getCell("B12").value = trabajador.secundaria;
    worksheet.getCell("B13").value = trabajador.carreraTecnicaComercial;

    // Hipervínculo al curriculum
    const curriculumUrl = construirUrlCvExterno(
      trabajador.noTrabajador,
      trabajador.tipoTrabajador,
    );
    // Busca la columna correspondiente, por ejemplo C20 (ajusta según plantilla)
    // Si la columna es variable, busca por nombre o posición
    worksheet.getCell("C20").value = {
      text: "Ver trayectoria",
      hyperlink: curriculumUrl,
    };

    const bachillerato = escolaridades.filter(
      (exp) =>
        String(exp.nivelAcademico || "").toUpperCase() === "BACHILLERATO",
    );
    bachillerato.forEach((exp, index) => {
      const fila = 14 + index; // Empieza en la fila 14, 15, 16…
      // Usando template literals con backticks:
      worksheet.getCell(`B${fila}`).value = exp.nombreTitulo;
      worksheet.getCell(`C${fila}`).value = exp.fechaObtencion;
      worksheet.getCell(`D${fila}`).value = exp.institucion;
      worksheet.getCell(`E${fila}`).value = exp.documentoAdquirido;
      worksheet.getCell(`F${fila}`).value = exp.estatus;
    });

    const EXPERIENCE_ROW = 26;
    const expMerges = [];
    worksheet.model.merges.forEach((range) => {
      const [start, end] = range.split(":");
      const startRow = +start.match(/\d+$/)[0];
      const endRow = +end.match(/\d+$/)[0];
      if (startRow === EXPERIENCE_ROW && endRow === EXPERIENCE_ROW) {
        expMerges.push(range);
      }
    });

    const experienciasPJ = experiencias;
    const totalExp = experienciasPJ.length;

    if (totalExp === 0) {
      const mergeRange = `A26:D26`;

      try {
        worksheet.unMergeCells(mergeRange);
      } catch (err) {}
      worksheet.mergeCells(mergeRange);
      worksheet.getCell("A26").value = "Sin Experiencias";
      worksheet.getCell("A26").alignment = {
        horizontal: "center",
      };
    }
    if (totalExp > 1) {
      worksheet.duplicateRow(EXPERIENCE_ROW, totalExp - 1, true);
    }

    expMerges.forEach((range) => {
      const [start, end] = range.split(":");
      const colStart = start.match(/^[A-Z]+/)[0];
      const colEnd = end.match(/^[A-Z]+/)[0];
      for (let i = 1; i < totalExp; i++) {
        worksheet.mergeCells(
          `${colStart}${EXPERIENCE_ROW + i}:${colEnd}${EXPERIENCE_ROW + i}`,
        );
      }
    });

    experienciasPJ.forEach((exp, idx) => {
      const r = EXPERIENCE_ROW + idx;
      worksheet.getCell(`A${r}`).value = exp.periodoInicio;
      worksheet.getCell(`B${r}`).value = exp.periodoFin;
      worksheet.getCell(`C${r}`).value =
        exp.institucion + " / " + exp.adscripcion;
      worksheet.getCell(`D${r}`).value = exp.cargo;
      worksheet.getCell(`E${r}`).value = exp.campoExperiencia;
    });
    ///////////////////////////////////////////////////

    const UPDATE_ROW = 22;
    const updateMerges = [];

    worksheet.model.merges.forEach((range) => {
      const [start, end] = range.split(":");
      const startRow = parseInt(start.match(/\d+$/)[0], 10);
      const endRow = parseInt(end.match(/\d+$/)[0], 10);
      if (startRow === UPDATE_ROW && endRow === UPDATE_ROW) {
        updateMerges.push(range);
      }
    });

    const totalActualizaciones = actualizaciones.length;
    if (totalActualizaciones > 1) {
      worksheet.duplicateRow(UPDATE_ROW, totalActualizaciones - 1, true);
    }

    if (totalActualizaciones === 0) {
      const mergeRange = `A22:D22`;

      try {
        worksheet.unMergeCells(mergeRange);
      } catch (err) {}
      worksheet.mergeCells(mergeRange);
      worksheet.getCell("B22").value = "Sin Actualizaciones";
      worksheet.getCell("B22").alignment = {
        horizontal: "center",
      };
    }

    updateMerges.forEach((range) => {
      const [start, end] = range.split(":");
      const colStart = start.match(/^[A-Z]+/)[0];
      const colEnd = end.match(/^[A-Z]+/)[0];
      for (let i = 1; i < totalActualizaciones; i++) {
        console.log(
          "rango de merge1: ",
          `${colStart}${UPDATE_ROW + i}:${colEnd}${UPDATE_ROW + i}`,
        );
        worksheet.mergeCells(
          `${colStart}${UPDATE_ROW + i}:${colEnd}${UPDATE_ROW + i}`,
        );
      }
    });

    actualizaciones.forEach((exp, idx) => {
      const r = UPDATE_ROW + idx;
      worksheet.getCell(`A${r}`).value = exp.tema;
      worksheet.getCell(`B${r}`).value = formatearFecha(exp.fecha);
      worksheet.getCell(`C${r}`).value = exp.institucion;
      worksheet.getCell(`D${r}`).value = exp.documento;
    });

    const PREPA_START_ROW = 14;
    const prepa = escolaridades.filter(
      (e) => String(e.nivelAcademico || "").toUpperCase() === "BACHILLERATO",
    );
    const prepaCount = prepa.length;

    if (prepaCount === 0) {
      const mergeRange = `B14:G14`;

      try {
        worksheet.unMergeCells(mergeRange);
      } catch (err) {}
      worksheet.mergeCells(mergeRange);
      worksheet.getCell("B14").value = "No";
      worksheet.getCell("B14").alignment = {
        horizontal: "center",
      };
      worksheet.getRow(PREPA_START_ROW).hidden = true;
    } else {
      if (prepaCount > 1) {
        worksheet.duplicateRow(PREPA_START_ROW, prepaCount - 1, true);
      }

      const start = PREPA_START_ROW;
      const end = PREPA_START_ROW + prepaCount - 1;

      // 4) Llenamos cada fila duplicada con los datos
      prepa.forEach((exp, i) => {
        const row = start + i;
        worksheet.getCell(`B${row}`).value = exp.nombreTitulo;
        worksheet.getCell(`C${row}`).value = exp.fechaObtencion;
        worksheet.getCell(`D${row}`).value = exp.institucion;
        worksheet.getCell(`E${row}`).value = exp.documentoAdquirido;
        worksheet.getCell(`F${row}`).value = exp.estatus;
      });

      // 5) Finalmente: si había >1, combinamos la columna A de start a end
      if (prepaCount > 1) {
        const mergeRange = `A${start}:A${end}`;

        try {
          // Intenta desfusionar (no lanza error si no existe)
          worksheet.unMergeCells(mergeRange);
        } catch (err) {
          // Ignora el error si el merge no existe
        }

        // Fusiona las celdas
        worksheet.mergeCells(mergeRange);
      }
    }

    let LIC_START_ROW = 15;
    const licenciaturas = escolaridades.filter(
      (e) => String(e.nivelAcademico || "").toUpperCase() === "LICENCIATURA",
    );
    let licCount = licenciaturas.length;
    // console.log("length lic:", licCount); // Depuración conteo de licenciaturas
    let acarreo = licCount;
    let licDesplazar = 15 + acarreo;

    // 1) Si no hay ninguna, borramos la fila base
    if (licCount === 0) {
      const mergeRange = `B${licDesplazar}:G${licDesplazar}`;

      try {
        worksheet.unMergeCells(mergeRange);
      } catch (err) {}
      worksheet.mergeCells(mergeRange);
      worksheet.getCell(`B${licDesplazar}`).value = "Sin licenciatura";
      worksheet.getCell(`B${licDesplazar}`).alignment = {
        horizontal: "center",
      };
      worksheet.getRow(LIC_START_ROW).hidden = true;
    } else {
      // 2) Si hay más de 1, duplicamos la fila base licCount-1 veces

      if (licCount === 1) {
        acarreo = acarreo - 1;
      }
      if (licCount > 1) {
        acarreo = acarreo - 1;
        // (filaBase, númeroDeDuplicados, insert => true para desplazar filas hacia abajo)
        worksheet.duplicateRow(LIC_START_ROW, licCount - 1, true);
      }

      const start = LIC_START_ROW;
      const end = LIC_START_ROW + licCount - 1;

      // 4) Llenamos cada fila duplicada con los datos
      licenciaturas.forEach((exp, i) => {
        const row = start + i;
        worksheet.getCell(`B${row}`).value = exp.nombreTitulo;
        worksheet.getCell(`C${row}`).value = exp.fechaObtencion;
        worksheet.getCell(`D${row}`).value = exp.institucion;
        worksheet.getCell(`E${row}`).value = exp.documentoAdquirido;
        worksheet.getCell(`F${row}`).value = exp.estatus;
        worksheet.getCell(`G${row}`).value = exp.cedulaProfesional;
      });

      // 5) Finalmente: si había >1, combinamos la columna A de start a end
      if (licCount > 1) {
        const mergeRange = `A${start}:A${end}`;

        try {
          // Intenta desfusionar (no lanza error si no existe)
          worksheet.unMergeCells(mergeRange);
        } catch (err) {
          // Ignora el error si el merge no existe
        }

        // Fusiona las celdas
        worksheet.mergeCells(mergeRange);
      }
    }
    ////////////////////////////////////////////////////////////////////////

    let MAE_START_ROW = 16 + acarreo; // ← ajusta esta fila al inicio de tu sección "Maestría"
    const maestrias = escolaridades.filter(
      (e) => String(e.nivelAcademico || "").toUpperCase() === "MAESTRIA",
    );
    let maeCount = maestrias.length;

    if (maeCount === 0) {
      const mergeRange = `B${MAE_START_ROW}:G${MAE_START_ROW}`;

      try {
        worksheet.unMergeCells(mergeRange);
      } catch (err) {}
      worksheet.mergeCells(mergeRange);

      worksheet.getCell(`B${MAE_START_ROW}`).value = "Sin maestria";
      worksheet.getCell(`B${MAE_START_ROW}`).alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: true,
      };
      worksheet.getRow(MAE_START_ROW).hidden = true;

      ///worksheet.getRow(MAE_START_ROW+1).hidden = true;
    } else {
      if (maeCount > 1) {
        acarreo = acarreo + maeCount - 1;
        worksheet.duplicateRow(MAE_START_ROW, maeCount - 1, true);
      }

      const start = MAE_START_ROW;
      const end = MAE_START_ROW + maeCount - 1;

      // Encabezado en A (si lo necesitas)

      maestrias.forEach((exp, i) => {
        const row = start + i;
        worksheet.getCell(`B${row}`).value = exp.nombreTitulo;
        worksheet.getCell(`C${row}`).value = exp.fechaObtencion;
        worksheet.getCell(`D${row}`).value = exp.institucion;
        worksheet.getCell(`E${row}`).value = exp.documentoAdquirido;
        worksheet.getCell(`F${row}`).value = exp.estatus;
        worksheet.getCell(`G${row}`).value = exp.cedulaProfesional;
      });

      if (maeCount > 1) {
        const mergeRange = `A${start}:A${end}`;
        try {
          worksheet.unMergeCells(mergeRange);
        } catch (err) {}
        try {
          worksheet.mergeCells(mergeRange);
        } catch (err) {
          console.warn(`No se pudo mergear ${mergeRange}:`, err);
        }
      }
    }
    ///////////////////////////////////////////////

    // console.log("acarreo mae: ", acarreo); // Depuración de desplazamiento de filas

    let DOC_START_ROW = 17 + acarreo; // ← ajusta esta fila al inicio de tu sección "Maestría"
    const doctorados = escolaridades.filter(
      (e) => String(e.nivelAcademico || "").toUpperCase() === "DOCTORADO",
    );
    let docCount = doctorados.length;
    if (docCount === 0) {
      const mergeRange = `B${DOC_START_ROW}:G${DOC_START_ROW}`;

      try {
        worksheet.unMergeCells(mergeRange);
      } catch (err) {}
      worksheet.mergeCells(mergeRange);

      worksheet.getCell(`B${DOC_START_ROW}`).value = "Sin doctorado";
      worksheet.getCell(`B${DOC_START_ROW}`).alignment = {
        horizontal: "center",
      };

      worksheet.getRow(DOC_START_ROW).hidden = true;
      ///worksheet.getRow(MAE_START_ROW+1).hidden = true;
    } else {
      if (docCount > 1) {
        acarreo = acarreo + docCount - 1;
        worksheet.duplicateRow(DOC_START_ROW, docCount - 1, true);
      }

      const start = DOC_START_ROW;
      const end = DOC_START_ROW + docCount - 1;

      // Encabezado en A (si lo necesitas)

      doctorados.forEach((exp, i) => {
        const row = start + i;
        worksheet.getCell(`B${row}`).value = exp.nombreTitulo;
        worksheet.getCell(`C${row}`).value = exp.fechaObtencion;
        worksheet.getCell(`D${row}`).value = exp.institucion;
        worksheet.getCell(`E${row}`).value = exp.documentoAdquirido;
        worksheet.getCell(`F${row}`).value = exp.estatus;
        worksheet.getCell(`G${row}`).value = exp.cedulaProfesional;
      });

      if (docCount > 1) {
        const mergeRange = `A${start}:A${end}`;
        try {
          worksheet.unMergeCells(mergeRange);
        } catch (err) {}
        try {
          worksheet.mergeCells(mergeRange);
        } catch (err) {
          console.warn(`No se pudo mergear ${mergeRange}:`, err);
        }
      }
    }
    // console.log("acarreo doc: ", acarreo);
    let ESP_START_ROW = 18 + acarreo; // ← ajusta esta fila al inicio de tu sección "Maestría"
    const especialidades = escolaridades.filter(
      (e) => String(e.nivelAcademico || "").toUpperCase() === "ESPECIALIDAD",
    );
    let espCount = especialidades.length;
    if (espCount === 0) {
      const mergeRange = `B${ESP_START_ROW}:G${ESP_START_ROW}`;

      try {
        worksheet.unMergeCells(mergeRange);
      } catch (err) {}

      try {
        worksheet.mergeCells(mergeRange);
      } catch (err) {}

      worksheet.getCell(`B${ESP_START_ROW}`).value = "Sin especialidad";
      worksheet.getCell(`B${ESP_START_ROW}`).alignment = {
        horizontal: "center",
      };

      worksheet.getRow(ESP_START_ROW).hidden = true;

      ///worksheet.getRow(MAE_START_ROW+1).hidden = true;
    } else {
      if (espCount > 1) {
        acarreo = acarreo + espCount - 1;
        worksheet.duplicateRow(ESP_START_ROW, espCount - 1, true);
      }

      const start = ESP_START_ROW;

      const end = ESP_START_ROW + espCount - 1;

      // Encabezado en A (si lo necesitas)

      especialidades.forEach((exp, i) => {
        const row = start + i;
        worksheet.getCell(`B${row}`).value = exp.nombreTitulo;
        worksheet.getCell(`C${row}`).value = exp.fechaObtencion;
        worksheet.getCell(`D${row}`).value = exp.institucion;
        worksheet.getCell(`E${row}`).value = exp.documentoAdquirido;
        worksheet.getCell(`F${row}`).value = exp.estatus;
        worksheet.getCell(`G${row}`).value = exp.cedulaProfesional;
      });

      if (espCount > 1) {
        const mergeRange = `A${start}:A${end}`;
        try {
          worksheet.unMergeCells(mergeRange);
        } catch (err) {}
        try {
          worksheet.mergeCells(mergeRange);
        } catch (err) {
          console.warn(`No se pudo mergear ${mergeRange}:`, err);
        }
      }
    }
    const rutaSalida = "./reportes/";
    const nombreArchivo = `${trabajador.nombreTrabajador}_CURRICULUM-EDITABLE.xlsx`;
    const rutaCompleta = path.join(rutaSalida, nombreArchivo);

    await workbook.xlsx.writeFile(rutaCompleta);
    //////////////////////////////////////////////////////////////////////
    const rutaPlantilla2 = "./plantillas/curriculum2.xlsx";

    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(rutaPlantilla2);
    const worksheet2 = workbook2.worksheets[0];

    worksheet2.getCell("B7").value = trabajador.nombreTrabajador;
    worksheet2.getCell("B11").value = trabajador.primaria;
    worksheet2.getCell("B12").value = trabajador.secundaria;
    worksheet2.getCell("B13").value = trabajador.carreraTecnicaComercial;

    bachillerato.forEach((exp, index) => {
      const fila = 14 + index;
      worksheet2.getCell(`B${fila}`).value = exp.nombreTitulo;
      worksheet2.getCell(`C${fila}`).value = exp.fechaObtencion;
      worksheet2.getCell(`D${fila}`).value = exp.institucion;
    });

    ///////////////////////////////////////////////////////////

    const expMerges2 = [];
    worksheet2.model.merges.forEach((range) => {
      const [start, end] = range.split(":");
      const startRow = +start.match(/\d+$/)[0];
      const endRow = +end.match(/\d+$/)[0];
      if (startRow === EXPERIENCE_ROW && endRow === EXPERIENCE_ROW) {
        expMerges2.push(range);
      }
    });

    if (totalExp === 0) {
      const mergeRange = `A26:E26`;

      try {
        worksheet2.unMergeCells(mergeRange);
      } catch (err) {}
      worksheet2.mergeCells(mergeRange);
      worksheet2.getCell("A26").value = "Sin Experiencias";
      worksheet2.getCell("A26").alignment = {
        horizontal: "center",
      };
    }
    if (totalExp > 1) {
      worksheet2.duplicateRow(EXPERIENCE_ROW, totalExp - 1, true);
    }

    expMerges2.forEach((range) => {
      const [start, end] = range.split(":");
      const colStart = start.match(/^[A-Z]+/)[0];
      const colEnd = end.match(/^[A-Z]+/)[0];
      for (let i = 1; i < totalExp; i++) {
        worksheet2.mergeCells(
          `${colStart}${EXPERIENCE_ROW + i}:${colEnd}${EXPERIENCE_ROW + i}`,
        );
      }
    });

    experienciasPJ.forEach((exp, idx) => {
      const r = EXPERIENCE_ROW + idx;
      worksheet2.getCell(`A${r}`).value = exp.periodoInicio;
      worksheet2.getCell(`B${r}`).value = exp.periodoFin;
      worksheet2.getCell(`C${r}`).value = exp.institucion;

      if (
        worksheet2.getCell(`C${r}`).value ==
        "PODER JUDICIAL DEL ESTADO DE HIDALGO"
      ) {
        worksheet2.getCell(`C${r}`).value =
          exp.institucion + "/" + exp.adscripcion;
      }
      worksheet2.getCell(`D${r}`).value = exp.cargo;
      worksheet2.getCell(`E${r}`).value = exp.campoExperiencia;
    });
    ///////////////////////////////////////////////////

    const updateMerges2 = [];
    worksheet2.model.merges.forEach((range) => {
      const [start, end] = range.split(":");
      const startRow = parseInt(start.match(/\d+$/)[0], 10);
      const endRow = parseInt(end.match(/\d+$/)[0], 10);
      if (startRow === UPDATE_ROW && endRow === UPDATE_ROW) {
        updateMerges2.push(range);
      }
    });

    if (totalActualizaciones > 1) {
      worksheet2.duplicateRow(UPDATE_ROW, totalActualizaciones - 1, true);
    }

    if (totalActualizaciones === 0) {
      const mergeRange = `A22:D22`;

      try {
        worksheet2.unMergeCells(mergeRange);
      } catch (err) {}
      worksheet2.mergeCells(mergeRange);
      worksheet2.getCell("B22").value = "Sin Actualizaciones";
      worksheet2.getCell("B22").alignment = {
        horizontal: "center",
      };
    }

    updateMerges2.forEach((range) => {
      const [start, end] = range.split(":");
      const colStart = start.match(/^[A-Z]+/)[0];
      const colEnd = end.match(/^[A-Z]+/)[0];
      for (let i = 1; i < totalActualizaciones; i++) {
        const mergeRange = `${colStart}${UPDATE_ROW + i}:${colEnd}${UPDATE_ROW + i}`;
        console.log("rango de merge1: ", mergeRange);
        try {
          worksheet2.unMergeCells(mergeRange);
        } catch (err) {}
        try {
          worksheet2.mergeCells(mergeRange);
        } catch (err) {
          console.warn(`No se pudo mergear ${mergeRange}:`, err.message);
        }
      }
    });

    actualizaciones.forEach((exp, idx) => {
      const r = UPDATE_ROW + idx;
      worksheet2.getCell(`A${r}`).value = exp.tema;
      worksheet2.getCell(`B${r}`).value = formatearFecha(exp.fecha);
      worksheet2.getCell(`C${r}`).value = exp.institucion;
      worksheet2.getCell(`D${r}`).value = exp.documento;
    });

    // 1) Si no hay ninguna, borramos la fila base
    if (prepaCount === 0) {
      const mergeRange = `B14:G14`;

      try {
        worksheet2.unMergeCells(mergeRange);
      } catch (err) {}
      worksheet2.mergeCells(mergeRange);
      worksheet2.getCell("B14").value = "No";
      worksheet2.getCell("B14").alignment = {
        horizontal: "center",
      };
      worksheet2.getRow(PREPA_START_ROW).hidden = true;
    } else {
      // 2) Si hay más de 1, duplicamos la fila base licCount-1 veces
      if (prepaCount > 1) {
        // (filaBase, númeroDeDuplicados, insert => true para desplazar filas hacia abajo)
        worksheet2.duplicateRow(PREPA_START_ROW, prepaCount - 1, true);
      }

      const start = PREPA_START_ROW;
      const end = PREPA_START_ROW + prepaCount - 1;

      prepa.forEach((exp, i) => {
        const row = start + i;
        worksheet2.getCell(`B${row}`).value = exp.nombreTitulo;
        worksheet2.getCell(`C${row}`).value = exp.fechaObtencion;
        worksheet2.getCell(`D${row}`).value = exp.institucion;
      });

      // 5) Finalmente: si había >1, combinamos la columna A de start a end
      if (prepaCount > 1) {
        const mergeRange = `A${start}:A${end}`;

        try {
          // Intenta desfusionar (no lanza error si no existe)
          worksheet2.unMergeCells(mergeRange);
        } catch (err) {
          // Ignora el error si el merge no existe
        }

        // Fusiona las celdas
        worksheet2.mergeCells(mergeRange);
      }
    }

    licCount = licenciaturas.length;
    console.log("lenght lic:", licCount);
    acarreo = licCount;

    console.log("acarreo lic: ", acarreo);
    licDesplazar = 15 + acarreo;

    // 1) Si no hay ninguna, borramos la fila base
    if (licCount === 0) {
      const mergeRange = `B${licDesplazar}:G${licDesplazar}`;

      try {
        worksheet2.unMergeCells(mergeRange);
      } catch (err) {}
      worksheet2.mergeCells(mergeRange);

      console.log("licdesplazar: ", licDesplazar);
      worksheet2.getCell(`B${licDesplazar}`).value = "Sin licenciatura";
      worksheet2.getCell(`B${licDesplazar}`).alignment = {
        horizontal: "center",
      };
      worksheet2.getRow(LIC_START_ROW).hidden = true;
    } else {
      // 2) Si hay más de 1, duplicamos la fila base licCount-1 veces

      if (licCount === 1) {
        acarreo = acarreo - 1;
      }
      if (licCount > 1) {
        acarreo = acarreo - 1;

        console.log("acarrero primero lic: ", acarreo);
        // (filaBase, númeroDeDuplicados, insert => true para desplazar filas hacia abajo)
        worksheet2.duplicateRow(LIC_START_ROW, licCount - 1, true);
      }

      const start = LIC_START_ROW;
      const end = LIC_START_ROW + licCount - 1;

      licenciaturas.forEach((exp, i) => {
        const row = start + i;
        worksheet2.getCell(`B${row}`).value = exp.nombreTitulo;
        worksheet2.getCell(`C${row}`).value = exp.fechaObtencion;
        worksheet2.getCell(`D${row}`).value = exp.institucion;
      });

      // 5) Finalmente: si había >1, combinamos la columna A de start a end
      if (licCount > 1) {
        const mergeRange = `A${start}:A${end}`;

        try {
          // Intenta desfusionar (no lanza error si no existe)
          worksheet2.unMergeCells(mergeRange);
        } catch (err) {
          // Ignora el error si el merge no existe
        }

        // Fusiona las celdas
        worksheet2.mergeCells(mergeRange);
      }
    }

    ////////////////////////////////////////////////////////////////////////

    MAE_START_ROW = 16 + acarreo; // ← ajusta esta fila al inicio de tu sección "Maestría"

    if (maeCount === 0) {
      console.log("MAE ", MAE_START_ROW);

      const mergeRange = `B${MAE_START_ROW}:G${MAE_START_ROW}`;

      try {
        worksheet2.unMergeCells(mergeRange);
      } catch (err) {}
      worksheet2.mergeCells(mergeRange);

      worksheet2.getCell(`B${MAE_START_ROW}`).value = "Sin maestria";
      worksheet2.getCell(`B${MAE_START_ROW}`).alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: true,
      };
      worksheet2.getRow(MAE_START_ROW).hidden = true;

      ///worksheet.getRow(MAE_START_ROW+1).hidden = true;
    } else {
      if (maeCount > 1) {
        acarreo = acarreo + maeCount - 1;
        worksheet2.duplicateRow(MAE_START_ROW, maeCount - 1, true);
      }

      const start = MAE_START_ROW;

      console.log(start, " starto ", licDesplazar);
      const end = MAE_START_ROW + maeCount - 1;

      maestrias.forEach((exp, i) => {
        const row = start + i;
        worksheet2.getCell(`B${row}`).value = exp.nombreTitulo;
        worksheet2.getCell(`C${row}`).value = exp.fechaObtencion;
        worksheet2.getCell(`D${row}`).value = exp.institucion;
      });

      if (maeCount > 1) {
        const mergeRange = `A${start}:A${end}`;
        try {
          worksheet2.unMergeCells(mergeRange);
        } catch (err) {}
        try {
          worksheet2.mergeCells(mergeRange);
        } catch (err) {
          console.warn(`No se pudo mergear ${mergeRange}:`, err);
        }
      }
    }

    ///////////////////////////////////////////////
    // console.log("acarreo mae: ", acarreo);

    DOC_START_ROW = 17 + acarreo; // ← ajusta esta fila al inicio de tu sección "Maestría"

    if (docCount === 0) {
      const mergeRange = `B${DOC_START_ROW}:G${DOC_START_ROW}`;

      try {
        worksheet2.unMergeCells(mergeRange);
      } catch (err) {}
      worksheet2.mergeCells(mergeRange);

      worksheet2.getCell(`B${DOC_START_ROW}`).value = "Sin doctorado";
      worksheet2.getCell(`B${DOC_START_ROW}`).alignment = {
        horizontal: "center",
      };

      worksheet2.getRow(DOC_START_ROW).hidden = true;
      ///worksheet.getRow(MAE_START_ROW+1).hidden = true;
    } else {
      if (docCount > 1) {
        acarreo = acarreo + docCount - 1;
        worksheet2.duplicateRow(DOC_START_ROW, docCount - 1, true);
      }

      const start = DOC_START_ROW;

      const end = DOC_START_ROW + docCount - 1;

      // Encabezado en A (si lo necesitas)

      doctorados.forEach((exp, i) => {
        const row = start + i;
        worksheet2.getCell(`B${row}`).value = exp.nombreTitulo;
        worksheet2.getCell(`C${row}`).value = exp.fechaObtencion;
        worksheet2.getCell(`D${row}`).value = exp.institucion;
      });

      if (docCount > 1) {
        const mergeRange = `A${start}:A${end}`;
        try {
          worksheet2.unMergeCells(mergeRange);
        } catch (err) {}
        try {
          worksheet2.mergeCells(mergeRange);
        } catch (err) {
          console.warn(`No se pudo mergear ${mergeRange}:`, err);
        }
      }
    }
    // console.log("acarreo doc: ", acarreo);

    ESP_START_ROW = 18 + acarreo; // ← ajusta esta fila al inicio de tu sección "Maestría"

    if (espCount === 0) {
      console.log("ESP ", ESP_START_ROW);

      const mergeRange = `B${ESP_START_ROW}:G${ESP_START_ROW}`;
      console.log("mergeRange: ", mergeRange);

      try {
        worksheet2.unMergeCells(mergeRange);
      } catch (err) {}

      try {
        worksheet2.mergeCells(mergeRange);
      } catch (err) {}

      worksheet2.getCell(`B${ESP_START_ROW}`).value = "Sin especialidad";
      worksheet2.getCell(`B${ESP_START_ROW}`).alignment = {
        horizontal: "center",
      };

      worksheet2.getRow(ESP_START_ROW).hidden = true;

      ///worksheet.getRow(MAE_START_ROW+1).hidden = true;
    } else {
      if (espCount > 1) {
        acarreo = acarreo + espCount - 1;
        worksheet2.duplicateRow(ESP_START_ROW, espCount - 1, true);
      }

      const start = ESP_START_ROW;

      const end = ESP_START_ROW + espCount - 1;

      // Encabezado en A (si lo necesitas)

      especialidades.forEach((exp, i) => {
        const row = start + i;
        worksheet2.getCell(`B${row}`).value = exp.nombreTitulo;
        worksheet2.getCell(`C${row}`).value = exp.fechaObtencion;
        worksheet2.getCell(`D${row}`).value = exp.institucion;
      });

      if (espCount > 1) {
        const mergeRange = `A${start}:A${end}`;
        try {
          worksheet2.unMergeCells(mergeRange);
        } catch (err) {}
        try {
          worksheet2.mergeCells(mergeRange);
        } catch (err) {
          console.warn(`No se pudo mergear ${mergeRange}:`, err);
        }
      }
    }

    if (licCount >= 1 || maeCount >= 1 || docCount >= 1 || espCount >= 1) {
      console.log("Si hay más de una");
      worksheet2.getRow(11).hidden = true;
      worksheet2.getRow(12).hidden = true;
      worksheet2.getRow(13).hidden = true;
      worksheet2.getRow(14).hidden = true;
    }

    const { exec } = require("child_process");

    const _uid = `${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    const rutaExcel = `./reportes/${trabajador.nombreTrabajador}_${_uid}_temp.xlsx`;
    const rutaPDF = `./reportes/${trabajador.nombreTrabajador}_${_uid}_CURRICULUM_TRANSPARENCIA.pdf`;
    const rutaHTML = `./reportes/${trabajador.nombreTrabajador}_${_uid}_temp.html`;
    // 1. Guardar Excel temporal
    await workbook2.xlsx.writeFile(rutaExcel);

    await excelToHtmlWithStyles(rutaExcel, rutaHTML);

    const browser = await puppeteer.launch({
      headless: "new",
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });
    const page = await browser.newPage();
    // Carga el HTML recién creado — convertir ruta Windows a URL file:/// válida
    const fileUrlHTML = `file:///${path.resolve(rutaHTML).replace(/\\/g, "/")}`;
    await page.goto(fileUrlHTML, {
      waitUntil: "networkidle0",
    });
    await page.pdf({
      path: rutaPDF,
      format: "letter",
      printBackground: true,
      margin: { top: "1cm", right: "1cm", bottom: "1cm", left: "1cm" },
    });
    await browser.close();

    // Limpiar archivos temporales de Excel y HTML (sin fallar si no existen)
    [rutaExcel, rutaHTML].forEach((file) => {
      try {
        if (fs.existsSync(file)) {
          fs.unlinkSync(file);
        }
      } catch (e) {
        console.warn(
          `No se pudo eliminar archivo temporal ${file}:`,
          e.message,
        );
      }
    });

    if (previewOnly) {
      const nombrePdfPreview = `${trabajador.nombreTrabajador}_CURRICULUM_TRANSPARENCIA.pdf`;
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader(
        "Content-Disposition",
        `inline; filename="${nombrePdfPreview}"`,
      );

      return res.sendFile(path.resolve(rutaPDF), (err) => {
        if (err) {
          console.error("Error al enviar preview de curriculum:", err);
        }

        [rutaPDF, rutaCompleta].forEach((file) => {
          try {
            if (fs.existsSync(file)) {
              fs.unlinkSync(file);
            }
          } catch (e) {
            console.warn(
              `No se pudo eliminar archivo temporal ${file}:`,
              e.message,
            );
          }
        });
      });
    }

    const archiver = require("archiver");

    // Crear el archivo ZIP
    const zipNombre = `reporte_${trabajador.nombreTrabajador}.zip`;
    const zipRuta = path.join("./reportes", zipNombre);
    const output = fs.createWriteStream(zipRuta);
    const archive = archiver("zip", { zlib: { level: 9 } });

    output.on("close", () => {
      console.log(`📦 ZIP creado con ${archive.pointer()} bytes`);
      res.download(zipRuta, (err) => {
        if (err) {
          if (err.code === "ECONNABORTED" || err.code === "ECONNRESET") {
            console.warn("El cliente canceló la descarga del ZIP.");
          } else {
            console.error("Error al enviar ZIP:", err);
          }
        } else {
          // Limpiar ZIP y Excel generado para este trabajador
          [zipRuta, rutaCompleta].forEach((file) => {
            try {
              if (fs.existsSync(file)) {
                fs.unlinkSync(file);
              }
            } catch (e) {
              console.warn(
                `No se pudo eliminar archivo generado ${file}:`,
                e.message,
              );
            }
          });
          ///fs.unlinkSync(rutaCompleta25);
        }
      });
    });

    archive.on("error", (err) => {
      throw err;
    });

    archive.pipe(output);
    const carpetaTrabajador = trabajador.nombreTrabajador; // sin modificaciones
    // Adjunta el Excel generado
    archive.file(rutaCompleta, {
      name: `${carpetaTrabajador}/${path.basename(rutaCompleta)}`,
    });
    archive.file(rutaPDF, {
      name: `${carpetaTrabajador}/${path.basename(rutaPDF)}`,
    });
    //archive.file(rutaCompleta25, { name: path.basename(rutaCompleta25) });

    // Adjunta todos los anexos del trabajador que existan en uploads
    const uploadsPath = path.join(__dirname, "uploads");
    const normalizarNombre = (valor) =>
      String(valor || "")
        .trim()
        .toUpperCase()
        .replace(/\s+/g, " ")
        .split(" ")
        .filter(Boolean)
        .sort()
        .join(" ");

    const nombreTrabajadorNormalizado = normalizarNombre(
      trabajador.nombreTrabajador,
    );

    const anexosTrabajador = fs.readdirSync(uploadsPath).filter((nombre) => {
      const ruta = path.join(uploadsPath, nombre);
      let stat;
      try {
        stat = fs.statSync(ruta);
      } catch (err) {
        return false;
      }
      if (!stat.isFile()) return false;

      const baseSinExt = path.parse(nombre).name;
      const idxUltimoGuionBajo = baseSinExt.lastIndexOf("_");
      if (idxUltimoGuionBajo <= 0) return false;

      const nombreEnArchivo = baseSinExt.slice(0, idxUltimoGuionBajo);
      const nombreArchivoNormalizado = normalizarNombre(nombreEnArchivo);
      return nombreArchivoNormalizado === nombreTrabajadorNormalizado;
    });

    anexosTrabajador.forEach((nombre) => {
      const ruta = path.join(uploadsPath, nombre);
      archive.file(ruta, { name: `${carpetaTrabajador}/${nombre}` });
    });

    await archive.finalize();

    // El PDF ya va incluido en el ZIP; puedes limpiarlo aquí
    try {
      if (fs.existsSync(rutaPDF)) {
        fs.unlinkSync(rutaPDF);
      }
    } catch (e) {
      console.warn(`No se pudo eliminar PDF temporal ${rutaPDF}:`, e.message);
    }
  } catch (error) {
    console.error("Error generando el reporte:", error);
    if (req.query.debug === "1") {
      return res
        .status(500)
        .send(`Error interno del servidor: ${error.message}`);
    }
    res.status(500).send("Error interno del servidor");
  }
});

app.get("/generar-reportes-todos", async (req, res) => {
  const archiver = require("archiver");
  try {
    // Obtener todos los trabajadores
    const [todosTrabajadores] = await db.query("SELECT * FROM trabajador");

    // Crear el archivo ZIP principal
    const zipNombre = `todos_trabajadores_${Date.now()}.zip`;
    const zipRuta = path.join("./reportes", zipNombre);
    const output = fs.createWriteStream(zipRuta);
    const archive = archiver("zip", { zlib: { level: 9 } });

    // Manejador para cuando se cierre el ZIP
    output.on("close", () => {
      console.log(`📦 ZIP creado con ${archive.pointer()} bytes`);
      res.download(zipRuta, (err) => {
        if (err) {
          if (err.code === "ECONNABORTED" || err.code === "ECONNRESET") {
            console.warn("El cliente canceló la descarga del ZIP general.");
          } else {
            console.error("Error al enviar ZIP:", err);
          }
        } else {
          // Limpiar archivos temporales después de enviar
          try {
            if (fs.existsSync(zipRuta)) {
              fs.unlinkSync(zipRuta);
            }
          } catch (e) {
            console.warn(
              `No se pudo eliminar ZIP general ${zipRuta}:`,
              e.message,
            );
          }

          // Eliminar archivos generados para cada trabajador
          todosTrabajadores.forEach((t) => {
            const base = `./reportes/${t.nombreTrabajador}_`;
            [
              `${base}CURRICULUM-EDITABLE.xlsx`,
              `${base}CURRICULUM_TRANSPARENCIA.pdf`,
            ].forEach((file) => {
              try {
                if (fs.existsSync(file)) {
                  fs.unlinkSync(file);
                }
              } catch (e) {
                console.warn(
                  `No se pudo eliminar archivo de ${t.nombreTrabajador} (${file}):`,
                  e.message,
                );
              }
            });
          });
        }
      });
    });

    archive.on("error", (err) => {
      throw err;
    });

    archive.pipe(output);

    // ===================================================================
    // INICIO DEL CICLO PARA CADA TRABAJADOR - AQUÍ VA TU CÓDIGO EXISTENTE
    // ===================================================================
    for (const trabajador of todosTrabajadores) {
      const noTrabajador = trabajador.noTrabajador;
      const nombreTrabajador = trabajador.nombreTrabajador;

      try {
        // ------ COMIENZA TU CÓDIGO ACTUAL (desde const rutaPlantilla...) ------
        const rutaPlantilla = "./plantillas/curriculum.xlsx";

        const [trabajadorRows] = await db.query(
          "SELECT * FROM trabajador WHERE idTrabajador = ?",
          [trabajador.idTrabajador],
        );
        const trabajadorData = trabajadorRows[0];
        const idTrabajador = trabajador.idTrabajador;

        if (!trabajadorData) {
          return res.status(404).send("No se encontró el trabajador");
        }

        const [hijos] = await db.query(
          "SELECT * FROM hijo WHERE idTrabajador = ?",
          [idTrabajador],
        );

        const [escolaridades] = await db.query(
          "SELECT * FROM escolaridad WHERE idTrabajador = ?",
          [idTrabajador],
        );

        const [actualizaciones] = await db.query(
          "SELECT * FROM actualizacionprofesional WHERE idTrabajador = ?",
          [idTrabajador],
        );
        ///const [experiencias] = await db.query('SELECT * FROM experienciapj WHERE idTrabajador = ?', [idTrabajador]);

        ///console.log(experiencias)

        const [experiencias] = await db.query(
          "SELECT * FROM experienciapj WHERE idTrabajador = ?",
          [idTrabajador],
        );

        console.log("Resultados:", experiencias);
        console.log("Total:", experiencias.length);
        console.log("ID del trabajador recibido:", idTrabajador);

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(rutaPlantilla);
        const worksheet = workbook.worksheets[0];
        worksheet.getCell("B7").value = trabajador.nombreTrabajador;
        worksheet.getCell("B11").value = trabajador.primaria;
        worksheet.getCell("B12").value = trabajador.secundaria;
        worksheet.getCell("B13").value = trabajador.carreraTecnicaComercial;

        const bachillerato = escolaridades.filter(
          (exp) => exp.nivelAcademico.toUpperCase() === "BACHILLERATO",
        );
        bachillerato.forEach((exp, index) => {
          const fila = 14 + index; // Empieza en la fila 14, 15, 16…
          // Usando template literals con backticks:
          worksheet.getCell(`B${fila}`).value = exp.nombreTitulo;
          worksheet.getCell(`C${fila}`).value = exp.fechaObtencion;
          worksheet.getCell(`D${fila}`).value = exp.institucion;
          worksheet.getCell(`E${fila}`).value = exp.documentoAdquirido;
          worksheet.getCell(`F${fila}`).value = exp.estatus;
        });

        /*worksheet.duplicateRow(29,1, true);
                    worksheet.getCell("A30").value = "HOLAAA";
            */

        ///////////////////////////////////////////////////////////

        const EXPERIENCE_ROW = 26;
        const expMerges = [];
        worksheet.model.merges.forEach((range) => {
          const [start, end] = range.split(":");
          const startRow = +start.match(/\d+$/)[0];
          const endRow = +end.match(/\d+$/)[0];
          if (startRow === EXPERIENCE_ROW && endRow === EXPERIENCE_ROW) {
            expMerges.push(range);
          }
        });

        const experienciasPJ = experiencias;
        const totalExp = experienciasPJ.length;

        if (totalExp === 0) {
          const mergeRange = `A26:D26`;

          try {
            worksheet.unMergeCells(mergeRange);
          } catch (err) {}
          worksheet.mergeCells(mergeRange);
          worksheet.getCell("A26").value = "Sin Experiencias";
          worksheet.getCell("A26").alignment = {
            horizontal: "center",
          };
        }
        if (totalExp > 1) {
          worksheet.duplicateRow(EXPERIENCE_ROW, totalExp - 1, true);
        }

        expMerges.forEach((range) => {
          const [start, end] = range.split(":");
          const colStart = start.match(/^[A-Z]+/)[0];
          const colEnd = end.match(/^[A-Z]+/)[0];
          for (let i = 1; i < totalExp; i++) {
            worksheet.mergeCells(
              `${colStart}${EXPERIENCE_ROW + i}:${colEnd}${EXPERIENCE_ROW + i}`,
            );
          }
        });

        experienciasPJ.forEach((exp, idx) => {
          const r = EXPERIENCE_ROW + idx;
          worksheet.getCell(`A${r}`).value = exp.periodoInicio;
          worksheet.getCell(`B${r}`).value = exp.periodoFin;
          worksheet.getCell(`C${r}`).value =
            exp.institucion + " / " + exp.adscripcion;
          worksheet.getCell(`D${r}`).value = exp.cargo;
          worksheet.getCell(`E${r}`).value = exp.campoExperiencia;
        });

        ///////////////////////////////////////////////////

        const UPDATE_ROW = 22;
        const updateMerges = [];

        worksheet.model.merges.forEach((range) => {
          const [start, end] = range.split(":");
          const startRow = parseInt(start.match(/\d+$/)[0], 10);
          const endRow = parseInt(end.match(/\d+$/)[0], 10);
          if (startRow === UPDATE_ROW && endRow === UPDATE_ROW) {
            updateMerges.push(range);
          }
        });

        const totalActualizaciones = actualizaciones.length;
        if (totalActualizaciones > 1) {
          worksheet.duplicateRow(UPDATE_ROW, totalActualizaciones - 1, true);
        }

        if (totalActualizaciones === 0) {
          const mergeRange = `A22:D22`;

          try {
            worksheet.unMergeCells(mergeRange);
          } catch (err) {}
          worksheet.mergeCells(mergeRange);
          worksheet.getCell("B22").value = "Sin Actualizaciones";
          worksheet.getCell("B22").alignment = {
            horizontal: "center",
          };
        }

        updateMerges.forEach((range) => {
          const [start, end] = range.split(":");
          const colStart = start.match(/^[A-Z]+/)[0];
          const colEnd = end.match(/^[A-Z]+/)[0];
          for (let i = 1; i < totalActualizaciones; i++) {
            console.log(
              "rango de merge1: ",
              `${colStart}${UPDATE_ROW + i}:${colEnd}${UPDATE_ROW + i}`,
            );
            worksheet.mergeCells(
              `${colStart}${UPDATE_ROW + i}:${colEnd}${UPDATE_ROW + i}`,
            );
          }
        });

        actualizaciones.forEach((exp, idx) => {
          const r = UPDATE_ROW + idx;
          worksheet.getCell(`A${r}`).value = exp.tema;
          worksheet.getCell(`B${r}`).value = formatearFecha(exp.fecha);
          worksheet.getCell(`C${r}`).value = exp.institucion;
          worksheet.getCell(`D${r}`).value = exp.documento;
          /*const mergeRange = `C${r}:D${r}`;
                        console.log(mergeRange);
                        try{
                        worksheet.mergeCells(mergeRange);
                        }catch(err){
                
                        }*/
        });

        const PREPA_START_ROW = 14;
        const prepa = escolaridades.filter(
          (e) => e.nivelAcademico.toUpperCase() === "BACHILLERATO",
        );
        const prepaCount = prepa.length;

        // 1) Si no hay ninguna, borramos la fila base
        if (prepaCount === 0) {
          const mergeRange = `B14:G14`;

          try {
            worksheet.unMergeCells(mergeRange);
          } catch (err) {}
          worksheet.mergeCells(mergeRange);
          worksheet.getCell("B14").value = "No";
          worksheet.getCell("B14").alignment = {
            horizontal: "center",
          };
          worksheet.getRow(PREPA_START_ROW).hidden = true;
        } else {
          // 2) Si hay más de 1, duplicamos la fila base licCount-1 veces
          if (prepaCount > 1) {
            // (filaBase, númeroDeDuplicados, insert => true para desplazar filas hacia abajo)
            worksheet.duplicateRow(LIC_START_ROW, licCount - 1, true);
          }

          const start = PREPA_START_ROW;
          const end = PREPA_START_ROW + prepaCount - 1;

          // 3) Asignamos el encabezado en la primera fila
          ///worksheet.getCell(`A${start}`).value = "hola";

          /*
                        worksheet.getCell(`A${start}`).alignment = {
                          vertical:   "middle",
                          horizontal: "center",
                          wrapText:   true
                        };*/

          // 4) Llenamos cada fila duplicada con los datos
          prepa.forEach((exp, i) => {
            const row = start + i;
            worksheet.getCell(`B${row}`).value = exp.nombreTitulo;
            worksheet.getCell(`C${row}`).value = exp.fechaObtencion;
            worksheet.getCell(`D${row}`).value = exp.institucion;
            worksheet.getCell(`E${row}`).value = exp.documentoAdquirido;
            worksheet.getCell(`F${row}`).value = exp.estatus;
          });

          // 5) Finalmente: si había >1, combinamos la columna A de start a end
          // 5) Finalmente: si había >1, combinamos la columna A de start a end
          if (prepaCount > 1) {
            const mergeRange = `A${start}:A${end}`;

            try {
              // Intenta desfusionar (no lanza error si no existe)
              worksheet.unMergeCells(mergeRange);
            } catch (err) {
              // Ignora el error si el merge no existe
            }

            // Fusiona las celdas
            worksheet.mergeCells(mergeRange);
          }
        }

        let LIC_START_ROW = 15;
        const licenciaturas = escolaridades.filter(
          (e) => e.nivelAcademico.toUpperCase() === "LICENCIATURA",
        );
        let licCount = licenciaturas.length;
        console.log("lenght lic:", licCount);
        let acarreo = licCount;

        console.log("acarreo lic: ", acarreo);
        let licDesplazar = 15 + acarreo;

        // 1) Si no hay ninguna, borramos la fila base
        if (licCount === 0) {
          const mergeRange = `B${licDesplazar}:G${licDesplazar}`;

          try {
            worksheet.unMergeCells(mergeRange);
          } catch (err) {}
          worksheet.mergeCells(mergeRange);

          console.log("licdesplazar: ", licDesplazar);
          worksheet.getCell(`B${licDesplazar}`).value = "Sin licenciatura";
          worksheet.getCell(`B${licDesplazar}`).alignment = {
            horizontal: "center",
          };
          worksheet.getRow(LIC_START_ROW).hidden = true;
        } else {
          // 2) Si hay más de 1, duplicamos la fila base licCount-1 veces

          if (licCount === 1) {
            acarreo = acarreo - 1;
          }
          if (licCount > 1) {
            acarreo = acarreo - 1;

            console.log("acarrero primero lic: ", acarreo);
            // (filaBase, númeroDeDuplicados, insert => true para desplazar filas hacia abajo)
            worksheet.duplicateRow(LIC_START_ROW, licCount - 1, true);
          }

          const start = LIC_START_ROW;
          const end = LIC_START_ROW + licCount - 1;

          // 3) Asignamos el encabezado en la primera fila
          ///worksheet.getCell(`A${start}`).value = "hola";

          /*
                        worksheet.getCell(`A${start}`).alignment = {
                          vertical:   "middle",
                          horizontal: "center",
                          wrapText:   true
                        };*/

          // 4) Llenamos cada fila duplicada con los datos
          licenciaturas.forEach((exp, i) => {
            const row = start + i;
            worksheet.getCell(`B${row}`).value = exp.nombreTitulo;
            worksheet.getCell(`C${row}`).value = exp.fechaObtencion;
            worksheet.getCell(`D${row}`).value = exp.institucion;
            worksheet.getCell(`E${row}`).value = exp.documentoAdquirido;
            worksheet.getCell(`F${row}`).value = exp.estatus;
            worksheet.getCell(`G${row}`).value = exp.cedulaProfesional;
          });

          // 5) Finalmente: si había >1, combinamos la columna A de start a end
          // 5) Finalmente: si había >1, combinamos la columna A de start a end
          if (licCount > 1) {
            const mergeRange = `A${start}:A${end}`;

            try {
              // Intenta desfusionar (no lanza error si no existe)
              worksheet.unMergeCells(mergeRange);
            } catch (err) {
              // Ignora el error si el merge no existe
            }

            // Fusiona las celdas
            worksheet.mergeCells(mergeRange);
          }
        }

        ////////////////////////////////////////////////////////////////////////

        let MAE_START_ROW = 16 + acarreo; // ← ajusta esta fila al inicio de tu sección "Maestría"
        const maestrias = escolaridades.filter(
          (e) => e.nivelAcademico.toUpperCase() === "MAESTRIA",
        );
        let maeCount = maestrias.length;

        if (maeCount === 0) {
          console.log("MAE ", MAE_START_ROW);

          const mergeRange = `B${MAE_START_ROW}:G${MAE_START_ROW}`;

          try {
            worksheet.unMergeCells(mergeRange);
          } catch (err) {}
          worksheet.mergeCells(mergeRange);

          worksheet.getCell(`B${MAE_START_ROW}`).value = "Sin maestria";
          worksheet.getCell(`B${MAE_START_ROW}`).alignment = {
            vertical: "middle",
            horizontal: "center",
            wrapText: true,
          };
          worksheet.getRow(MAE_START_ROW).hidden = true;

          ///worksheet.getRow(MAE_START_ROW+1).hidden = true;
        } else {
          if (maeCount > 1) {
            acarreo = acarreo + maeCount - 1;
            worksheet.duplicateRow(MAE_START_ROW, maeCount - 1, true);
          }

          const start = MAE_START_ROW;

          console.log(start, " starto ", licDesplazar);
          const end = MAE_START_ROW + maeCount - 1;

          // Encabezado en A (si lo necesitas)
          //  worksheet.getCell(`A${start}`).value = "MAESTRÍA";
          //  worksheet.getCell(`A${start}`).alignment = { vertical:"middle", horizontal:"center", wrapText:true };

          maestrias.forEach((exp, i) => {
            const row = start + i;
            worksheet.getCell(`B${row}`).value = exp.nombreTitulo;
            worksheet.getCell(`C${row}`).value = exp.fechaObtencion;
            worksheet.getCell(`D${row}`).value = exp.institucion;
            worksheet.getCell(`E${row}`).value = exp.documentoAdquirido;
            worksheet.getCell(`F${row}`).value = exp.estatus;
            worksheet.getCell(`G${row}`).value = exp.cedulaProfesional;
          });

          if (maeCount > 1) {
            const mergeRange = `A${start}:A${end}`;
            try {
              worksheet.unMergeCells(mergeRange);
            } catch (err) {}
            try {
              worksheet.mergeCells(mergeRange);
            } catch (err) {
              console.warn(`No se pudo mergear ${mergeRange}:`, err);
            }
          }
        }

        ///////////////////////////////////////////////

        console.log("acarreo mae: ", acarreo);

        let DOC_START_ROW = 17 + acarreo; // ← ajusta esta fila al inicio de tu sección "Maestría"
        const doctorados = escolaridades.filter(
          (e) => e.nivelAcademico.toUpperCase() === "DOCTORADO",
        );
        let docCount = doctorados.length;

        if (docCount === 0) {
          console.log("DOC ", DOC_START_ROW);

          const mergeRange = `B${DOC_START_ROW}:G${DOC_START_ROW}`;
          console.log("mergeRange: ", mergeRange);

          try {
            worksheet.unMergeCells(mergeRange);
          } catch (err) {}
          worksheet.mergeCells(mergeRange);

          worksheet.getCell(`B${DOC_START_ROW}`).value = "Sin doctorado";
          worksheet.getCell(`B${DOC_START_ROW}`).alignment = {
            horizontal: "center",
          };

          worksheet.getRow(DOC_START_ROW).hidden = true;
          ///worksheet.getRow(MAE_START_ROW+1).hidden = true;
        } else {
          if (docCount > 1) {
            acarreo = acarreo + docCount - 1;
            worksheet.duplicateRow(DOC_START_ROW, docCount - 1, true);
          }

          const start = DOC_START_ROW;

          const end = DOC_START_ROW + docCount - 1;

          // Encabezado en A (si lo necesitas)
          //  worksheet.getCell(`A${start}`).value = "MAESTRÍA";
          //  worksheet.getCell(`A${start}`).alignment = { vertical:"middle", horizontal:"center", wrapText:true };

          doctorados.forEach((exp, i) => {
            const row = start + i;
            worksheet.getCell(`B${row}`).value = exp.nombreTitulo;
            worksheet.getCell(`C${row}`).value = exp.fechaObtencion;
            worksheet.getCell(`D${row}`).value = exp.institucion;
            worksheet.getCell(`E${row}`).value = exp.documentoAdquirido;
            worksheet.getCell(`F${row}`).value = exp.estatus;
            worksheet.getCell(`G${row}`).value = exp.cedulaProfesional;
          });

          if (docCount > 1) {
            const mergeRange = `A${start}:A${end}`;
            try {
              worksheet.unMergeCells(mergeRange);
            } catch (err) {}
            try {
              worksheet.mergeCells(mergeRange);
            } catch (err) {
              console.warn(`No se pudo mergear ${mergeRange}:`, err);
            }
          }
        }

        console.log("acarreo doc: ", acarreo);

        let ESP_START_ROW = 18 + acarreo; // ← ajusta esta fila al inicio de tu sección "Maestría"
        const especialidades = escolaridades.filter(
          (e) => e.nivelAcademico.toUpperCase() === "ESPECIALIDAD",
        );
        let espCount = especialidades.length;

        if (espCount === 0) {
          console.log("ESP ", ESP_START_ROW);

          const mergeRange = `B${ESP_START_ROW}:G${ESP_START_ROW}`;
          console.log("mergeRange: ", mergeRange);

          try {
            worksheet.unMergeCells(mergeRange);
          } catch (err) {}

          try {
            worksheet.mergeCells(mergeRange);
          } catch (err) {}

          worksheet.getCell(`B${ESP_START_ROW}`).value = "Sin especialidad";
          worksheet.getCell(`B${ESP_START_ROW}`).alignment = {
            horizontal: "center",
          };

          worksheet.getRow(ESP_START_ROW).hidden = true;

          ///worksheet.getRow(MAE_START_ROW+1).hidden = true;
        } else {
          if (espCount > 1) {
            acarreo = acarreo + espCount - 1;
            worksheet.duplicateRow(ESP_START_ROW, docCount - 1, true);
          }

          const start = ESP_START_ROW;

          const end = ESP_START_ROW + docCount - 1;

          // Encabezado en A (si lo necesitas)
          //  worksheet.getCell(`A${start}`).value = "MAESTRÍA";
          //  worksheet.getCell(`A${start}`).alignment = { vertical:"middle", horizontal:"center", wrapText:true };

          especialidades.forEach((exp, i) => {
            const row = start + i;
            worksheet.getCell(`B${row}`).value = exp.nombreTitulo;
            worksheet.getCell(`C${row}`).value = exp.fechaObtencion;
            worksheet.getCell(`D${row}`).value = exp.institucion;
            worksheet.getCell(`E${row}`).value = exp.documentoAdquirido;
            worksheet.getCell(`F${row}`).value = exp.estatus;
            worksheet.getCell(`G${row}`).value = exp.cedulaProfesional;
          });

          if (espCount > 1) {
            const mergeRange = `A${start}:A${end}`;
            try {
              worksheet.unMergeCells(mergeRange);
            } catch (err) {}
            try {
              worksheet.mergeCells(mergeRange);
            } catch (err) {
              console.warn(`No se pudo mergear ${mergeRange}:`, err);
            }
          }
        }

        const rutaSalida = "./reportes/";
        const nombreArchivo = `${trabajador.nombreTrabajador}_CURRICULUM-EDITABLE.xlsx`;
        const rutaCompleta = path.join(rutaSalida, nombreArchivo);

        await workbook.xlsx.writeFile(rutaCompleta);

        ////////////////////////////////////////////////////////////////////
        /////////////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////

        const rutaPlantilla2 = "./plantillas/curriculum2.xlsx";

        const workbook2 = new ExcelJS.Workbook();
        await workbook2.xlsx.readFile(rutaPlantilla2);
        const worksheet2 = workbook2.worksheets[0];

        worksheet2.getCell("B7").value = trabajador.nombreTrabajador;
        worksheet2.getCell("B11").value = trabajador.primaria;
        worksheet2.getCell("B12").value = trabajador.secundaria;
        worksheet2.getCell("B13").value = trabajador.carreraTecnicaComercial;

        bachillerato.forEach((exp, index) => {
          const fila = 14 + index; // Empieza en la fila 14, 15, 16…
          // Usando template literals con backticks:
          worksheet2.getCell(`B${fila}`).value = exp.nombreTitulo;
          worksheet2.getCell(`C${fila}`).value = exp.fechaObtencion;
          worksheet2.getCell(`D${fila}`).value = exp.institucion;
        });

        /*worksheet.duplicateRow(29,1, true);
                    worksheet.getCell("A30").value = "HOLAAA";
            */

        ///////////////////////////////////////////////////////////

        worksheet2.model.merges.forEach((range) => {
          const [start, end] = range.split(":");
          const startRow = +start.match(/\d+$/)[0];
          const endRow = +end.match(/\d+$/)[0];
          if (startRow === EXPERIENCE_ROW && endRow === EXPERIENCE_ROW) {
            expMerges.push(range);
          }
        });

        if (totalExp === 0) {
          const mergeRange = `A26:E26`;

          try {
            worksheet2.unMergeCells(mergeRange);
          } catch (err) {}
          worksheet2.mergeCells(mergeRange);
          worksheet2.getCell("A26").value = "Sin Experiencias";
          worksheet2.getCell("A26").alignment = {
            horizontal: "center",
          };
        }
        if (totalExp > 1) {
          worksheet2.duplicateRow(EXPERIENCE_ROW, totalExp - 1, true);
        }

        expMerges.forEach((range) => {
          const [start, end] = range.split(":");
          const colStart = start.match(/^[A-Z]+/)[0];
          const colEnd = end.match(/^[A-Z]+/)[0];
          for (let i = 1; i < totalExp; i++) {
            worksheet2.mergeCells(
              `${colStart}${EXPERIENCE_ROW + i}:${colEnd}${EXPERIENCE_ROW + i}`,
            );
          }
        });

        experienciasPJ.forEach((exp, idx) => {
          const r = EXPERIENCE_ROW + idx;
          worksheet2.getCell(`A${r}`).value = exp.periodoInicio;
          worksheet2.getCell(`B${r}`).value = exp.periodoFin;
          worksheet2.getCell(`C${r}`).value = exp.institucion;

          if (
            worksheet2.getCell(`C${r}`).value ==
            "PODER JUDICIAL DEL ESTADO DE HIDALGO"
          ) {
            worksheet2.getCell(`C${r}`).value =
              exp.institucion + "/" + exp.adscripcion;
          }
          worksheet2.getCell(`D${r}`).value = exp.cargo;
          worksheet2.getCell(`E${r}`).value = exp.campoExperiencia;
        });

        ///////////////////////////////////////////////////

        worksheet2.model.merges.forEach((range) => {
          const [start, end] = range.split(":");
          const startRow = parseInt(start.match(/\d+$/)[0], 10);
          const endRow = parseInt(end.match(/\d+$/)[0], 10);
          if (startRow === UPDATE_ROW && endRow === UPDATE_ROW) {
            updateMerges.push(range);
          }
        });

        if (totalActualizaciones > 1) {
          worksheet2.duplicateRow(UPDATE_ROW, totalActualizaciones - 1, true);
        }

        if (totalActualizaciones === 0) {
          const mergeRange = `A22:D22`;

          try {
            worksheet2.unMergeCells(mergeRange);
          } catch (err) {}
          worksheet2.mergeCells(mergeRange);
          worksheet2.getCell("B22").value = "Sin Actualizaciones";
          worksheet2.getCell("B22").alignment = {
            horizontal: "center",
          };
        }

        updateMerges.forEach((range) => {
          const [start, end] = range.split(":");
          const colStart = start.match(/^[A-Z]+/)[0];
          const colEnd = end.match(/^[A-Z]+/)[0];
          for (let i = 1; i < totalActualizaciones; i++) {
            console.log(
              "rango de merge1: ",
              `${colStart}${UPDATE_ROW + i}:${colEnd}${UPDATE_ROW + i}`,
            );
            worksheet2.mergeCells(
              `${colStart}${UPDATE_ROW + i}:${colEnd}${UPDATE_ROW + i}`,
            );
          }
        });

        actualizaciones.forEach((exp, idx) => {
          const r = UPDATE_ROW + idx;
          worksheet2.getCell(`A${r}`).value = exp.tema;
          worksheet2.getCell(`B${r}`).value = formatearFecha(exp.fecha);
          worksheet2.getCell(`C${r}`).value = exp.institucion;
          worksheet2.getCell(`D${r}`).value = exp.documento;
          /*const mergeRange = `C${r}:D${r}`;
                        console.log(mergeRange);
                        try{
                        worksheet.mergeCells(mergeRange);
                        }catch(err){
                
                        }*/
        });

        // 1) Si no hay ninguna, borramos la fila base
        if (prepaCount === 0) {
          const mergeRange = `B14:G14`;

          try {
            worksheet2.unMergeCells(mergeRange);
          } catch (err) {}
          worksheet2.mergeCells(mergeRange);
          worksheet2.getCell("B14").value = "No";
          worksheet2.getCell("B14").alignment = {
            horizontal: "center",
          };
          worksheet2.getRow(PREPA_START_ROW).hidden = true;
        } else {
          // 2) Si hay más de 1, duplicamos la fila base licCount-1 veces
          if (prepaCount > 1) {
            // (filaBase, númeroDeDuplicados, insert => true para desplazar filas hacia abajo)
            worksheet2.duplicateRow(LIC_START_ROW, licCount - 1, true);
          }

          const start = PREPA_START_ROW;
          const end = PREPA_START_ROW + prepaCount - 1;

          // 3) Asignamos el encabezado en la primera fila
          ///worksheet.getCell(`A${start}`).value = "hola";

          /*
                        worksheet.getCell(`A${start}`).alignment = {
                          vertical:   "middle",
                          horizontal: "center",
                          wrapText:   true
                        };*/

          // 4) Llenamos cada fila duplicada con los datos
          prepa.forEach((exp, i) => {
            const row = start + i;
            worksheet2.getCell(`B${row}`).value = exp.nombreTitulo;
            worksheet2.getCell(`C${row}`).value = exp.fechaObtencion;
            worksheet2.getCell(`D${row}`).value = exp.institucion;
          });

          // 5) Finalmente: si había >1, combinamos la columna A de start a end
          // 5) Finalmente: si había >1, combinamos la columna A de start a end
          if (prepaCount > 1) {
            const mergeRange = `A${start}:A${end}`;

            try {
              // Intenta desfusionar (no lanza error si no existe)
              worksheet2.unMergeCells(mergeRange);
            } catch (err) {
              // Ignora el error si el merge no existe
            }

            // Fusiona las celdas
            worksheet2.mergeCells(mergeRange);
          }
        }

        licCount = licenciaturas.length;
        console.log("lenght lic:", licCount);
        acarreo = licCount;

        console.log("acarreo lic: ", acarreo);
        licDesplazar = 15 + acarreo;

        // 1) Si no hay ninguna, borramos la fila base
        if (licCount === 0) {
          const mergeRange = `B${licDesplazar}:G${licDesplazar}`;

          try {
            worksheet2.unMergeCells(mergeRange);
          } catch (err) {}
          worksheet2.mergeCells(mergeRange);

          console.log("licdesplazar: ", licDesplazar);
          worksheet2.getCell(`B${licDesplazar}`).value = "Sin licenciatura";
          worksheet2.getCell(`B${licDesplazar}`).alignment = {
            horizontal: "center",
          };
          worksheet2.getRow(LIC_START_ROW).hidden = true;
        } else {
          // 2) Si hay más de 1, duplicamos la fila base licCount-1 veces

          if (licCount === 1) {
            acarreo = acarreo - 1;
          }
          if (licCount > 1) {
            acarreo = acarreo - 1;

            console.log("acarrero primero lic: ", acarreo);
            // (filaBase, númeroDeDuplicados, insert => true para desplazar filas hacia abajo)
            worksheet2.duplicateRow(LIC_START_ROW, licCount - 1, true);
          }

          const start = LIC_START_ROW;
          const end = LIC_START_ROW + licCount - 1;

          // 3) Asignamos el encabezado en la primera fila
          ///worksheet.getCell(`A${start}`).value = "hola";

          /*
                        worksheet.getCell(`A${start}`).alignment = {
                          vertical:   "middle",
                          horizontal: "center",
                          wrapText:   true
                        };*/

          // 4) Llenamos cada fila duplicada con los datos
          licenciaturas.forEach((exp, i) => {
            const row = start + i;
            worksheet2.getCell(`B${row}`).value = exp.nombreTitulo;
            worksheet2.getCell(`C${row}`).value = exp.fechaObtencion;
            worksheet2.getCell(`D${row}`).value = exp.institucion;
          });

          // 5) Finalmente: si había >1, combinamos la columna A de start a end
          // 5) Finalmente: si había >1, combinamos la columna A de start a end
          if (licCount > 1) {
            const mergeRange = `A${start}:A${end}`;

            try {
              // Intenta desfusionar (no lanza error si no existe)
              worksheet2.unMergeCells(mergeRange);
            } catch (err) {
              // Ignora el error si el merge no existe
            }

            // Fusiona las celdas
            worksheet2.mergeCells(mergeRange);
          }
        }

        ////////////////////////////////////////////////////////////////////////

        MAE_START_ROW = 16 + acarreo; // ← ajusta esta fila al inicio de tu sección "Maestría"

        if (maeCount === 0) {
          console.log("MAE ", MAE_START_ROW);

          const mergeRange = `B${MAE_START_ROW}:G${MAE_START_ROW}`;

          try {
            worksheet2.unMergeCells(mergeRange);
          } catch (err) {}
          worksheet2.mergeCells(mergeRange);

          worksheet2.getCell(`B${MAE_START_ROW}`).value = "Sin maestria";
          worksheet2.getCell(`B${MAE_START_ROW}`).alignment = {
            vertical: "middle",
            horizontal: "center",
            wrapText: true,
          };
          worksheet2.getRow(MAE_START_ROW).hidden = true;

          ///worksheet.getRow(MAE_START_ROW+1).hidden = true;
        } else {
          if (maeCount > 1) {
            acarreo = acarreo + maeCount - 1;
            worksheet2.duplicateRow(MAE_START_ROW, maeCount - 1, true);
          }

          const start = MAE_START_ROW;

          console.log(start, " starto ", licDesplazar);
          const end = MAE_START_ROW + maeCount - 1;

          // Encabezado en A (si lo necesitas)
          //  worksheet.getCell(`A${start}`).value = "MAESTRÍA";
          //  worksheet.getCell(`A${start}`).alignment = { vertical:"middle", horizontal:"center", wrapText:true };

          maestrias.forEach((exp, i) => {
            const row = start + i;
            worksheet2.getCell(`B${row}`).value = exp.nombreTitulo;
            worksheet2.getCell(`C${row}`).value = exp.fechaObtencion;
            worksheet2.getCell(`D${row}`).value = exp.institucion;
          });

          if (maeCount > 1) {
            const mergeRange = `A${start}:A${end}`;
            try {
              worksheet2.unMergeCells(mergeRange);
            } catch (err) {}
            try {
              worksheet2.mergeCells(mergeRange);
            } catch (err) {
              console.warn(`No se pudo mergear ${mergeRange}:`, err);
            }
          }
        }

        ///////////////////////////////////////////////

        console.log("acarreo mae: ", acarreo);

        DOC_START_ROW = 17 + acarreo; // ← ajusta esta fila al inicio de tu sección "Maestría"

        if (docCount === 0) {
          console.log("DOC ", DOC_START_ROW);

          const mergeRange = `B${DOC_START_ROW}:G${DOC_START_ROW}`;
          console.log("mergeRange: ", mergeRange);

          try {
            worksheet2.unMergeCells(mergeRange);
          } catch (err) {}
          worksheet2.mergeCells(mergeRange);

          worksheet2.getCell(`B${DOC_START_ROW}`).value = "Sin doctorado";
          worksheet2.getCell(`B${DOC_START_ROW}`).alignment = {
            horizontal: "center",
          };

          worksheet2.getRow(DOC_START_ROW).hidden = true;
          ///worksheet.getRow(MAE_START_ROW+1).hidden = true;
        } else {
          if (docCount > 1) {
            acarreo = acarreo + docCount - 1;
            worksheet2.duplicateRow(DOC_START_ROW, docCount - 1, true);
          }

          const start = DOC_START_ROW;

          const end = DOC_START_ROW + docCount - 1;

          // Encabezado en A (si lo necesitas)
          //  worksheet.getCell(`A${start}`).value = "MAESTRÍA";
          //  worksheet.getCell(`A${start}`).alignment = { vertical:"middle", horizontal:"center", wrapText:true };

          doctorados.forEach((exp, i) => {
            const row = start + i;
            worksheet2.getCell(`B${row}`).value = exp.nombreTitulo;
            worksheet2.getCell(`C${row}`).value = exp.fechaObtencion;
            worksheet2.getCell(`D${row}`).value = exp.institucion;
          });

          if (docCount > 1) {
            const mergeRange = `A${start}:A${end}`;
            try {
              worksheet.unMergeCells(mergeRange);
            } catch (err) {}
            try {
              worksheet.mergeCells(mergeRange);
            } catch (err) {
              console.warn(`No se pudo mergear ${mergeRange}:`, err);
            }
          }
        }

        console.log("acarreo doc: ", acarreo);

        ESP_START_ROW = 18 + acarreo; // ← ajusta esta fila al inicio de tu sección "Maestría"

        if (espCount === 0) {
          console.log("ESP ", ESP_START_ROW);

          const mergeRange = `B${ESP_START_ROW}:G${ESP_START_ROW}`;
          console.log("mergeRange: ", mergeRange);

          try {
            worksheet2.unMergeCells(mergeRange);
          } catch (err) {}

          try {
            worksheet2.mergeCells(mergeRange);
          } catch (err) {}

          worksheet2.getCell(`B${ESP_START_ROW}`).value = "Sin especialidad";
          worksheet2.getCell(`B${ESP_START_ROW}`).alignment = {
            horizontal: "center",
          };

          worksheet2.getRow(ESP_START_ROW).hidden = true;

          ///worksheet.getRow(MAE_START_ROW+1).hidden = true;
        } else {
          if (espCount > 1) {
            acarreo = acarreo + espCount - 1;
            worksheet2.duplicateRow(ESP_START_ROW, docCount - 1, true);
          }

          const start = ESP_START_ROW;

          const end = ESP_START_ROW + docCount - 1;

          // Encabezado en A (si lo necesitas)
          //  worksheet.getCell(`A${start}`).value = "MAESTRÍA";
          //  worksheet.getCell(`A${start}`).alignment = { vertical:"middle", horizontal:"center", wrapText:true };

          especialidades.forEach((exp, i) => {
            const row = start + i;
            worksheet2.getCell(`B${row}`).value = exp.nombreTitulo;
            worksheet2.getCell(`C${row}`).value = exp.fechaObtencion;
            worksheet2.getCell(`D${row}`).value = exp.institucion;
          });

          if (espCount > 1) {
            const mergeRange = `A${start}:A${end}`;
            try {
              worksheet2.unMergeCells(mergeRange);
            } catch (err) {}
            try {
              worksheet2.mergeCells(mergeRange);
            } catch (err) {
              console.warn(`No se pudo mergear ${mergeRange}:`, err);
            }
          }
        }

        if (licCount >= 1 || maeCount >= 1 || docCount >= 1 || espCount >= 1) {
          console.log("Si hay más de una");
          worksheet2.getRow(11).hidden = true;
          worksheet2.getRow(12).hidden = true;
          worksheet2.getRow(13).hidden = true;
          worksheet2.getRow(14).hidden = true;
        }

        const { exec } = require("child_process");

        const _uid = `${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
        const rutaExcel = `./reportes/${trabajador.nombreTrabajador}_${_uid}_temp.xlsx`;
        const rutaPDF = `./reportes/${trabajador.nombreTrabajador}_${_uid}_CURRICULUM_TRANSPARENCIA.pdf`;
        const rutaHTML = `./reportes/${trabajador.nombreTrabajador}_${_uid}_temp.html`;
        // 1. Guardar Excel temporal
        await workbook2.xlsx.writeFile(rutaExcel);

        await excelToHtmlWithStyles(rutaExcel, rutaHTML);

        const browser = await puppeteer.launch({
          headless: "new",
          args: ["--no-sandbox", "--disable-setuid-sandbox"],
        });
        const page = await browser.newPage();
        // Carga el HTML recién creado — convertir ruta Windows a URL file:/// válida
        const fileUrlHTML = `file:///${path.resolve(rutaHTML).replace(/\\/g, "/")}`;
        await page.goto(fileUrlHTML, {
          waitUntil: "networkidle0",
        });
        await page.pdf({
          path: rutaPDF,
          format: "letter",
          printBackground: true,
          margin: { top: "1cm", right: "1cm", bottom: "1cm", left: "1cm" },
        });
        await browser.close();

        fs.unlinkSync(rutaExcel);
        fs.unlinkSync(rutaHTML);

        // ... TODO TU CÓDIGO ACTUAL PARA GENERAR EXCEL Y PDF ...
        // (mantén intacta toda la lógica de generación de documentos)

        // Al final de tu código actual, tendrás estas rutas:
        let rutaExcel2 = `./reportes/${nombreTrabajador}_CURRICULUM-EDITABLE.xlsx`;
        let rutaPDF2 = `./reportes/${nombreTrabajador}_CURRICULUM_TRANSPARENCIA.pdf`;
        // ------ FIN DE TU CÓDIGO ACTUAL ------

        // Crear carpeta en el ZIP para este trabajador
        const carpetaTrabajador = nombreTrabajador.replace(/[^\w\s]/gi, "");

        // Agregar archivos generados al ZIP
        archive.file(rutaExcel2, {
          name: `${carpetaTrabajador}/${nombreTrabajador}_CURRICULUM-EDITABLE.xlsx`,
        });

        archive.file(rutaPDF2, {
          name: `${carpetaTrabajador}/${nombreTrabajador}_CURRICULUM_TRANSPARENCIA.pdf`,
        });

        // Agregar documentos adicionales con nombre del trabajador
        const documentos = [
          "INE.pdf",
          "ACTA-NACIMIENTO.pdf",
          "COMPROBANTE-DOMICILIO.pdf",
          "RFC.pdf",
          "GRADO-ESTUDIOS.pdf",
        ];

        documentos.forEach((doc) => {
          const nombreArchivo = `${nombreTrabajador}_${doc}`;
          const rutaDoc = path.join("./uploads", nombreArchivo);

          if (fs.existsSync(rutaDoc)) {
            archive.file(rutaDoc, {
              name: `${carpetaTrabajador}/${nombreArchivo}`,
            });
          }
        });
      } catch (error) {
        console.error(`Error procesando ${nombreTrabajador}:`, error);
        // Continuar con el siguiente trabajador aunque falle uno
      }
    }
    // ===================================================================
    // FIN DEL CICLO PARA CADA TRABAJADOR
    // ===================================================================

    // Finalizar el ZIP después de procesar todos los trabajadores
    await archive.finalize();
  } catch (error) {
    console.error("Error general en el endpoint:", error);
    res.status(500).send("Error interno del servidor");
  }
});

/**
 * ================================
 * RUTAS DE ADMINISTRADOR
 * ================================
 */

// GET - Mostrar login de administrador
app.get("/admin/login", (req, res) => {
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
    return res.status(400).send("Debes enviar fechaInicio y fechaFin");
  }

  const inicio = dayjs(fechaInicio);
  const fin = dayjs(fechaFin);

  if (!inicio.isValid() || !fin.isValid()) {
    return res.status(400).send("Las fechas enviadas no son validas");
  }

  if (inicio.isAfter(fin, "day")) {
    return res
      .status(400)
      .send("La fecha de inicio no puede ser mayor a la fecha de termino");
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
    return res.status(500).send("Error al generar el reporte de transparencia");
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
