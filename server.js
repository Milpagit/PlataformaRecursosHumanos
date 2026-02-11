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

app.locals.formatearFecha = function (fecha) {
  if (!fecha) return "N/A";

  // Convierte la fecha ISO a hora local sin perder el día
  const f = dayjs.utc(fecha).tz("America/Mexico_City", true);

  if (!f.isValid()) return "N/A";

  return f.format("D [de] MMMM [de] YYYY"); // Ejemplo: 29 de marzo de 1984
};

////////////////////////////////////////////////////////////////////////////////

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
    secret: "Vr#Rno@jyc%j5Kckh978SM", // Llave unica de la sesión
    resave: false,
    saveUninitialized: false,
  }),
);

app.get("/acuse/:noTrabajador", async (req, res) => {
  const { noTrabajador } = req.params;

  try {
    // ----------------------------------------------------
    // 1. SEGURIDAD (De tu código original)
    // ----------------------------------------------------
    // Verifica si existe la sesión (si usas express-session)
    if (req.session && req.session.idTrabajador) {
      // Validación extra: Que el usuario logueado sea el dueño del acuse
      // Si quieres permitir que RRHH vea cualquier acuse, comenta este if
      /*
        if (String(req.session.noTrabajador) !== String(noTrabajador)) {
           return res.redirect("/login");
        }
        */
    }
    // Nota: Si no tienes configurado express-session en app.js,
    // borra el bloque if de arriba para que no te de error.

    // ----------------------------------------------------
    // 2. CONSULTAS SQL CORREGIDAS (Sin Vistas ni Funciones raras)
    // ----------------------------------------------------

    // A) Datos del Trabajador
    const [rows] = await db.query(
      "SELECT * FROM trabajador WHERE noTrabajador = ?",
      [noTrabajador],
    );

    if (!rows || rows.length === 0) {
      return res.status(404).send("Trabajador no encontrado");
    }
    const trabajador = rows[0];
    // Usamos el ID real para las siguientes búsquedas
    const idReal = trabajador.idTrabajador;

    // B) Hijos (Cálculo de edad manual, sin dbo.CalcularEdad)
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

    // ----------------------------------------------------
    // 3. HELPER Y RENDERIZADO
    // ----------------------------------------------------

    // Función necesaria para que acuse.ejs formatee las fechas
    const formatearFecha = (f) => {
      if (!f) return "N/A";
      const fecha = new Date(f);
      return fecha.toISOString().split("T")[0];
    };

    res.render("acuse", {
      trabajador,
      hijos,
      escolaridad,
      experienciaPJ, // Ahora lleva los datos correctos de la tabla experienciapj
      actualizacion,
      noTrabajador,
      formatearFecha, // ¡Importante pasar esto!
    });
  } catch (error) {
    console.error("Error en /acuse:", error);
    res.status(500).send("Error interno al generar el acuse: " + error.message);
  }
});
app.get("/login", (req, res) => {
  res.render("login", { error: null });
});

app.get("/index", async (req, res) => {
  if (!req.session.user) {
    return res.redirect("/login");
  }

  const nombreTrabajador = req.session.user.nombreTrabajador;
  const uploadDir = path.join(__dirname, "uploads");

  // Función helper para verificar si existe el PDF
  const checkFile = (suffix) =>
    fs.existsSync(path.join(uploadDir, `${nombreTrabajador}_${suffix}.pdf`));

  // Consultar experiencia laboral existente
  const [experienciaPJ] = await db.query(
    "SELECT * FROM experienciapj WHERE idTrabajador = ?",
    [req.session.user.idTrabajador],
  );

  // Consultar hijos existentes
  const [hijos] = await db.query("SELECT * FROM hijo WHERE idTrabajador = ?", [
    req.session.user.idTrabajador,
  ]);

  res.render("index", {
    user: req.session.user,
    experienciaPJ,
    hijos,
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

app.get("/existe/:noTrabajador", async (req, res) => {
  const noTrabajador = req.params.noTrabajador;

  try {
    const [result] = await db.query(
      "SELECT COUNT(*) AS total FROM trabajador WHERE noTrabajador = ?",
      [noTrabajador],
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

//function formatearFecha(fecha) {
//if (!fecha || fecha === '0000-00-00') {
//   return 'N/A';  // O puedes poner simplemente ''
// }

// // Si es un objeto Date, lo formateamos
// const opciones = { day: '2-digit', month: 'long', year: 'numeric' };
// const fechaFormateada = new Date(fecha).toLocaleDateString('es-MX', opciones);

// // Pone la primera letra de mes en mayúscula (por si acaso)
//  return fechaFormateada.charAt(0).toUpperCase() + fechaFormateada.slice(1);
//}

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

    console.log("Papa: " + papa + " y " + sexo);
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
      comunidadIndigena =
        req.body.otraComunidad?.trim().toUpperCase() || "OTRO";
    }

    let familiaLinguistica =
      req.body.hablaLenguaIndigena === "1"
        ? req.body.familiaLinguistica?.trim().toUpperCase() || "N/A"
        : "N/A";

    if (familiaLinguistica === "OTRA") {
      familiaLinguistica = req.body.otraLengua?.trim().toUpperCase() || "OTRO";
    }

    let tipoDiscapacidad =
      req.body.discapacidad === "1"
        ? req.body.tipoDiscapacidad?.trim().toUpperCase() || "NO"
        : "No";

    let discapacidadEspecificada = req.body.discapacidadTexto;

    if (req.body.discapacidad === "0") {
      discapacidadEspecificada = "N/A";
    }
    ///let carreraTecnica = req.body.tecnica_institucion;
    const vigenciaINE = req.body.vigenciaINE;
    const tieneTecnica = req.body.tieneTecnica === "1";
    const carreraTecnicaComercial = tieneTecnica
      ? req.body.tecnica_institucion.trim().toUpperCase()
      : "No";

    const tieneBachillerato = req.body.tieneBachillerato === "1";
    const prepaTitulo = tieneBachillerato
      ? req.body.prepa_titulo.trim().toUpperCase()
      : null;
    const prepaFecha = tieneBachillerato
      ? convertirFechaTexto(req.body.prepa_fecha)
      : null;
    const prepaInst = tieneBachillerato
      ? req.body.prepa_institucion.trim().toUpperCase()
      : null;
    const prepaDocumento = tieneBachillerato ? req.body.prepa_documento : null;
    const prepaEstatus = tieneBachillerato ? req.body.prepa_estatus : null;

    let dominioLSM = req.body.nivelLSM;

    if (lenguaSenias === "No") {
      dominioLSM = "N/A";
    }

    const query =
      "UPDATE trabajador SET  fechaIngreso=?, adscripcionActual=?,cargoActual=?,sexoTrabajador = ?,fechaNacimiento=?,lugarNacimiento=?,estadoCivil=?,nombreConyuge=?,fechaNacimientoConyuge=?,sexoConyuge=?,tipoSangre=?,vigenciaINE=?,calleNumero=?,colonia=?,municipio=?,estado=?,codigoPostal=?,actualCalleNumero=?,actualColonia=?,actualMunicipio=?,actualEstado=?, actualCP=?,  noTelefono = ?,telefonoCasa=?,telefonoFamiliar=?,parentescoFamiliar=?, correoElectronico=?,comunidadIndigena=?,familiaLinguistica=?,tipoDiscapacidad=?,discapacidadEspecificada=?,lenguajeSenias=?,dominioLSM=?,tieneEnfermedadCronica=?,tipoCronica=?, es_padre_madre=? , embarazada = ?, cantidadHijos = ?,primaria=?,secundaria=?, carreraTecnicaComercial=? WHERE idTrabajador = ?";

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
      console.log("nacimiento Conyuge" + fechaFormateada);
      console.log("nacimiento Conyuge" + fechaFormateada);
      console.log("Fecha de ingreso" + fechaIngresoFormateada);
      console.log("Fecha de nacimiento" + fechaNacimientoFormateada);

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
        telefonocasa,
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
        cantidadHijos,
        primaria,
        secundaria_institucion,
        carreraTecnicaComercial,
        idTrabajador,
      ]);
      console.log("✅ Actualización exitosa:", result);

      const deleteQuery = "DELETE FROM hijo WHERE idTrabajador = ?";
      await db.query(deleteQuery, [idTrabajador]);
      console.log(
        `🗑️ Hijos anteriores eliminados para trabajador ${idTrabajador}`,
      );

      // ... después de borrar los hijos anteriores ...

      if (cantidadHijos && cantidadHijos > 0) {
        for (let i = 0; i < cantidadHijos; i++) {
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
            console.log(`✅ Hijo ${i + 1} guardado con edad: ${edad}`);
          }
        }
      }
      const deleteQuery2 = "DELETE FROM escolaridad WHERE idTrabajador = ?";
      await db.query(deleteQuery2, [idTrabajador]);

      await db.query("DELETE FROM escolaridad WHERE idTrabajador = ?", [
        idTrabajador,
      ]);

      const niveles = [
        "licenciatura",
        "maestria",
        "doctorado",
        "posdoctorado",
        "especialidad",
      ];
      const maximos = {
        licenciatura: 3,
        maestria: 3,
        doctorado: 2,
        posdoctorado: 2,
        especialidad: 3,
      };

      const insertQuery = `
        INSERT INTO escolaridad
            (nivelAcademico, nombreTitulo, fechaObtencion, institucion,
            documentoAdquirido, estatus, cedulaProfesional, idTrabajador)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        `;

      if (req.body.tieneBachillerato === "1") {
        await db.query(insertQuery, [
          "BACHILLERATO",
          req.body.prepa_titulo.trim().toUpperCase(),
          req.body.prepa_fecha,
          req.body.prepa_institucion.trim().toUpperCase(),
          req.body.prepa_documento,
          req.body.prepa_estatus,
          null, // cedula
          idTrabajador,
        ]);
      }

      for (let nivel of niveles) {
        const cantidadRaw = req.body[`${nivel}Cantidad`];
        const cantidad = Math.min(
          maximos[nivel],
          Math.max(0, parseInt(cantidadRaw, 10) || 0),
        );

        for (let i = 1; i <= cantidad; i++) {
          const tituloField = `${nivel}_titulo${i}`;
          const fechaField = `${nivel}_fecha${i}`;
          const institucionField = `${nivel}_institucion${i}`;
          const documentoField = `${nivel}_documento${i}`;
          const estatusField = `${nivel}_estatus${i}`;
          const cedulaField = `${nivel}_cedula${i}`;

          const titulo = req.body[tituloField]?.trim() || null;
          const fechaRaw = req.body[fechaField] || null; // yyyy-mm-dd
          const institucion = req.body[institucionField]?.trim() || null;
          const documento = req.body[documentoField] || null;
          const estatus = req.body[estatusField] || null;
          const cedula = req.body[cedulaField]?.trim() || null;

          if (titulo || fechaRaw || institucion) {
            const fecha = fechaRaw;
            await db.query(insertQuery, [
              nivel.toUpperCase(), // nivelAcademico
              titulo,
              fecha,
              institucion,
              documento,
              estatus,
              cedula,
              idTrabajador,
            ]);
          }
        }
      }

      console.log(`Escolaridad eliminada para trabajador ${idTrabajador}`);

      const escolaridades = [
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
      ];

      // Recorremos cada escolaridad y guardamos solo si tienen algún dato real
      for (let escolaridad of escolaridades) {
        console.log(escolaridad.fecha);
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

          console.log(`✅ Escolaridad ${escolaridad.nivel} guardada`);
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

        const hayDatos = tema || fecha || institucion || documento;

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
          console.log(`✅ Experiencia insertada: ${exp.puesto}`);
        }
      }

      const idTrab = idTrabajador;
      req.session.idTrabajador = idTrab;
      res.redirect(`/acuse/${noTrabajador}`);
    } catch (err) {
      console.error("❌ Error al guardar:", err);
      res.send("Error al guardar");
    }
  },
);

app.post("/login", async (req, res) => {
  console.log("🚀 Entró al POST /login");

  const curp = req.body.curp?.toUpperCase();
  const rfc = req.body.rfc?.toUpperCase();
  const password = req.body.password;

  try {
    console.log("🔍 Antes del query");

    const [results] = await db.query(
      "SELECT * FROM trabajador WHERE curp = ? AND rfc = ?",
      [curp, rfc],
    );

    console.log("✅ Después del query");
    console.log("Resultado de la consulta:", results);

    if (results.length === 0) {
      return res.send(`
              <script>
                alert('CURP o RFC no encontrado. Verifica los datos.');
                window.location.href = '/login';
              </script>
            `);
    }

    const usuario = results[0];

    if (usuario.password !== password) {
      return res.send(`
                <script>
                    alert('Contraseña incorrecta. Intenta nuevamente.');
                    window.location.href = '/login';
                </script>
            `);
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
    return res.send("Error consultando la base");
  }
});

app.get("/reporte/:noTrabajador", async (req, res) => {
  const noTrabajador = req.params.noTrabajador;

  const rutaPlantilla = "./plantillas/curriculum.xlsx";

  try {
    const [trabajadorRows] = await db.query(
      "SELECT * FROM trabajador WHERE noTrabajador = ?",
      [noTrabajador],
    );
    const trabajador = trabajadorRows[0];
    const idTrabajador = trabajador.idTrabajador;

    if (!trabajador) {
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

    console.log(experiencias);

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
      (e) => e.nivelAcademico.toUpperCase() === "BACHILLERATO",
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
        worksheet.duplicateRow(LIC_START_ROW, licCount - 1, true);
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

    const rutaExcel = `./reportes/${trabajador.nombreTrabajador}_temp.xlsx`;
    const rutaPDF = `./reportes/${trabajador.nombreTrabajador}_CURRICULUM_TRANSPARENCIA.pdf`;
    const rutaHTML = `./reportes/${trabajador.nombreTrabajador}_temp.html`;
    // 1. Guardar Excel temporal
    await workbook2.xlsx.writeFile(rutaExcel);

    excelToHtmlWithStyles(rutaExcel, rutaHTML);

    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    // Carga el HTML recién creado (asegúrate de resolver la ruta absoluta)
    await page.goto(`file://${path.resolve(rutaHTML)}`, {
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
          console.error("Error al enviar ZIP:", err);
        } else {
          fs.unlinkSync(zipRuta);
          fs.unlinkSync(rutaCompleta);
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

    // Archivos PDF
    const uploadsPath = path.join(__dirname, "uploads");
    const pdfs = [
      `${trabajador.nombreTrabajador}_INE.pdf`,
      `${trabajador.nombreTrabajador}_ACTA-NACIMIENTO.pdf`,
      `${trabajador.nombreTrabajador}_COMPROBANTE-DOMICILIO.pdf`,
      `${trabajador.nombreTrabajador}_RFC.pdf`,
      `${trabajador.nombreTrabajador}_GRADO-ESTUDIOS.pdf`,
    ];

    pdfs.forEach((nombre) => {
      const ruta = path.join(uploadsPath, nombre);
      if (fs.existsSync(ruta)) {
        archive.file(ruta, { name: `${carpetaTrabajador}/${nombre}` });
      }
    });

    await archive.finalize();

    fs.unlinkSync(rutaPDF);
  } catch (error) {
    console.error("Error generando el reporte:", error);
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
          console.error("Error al enviar ZIP:", err);
        } else {
          // Limpiar archivos temporales después de enviar
          fs.unlinkSync(zipRuta);

          // Eliminar archivos generados para cada trabajador
          todosTrabajadores.forEach((t) => {
            const base = `./reportes/${t.nombreTrabajador}_`;
            try {
              fs.unlinkSync(`${base}CURRICULUM-EDITABLE.xlsx`);
              fs.unlinkSync(`${base}CURRICULUM_TRANSPARENCIA.pdf`);
            } catch (e) {
              console.warn(
                `No se pudo eliminar archivo de ${t.nombreTrabajador}:`,
                e.message,
              );
            }
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
          "SELECT * FROM trabajador WHERE noTrabajador = ?",
          [noTrabajador],
        );
        const trabajador = trabajadorRows[0];
        const idTrabajador = trabajador.idTrabajador;

        if (!trabajador) {
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

        const rutaExcel = `./reportes/${trabajador.nombreTrabajador}_temp.xlsx`;
        const rutaPDF = `./reportes/${trabajador.nombreTrabajador}_CURRICULUM_TRANSPARENCIA.pdf`;
        const rutaHTML = `./reportes/${trabajador.nombreTrabajador}_temp.html`;
        // 1. Guardar Excel temporal
        await workbook2.xlsx.writeFile(rutaExcel);

        excelToHtmlWithStyles(rutaExcel, rutaHTML);

        const browser = await puppeteer.launch();
        const page = await browser.newPage();
        // Carga el HTML recién creado (asegúrate de resolver la ruta absoluta)
        await page.goto(`file://${path.resolve(rutaHTML)}`, {
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

  // Usar file:// para cargar el archivo local
  await page.goto(`file://${path.resolve(htmlPath)}`, {
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
  host: "smtp.office365.com",
  port: 587,
  secure: false,
  auth: {
    user: "desarrollo.sistemas@pjhidalgo.gob.mx",
    pass: "Sub20Pach",
  },
});

// GET mostrar formulario
app.get("/recuperar-password", (req, res) => {
  res.render("recuperar-password", { error: null, mensaje: null });
});

app.post("/recuperar-password", async (req, res) => {
  const { correoe, curp, action } = req.body;

  try {
    // Recuperar por correo
    if (action === "recuperar-clave") {
      if (!correoe) {
        return res.render("recuperar-password", {
          error: "Debe ingresar un correo electrónico",
          mensaje: null,
        });
      }

      const result = await textdb.query(
        "EXEC sp_RecuperarClavePorCorreo @Correo",
        {
          Correo: { type: sql.VarChar(200), value: correoe.trim() },
        },
      );

      const row = result[0]; // db.query devuelve recordset directamente

      if (!row || row.Estado === "NO_ENCONTRADO") {
        return res.render("recuperar-password", {
          error: "No se encontró trabajador con ese correo",
          mensaje: null,
        });
      }
      if (row.Estado === "SIN_CORREO") {
        return res.render("recuperar-password", {
          error: "El trabajador existe pero no tiene correo registrado",
          mensaje: null,
        });
      }

      if (row.Estado === "OK") {
        await transporter.sendMail({
          from: '"Soporte Recursos Humanos" <desarrollo.sistemas@pjhidalgo.gob.mx>',
          to: correoe.trim(),
          subject: "Recuperación de clave",
          html: `<p>Su clave de acceso al módulo de actualización de datos personales de los Trabajadores del Poder Judicial de Hidalgo es: <b>${row.Clave}</b></p>`,
        });

        return res.render("recuperar-password", {
          error: null,
          mensaje: `La clave fue enviada al correo ${correoe.trim()}`,
        });
        //return res.send(`<script>alert('La clave fue enviada al correo: ${correoe.trim()}.');</script>`);
      }
    }

    // Recuperar por CURP
    else if (action === "recuperar-correo") {
      if (!curp) {
        return res.render("recuperar-password", {
          error: "Debe ingresar la CURP",
          mensaje: null,
        });
      }

      const result = await textdb.query(
        "EXEC sp_RecuperarCorreoPorCurp @Curp",
        {
          Curp: { type: sql.VarChar(18), value: curp.trim() },
        },
      );

      const row = result[0];

      if (!row || row.Estado === "NO_ENCONTRADO") {
        return res.render("recuperar-password", {
          error: "No se encontró usuario con esa CURP",
          mensaje: null,
        });
      }
      if (row.Estado === "SIN_CORREO") {
        return res.render("recuperar-password", {
          error: "El trabajador existe pero no tiene correo registrado",
          mensaje: null,
        });
      }

      if (row.Estado === "OK") {
        return res.render("recuperar-password", {
          error: null,
          mensaje: `El correo registrado para la CURP ${curp.trim()} es ${row.Correo}`,
        });
      }
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
