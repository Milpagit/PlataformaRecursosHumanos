const express = require("express");
const path = require("path");
const db = require("./config/db"); // Importamos tu conexión unificada
const app = express();

// Configuración de Middlewares
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "public")));

// Configuración del motor de plantillas
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

// Funciones globales disponibles en todas las vistas EJS
app.locals.formatearFecha = function (fecha) {
  if (!fecha) return "N/A";

  // Si es un string tipo "2024-10-17", formatearlo
  if (typeof fecha === "string" && fecha.includes("-")) {
    const partes = fecha.split("-");
    if (partes[0].length === 4) {
      // Formato YYYY-MM-DD
      const [anio, mes, dia] = partes;
      const meses = [
        "",
        "Enero",
        "Febrero",
        "Marzo",
        "Abril",
        "Mayo",
        "Junio",
        "Julio",
        "Agosto",
        "Septiembre",
        "Octubre",
        "Noviembre",
        "Diciembre",
      ];
      return `${dia} de ${meses[parseInt(mes)]} de ${anio}`;
    }
  }

  // Si es un objeto Date
  if (fecha instanceof Date) {
    const opciones = { day: "2-digit", month: "long", year: "numeric" };
    return fecha.toLocaleDateString("es-MX", opciones);
  }

  // Intentar crear un Date si es string de otro formato
  try {
    const d = new Date(fecha);
    if (!isNaN(d.getTime())) {
      const opciones = { day: "2-digit", month: "long", year: "numeric" };
      return d.toLocaleDateString("es-MX", opciones);
    }
  } catch (e) {
    // Ignorar error
  }

  return String(fecha); // Devolver como string al final
};

/**
 * RUTAS DE ACCESO (Login y Registro)
 */

// Pantalla principal (Login)
app.get("/", (req, res) => {
  res.render("login");
});

app.get("/login", (req, res) => {
  res.render("login");
});

// Procesar Login y Cargar datos para Autocompletado
app.post("/login", async (req, res) => {
  const { curp, rfc } = req.body;
  try {
    // Buscamos si el trabajador ya existe en la base de datos
    const [rows] = await db.query(
      "SELECT * FROM trabajador WHERE curp = ? AND rfc = ?",
      [curp, rfc],
    );

    if (rows && rows.length > 0) {
      // Si existe, mandamos sus datos al formulario (index.ejs) para autocompletar
      // Nota: index.ejs usa la variable "user" [cite: 10, 51]
      res.render("index", { user: rows[0] });
    } else {
      res.render("login", {
        error: "No se encontró el registro o los datos son incorrectos.",
      });
    }
  } catch (error) {
    console.error("Error en login:", error);
    res.status(500).send("Error interno del servidor");
  }
});

/**
 * RUTAS DEL FORMULARIO Y REGISTRO
 */

// Guardar o Actualizar Datos
app.post("/guardar", async (req, res) => {
  try {
    const datos = req.body;

    // CORRECCIÓN DE FECHAS: Convertir DD-MM-YYYY a YYYY-MM-DD para SQL Server
    const formatearFecha = (f) => (f ? f.split("-").reverse().join("-") : null);

    const fechaIngresoSQL = formatearFecha(datos.fechaIngreso);
    const fechaNacimientoSQL = formatearFecha(datos.fechaNacimiento);

    // Aquí iría tu lógica de UPDATE o INSERT en la base de datos
    // Ejemplo simplificado:
    /*
        await db.query('UPDATE trabajador SET fechaIngreso = ?, adscripcionActual = ? WHERE noTrabajador = ?', 
        [fechaIngresoSQL, datos.adscripcionActual, datos.noTrabajador]);
        */

    res.render("success");
  } catch (error) {
    console.error("Error al guardar:", error);
    res.status(500).send("Error al procesar el formulario");
  }
});

/**
 * RUTAS DE REPORTES Y ACUSES
 */

// Generar Acuse (Ruta que te daba error 500)
app.get("/acuse/:noTrabajador/:tipoTrabajador", async (req, res) => {
  const { noTrabajador, tipoTrabajador } = req.params;
  try {
    const [rows] = await db.query(
      "SELECT * FROM trabajador WHERE noTrabajador = ? AND tipoTrabajador = ?",
      [noTrabajador, tipoTrabajador],
    );

    if (rows && rows.length > 0) {
      const trabajador = rows[0];
      const idReal = trabajador.idTrabajador;

      // Helper para formatear fechas en la vista del acuse
      const formatearFecha = (f) =>
        f ? new Date(f).toLocaleDateString("es-MX") : "N/A";

      // Consultas adicionales usando destructuring [ ] para obtener el recordset correctamente
      const [hijos] = await db.query(
        "SELECT inicialesHijo, sexoHijo, fechaNacimientoHijo, (CONVERT(int,CONVERT(char(8),GETDATE(),112))-CONVERT(int,CONVERT(char(8),fechaNacimientoHijo,112)))/10000 AS edadHijo FROM hijo WHERE idTrabajador = ?",
        [idReal],
      );

      const [escolaridad] = await db.query(
        "SELECT * FROM escolaridad WHERE idTrabajador = ?",
        [idReal],
      );

      const [experienciaPJ] = await db.query(
        "SELECT institucion, adscripcion, cargo, campoExperiencia, periodoInicio, periodoFin FROM experienciapj WHERE idTrabajador = ?",
        [idReal],
      );

      const [actualizacion] = await db.query(
        "SELECT * FROM actualizacionprofesional WHERE idTrabajador = ?",
        [idReal],
      );

      // acuse.ejs espera la variable "trabajador" [cite: 369, 370]
      res.render("acuse", {
        trabajador,
        formatearFecha,
        escolaridad,
        actualizacion,
        experienciaPJ,
        hijos,
      });
    } else {
      res.status(404).send("Trabajador no encontrado");
    }
  } catch (error) {
    console.error("Error al generar acuse:", error);
    res.status(500).send("Error interno al generar el acuse");
  }
});

// Buscador de Reportes
app.get("/buscador", (req, res) => {
  res.render("buscador");
});

// Verificar existencia para el buscador
app.get("/existe/:noTrabajador/:tipoTrabajador", async (req, res) => {
  try {
    const [rows] = await db.query(
      "SELECT 1 FROM trabajador WHERE noTrabajador = ? AND tipoTrabajador = ?",
      [req.params.noTrabajador, req.params.tipoTrabajador],
    );
    res.json({ existe: rows.length > 0 });
  } catch (error) {
    res.status(500).json({ error: true });
  }
});

/**
 * OTRAS VISTAS
 */
app.get("/registrar-password", (req, res) =>
  res.render("registrarPassword", { error: null }),
);
app.get("/recuperar-password", (req, res) =>
  res.render("recuperar-password", { error: null, mensaje: null }),
);

// Iniciar servidor (solo para entorno de pruebas; en producción se usa server.js)
const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Servidor corriendo en http://localhost:${PORT}`);
});
