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
app.get("/acuse/:noTrabajador", async (req, res) => {
  const { noTrabajador } = req.params;
  try {
    const [rows] = await db.query(
      "SELECT * FROM trabajador WHERE noTrabajador = ?",
      [noTrabajador],
    );

    if (rows && rows.length > 0) {
      // Helper para formatear fechas en la vista del acuse
      const formatearFecha = (f) =>
        f ? new Date(f).toLocaleDateString("es-MX") : "N/A";

      // acuse.ejs espera la variable "trabajador" [cite: 369, 370]
      res.render("acuse", {
        trabajador: rows[0],
        formatearFecha,
        escolaridad: [], // Debes consultar estas tablas también si las necesitas
        actualizacion: [],
        experienciaPJ: [],
        hijos: [],
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

// Verificar existencia para el buscador [cite: 406, 407]
app.get("/existe/:noTrabajador", async (req, res) => {
  try {
    const [rows] = await db.query(
      "SELECT 1 FROM trabajador WHERE noTrabajador = ?",
      [req.params.noTrabajador],
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

// Iniciar servidor
const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Servidor corriendo en http://localhost:${PORT}`);
});
