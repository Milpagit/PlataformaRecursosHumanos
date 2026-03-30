# Plataforma de Recursos Humanos — Poder Judicial del Estado de Hidalgo

Sistema web para que los trabajadores del Poder Judicial del Estado de Hidalgo registren y actualicen su información personal, académica y laboral. Permite generar currículums en Excel y PDF, cargar documentos oficiales (INE, RFC, etc.) y administrar el padrón de trabajadores desde un panel de administración.

---

## Tabla de contenido

1. [Requisitos](#requisitos)
2. [Instalación](#instalación)
3. [Variables de entorno](#variables-de-entorno)
4. [Estructura del proyecto](#estructura-del-proyecto)
5. [Arquitectura y funcionamiento](#arquitectura-y-funcionamiento)
6. [Rutas del sistema](#rutas-del-sistema)
7. [Base de datos](#base-de-datos)
8. [Generación de reportes](#generación-de-reportes)
9. [Panel de administración](#panel-de-administración)
10. [Scripts disponibles](#scripts-disponibles)

---

## Requisitos

- Node.js v18 o superior
- SQL Server (instancia local o remota)
- Cuenta de correo SMTP (Office 365 por defecto)

---

## Instalación

```bash
# 1. Clonar el repositorio
git clone <url-del-repositorio>
cd PlataformaRecursosHumanos

# 2. Instalar dependencias
npm install

# 3. Crear el archivo .env con las variables necesarias (ver sección siguiente)

# 4. Iniciar en modo desarrollo
npm run dev

# 5. Iniciar en producción
npm start
```

El servidor escucha en el puerto **3001** por defecto (`http://localhost:3001`).

---

## Variables de entorno

Crear un archivo `.env` en la raíz del proyecto con los siguientes valores:

```env
# Base de datos SQL Server
DB_USER=usuario_sql
DB_PASSWORD=contraseña_sql
DB_HOST=localhost
DB_PORT=1433
DB_NAME=nombre_bd

# Sesiones
SESSION_SECRET=clave_secreta_larga_y_aleatoria

# URL base de la aplicacion (usada en hipervinculos de reportes)
BASE_URL=http://localhost:3001

# Correo SMTP (para envio de contraseñas y notificaciones)
SMTP_HOST=smtp.office365.com
SMTP_PORT=587
SMTP_USER=correo@dominio.mx
SMTP_PASS=contraseña_correo
```

---

## Estructura del proyecto

```
PlataformaRecursosHumanos/
├── server.js          # Punto de entrada principal — todas las rutas y lógica del servidor
├── app.js             # Configuración alternativa/auxiliar de Express (rutas secundarias)
├── config/
│   └── db.js          # Pool de conexión a SQL Server con mssql
├── views/             # Plantillas EJS para trabajadores
│   ├── login.ejs
│   ├── index.ejs      # Formulario principal del trabajador
│   ├── acuse.ejs      # Acuse de recibo tras guardar
│   ├── buscador.ejs
│   ├── insert.ejs
│   ├── progreso.ejs
│   ├── recuperar-password.ejs
│   ├── registrarPassword.ejs
│   ├── success.ejs
│   └── admin/
│       ├── login.ejs
│       ├── dashboard.ejs
│       └── usuario-detalle.ejs
├── public/            # Archivos estáticos (CSS, JS, imágenes)
│   ├── css/
│   │   ├── styles.css
│   │   └── tailwind.css
│   └── js/
│       └── formulario.js   # Lógica del formulario del trabajador en el cliente
├── plantillas/        # Archivos Excel base para generacion de curriculum (.xlsx)
├── uploads/           # PDFs subidos por los trabajadores (INE, RFC, etc.)
├── reportes/          # Reportes Excel y PDF generados dinámicamente (temporales)
├── tailwind.config.js
└── package.json
```

---

## Arquitectura y funcionamiento

### Stack tecnológico

| Componente         | Tecnología                                     |
| ------------------ | ---------------------------------------------- |
| Servidor           | Node.js + Express                              |
| Vistas             | EJS (servidor) + Tailwind CSS                  |
| Base de datos      | SQL Server (driver `mssql`)                    |
| Sesiones           | `express-session`                              |
| Subida de archivos | `multer` (disco local)                         |
| Reportes Excel     | `ExcelJS` sobre plantillas `.xlsx`             |
| Reportes PDF       | `Puppeteer` (headless Chrome)                  |
| Descarga ZIP       | `archiver`                                     |
| Envío de correo    | `nodemailer`                                   |
| Fechas             | `dayjs` con zona horaria `America/Mexico_City` |

### Flujo del trabajador

1. El trabajador ingresa con **CURP + RFC + contraseña** en `/login`.
2. Se crea una sesión con `req.session.user` que contiene todos los datos del trabajador.
3. En `/index` el trabajador llena o actualiza su información personal, familiar, académica, laboral y sube documentos PDF.
4. Al enviar el formulario (`POST /guardar`), el sistema:
   - Actualiza la tabla `trabajador`
   - Elimina y vuelve a insertar registros en `hijo`, `escolaridad`, `actualizacionprofesional` y `experienciapj`
   - Guarda los archivos subidos en la carpeta `uploads/` con el nombre `{nombreTrabajador}_{tipo}.pdf`
5. Se redirige al acuse de recibo en `/acuse/:noTrabajador/:tipoTrabajador`.

### Autenticación y sesiones

- La sesión del trabajador se guarda en `req.session.user`.
- La sesión del administrador se guarda en `req.session.admin`.
- Las rutas de administrador están protegidas por el middleware `verificarAdmin`.
- Un trabajador solo puede ver su propio acuse y curriculum; el administrador puede acceder a todos.

---

## Rutas del sistema

### Rutas del trabajador

| Método | Ruta                                        | Descripción                                           |
| ------ | ------------------------------------------- | ----------------------------------------------------- |
| GET    | `/`                                         | Redirige a `/login`                                   |
| GET    | `/login`                                    | Formulario de inicio de sesión                        |
| POST   | `/login`                                    | Autenticar con CURP + RFC + contraseña                |
| GET    | `/index`                                    | Formulario principal (requiere sesión)                |
| POST   | `/guardar`                                  | Guardar datos del trabajador (con subida de archivos) |
| GET    | `/acuse/:noTrabajador/:tipoTrabajador`      | Acuse de recibo de los datos guardados                |
| GET    | `/curriculum/:noTrabajador/:tipoTrabajador` | Redirige a la vista del reporte                       |
| GET    | `/curriculum_transparencia/:no/:tipo`       | Enlace público del curriculum (sin autenticación)     |
| GET    | `/reporte/:noTrabajador/:tipoTrabajador`    | Generar y descargar reporte ZIP (Excel + PDF)         |
| GET    | `/registrar-password`                       | Formulario para registrar contraseña inicial          |
| GET    | `/existe/:noTrabajador/:tipoTrabajador`     | Verificar si un trabajador existe (JSON)              |
| GET    | `/buscador`                                 | Página de búsqueda pública                            |
| GET    | `/success`                                  | Página de éxito genérica                              |
| GET    | `/logout`                                   | Cerrar sesión del trabajador                          |

### Rutas del panel de administración

| Método | Ruta                           | Descripción                                         |
| ------ | ------------------------------ | --------------------------------------------------- |
| GET    | `/admin/login`                 | Formulario de login administrador                   |
| POST   | `/admin/login`                 | Autenticar administrador                            |
| GET    | `/admin/dashboard`             | Panel principal con lista de trabajadores           |
| POST   | `/admin/crear-usuario`         | Dar de alta a un nuevo trabajador                   |
| POST   | `/admin/buscar-trabajadores`   | Buscar trabajadores por nombre/número               |
| GET    | `/admin/usuario/:idTrabajador` | Ver detalle de un trabajador                        |
| POST   | `/admin/generar-acuse`         | Generar acuse desde el panel admin                  |
| POST   | `/admin/editar-datos-criticos` | Editar datos sensibles (noTrabajador, CURP, RFC...) |
| POST   | `/admin/eliminar-trabajador`   | Eliminar un trabajador del sistema                  |
| GET    | `/admin/reporte-transparencia` | Reporte global para transparencia                   |
| GET    | `/admin/logout`                | Cerrar sesión del administrador                     |

---

## Base de datos

La conexión se configura en `config/db.js` usando un pool de `mssql`. Las tablas principales son:

| Tabla                      | Descripción                                                              |
| -------------------------- | ------------------------------------------------------------------------ |
| `trabajador`               | Datos personales, laborales y de contacto del trabajador                 |
| `hijo`                     | Hijos del trabajador (iniciales, sexo, fecha de nacimiento)              |
| `escolaridad`              | Estudios (bachillerato, licenciatura, maestría, doctorado, especialidad) |
| `experienciapj`            | Experiencia laboral en el Poder Judicial u otras instituciones           |
| `actualizacionprofesional` | Cursos y actualizaciones profesionales                                   |

El módulo `config/db.js` expone dos métodos:

- `db.query(sql, params)` — Ejecuta una consulta parametrizada. Si `params` es array usa `?` como placeholder; si es objeto usa `@nombre`.
- `db.getConnection()` — Devuelve el pool directo para operaciones avanzadas.

---

## Generación de reportes

La ruta `GET /reporte/:noTrabajador/:tipoTrabajador` genera el curriculum completo:

1. Lee la plantilla `plantillas/curriculum.xlsx` con ExcelJS.
2. Rellena los datos del trabajador, escolaridades, actualizaciones y experiencias, duplicando filas dinámicamente para acomodar múltiples registros.
3. Guarda el Excel resultante en `reportes/`.
4. Usa Puppeteer para renderizar `plantillas/curriculum2.xlsx` (segunda hoja) como PDF.
5. Empaqueta el Excel editable + el PDF + los documentos oficiales subidos (`uploads/`) en un archivo ZIP.
6. Envía el ZIP al cliente y elimina los archivos temporales.

> **Importante:** La carpeta `reportes/` contiene archivos temporales. No deben commitearse al repositorio.

---

## Panel de administración

El administrador accede por `/admin/login`. Sus credenciales se verifican contra una tabla de administradores en la base de datos. Desde el dashboard puede:

- Buscar y listar trabajadores registrados.
- Crear nuevos trabajadores con sus datos iniciales.
- Ver el detalle completo de cualquier trabajador.
- Editar datos sensibles (CURP, RFC, número de trabajador, contraseña).
- Eliminar trabajadores junto con todos sus registros relacionados.
- Generar acuses individuales.
- Descargar el reporte de transparencia global.

---

## Scripts disponibles

```bash
npm start      # Inicia el servidor con node server.js
npm run dev    # Inicia con nodemon (reinicia automáticamente al guardar cambios)
```

---

## Notas para el desarrollador

- El archivo principal con **toda la lógica del servidor** es `server.js`. El archivo `app.js` contiene una configuración auxiliar y algunas rutas secundarias (recuperar contraseña, etc.).
- Los formularios del cliente están en `public/js/formulario.js`. Este archivo controla la visibilidad dinámica de secciones del formulario según las respuestas del trabajador.
- Las fechas siempre se manejan en zona horaria `America/Mexico_City` usando `dayjs`. Los helpers `convertirFecha` y `convertirFechaTexto` en `server.js` normalizan los diferentes formatos que llegan del formulario antes de guardarlos en SQL Server.
- Los archivos subidos se nombran `{nombreTrabajador}_{tipodocumento}.pdf`. Si el mismo trabajador sube un nuevo documento, el anterior se sobreescribe automáticamente.
- Las variables de entorno **nunca deben commitearse**. Crear siempre el `.env` localmente con base en la tabla de la sección [Variables de entorno](#variables-de-entorno).
