const sql = require("mssql");
//Entorno de desarrollo
const config = {
  user: "JesusML",
  password: "M1lp426",
  server: "192.168.73.5",
  port: 14357,
  database: "PersonalPrue",
  options: {
    encrypt: false,
    trustServerCertificate: true,
  },
};
/* Produccion 
const config = {
  user: 'P@rents25',
  password: 'ChildrenRH25',
  server: '192.168.73.5',
  port: 14357,
  database: 'PersonalRH',
  options: {
    encrypt: false,
    trustServerCertificate: true,
  }
};
*/

const pool = new sql.ConnectionPool(config);
const poolConnect = pool.connect();

module.exports = {
  // Obtener el pool conectado
  getConnection: async () => {
    await poolConnect;
    return pool;
  },

  // Ejecutar consultas con parámetros nombrados
  query: async (query, params = {}) => {
    if (!query) throw new Error("Se requiere un query válido.");

    await poolConnect;
    const request = pool.request();

    // Agregar parámetros
    for (const [name, { type, value }] of Object.entries(params)) {
      request.input(name, type, value);
    }

    const result = await request.query(query);
    return result.recordset;
  },
};
