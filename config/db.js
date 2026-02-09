const sql = require("mssql");

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

const pool = new sql.ConnectionPool(config);
const poolConnect = pool.connect();

module.exports = {
  // Para obtener el pool directamente si necesitas hacer algo muy específico
  getConnection: async () => {
    await poolConnect;
    return pool;
  },

  // Método unificado que detecta si mandas un Array [] o un Objeto {}
  query: async (queryStr, params = []) => {
    await poolConnect;
    const request = pool.request();

    if (Array.isArray(params)) {
      // ESTILO POSICIONAL (Sustituye los '?' por @paramN)
      params.forEach((param, index) => {
        request.input(`param${index}`, param);
      });
      let i = 0;
      const transformedQuery = queryStr.replace(/\?/g, () => `@param${i++}`);
      const result = await request.query(transformedQuery);
      // Retornamos [rows] para mantener compatibilidad con tu código actual
      return [result.recordset];
    } else {
      // ESTILO NOMBRADO (Usa objetos como { nombre: valor })
      for (const [name, value] of Object.entries(params)) {
        // Si el valor viene como { type, value }, usamos el tipo específico
        if (value && typeof value === "object" && "type" in value) {
          request.input(name, value.type, value.value);
        } else {
          request.input(name, value);
        }
      }
      const result = await request.query(queryStr);
      return result.recordset;
    }
  },
};
