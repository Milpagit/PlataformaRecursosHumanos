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
/*
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
  // Método getConnection() para obtener el pool ya conectado
  getConnection: async () => {
    await poolConnect;
    return pool;
  },

  // Método para ejecutar consultas con parámetros tipo ?
  query: async (query, params = []) => {
    await poolConnect;
    const request = pool.request();

    // Reemplazar los parámetros ? por @param0, @param1, ...
    params.forEach((param, index) => {
      request.input(`param${index}`, param);
    });

    let paramIndex = 0;
    const transformedQuery = query.replace(
      /\?/g,
      () => `@param${paramIndex++}`,
    );

    const result = await request.query(transformedQuery);
    return [result.recordset];
  },
};
