const express = require('express'); //"Esto es como un import numpy as np en python"
const app = express(); //"Se crea la aplicación, app llama a express para utilizarlo"
"^"
"l"
"Es como abrir un documento nuevo de word para trabajar en el"


const path = require('path'); //path es un modulo nativo de node.js para manejar rutas archivos
/*
const conexion = mysql.createConnection({
  host: "localhost" ,
  database: 'trabajadores_servicio',
  user: 'root',
  password: ''
});*/

app.set('view engine', 'ejs');  //view engine es una opcion especial, ejf es Embedded JavaScript
// Ruta inicial para probar el servidor


app.use(express.static(path.join(__dirname, 'public')));

//permite a express servir archivos estaticos desde la carpeta public. Es decir que el
//servidor los envia directamente al navegador 
//los archivos que no cambian dinamicamente cson el css, javascript, imagenes, fuentes, etc.

//Servir significa enviar o entregar los archivos estaticos cuando los pida el navegador
"__dirname es una variable especial de Node.js, "
"la carpeta donde esta el archivo actual, se une a public"

/*
conexion.connect(function(error){
  if(error){
    console.log(error);
  }
  else{
    console.log("Conexión exitosa");
  }
});

conexion.end
*/

async function obtenerTrabajadores() {
  try {
    const conexion = await conectarDB();
    
    const [rows] = await conexion.execute('SELECT * FROM trabajador');
    
    console.log("📋 Lista de trabajadores:", [rows]); // 🔍 Imprime los resultados

    console.log(rows[0]);
    
    await conexion.end();
  } catch (error) {
    console.error("❌ Error al obtener trabajadores:", error);
  }
}






"************************"
//Define una ruta GET para la URL / la pagina principal

//req manda una solicitud que hace el usuario, res es lo que va a responer el servidor
app.get('/', (request, response) => {
  //response.send('¡Hola, Isac! El servidor está funcionando correctamente.');
    response.render('index')
});

//El segundo parametro de app.get es la funcion callback que se ejecuta cuando alguien visita la ruta

"app.get es para manejar las solicitudes get"
"su primer parametro es la ruta GET, en este caso es la principal por eso no lleva nada más"





// Iniciar el servidor en el puerto 3000
app.listen(3000, () => console.log('Servidor corriendo en http://localhost:3000'));


//app.listen(3000, function() {
  //console.log('Servidor corriendo en http://localhost:3000');
//});



// app.listen(puerto , callback) el callback es opcional

"El numero de puero es en donde quiero el servidor"
"El callback es la funcion que quiero que se ejecute cuando el servidor empieza a funcionar"


//npx tailwindcss-cli -i ./public/css/tailwind.css -o ./public/css/styles.css --watch"


"node app.js para ejecutarlo"


//npx tailwindcss -i ./public/css/tailwind.css -o ./public/css/styles.css --watch
// .\ngrok.exe http 3000
// node app.js



obtenerTrabajadores();