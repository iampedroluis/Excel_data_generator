require('dotenv').config();

const xlsx = require("xlsx");
const path = require("path");
const os = require("os");
const readline = require("readline");


console.log('\x1b[33m%s\x1b[0m %s \x1b[33m%s\x1b[0m', ' ⚠ ', 'IMPORTANTE : RECUERDA NO TENER ABIERTO EL ARCHIVO EXCEL MIENTRAS CORRES EL SCRIPT', ' ⚠ ');

// Obtener el nombre de usuario desde las variables de entorno
const userName = process.env.USER_NAME;

// Verifica si la variable está definida
if (!userName) {
  console.error('\x1b[31m%s\x1b[0m ', ' ❌ ', 'El nombre de usuario no está definido en las variables de entorno.');
  console.error('\x1b[36m%s\x1b[0m ', ' ❕ ', `Puede añadirlo en el archivo .env de la carpeta raiz: 
    USER_NAME=TU_USERNAME` );
  process.exit(1);
}

// Configuración del módulo readline para obtener input del usuario
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// Conjunto para almacenar NMUs únicos
const nmuSet = new Set();

// Función para generar un NMU único
const generarNMU = () => {
  let nmu;
  do {
    const timestamp = Date.now().toString(); // Timestamp actual
    const randomChar = Math.random() < 0.5 ? 'N' : 'I'; // Generar 'N' o 'I'
    nmu = `P${timestamp.slice(-7)}${randomChar}`; // Últimos 7 dígitos del timestamp + 'N' o 'I'
  } while (nmuSet.has(nmu)); // Verificar si ya existe
  nmuSet.add(nmu); // Añadir NMU al conjunto
  return nmu;
};

// Función para formatear la fecha en 'YYYY-MM-DD'
const formatearFecha = (fecha) => {
  const year = fecha.getFullYear();
  const month = String(fecha.getMonth() + 1).padStart(2, '0'); // Mes de 2 dígitos
  const day = String(fecha.getDate()).padStart(2, '0'); // Día de 2 dígitos
  return `${year}-${month}-${day}`; // Retornar en el formato 'YYYY-MM-DD'
};

// Función para generar datos
const generarDatos = (cantidad) => {
  let data = [];
  
  for (let i = 0; i < cantidad; i++) {
    let thenmu = generarNMU();
    let fila = {
      "Nombre": `SAMSUNG PM ${thenmu}`, 
      "NMU": thenmu,
      "Código de Familia": `familycodex`, 
      "Marca": "Motorola",
      "Modelo": `X${55 + i}`,
      "Gama": "High",
      "SubGama": "High End A",
      "Tipo de SimCard": "Nano SIM",
      "Período de Garantía": "12 meses",
      "Susceptible de servicio técnico": "Si",
      "Orden de Entrega": "true",
      "Fecha Orden de Entrega": formatearFecha(new Date(Date.now() + 5 * 24 * 60 * 60 * 1000)) // Sumar 5 días a la fecha actual
    };
    data.push(fila);
  }
  return data;
};




// Preguntar cuántos datos quiere generar
rl.question('Cantidad de productos:  ', (respuesta) => {
  const cantidadDatos = parseInt(respuesta);

  if (isNaN(cantidadDatos) || cantidadDatos <= 0) {
    console.log('Por favor, ingresa un número válido.');
    rl.close();
    return;
  }


  // Generar los datos
  let datos = generarDatos(cantidadDatos);

  // Crear el libro de trabajo y la hoja
  let workbook = xlsx.utils.book_new();
  let worksheet = xlsx.utils.json_to_sheet(datos);

  // Aplicar formato de texto a la columna "Fecha Orden de Entrega"
  const range = xlsx.utils.decode_range(worksheet['!ref']);
  for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const cellRef = xlsx.utils.encode_cell({r: row, c: 11}); // Columna 11 es "Fecha Orden de Entrega"
    if (worksheet[cellRef]) {
      worksheet[cellRef].z = '@'; // Establecer formato de texto
    }
  }

  // Cargar el archivo existente 'CARGA_MASIVA_GDC.xlsx'
  const cargaMasivaPath = path.join("C:", "Users", userName, "Desktop", "dev", "Excel_data_generator", "excel_carga_masiva.xlsx");
  let cargaMasivaWorkbook = xlsx.readFile(cargaMasivaPath);

  // la hoja de trabajo a modificar es la primera hoja del libro
  let cargaMasivaSheetName = cargaMasivaWorkbook.SheetNames[0];
  let cargaMasivaSheet = cargaMasivaWorkbook.Sheets[cargaMasivaSheetName];

  // Limpiar los datos existentes en la hoja de trabajo 
  const cargaMasivaRange = xlsx.utils.decode_range(cargaMasivaSheet['!ref']);
  for (let row = cargaMasivaRange.s.r + 1; row <= cargaMasivaRange.e.r; row++) {
    for (let col = cargaMasivaRange.s.c; col <= cargaMasivaRange.e.c; col++) {
      const cellRef = xlsx.utils.encode_cell({ r: row, c: col });
      delete cargaMasivaSheet[cellRef]; // Eliminar cada celda
    }
  }

  // Insertar los nuevos datos generados en la hoja de trabajo
  xlsx.utils.sheet_add_json(cargaMasivaSheet, datos, { origin: "A1" });

  // Sobrescribir el archivo existente con los nuevos datos
  xlsx.writeFile(cargaMasivaWorkbook, cargaMasivaPath);



  console.log('%s \x1b[32m%s\x1b[0m',  'archivo creado con éxito', ' ✔ ');


  console.log(`
    se crearon:  ${cantidadDatos} productos.
    ...
    ..
    .
    ruta del archivo:
    ${cargaMasivaPath}
     `);
  
  
    console.log('\x1b[33m%s\x1b[0m', `⚠ ` + " NOTA: RECUERDA NO TENER ABIERTO EL ARCHIVO EXCEL MIENTRAS CORRES EL SCRIPT");

  // Cerrar readline
  rl.close();

  console.log(`
  `);

});
