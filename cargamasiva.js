require('dotenv').config();
const xlsx = require("xlsx");
const path = require("path");
const readline = require("readline");

console.log('\x1b[33m%s\x1b[0m %s \x1b[33m%s\x1b[0m', ' ⚠ ', 'IMPORTANTE : RECUERDA NO TENER ABIERTO EL ARCHIVO EXCEL MIENTRAS CORRES EL SCRIPT', ' ⚠ ');

// Obtener el nombre de usuario desde las variables de entorno
const userName = process.env.USER_NAME;

// Verifica si la variable está definida
if (!userName) {
  console.error('\x1b[31m%s\x1b[0m ', ' ❌ ', 'El nombre de usuario no está definido en las variables de entorno.');
  console.error('\x1b[36m%s\x1b[0m ', ' ❕ ', `Puede añadirlo en el archivo .env de la carpeta raiz: USER_NAME=TU_USERNAME`);
  process.exit(1);
}

// Configuración del módulo readline para obtener input del usuario
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// Función para generar un NMU único
const generarNMU = (nmuSet, caracter) => {
  let nmu;
  const timestamp = Date.now().toString(); // Timestamp actual

  nmu = `P${timestamp.slice(-7)}${caracter}`; // Últimos 7 dígitos del timestamp + carácter proporcionado
  return nmuSet.has(nmu) ? null : nmu; // Retorna nmu si es único, de lo contrario null
};

// Función para formatear la fecha en 'YYYY-MM-DD'
const formatearFecha = (fecha) => {
  const year = fecha.getFullYear();
  const month = String(fecha.getMonth() + 1).padStart(2, '0'); // Mes de 2 dígitos
  const day = String(fecha.getDate()).padStart(2, '0'); // Día de 2 dígitos
  return `${year}-${month}-${day}`; // Retornar en el formato 'YYYY-MM-DD'
};

// Función para generar datos
const generarDatos = (cantidad, nmuSet, caracter) => {
  let data = [];
  
  while (data.length < cantidad) {
    let thenmu = generarNMU(nmuSet, caracter);
    if (thenmu) {
      let fila = {
        "Nombre": `SAMSUNG PM ${thenmu}`, 
        "NMU": thenmu,
        "Código de Familia": `familycodex`, 
        "Marca": "Motorola",
        "Modelo": `X${55 + data.length}`,
        "Gama": "High",
        "SubGama": "High End A",
        "Tipo de SimCard": "Nano SIM",
        "Período de Garantía": "12 meses",
        "Susceptible de servicio técnico": "Si",
        "Orden de Entrega": "true",
        "Fecha Orden de Entrega": formatearFecha(new Date(Date.now() + 5 * 24 * 60 * 60 * 1000)) // Sumar 5 días a la fecha actual
      };
      data.push(fila);
      nmuSet.add(thenmu); // Añadir NMU al conjunto
    }
  }
  return data;
};

// Función para cargar NMUs existentes del archivo
const cargarNMUsExistentes = (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet);
  const nmuSet = new Set(data.map(item => item.NMU));
  return nmuSet;
};

// Función para preguntar por la finalización del NMU
const preguntarFinalizacion = (callback) => {
  rl.question('¿Desea finalizar el NMU con I, N o R? (I/N/R): ', (respuesta) => {
    if (['I', 'N', 'R'].includes(respuesta.toUpperCase())) {
      callback(respuesta.toUpperCase());
    } else {
      console.log('Por favor, elige una opción válida: I, N o R.');
      preguntarFinalizacion(callback); // Volver a preguntar
    }
  });
};

// Función para preguntar la cantidad de productos
const preguntarCantidad = (callback) => {
  rl.question('Cantidad de productos:  ', (respuesta) => {
    const cantidadDatos = parseInt(respuesta);
    if (!isNaN(cantidadDatos) && cantidadDatos > 0) {
      callback(cantidadDatos);
    } else {
      console.log('Por favor, ingresa un número válido.');
      preguntarCantidad(callback); // Volver a preguntar
    }
  });
};

// Preguntar la cantidad de productos
preguntarCantidad((cantidadDatos) => {
  // Cargar NMUs existentes
  const cargaMasivaPath = path.join("C:", "Users", userName, "Desktop", "dev", "Excel_data_generator", "excel_carga_masiva.xlsx");
  const nmuSet = cargarNMUsExistentes(cargaMasivaPath);

  // Preguntar por la finalización del NMU
  preguntarFinalizacion((respuestaFinalizacion) => {
    console.log(`Se finaliza el NMU con la opción: ${respuestaFinalizacion}`);

    // Generar los datos
    let datos = generarDatos(cantidadDatos, nmuSet, respuestaFinalizacion);

    // Crear el libro de trabajo y la hoja
    let workbook = xlsx.utils.book_new();
    let worksheet = xlsx.utils.json_to_sheet(datos);

    // Aplicar formato de texto a la columna "Fecha Orden de Entrega"
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
      const cellRef = xlsx.utils.encode_cell({ r: row, c: 10 }); // Columna 10 es "Fecha Orden de Entrega"
      if (worksheet[cellRef]) {
        worksheet[cellRef].z = '@'; // Establecer formato de texto
      }
    }

    // Cargar el archivo existente
    let cargaMasivaWorkbook = xlsx.readFile(cargaMasivaPath);
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

    console.log('%s \x1b[32m%s\x1b[0m', 'archivo creado con éxito', ' ✔ ');

    console.log(`

      se crearon:  ${cantidadDatos} productos.

      ruta del archivo:
      
      ${cargaMasivaPath}
    `);
    
    console.log('\x1b[33m%s\x1b[0m', `⚠ ` + " NOTA: RECUERDA NO TENER ABIERTO EL ARCHIVO EXCEL MIENTRAS CORRES EL SCRIPT");

    // Cerrar readline
    rl.close();
  });
});
