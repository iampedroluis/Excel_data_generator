
# Generador de Datos Excel

Este proyecto es un generador de datos en Excel que permite crear y actualizar un archivo llamado `excel_carga_masiva.xlsx` con información sobre productos. Utiliza Node.js y la biblioteca `xlsx` para manejar archivos Excel.

## Requisitos

- [Node.js](https://nodejs.org/) instalado en tu computadora.
- Un archivo Excel llamado `excel_carga_masiva.xlsx` en la siguiente ruta: `C:\Users\<TU_USERNAME>\Desktop\dev\Excel_data_generator`.

## Instalación

1. Clona este repositorio en tu computadora:

   ```bash
   git clone <URL_DEL_REPOSITORIO>

#Navega al directorio del proyecto:


cd <NOMBRE_DEL_DIRECTORIO>

Instala las dependencias necesarias:

npm install dotenv xlsx

Crea un archivo .env en la raíz del proyecto y añade tu nombre de usuario:

USER_NAME=TU_USERNAME
Asegúrate de reemplazar TU_USERNAME con tu nombre de usuario real en el sistema.

Uso
Cierra el archivo Excel excel_carga_masiva.xlsx si está abierto. IMPORTANTE: El script no funcionará correctamente si el archivo está abierto.

Ejecuta el script:

node <NOMBRE_DEL_SCRIPT>.js

Reemplaza <NOMBRE_DEL_SCRIPT> con el nombre del archivo JavaScript que contiene el código.

Cuando se te pregunte, ingresa la cantidad de productos que deseas generar.

#Notas:

El script generará un conjunto único de NMUs (números de modelo únicos) y los insertará en la hoja de trabajo del archivo Excel.
La ruta del archivo Excel actualizado se mostrará en la consola una vez que el script se ejecute correctamente.
Contribuciones
Si deseas contribuir a este proyecto, siéntete libre de hacer un fork del repositorio y enviar un pull request.

Licencia
Este proyecto está bajo la Licencia MIT.





