# Excel_data_generator

# Proyecto de Manipulación de Archivos Excel en Node.js

Este proyecto permite generar y modificar archivos de Excel, reemplazando datos existentes en un archivo de Excel con datos generados dinámicamente.

## Requisitos

- Node.js (versión 12 o superior)
- npm (gestor de paquetes de Node.js)

## Paquetes necesarios

Antes de ejecutar el proyecto, asegúrate de instalar los siguientes módulos utilizando `npm`:

1. **xlsx**: Para manipular y generar archivos Excel.

### Instalación de dependencias

Para instalar los paquetes necesarios, ejecuta el siguiente comando en la terminal:

```bash
npm install xlsx

Uso
El proyecto genera datos y reemplaza los existentes en un archivo Excel (CARGA_MASIVA_GDC.xlsx) ubicado en el escritorio. Sigue estos pasos para ejecutar el código:

Clona o descarga este repositorio en tu máquina local.
Asegúrate de tener un archivo llamado CARGA_MASIVA_GDC.xlsx en tu escritorio.
Ejecuta el script usando Node.js:
node index.js

Esto generará nuevos datos y reemplazará los datos antiguos en el archivo CARGA_MASIVA_GDC.xlsx en el escritorio.

Funciones principales
generarNMU: Genera un identificador único para cada entrada.
formatearFecha: Formatea las fechas en el formato YYYY-MM-DD.
generarDatos: Genera un conjunto de datos simulados para insertar en el archivo Excel.
Manipulación de archivos Excel: El script carga el archivo Excel existente, limpia los datos antiguos y los reemplaza por los nuevos generados.
