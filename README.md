
# Generador de Datos Excel

Este proyecto es un generador de datos en Excel que permite crear y actualizar un archivo llamado `excel_carga_masiva.xlsx` con información sobre productos. Utiliza Node.js y la biblioteca `xlsx` para manejar archivos Excel.

## Requisitos

- [Node.js](https://nodejs.org/) instalado en tu computadora.
- Al Clonar el repositorio hacerlo en el Escritorio dentro de una carpeta dev (IMPORTANTE)
- Un archivo Excel llamado `excel_carga_masiva.xlsx` en la siguiente ruta: `C:\Users\<TU_USERNAME>\Desktop\dev\Excel_data_generator`.

## Instalación

1. Clona este repositorio en tu computadora:

```bash
git clone https://github.com/iampedroluis/Excel_data_generator.git
```

2. Navega al directorio del proyecto:

```bash
cd C:\Users\<TU_USERNAME>\Desktop\dev\Excel_data_generator
````

3. Instala las dependencias necesarias:

```bash
npm install
```

4. Crea un archivo .env en la raíz del proyecto y añade tu nombre de usuario:

```code

USER_NAME=TU_USERNAME

````

__* Asegúrate de reemplazar TU_USERNAME con tu nombre de usuario real en el sistema.__

## Uso

⚠  Cierra el archivo Excel `excel_carga_masiva.xlsx` si está abierto. IMPORTANTE: El script no funcionará correctamente si el archivo está abierto.

# Ejecuta el script:

```bash
node cargamasiva.js
```
__* Completar la cantidad de Terminales a Crear.__




# Notas:
El script generará un conjunto único de NMUs (números de modelo únicos) y los insertará en la hoja de trabajo del archivo Excel.
La ruta del archivo Excel actualizado se mostrará en la consola una vez que el script se ejecute correctamente.
Contribuciones
Si deseas contribuir a este proyecto, siéntete libre de hacer un fork del repositorio y enviar un pull request.





# EVIDENCIAS
![image](https://github.com/user-attachments/assets/ca444743-9c32-47cf-ba5c-dfa97b5c78b4)
![image](https://github.com/user-attachments/assets/9202c4c7-595c-4951-98fe-44a33b3d6cf6)


















## Licencia
`Este proyecto está bajo la Licencia MIT.`





