Sistema de Gestion de Gimnasio

Trabajo Final Integrador
Algoritmos y Estructuras de Datos
2026

Integrantes:
Juan Cruz Díaz
Marcos Audisio
German Tomas Catuogno
Mariano Valentino Lazarte
Lisandro Fabrizio Diaz

Comision: C

Descripcion general del sistema
Nuestro sistema es una aplicacion de gestión de gimnasio desarrollada en Python que permite administrar de forma integral la informacion de socios, actividades y cuotas
Da las herramientas para registrar y dar de baja socios, gestionar distintos tipos de membresias, controlar pagos, inscribir personas en actividades, registrar asistencias y consultar estadisticas basicas, ademas, toda la informacion se almacena de forma permanente en un archivo excel, lo que permite conservar los datos entre distintas ejecuciones del programa sin necesidad de volver a cargarlos

Funcionalidades principales
Gestion de socios: alta, baja listado y filtrado (por membresia estado o actividad)
Gestion de cuotas: registro de pagos y consulta del estado de cuotas por socio
Gestion de actividades: alta de actividades, inscripcion de socios y listado con control de cupo
Control de asistencia: registro de asistencias (validando inscripcion previa) y estadisticas basicas de concurrencia (por actividad y por socio)
Persistencia de datos: toda la informacion se guarda automaticamente en un archivo excel (gimnasio_datos.xlsx), que se crea la primera vez que se ejecuta el programa y se lee o actualiza en las ejecuciones siguientes

herramientas utilizadas
Python 3.13
OpenPyXL (lectura, escritura y persistencia de datos en formato excel(.xlsx)) 

Instrucciones de ejecucion: **PARA QUE EL PROGRAMA CORRA BIEN DEBE EJECUTARSE EN LA MISMA CARPETA QUE "build" Y "dist"**
Abrir el archivo .exe compilado,
  **El cual se encuentra en la carpeta dist/**                              
o bien ejecutar el codigo directamente con Python:
   python gymnasioC7.py

Interactuar con la consola que se genera: el sistema muestra un menu principal desde el cual se accede a las distintas gestiones, socios, cuotas, actividades y asistencias
La primera vez que se ejecuta el programa, este detecta que no existe el archivo excel de datos (gimnasio_datos.xlsx) y lo crea automaticamente, en las ejecuciones siguientes, el programa lee ese archivo para recuperar toda la informacion cargada previamente
El uso es directo: cada menu indica las opciones disponibles mediante numeros, y el sistema guia al usuario con validaciones y mensajes claros ante errores de carga

**Nota: si se usa el .exe, el archivo gimnasio_datos.xlsx se genera en la misma carpeta donde se encuentra el ejecutable**

--

**Uso de Inteligencia Artificial**
Utilizamos principalmente Claude (Anthropic) como herramienta de apoyo durante el desarrollo, especificamente para:
Entender el funcionamiento y la sintaxis de la libreria OpenPyXL (lectura, escritura y persistencia de datos en archivos excel)
Depurar errores puntuales relacionados con el manejo de esta libreria, dada la falta de experiencia previa del grupo con librerias externas de Python
Todas las decisiones de diseño, la logica del sistema y las validaciones fueron entendidas y revisadas por el equipo antes de agregarse al proyecto






