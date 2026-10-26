# Tracking de Unidades de Transporte — SQL Server + Excel VBA

Proyecto base para registrar unidades, programar viajes, guardar novedades de tracking y consultar el estado desde Excel.

## Componentes
- `sql/01_CrearBaseDatos.sql`: crea la base `TrackingTransportes`, tablas y vista de tracking.
- `excel/Plantilla_Tracking_Transportes.xlsx`: plantilla de formularios y consulta.
- `vba/TrackingTransportes.bas`: macros VBA para insertar y consultar datos en SQL Server.

## Requisitos
- Windows + Microsoft Excel de escritorio (VBA no funciona en Excel web).
- SQL Server Express/Developer/Standard y SQL Server Management Studio (SSMS).
- Microsoft OLE DB Driver for SQL Server (`MSOLEDBSQL`).
- Permiso de conexión y escritura/lectura sobre la base de datos.

## Instalación
1. Abre SSMS y ejecuta `sql/01_CrearBaseDatos.sql`.
2. Abre la plantilla de Excel y guárdala como **Libro de Excel habilitado para macros (*.xlsm)**.
3. Pulsa `Alt + F11` para abrir el editor VBA.
4. En el editor: **File > Import File...** y selecciona `vba/TrackingTransportes.bas`.
5. En el módulo `TrackingTransportes`, cambia `SERVER_NAME = "localhost"` por el nombre de tu servidor/instancia, por ejemplo `MI-PC\SQLEXPRESS`.
6. En Excel, habilita macros solo para este archivo y ejecuta `ProbarConexionSQL` desde `Alt + F8`.
7. Para crear botones: activa la pestaña **Desarrollador** > **Insertar** > **Botón (Control de formulario)**. Dibuja cada botón y asigna la macro indicada:
   - `Registro`: **Guardar unidad** → `RegistrarUnidad`
   - `Registro`: **Guardar viaje** → `RegistrarViaje`
   - `Tracking`: **Guardar evento** → `RegistrarEventoTracking`
   - `Consultas`: **Consultar tracking desde SQL** → `ConsultarTracking`
   - Cualquier hoja: **Probar conexión** → `ProbarConexionSQL`
   - Cualquier hoja: **Limpiar formulario** → `LimpiarFormulario`

## Uso de la plantilla
### Hoja Registro
- Unidad: placa, tipo, marca, modelo, capacidad en toneladas y estado.
- Viaje: código único, placa existente, conductor, origen, destino, fecha de salida, documento, cliente, ETA, estado y observaciones.
- Registra primero la unidad; luego registra el viaje asociado a su placa.

### Hoja Tracking
- Ingresa código de viaje, ubicación, estado, fecha/hora, kilometraje y observación.
- Cada evento queda en el historial `dbo.EventosTracking`. La macro actualiza también el estado del viaje.

### Hoja Consultas
- El botón de consulta carga el tracking actual desde `dbo.vw_TrackingActual`.

## Estados permitidos
- Unidad: `Disponible`, `En ruta`, `Mantenimiento`, `Inactivo`.
- Viaje/evento: `Programado`, `En tránsito`, `En destino`, `Entregado`, `Incidencia`, `Cancelado`.

## Seguridad y notas
- La conexión usa autenticación integrada de Windows; no se guarda una contraseña en el archivo.
- La cuenta de Windows que abre Excel debe tener permisos en SQL Server.
- Los datos de ubicación son ingresados manualmente. Esto **no** es GPS en tiempo real; para ello se requiere integrar una fuente GPS/API del proveedor de telemetría.
- El código usa consultas parametrizadas en las inserciones para reducir riesgos de SQL injection.
- Para uso productivo, agrega roles de base de datos, auditoría, copias de seguridad y validación de datos.

## Publicación en GitHub
1. Descomprime el ZIP.
2. Crea un repositorio público o privado en GitHub.
3. Sube las carpetas `sql`, `excel`, `vba` y este `README.md`.
4. No subas contraseñas, cadenas con credenciales ni datos reales de conductores/vehículos.
