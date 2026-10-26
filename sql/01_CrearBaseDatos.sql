/* Tracking de Unidades de Transporte - SQL Server
   Ejecutar en SQL Server Management Studio (SSMS).
*/
IF DB_ID(N'TrackingTransportes') IS NULL
    CREATE DATABASE TrackingTransportes;
GO
USE TrackingTransportes;
GO

IF OBJECT_ID('dbo.Unidades', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Unidades (
        UnidadID INT IDENTITY(1,1) PRIMARY KEY,
        Placa NVARCHAR(15) NOT NULL UNIQUE,
        TipoUnidad NVARCHAR(40) NOT NULL,
        Marca NVARCHAR(50) NULL,
        Modelo NVARCHAR(50) NULL,
        CapacidadTon DECIMAL(10,2) NULL,
        Estado NVARCHAR(20) NOT NULL CONSTRAINT DF_Unidades_Estado DEFAULT 'Disponible',
        FechaRegistro DATETIME2 NOT NULL CONSTRAINT DF_Unidades_Fecha DEFAULT SYSDATETIME(),
        CONSTRAINT CK_Unidades_Estado CHECK (Estado IN ('Disponible','En ruta','Mantenimiento','Inactivo'))
    );
END;
GO

IF OBJECT_ID('dbo.Viajes', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Viajes (
        ViajeID INT IDENTITY(1,1) PRIMARY KEY,
        CodigoViaje NVARCHAR(30) NOT NULL UNIQUE,
        UnidadID INT NOT NULL,
        Conductor NVARCHAR(120) NOT NULL,
        DocumentoConductor NVARCHAR(20) NULL,
        Origen NVARCHAR(150) NOT NULL,
        Destino NVARCHAR(150) NOT NULL,
        Cliente NVARCHAR(120) NULL,
        FechaSalida DATETIME2 NOT NULL,
        ETA DATETIME2 NULL,
        FechaLlegada DATETIME2 NULL,
        Estado NVARCHAR(25) NOT NULL CONSTRAINT DF_Viajes_Estado DEFAULT 'Programado',
        Observaciones NVARCHAR(500) NULL,
        UsuarioRegistro NVARCHAR(100) NULL,
        FechaRegistro DATETIME2 NOT NULL CONSTRAINT DF_Viajes_Fecha DEFAULT SYSDATETIME(),
        CONSTRAINT FK_Viajes_Unidades FOREIGN KEY (UnidadID) REFERENCES dbo.Unidades(UnidadID),
        CONSTRAINT CK_Viajes_Estado CHECK (Estado IN ('Programado','En tránsito','En destino','Entregado','Incidencia','Cancelado'))
    );
END;
GO

IF OBJECT_ID('dbo.EventosTracking', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.EventosTracking (
        EventoID BIGINT IDENTITY(1,1) PRIMARY KEY,
        ViajeID INT NOT NULL,
        FechaEvento DATETIME2 NOT NULL CONSTRAINT DF_Eventos_Fecha DEFAULT SYSDATETIME(),
        Ubicacion NVARCHAR(180) NOT NULL,
        Estado NVARCHAR(25) NOT NULL,
        Kilometraje INT NULL,
        Observacion NVARCHAR(500) NULL,
        RegistradoPor NVARCHAR(100) NULL,
        CONSTRAINT FK_Eventos_Viajes FOREIGN KEY (ViajeID) REFERENCES dbo.Viajes(ViajeID),
        CONSTRAINT CK_Eventos_Estado CHECK (Estado IN ('Programado','En tránsito','En destino','Entregado','Incidencia','Cancelado'))
    );
END;
GO

CREATE OR ALTER VIEW dbo.vw_TrackingActual AS
WITH UltimoEvento AS (
    SELECT e.*, ROW_NUMBER() OVER (PARTITION BY e.ViajeID ORDER BY e.FechaEvento DESC, e.EventoID DESC) AS rn
    FROM dbo.EventosTracking e
)
SELECT v.ViajeID, v.CodigoViaje, u.Placa, v.Conductor, v.Cliente, v.Origen, v.Destino,
       v.FechaSalida, v.ETA, v.FechaLlegada,
       COALESCE(ue.Estado, v.Estado) AS EstadoActual,
       ue.FechaEvento AS UltimaActualizacion, ue.Ubicacion AS UltimaUbicacion,
       ue.Kilometraje, ue.Observacion
FROM dbo.Viajes v
JOIN dbo.Unidades u ON u.UnidadID = v.UnidadID
LEFT JOIN UltimoEvento ue ON ue.ViajeID = v.ViajeID AND ue.rn = 1;
GO

/* Datos de ejemplo opcionales */
IF NOT EXISTS (SELECT 1 FROM dbo.Unidades WHERE Placa = 'ABC-123')
    INSERT dbo.Unidades (Placa, TipoUnidad, Marca, Modelo, CapacidadTon, Estado)
    VALUES ('ABC-123','Camión','Volvo','FH',20,'Disponible');
GO
