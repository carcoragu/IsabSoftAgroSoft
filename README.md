# IsabSoft | Plataforma Agroganadera

Plataforma inteligente para gestión ganadera con IA conversacional, modo offline y sincronización con Access 2010.

## Productos

| Módulo | Descripción |
|---|---|
| **IsabSoft \| AgroSoft** | ERP ganadero: compra, venta, faena, stock |
| **IsabSoft \| ReproGan** | Gestión reproductiva: inseminación, protocolos, preñez |
| **IsabSoft \| IA Chat** | Asistente conversacional offline-first |

## Requisitos

- Windows 10/11
- IIS activado (Panel de control → Características de Windows)
- Microsoft Access 2010 o superior instalado
- Microsoft ACE OLEDB 12.0 (viene con Office)

## Instalación

### 1. Clonar o descargar el repositorio

```
https://github.com/TU_USUARIO/isabsoft
```

### 2. Configurar la base de datos

```
a. Copiá config.example.asp → config.asp
b. Editá config.asp con la ruta de tu base de datos:
   Const DB_RUTA = "C:\tu\ruta\isabsoft.accdb"
c. Creá una base de datos Access vacía en esa ruta
```

### 3. Crear las tablas automáticamente

```
Abrí en el navegador: http://localhost/isabsoft/instalador.asp
```

Esto crea todas las tablas en tu base de datos Access.

> ⚠️ Ejecutar una sola vez. Luego eliminá o renombrá instalador.asp.

### 4. Configurar IIS

```
a. Abrí el Administrador de IIS
b. Sitios → Default Web Site → Agregar aplicación
c. Alias: isabsoft
d. Ruta física: C:\Users\HOME\isabsoft (o donde clonaste)
e. Asegurate de que ASP esté habilitado
```

### 5. Acceder a la aplicación

```
http://localhost/isabsoft/index.html
```

Desde celular en la misma red WiFi:
```
http://IP-DE-TU-PC/isabsoft/index.html
```

## Estructura del proyecto

```
isabsoft/
├── index.html              ← Aplicación principal
├── config.asp              ← Configuración local (NO en GitHub)
├── config.example.asp      ← Plantilla de configuración
├── instalador.asp          ← Crea las tablas en Access
├── .gitignore
├── README.md
└── api/
    ├── registrar.asp       ← Guarda registros en Access
    └── sync.asp            ← Procesa cola offline
```

## Tablas de la base de datos

| Tabla | Descripción |
|---|---|
| `Animales` | Stock general de animales |
| `ComprasAnimales` | Compras de bovinos |
| `VentasAnimales` | Ventas de bovinos |
| `ComprasInsumos` | Compras de balanceado, medicina, etc. |
| `Inseminaciones` | Registro reproductivo |
| `ColaSyncOffline` | Cola de sincronización offline |
| `Clientes` | Directorio de clientes |
| `Proveedores` | Directorio de proveedores |

## Funcionamiento offline

1. Sin internet → los registros se guardan en IndexedDB del navegador
2. Al reconectar → el sistema detecta la conexión automáticamente
3. Sincronización → envía la cola pendiente a Access via ASP

## Contacto

IsabSoft · Paraguay · +595 981 153477
