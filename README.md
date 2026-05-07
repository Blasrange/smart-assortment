# Smart Assortment (Surtidor Logistico)

Aplicacion Angular para gestion operativa de surtido con tres vistas principales:

- Inicio
- Productos
- Ventas (analisis de faltantes y sugerencias de resurtido)

## Stack Tecnologico

- Angular 21 (standalone components)
- PrimeNG 21 + PrimeIcons + PrimeFlex
- Tema PrimeUIX Aura
- Signals de Angular para estado local
- SSR habilitado con @angular/ssr + Express
- XLSX para lectura y exportacion de Excel

## Requisitos

- Node.js 20+
- npm 10+

## Instalacion Y Ejecucion

1. Instalar dependencias:

```bash
npm install
```

2. Levantar entorno de desarrollo:

```bash
npm start
```

3. Abrir en navegador:

http://localhost:4200/

## Scripts Disponibles

- npm start: inicia servidor de desarrollo (ng serve)
- npm run build: compila aplicacion
- npm run watch: compilacion en modo watch para desarrollo
- npm test: ejecuta pruebas unitarias
- npm run serve:ssr:surtidor_logistico: ejecuta build SSR en Node

## Estructura Funcional

### Inicio

Vista de portada con acceso al flujo operativo del sistema.

### Productos

Modulo conectado a API publica DummyJSON para operaciones CRUD basicas:

- Listado con filtros y ordenamiento
- Creacion de producto
- Edicion de producto
- Inactivacion logica (PATCH active: false)

Servicio: src/app/service/api.products.ts

### Ventas (Modulo Principal de Analisis)

El modulo de ventas trabaja con 2 archivos de entrada:

- Archivo de ventas (CSV/XLSX/XLS/TXT)
- Archivo de inventario (CSV/XLSX/XLS/TXT)

Incluye validacion de columnas, analisis de cobertura y exportacion a Excel.

#### Flujo Funcional

1. Carga de archivos por click o drag and drop.
2. Validacion de formato y columnas requeridas.
3. Mapeo flexible de encabezados (acepta aliases).
4. Agrupacion por SKU de ventas e inventario.
5. Clasificacion de ubicaciones por picking/reserva.
6. Generacion de:
	 - sugerencias de surtido
	 - productos faltantes
	 - metricas de cobertura
7. Exportacion a Excel en dos hojas:
	 - Surtido
	 - Faltantes

#### Columnas Esperadas (Con Alias)

Ventas:

- material
- descripcion
- cantidadConfirmada

Inventario:

- sku
- descripcion
- localizacion
- disponible
- estado

Nota: el mapeo contempla multiples nombres de columna para cada campo (por ejemplo SKU/Codigo/Material, etc.), por lo que no depende de un solo encabezado fijo.

#### Reglas De Negocio Clave

- Solo se considera inventario con estado valido (o sin estado reconocible) y se excluyen ubicaciones ignoradas.
- El stock se separa entre picking y reserva.
- Si picking no cubre ventas, se propone surtido desde reserva.
- La seleccion de ubicaciones usa criterio FEFO (prioriza menor FPC y vencimiento mas cercano).
- Si no hay inventario o no hay reserva suficiente, se marcan faltantes.
- Para alta rotacion (umbral >= 10 unidades), se permite sugerir estibas completas.

#### Salida Del Analisis

- Resumen general del proceso
- Tabla de faltantes por SKU
- Tabla de surtido sugerido por SKU y ubicacion
- KPIs:
	- total productos
	- productos a surtir
	- productos faltantes
	- porcentaje de cobertura

## Arquitectura De Frontend

- Navegacion principal con menubar y rutas standalone.
- Rutas actuales:
	- /
	- /products
	- /sales
- Internacionalizacion parcial de textos de filtros PrimeNG a espanol.

## Estado Actual Del Proyecto

- El menu muestra una seccion Reportes, pero no existe ruta funcional para ese modulo aun.
- Existe una prueba en app.spec.ts que espera un titulo antiguo y puede requerir ajuste para reflejar el template actual.
- En sales.ts hay un metodo placeholder analizarArchivos() no utilizado.

## Recomendaciones De Mejora

1. Crear modulo real de Reportes con rutas y pantallas.
2. Unificar tipado de modelos (ventas, inventario, sugerencias) en archivos compartidos.
3. Agregar pruebas unitarias al algoritmo de analisis de ventas.
4. Corregir y ampliar pruebas de interfaz para Home, Productos y Ventas.
5. Parametrizar constantes de negocio (estados validos, niveles de picking/reserva, umbral de alta rotacion).