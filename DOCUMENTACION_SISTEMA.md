# SISTEMA DE ÓRDENES DE COMPRAS RST
## Manual de Funcionalidades

---

## 📋 TABLA DE CONTENIDOS
1. Descripción General
2. Módulo de Login
3. Pantalla Principal
4. Órdenes de Compra
5. Gestión de Proveedores
6. Gestión de Usuarios
7. Gestión de Inventarios
8. Términos y Condiciones
9. Información Técnica

---

## 🎯 DESCRIPCIÓN GENERAL

**OCXRST (Orden de Compras RST)** es una aplicación de escritorio desarrollada en **Java** que permite gestionar de forma integral:

- ✅ Creación y generación de órdenes de compra en PDF
- ✅ Administración de proveedores
- ✅ Gestión de usuarios y permisos
- ✅ Control de inventario
- ✅ Autenticación segura
- ✅ Términos y condiciones de la empresa

**Tecnologías utilizadas:**
- **Lenguaje:** Java 17+
- **Interfaz:** Swing (GUI de escritorio)
- **Base de datos:** Archivos Excel (.xlsx)
- **Generación de PDFs:** iText 7
- **Build:** Maven

---

## 🔐 MÓDULO DE LOGIN (INICIARSESION)

### Funcionalidades:
- Autenticación de usuarios
- Validación de credenciales contra archivo Excel
- Interfaz visual con logo de la empresa

### Características técnicas:
- **Archivo de datos:** `USUARIOS.xlsx`
- **Validación:** Usuario y contraseña
- **Look and Feel:** Nimbus (interfaz moderna)
- **Icono personalizado:** Logo PDF2.jpeg

### Proceso de login:
1. Usuario ingresa nombre de usuario
2. Usuario ingresa contraseña
3. El sistema valida contra la base de datos Excel
4. Si es correcto → Acceso a menú principal
5. Si es incorrecto → Mensaje de error

---

## 🏠 PANTALLA PRINCIPAL (PRINCIPAL)

Menú central con acceso a todos los módulos:

### Botones disponibles:
1. **ORDEN DE COMPRAS** → Acceso a crear/editar órdenes
2. **PROVEEDORES** → Gestión de proveedores
3. **USUARIOS** → Gestión de usuarios del sistema
4. **INVENTARIOS** → Gestión de inventario
5. **TÉRMINOS Y CONDICIONES** → Ver/editar términos
6. **SALIR** → Cerrar aplicación

### Diseño:
- Fondo blanco profesional
- Logo empresarial (250x250 px)
- Color corporativo: Azul oscuro (RGB: 0, 95, 131)
- Interfaz intuitiva y centrada en pantalla

---

## 📄 MÓDULO DE ÓRDENES DE COMPRA (ORDENCOMPRA)

### Descripción:
Este es el módulo principal del sistema. Permite crear, editar y generar órdenes de compra con generación automática de PDFs.

### Características principales:

#### 1. **Información del Encabezado:**
- Número de orden (generado automáticamente)
- Fecha de documento
- Usuario solicitante (cargado automáticamente)
- Área/Departamento
- Proyecto (opcional)
- Cotización (número de referencia)

#### 2. **Información del Proveedor:**
- **Búsqueda automática:** Campo con autocompletado
- **Datos cargados:**
  - ID Proveedor
  - Nombre/Razón Social
  - RFC
  - Teléfono
  - Email
  - Dirección
  - Contacto
  - Método de pago
  - Tiempo de entrega

#### 3. **Tabla de Productos:**
- Código del producto
- Descripción
- Unidad (pieza, kg, metro, etc.)
- Precio unitario
- Cantidad
- Total por línea
- **Autocompletado de productos** desde inventario

#### 4. **Información de Pago:**
- Tipo de CFDI (Catálogo de Códigos)
- Moneda (pesos, dólares, euros, etc.)
- Forma de pago (efectivo, cheque, transferencia, etc.)
- Método de pago (varios disponibles)
- Descuento (opcional)
- Subtotal (calculado automáticamente)
- IVA (calculado automáticamente)
- **TOTAL (calculado automáticamente)**

#### 5. **Funciones principales:**
- **Crear orden:** Nuevo documento
- **Editar orden:** Modificar datos existentes
- **Generar PDF:** Exportar orden en formato PDF
- **Buscar orden:** Localizar por número
- **Eliminar:** Cancelar orden
- **Guardar:** Persistir datos en Excel

#### 6. **Validaciones:**
- Cálculo automático de totales
- Validación de campos requeridos
- Generación automática de número de orden
- Caché de inventario y proveedores para búsqueda rápida

### Archivos de datos:
- `INVENTARIOS.xlsx` → Productos disponibles
- `PROVEEDORES.xlsx` → Información de proveedores
- `USUARIOS.xlsx` → Datos de usuarios

---

## 🏢 MÓDULO DE PROVEEDORES (PROVEEDORES)

### Funcionalidades:
Gestión completa del catálogo de proveedores con la siguiente información:

### Campos de datos:
- **ID_PROVEEDOR** → Identificador único
- **NOMBRE_RAZON_SOCIAL** → Nombre completo o razón social
- **RFC** → Registro Federal de Contribuyentes
- **TELÉFONO** → Número de contacto
- **CORREO** → Email de contacto
- **DIRECCIÓN** → Domicilio fiscal
- **CONTACTO** → Nombre del contacto principal
- **PAGO** → Saldo pendiente/historial
- **MÉTODO_PAGO** → Forma de pago habitual
- **TIEMPO_ENTREGA** → Días de entrega
- **FORMA_PAGO** → Condiciones de pago
- **ACTIVO** → Estado del proveedor (sí/no)

### Operaciones disponibles:
1. **CREAR** → Agregar nuevo proveedor
2. **EDITAR** → Modificar datos existentes
3. **ELIMINAR** → Dar de baja proveedor
4. **BUSCAR** → Localizar proveedor por campos
5. **VOLVER** → Regresar al menú principal

### Persistencia:
Todos los datos se guardan en `PROVEEDORES.xlsx`

---

## 👥 MÓDULO DE USUARIOS (USUARIOS)

### Funcionalidades:
Administración de usuarios del sistema con control de acceso.

### Campos de datos:
- **USUARIO** → Nombre de usuario (login)
- **CONTRASEÑA** → Contraseña encriptada
- **NOMBRE** → Nombre completo
- **ÁREA** → Área o departamento
- **PUESTO** → Posición en la empresa
- **CORREO** → Email corporativo
- **FIRMA** → Firma digital/imagen

### Operaciones disponibles:
1. **CREAR** → Registrar nuevo usuario
2. **EDITAR** → Modificar datos del usuario
3. **ELIMINAR** → Dar de baja usuario
4. **BUSCAR** → Localizar usuario
5. **VOLVER** → Regresar al menú principal

### Características de seguridad:
- Validación de usuario único
- Gestión de permisos por área
- Registro de firma para documentos

### Persistencia:
Todos los datos se guardan en `USUARIOS.xlsx`

---

## 📦 MÓDULO DE INVENTARIOS (INVENTARIOS)

### Funcionalidades:
Control del catálogo de productos disponibles para crear órdenes.

### Campos de datos:
- **CÓDIGO** → SKU o código de producto
- **DESCRIPCIÓN** → Detalle del producto
- **UNIDAD** → Unidad de medida (pieza, kg, metro, etc.)
- **PRECIO_UNITARIO** → Costo unitario
- **CANTIDAD** → Stock disponible
- **ACTIVO** → Disponible para compra (sí/no)

### Operaciones disponibles:
1. **CREAR** → Agregar nuevo producto
2. **EDITAR** → Modificar datos del producto
3. **ELIMINAR** → Dar de baja producto
4. **BUSCAR** → Localizar por código o descripción
5. **VOLVER** → Regresar al menú principal

### Características:
- Búsqueda rápida por código
- Control de disponibilidad
- Actualización de precios
- Caché en memoria para autocompletado rápido

### Persistencia:
Todos los datos se guardan en `INVENTARIOS.xlsx`

---

## 📋 MÓDULO DE TÉRMINOS Y CONDICIONES (TERMINOS)

### Funcionalidades:
Visualización y edición de términos y condiciones de la empresa.

### Características:
- Editor de texto (JTextPane)
- Lectura de archivo: `Términos y Condiciones.txt`
- Guardado automático de cambios
- Persistencia en dos ubicaciones:
  - `target/classes/` (ejecución inmediata)
  - `src/main/java/` (persistencia ante recompilaciones)

### Botones:
- **GUARDAR** → Persistir cambios en archivo
- **VOLVER** → Regresar al menú principal

---

## 🔧 INFORMACIÓN TÉCNICA

### Estructura del proyecto:
```
OrdendeComprasRST/
├── src/main/java/com/mycompany/ocxrst/
│   ├── INICIARSESION.java      (Login)
│   ├── PRINCIPAL.java          (Menú principal)
│   ├── ORDENCOMPRA.java        (Órdenes de compra)
│   ├── PROVEEDORES.java        (Gestión proveedores)
│   ├── USUARIOS.java           (Gestión usuarios)
│   ├── INVENTARIOS.java        (Gestión inventarios)
│   ├── TERMINOS.java           (Términos y condiciones)
│   ├── IconoVentanaUtil.java   (Utilería de iconos)
│   └── BASES/
│       ├── USUARIOS.xlsx
│       ├── PROVEEDORES.xlsx
│       ├── INVENTARIOS.xlsx
│       └── Términos y Condiciones.txt
├── src/main/resources/com/mycompany/ocxrst/
│   ├── IMAGENES/
│   │   ├── Logo PDF2.jpeg
│   │   └── COPA.jpeg
│   └── TERMINOS/
│       └── Términos y Condiciones.txt
├── pom.xml
└── target/

```

### Dependencias principales:
1. **iText 7** (7.2.5) → Generación de PDFs
2. **Apache POI** (5.5.1) → Lectura/escritura Excel
3. **JCalendar** (1.4) → Componentes de fecha
4. **NetBeans AbsoluteLayout** → Layouts avanzados
5. **Java Swing** → Interfaz gráfica

### Base de datos:
- **Tipo:** Archivos Excel (.xlsx)
- **Ubicación:** `C:\OCXRST\OrdendeComprasRST\src\main\java\com\mycompany\ocxrst\BASES\`
- **Archivos:**
  - `USUARIOS.xlsx` - Credenciales y datos de usuarios
  - `PROVEEDORES.xlsx` - Catálogo de proveedores
  - `INVENTARIOS.xlsx` - Catálogo de productos

### Características de desarrollador:
- **Java Version:** 17+
- **Build Tool:** Maven
- **IDE:** NetBeans
- **Look and Feel:** Nimbus
- **Logging:** Java.util.logging
- **Encoding:** UTF-8

### Flujo de datos:
```
USUARIO → INICIARSESION → Validación Excel
        ↓
    PRINCIPAL → ORDENCOMPRA
                    ↓
            ┌─── PROVEEDORES (búsqueda)
            ├─── INVENTARIOS (búsqueda)
            └─── Genera PDF (iText)
                    ↓
            Persistencia en Excel
```

### Seguridad:
- Autenticación por usuario/contraseña
- Validación de credenciales en archivo Excel
- Registro de usuario actual en la sesión
- Icono personalizado para identificación visual

---

## 📊 FLUJO DE TRABAJO TÍPICO

### Crear una Orden de Compra:
1. Iniciar sesión con usuario y contraseña
2. Ir a "ORDEN DE COMPRAS" desde el menú principal
3. Completar información de encabezado:
   - Número (auto-generado)
   - Fecha
   - Usuario solicitante (auto-cargado)
   - Área
   - Proyecto (opcional)
4. Buscar y seleccionar proveedor (autocompletado)
5. Agregar productos:
   - Seleccionar producto del inventario (autocompletado)
   - Ingresar cantidad
   - El sistema calcula precio total automáticamente
6. Seleccionar forma de pago y método
7. Seleccionar CFDI y moneda
8. Revisar totales (automáticamente calculados)
9. Generar PDF
10. Guardar en base de datos

### Gestionar Proveedores:
1. Desde el menú principal → "PROVEEDORES"
2. Crear/Editar/Eliminar proveedores
3. Los cambios se sincronizar automáticamente
4. Regresar al menú principal

### Gestionar Inventarios:
1. Desde el menú principal → "INVENTARIOS"
2. Crear/Editar/Eliminar productos
3. Actualizar precios y stock
4. Los cambios se usan automáticamente en órdenes

---

## 📝 NOTAS ADICIONALES

- Todos los archivos de datos se encuentran en formato Excel (.xlsx)
- La aplicación genera PDFs profesionales de órdenes de compra
- El sistema usa caché en memoria para búsquedas rápidas
- Las fechas se seleccionan con calendario gráfico
- Los totales se calculan automáticamente
- La aplicación mantiene el usuario actual en sesión
- Interfaz responsive y centrada en pantalla
- Idioma: Español

---

**Versión:** 1.0-SNAPSHOT
**Autor:** Maria Luisa Martinez
**Fecha:** Marzo 2026
**Estado:** Producción

