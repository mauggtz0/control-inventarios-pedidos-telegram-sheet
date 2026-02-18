# 📦 Sistema Profesional de Control de Pedidos e Inventario  
Google Sheets + Apps Script + Telegram

Sistema integral desarrollado en Google Sheets con automatización avanzada mediante Google Apps Script.

---

## 🚀 Funcionalidades

✅ Control completo de pedidos  
✅ Generación automática de tickets  
✅ Control de facturación  
✅ Control de surtido  
✅ Salida automática de inventario  
✅ Kardex automático  
✅ Inventario resumen en tiempo real  
✅ Registro de entradas sin duplicar  
✅ Reporte diario automático  
✅ Envío de tickets a Telegram  

---

## 🧠 Flujo del Sistema

Pedido → Facturación → Surtido → Salida a reparto →  
Descuento automático de inventario → Kardex →  
Inventario actualizado → Cierre con documento recibido

---

## 🗂 Estructura de Hojas

### PEDIDOS_CONTROL
Control principal del flujo operativo.

- Ticket automático
- Timestamps automáticos
- Colores por estado
- Hasta 10 productos por pedido
- Descuento automático al marcar “SALIO_A_REPARTO”

---

### CATALOGO_PRODUCTOS
Lista maestra de productos activos.

- Producto
- Stock inicial
- Activo (SI/NO)

---

### ENTRADAS
Registro manual de compras o entradas.

- Evita duplicados
- Actualiza Kardex
- Recalcula inventario

---

### KARDEX (Automático)
Registro completo de movimientos:

- ENTRADA_INICIAL
- ENTRADA_COMPRA
- SALIDA_PEDIDO

---

### INVENTARIO_RESUMEN (Automático)

- Existencia actual
- Total entradas
- Total salidas

---

## 🤖 Integración Telegram

Permite enviar el ticket estructurado directamente a Telegram.

Formato enviado:

🧾 TICKET  
👤 Cliente  
📦 Productos  
📌 Estado  
🚚 Repartidor  

---

## 🛠 Tecnologías

- Google Sheets
- Google Apps Script
- Telegram Bot API
- PropertiesService
- UrlFetchApp

---

## ⚙️ Instalación

1. Crear Google Sheets
2. Ir a Extensiones → Apps Script
3. Copiar CODE.gs
4. Guardar
5. Autorizar permisos
6. Configurar token Telegram

---

## 🔐 Seguridad

- Token Telegram guardado en PropertiesService
- Prevención de duplicados
- Control de inventario protegido

---

## 📈 Beneficios

- Inventario en tiempo real
- Kardex automático
- Control operativo profesional
- Reducción de errores manuales
- Reportes instantáneos
- Integración directa con mensajería

---

## 👨‍💻 Autor

Sistema desarrollado para distribución veterinaria y control logístico profesional.
