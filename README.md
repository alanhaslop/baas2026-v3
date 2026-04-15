# BAAS 2026 — Gestor de vtas v3

Sistema completo de gestión de ventas para el congreso BAAS 2026.  
Proyecto limpio e independiente de versiones anteriores.

## Estructura

```
baas2026-v3/
├── frontend/
│   └── index.html     ← Toda la UI (GitHub Pages)
└── backend/
    └── Code.gs        ← Google Apps Script completo
```

## Configuración antes del congreso

### 1. Google Sheet nueva
Crear una hoja vacía en Google Drive. Copiar su ID de la URL.  
El script crea las pestañas automáticamente al primer uso.

### 2. Carpeta en Drive para PDFs
Crear carpeta en Drive. Copiar su ID de la URL.

### 3. Apps Script
1. Ir a script.google.com → Nuevo proyecto
2. Nombrar: `BAAS2026-v3`
3. Pegar el contenido de `backend/Code.gs`
4. Completar el bloque `CONFIG` al inicio con los datos reales
5. Guardar (Ctrl+S)
6. Ejecutar `instalarTrigger` una vez (para el alerta diario de reservas)
7. Desplegar → Nueva implementación → Aplicación web
   - Ejecutar como: **Yo**
   - Acceso: **Cualquier persona**
8. Copiar la URL del despliegue

> ⚠️ Siempre "Nueva versión" al desplegar cambios — no solo guardar.

### 4. Frontend
1. Crear repo `baas2026-v3` en GitHub (privado)
2. Subir contenido de `frontend/` a la rama `main`
3. Settings → Pages → Branch: main / folder: /frontend
4. En `index.html`, reemplazar `TU_URL_DEL_APPS_SCRIPT_AQUI` con la URL del paso 3.8

## Roles y acceso

| Rol | Contraseña (configurar en Code.gs) | Pantallas |
|-----|-----------------------------------|-----------|
| Operador | `PASSWORD_OPERADOR` | Formulario, reservas, modificaciones, cierre |
| Administrativa | `PASSWORD_ADMIN` | Panel kanban, modificar al retirar, cierre |

## Flujos cubiertos

- Venta directa con confirmación inmediata
- Reserva con pago diferido (stock comprometido)
- Cierre de reserva (con ajuste de monto)
- Modificación de reserva antes del cobro
- Modificación al retirar (nuevo recibo)
- Cancelación con código supervisor
- Panel de despacho kanban (4 estados)
- Alerta diaria de reservas vencidas (7:30 AM)
- Cierre del día con arqueo completo
- WhatsApp en 1 click en cada momento del flujo
- Reimprimir comprobante desde cualquier pantalla

## Comprobantes

| Tipo | Marca visual |
|------|-------------|
| Venta directa | Sin marca |
| Reserva | Watermark "PENDIENTE DE PAGO" |
| Cierre de reserva | Referencia "Cancela reserva #XXX" |
| Modificación | Referencia "Modifica pedido #XXX" |
| Anulación | Watermark "ANULADO" |

## Comandos para subir el repo

```bash
cd baas2026-v3
git init
git branch -m main
git add .
git commit -m "feat: sistema completo BAAS 2026 v3"
git remote add origin https://github.com/alanhaslop/baas2026-v3.git
git push -u origin main
```
