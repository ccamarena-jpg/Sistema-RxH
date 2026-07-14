# Sistema RxH — Registro de Pagos personal de campo (TT Audit)

Plataforma web para gestionar el ciclo de pagos por Recibo por Honorarios del
personal de campo: registro, aprobación de cuentas, emisión de RxH, contabilidad,
dashboard y carga masiva. Es una app de un solo archivo (`index.html`) con backend
en Google Sheets vía Apps Script.

## Despliegue automático

El flujo es:

```
editas index.html  →  commit + push a GitHub  →  Vercel despliega solo
```

### Formas de publicar cambios

- **Automático (con Claude Code):** al terminar cada sesión de cambios se hace
  `commit` + `push` solo (configurado en `.claude/settings.json`).
- **Manual (editando a mano):** doble clic en **`deploy.bat`**. Sube tus cambios
  a GitHub; Vercel despliega en segundos.

### Enlaces

- Repositorio: https://github.com/ccamarena-jpg/Sistema-RxH
- Producción (Vercel): _pendiente de conectar en vercel.com_

## Estructura

| Archivo | Qué es |
|---------|--------|
| `index.html` | La aplicación completa (HTML + CSS + JS) |
| `apps-script.gs` | Backend (Google Apps Script) con verificación de identidad |
| `vercel.json` / `netlify.toml` | Config de despliegue estático |
| `deploy.bat` | Script de publicación manual (doble clic) |
| `.gitattributes` | Normaliza saltos de línea (LF) |

## Seguridad del backend

El frontend envía en cada petición el **token de identidad de Google** del usuario
logueado. El Apps Script (`apps-script.gs`) lo verifica contra Google
(`verificar_`): confirma que sea una cuenta **@ttaudit.com** verificada y no
expirada; si no, responde `{error:'no-autorizado'}` y no ejecuta ninguna acción.

La sesión de Google dura ~1 hora: al expirar, la app avisa y recarga para volver
a iniciar sesión.

### Cómo actualizar el backend (redeploy)

El orden importa para no romper la app en producción:

1. **Primero** publica el frontend (ya lo hace el auto-push). El nuevo `index.html`
   envía el token pero sigue funcionando con el backend viejo (el parámetro extra
   se ignora), así que no hay corte.
2. **Luego** abre el Google Sheet → *Extensiones → Apps Script*, pega el contenido
   de `apps-script.gs` (reemplazando todo) y ve a *Implementar → Gestionar
   implementaciones → editar (lápiz) → Versión: Nueva → Implementar*. Mantén
   **Ejecutar como: Yo** y **Quién tiene acceso: Cualquiera**.
3. Prueba: inicia sesión, verifica que carga la data y que puedes aprobar/registrar.

### Calendario de asistencia (columna nueva)

El backend nuevo usa la columna **AL "DIAS DETALLE"** en la hoja `REGISTROS`,
donde se guarda el patrón de días trabajados (ej. `1:1,2:0.5,15:1`). El encabezado
se crea **automáticamente** (`asegurarHeaderDiasDetalle_`) al guardar o leer registros;
también puedes ejecutar la función `configurarHoja()` una vez desde el editor de Apps
Script para crearlo de inmediato. La pestaña
**Asistencia** muestra el calendario por quincena (con filtros de ejecutivo, cliente,
proyecto y supervisor); el supervisor marca los días al registrar y se autocompletan
los "Días trabajados" (soporta medio día = ½).
