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
| `vercel.json` | Config de despliegue estático en Vercel |
| `deploy.bat` | Script de publicación manual (doble clic) |
| `.gitattributes` | Normaliza saltos de línea (LF) |
