# LeadMailer v2.3.0

## 🚀 Novedades principales

- Vista de `Leads` optimizada para ganar espacio (columnas ajustadas).
- En `Configuración` se añade `Nombre` y `Apellidos` de captadora.
- La captadora queda anexada al lead tratado (email, WhatsApp y cambio de etiqueta).
- Regla de duplicados actualizada: solo cuenta como duplicado si coincide `Email + Curso`.
- En `Alumnos`, el alta manual pasa a modal para liberar espacio del grid.
- El alta manual puede usar el curso seleccionado en el filtro de alumnos.
- Nuevo seguimiento por teléfono en leads:
  - check de llamado,
  - guardado automático de fecha/hora de llamada.

## 📤 Reportes de leads

- Envío de reporte al cerrar con confirmación (`Sí/No/Cancelar`).
- En `Configuración`:
  - email destino del reporte de cierre,
  - botón `Enviar reporte ahora`,
  - botón `Descargar reporte`.
- Nuevo filtro configurable `Fecha desde reporte` para limitar qué leads se incluyen en:
  - exportación CSV,
  - exportación XLSX,
  - envío manual del reporte,
  - envío al cerrar,
  - descarga manual.
- Nombre de archivo del reporte actualizado a:
  - `Leads_Asturias_<Nombre Apellidos Captadora>.xlsx`

## 📊 Exportaciones

- CSV/XLSX de leads ahora incluyen también:
  - `Etiqueta` del lead,
  - `LlamadoTelefono` (`Sí/No`),
  - `FechaLlamadaTelefono`.

## 🛠️ Instalador y versión

- Versión de app actualizada a `2.3.0`.
- Script de build del instalador actualizado para usar `2.3.0` por defecto.
- Script de Inno Setup actualizado a `2.3.0`.

## ✅ Compatibilidad

- Proyecto sobre `.NET 8` (WPF).
