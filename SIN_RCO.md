# Identificación de conductores - Sin RCO

Actualización del 14 de septiembre de 2026 sobre la versión vigente del repositorio.

- Nuevo logo proporcionado por el usuario, en la misma cabecera y sin deformación.
- Resumen visual de Sin RCO, porcentaje y comparación con el período anterior.
- Explicación del efecto sobre los KPIs y detalle de casos por regularizar.
- Alertas Sin RCO incluidas en totales, tipos de evento y puntos de riesgo operacional.
- Ranking, conteo de conductores y reincidencia calculados con conductores identificados.
- La reincidencia conserva el criterio vigente de dos o más eventos.
- Los casos Sin RCO no se agrupan como una persona ni activan por sí solos la regla de riesgo crítico por reincidencia de fatiga sin cumplimiento.
- Los marcadores vacíos o sin identificación se muestran separados de los Sin RCO explícitos.
- El porcentaje se refiere a alertas, no a personas ni viajes; sin base se muestra No aplica.
- Se conservan informes mensuales, análisis de sensores, detalle completo y envío de correo.

Detección en la columna Conductor: coincidencia completa de Sin RCO, ignorando mayúsculas, espacios y puntuación. No se buscan coincidencias parciales en nombres.

Verificación: 14 pruebas unitarias; generación de informes global y por transportista, semanales y mensuales, y caso con solo Sin RCO. No se enviaron correos en las pruebas.
