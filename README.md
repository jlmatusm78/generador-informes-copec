# Generador Automático de Informes COPEC con envío por correo

Aplicación Streamlit para generar informes semanales desde un Excel con hojas `GUARDIAN` y/o `FLOTAGO`, y enviarlos automáticamente a cada transportista.

## Funciones

- Informe Global COPEC en PDF.
- Informe individual por cada transportista con alertas.
- ZIP automático con todos los informes.
- Evolución semanal y por tipo de alerta.
- Semáforo ejecutivo, reincidencia y gestión de fatiga.
- Envío por Gmail o Google Workspace mediante SMTP seguro.
- Tabla editable de destinatarios.
- Vista previa de cada correo.
- Envío de prueba antes del envío masivo.
- Selección de informes a enviar.
- Protección contra envíos duplicados dentro de la sesión.
- Registro descargable en CSV.
- Riesgo operacional explicado en cada informe, sin índice genérico de criticidad.
- Índice de reincidencia separado del riesgo operacional.
- Comparación predeterminada de 4 semanas completas, de lunes a domingo; selección de 4, 5 o 6 semanas.
- Comparación mensual configurable hasta 12 meses según la información disponible.
- Ranking configurable de conductores, con Top 10 predeterminado.
- PDF por transportista con anexo completo de eventos.
- Excel por transportista con hojas Resumen, Ranking conductores, Evolución y Detalle eventos.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Configurar Gmail en Streamlit Community Cloud

En la aplicación desplegada abre:

`Manage app > Settings > Secrets`

Pega lo siguiente, reemplazando los datos:

```toml
[gmail]
host = "smtp.gmail.com"
port = 587
username = "correo@tuempresa.cl"
password = "CONTRASENA_DE_APLICACION"
sender_email = "correo@tuempresa.cl"
sender_name = "Torre de Control COPEC"
use_tls = true
use_ssl = false
```

### Contraseña de aplicación

Para una cuenta Gmail o Google Workspace:

1. Activa la verificación en dos pasos en la cuenta remitente.
2. Crea una contraseña de aplicación para correo.
3. Usa esa contraseña en `password`; no uses la contraseña normal de la cuenta.
4. No subas `secrets.toml` a GitHub. El proyecto incluye `.gitignore` para evitarlo.

Si la organización bloquea contraseñas de aplicación, el administrador de Google Workspace deberá habilitarlas o proporcionar un relay SMTP autorizado.

## Tabla de destinatarios

Carga un Excel o CSV con estas columnas:

| Transportista | Para | CC | Activo |
|---|---|---|---|
| COPEC | gerencia@empresa.cl | operaciones@empresa.cl | SI |
| TRANSPORTES ACUBAR LIMITADA | prevencion@acubar.cl | supervisor@empresa.cl | SI |

- Usa `COPEC` para el destinatario del informe global.
- Puedes ingresar varios correos separados por punto y coma.
- `Activo = NO` deja el informe desmarcado inicialmente.
- La tabla puede corregirse directamente desde la aplicación antes del envío.

## Flujo recomendado

1. Subir el registro Guardian/FlotaGo.
2. Seleccionar la semana.
3. Generar informes.
4. Abrir la pestaña **Envío por correo**.
5. Cargar la tabla de destinatarios.
6. Revisar la vista previa.
7. Enviar un correo de prueba.
8. Marcar la confirmación.
9. Enviar los informes seleccionados.
10. Descargar el registro de envíos.

## Seguridad

- Los registros operacionales se cargan temporalmente en la sesión de Streamlit.
- Las credenciales se guardan en Streamlit Secrets, no en GitHub.
- La aplicación no almacena permanentemente las contraseñas ni los Excel cargados.

## Informes semanales y mensuales

La aplicación permite seleccionar **Semanal** o **Mensual** antes de generar los informes.

Todos los PDF incluyen la cabecera institucional COPEC 90 años en la esquina
superior izquierda, con proporción protegida y margen reservado para el contenido.

- Informe Global COPEC: 6 páginas para toma de decisiones.
- Informe por transportista: 6 páginas con la misma estructura ejecutiva.
- Hallazgos automáticos, evolución por tipo de alerta y variación contra el período anterior.
- Concentración del riesgo, transportistas/conductores prioritarios y plan de acción con responsables y plazos.
- Análisis técnico de Sensor Tapado y Sensor desalineado: evolución, tractos afectados, reincidencia y prioridades por transportista.
- Tendencias configurables de 3, 4, 6 o 12 períodos.
- Ranking ponderado de conductores por severidad.
- Comparación con el período anterior.
- Tendencias por tipo de alerta.
- Indicadores de fatiga, reincidencia y concentración del riesgo.
- Correos y nombres de archivos adaptados al modo semanal o mensual.

El riesgo operacional pondera Fatiga sin cumplimiento (10), Fatiga pendiente (8), Sin cinturón y Conductor fumando (6), y Fatiga con cumplimiento (3). El informe muestra el cálculo, la regla aplicada y el nivel resultante.

Los nombres históricos se aceptan como alias internos, pero todos los informes,
tablas y archivos exportados muestran únicamente `Sensor desalineado` y
`Sensor Tapado`.
