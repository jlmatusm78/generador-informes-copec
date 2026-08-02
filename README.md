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

- Informe Global COPEC: 4 páginas ejecutivas.
- Informe por transportista: 3 páginas.
- Tendencias configurables de 3, 4, 6 o 12 períodos.
- Ranking ponderado de conductores por severidad.
- Comparación con el período anterior.
- Tendencias por tipo de alerta.
- Indicadores de fatiga, reincidencia y concentración del riesgo.
- Correos y nombres de archivos adaptados al modo semanal o mensual.

El puntaje de criticidad pondera Fatiga (5), Uso celular (4), Sin cinturón y Tapado de cámara (3), Fumar y Sin conductor (2), Cámara desalineada y Bostezo (1).
