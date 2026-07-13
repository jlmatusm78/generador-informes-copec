# Generador Automático de Informes COPEC

Aplicación Streamlit para generar informes semanales a partir de un Excel con hojas `GUARDIAN` y/o `FLOTAGO`.

## Funciones

- Informe Global COPEC en PDF.
- Informe individual por cada transportista con alertas.
- ZIP automático con todos los informes.
- Evolución semanal y por tipo de alerta.
- Semáforo ejecutivo.
- Rankings de empresas y conductores.
- Fatiga y cumplimiento del descanso mínimo.
- Indicador de reincidencia, sin IRO.
- Matriz empresa vs tipo de alerta.

## Ejecutar localmente

1. Instalar Python 3.11 o superior.
2. Abrir una terminal en esta carpeta.
3. Ejecutar:

```bash
pip install -r requirements.txt
streamlit run app.py
```

4. Abrir la dirección que muestra Streamlit, normalmente `http://localhost:8501`.

## Streamlit Community Cloud

Sube esta carpeta a un repositorio de GitHub y selecciona `app.py` como archivo principal al crear la aplicación.

## Formato esperado

El Excel debe contener al menos una hoja llamada `GUARDIAN` o `FLOTAGO`. Las columnas principales utilizadas son:

- ID
- Fecha
- Transportista
- Conductor
- Tracto
- Incidente
- Conductor se detine mínimo 15 minutos

La aplicación tolera columnas faltantes no críticas y normaliza automáticamente los tipos de alerta.
