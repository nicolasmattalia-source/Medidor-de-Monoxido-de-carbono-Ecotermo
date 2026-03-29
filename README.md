# App de corrección de ensayos de Monóxido de Carbono

Esta app permite:

1. Cargar datos de ensayo del medidor de CO (`.xlsx` o `.csv`).
2. Cargar planillas con correcciones por medidor/modelo (factor y offset).
3. Calcular automáticamente `co_corregido` por fila.
4. Cargar una plantilla de Excel y completar celdas específicas.
5. Descargar el Excel final completo.

## Ejecutar

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Flujo de uso

- En la barra lateral, subir:
  - Archivo de ensayo.
  - Archivo de correcciones.
  - Plantilla de Excel.
- Seleccionar columnas clave:
  - ID/modelo del ensayo.
  - Columna de CO sin corregir.
  - ID/modelo, factor y offset de la planilla de correcciones.
- Revisar la vista previa y métricas.
- Configurar celdas de destino en la plantilla.
- Generar y descargar el Excel final.

## Supuestos de cálculo

La corrección se realiza con:

```text
co_corregido = lectura_cruda * factor + offset
```

Si una fila no tiene regla de corrección, la app usa `factor=1` y `offset=0`.
