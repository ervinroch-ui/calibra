# Calibra — Cartas de control de calidad

**Proyecto Flask listo para ejecutar** que permite cargar tantos Excel/CSV como necesites, analizar una columna numérica, establecer límites de aceptación y consultar cartas de control **I (individuales)** y **MR (rango móvil)**. Incluye interfaz adaptable al celular, historial, alertas y descarga de CSV y PowerPoint (.pptx).

## ¿Cómo probarlo ya en Windows?

1. Descomprimí este proyecto y abrí su carpeta en VS Code.
2. Abrí una terminal en la carpeta donde está `app.py`.
3. Ejecutá:

   ```powershell
   py -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   python app.py
   ```

4. Abrí **http://127.0.0.1:5000** en el navegador.
5. Probá con **`ejemplo_mediciones.xlsx`**. Elegí la hoja `Producción`, la columna **`Medición (mm)`**, las etiquetas **`Lote`**, mínimo **45** y máximo **55**.

Este ejemplo contiene **48 mediciones y 2 fuera de especificación** (una bajo el mínimo y otra sobre el máximo); la aplicación debe detectarlas.

> Si PowerShell bloquea la activación del entorno, usá sin activar: `.venv\Scripts\python -m pip install -r requirements.txt` y `.venv\Scripts\python app.py`.

## Informe G16 de una sola vez

El informe automático intenta procesar 9 ensayos de las hojas: **VOLUMETRÍA (Va, VMA, VFA, P/A)**, **COMPACTACION**, **TEMPERATURA**, **GRANULOMETRIA (Tamiz 1/2")**, **RUEDA DE HAMBURGO** y **TENSION DIAM Y RETENIDA (TSR)**.

Límites utilizados:

| Ensayo | Mínimo | Máximo | Procedencia |
|---|---:|---:|---|
| Va (%) | 3 | 5 | Referencias aportadas (Grupo 16) |
| VMA (%) | 14 | Sin límite | Referencias aportadas (Grupo 16) |
| VFA (%) | 65 | 75 | Referencias aportadas (Grupo 16) |
| Polvo/Asfalto | 0.6 | 1.2 | Referencias aportadas (Grupo 16) |
| Compactación (%) | 92 | 96 | Referencias aportadas (Grupo 16) |
| Temperatura (°C) | 145 | 165 | Campo de rangos del Excel Grupo 16 |
| Tamiz 1/2", Hamburgo y TSR | Sin confirmar | Sin confirmar | Se muestran gráficas, NO se califica cumplimiento |

Estas reglas están dirigidas al archivo **Grupo 16 CORASCO–LLANSA**. Antes de usarlas en otro proyecto, corroborá los límites técnicos correspondientes. No se modifican los datos o fórmulas del Excel original. Las cifras pueden variar respecto a presentaciones elaboradas con otros cortes o criterios de depuración; Calibra muestra filas válidas y omitidas.

## Nuevo: filtro por fechas (octubre y cualquier otro período)

En **Paso 02**, debajo de los límites, encontrás **Filtrar por período**. Elegí la columna **Fecha** y completá **Desde** y **Hasta**; el botón **Octubre 2026** coloca automáticamente `2026-10-01` y `2026-10-31`. Las dos fechas se incluyen. Para volver a considerar todas las fechas, pulsá **Quitar filtro**.

El filtro se aplica en el servidor **antes de calcular las cartas I–MR y antes de generar cualquier PowerPoint**. También se aplica a los **9 ensayos del informe G16 completo**, y se conserva dentro del historial, biblioteca y exportaciones. Se reconocen fechas numéricas de Excel (incluidos libros con sistema de fechas 1904) y fechas de texto ISO o DD/MM/AAAA. Cuando no hay al menos 3 mediciones válidas en el período, el análisis individual informa el motivo; en el informe G16 se omite el ensayo afectado y se explica. Sin fechas seleccionadas, el comportamiento anterior continúa igual.

**No se modifica el archivo Excel**: el programa archiva los bytes originales y filtra una copia temporal de filas en memoria. Los límites estadísticos son **recalculados solo con las mediciones del período**, por lo que los resultados pueden diferir de los del histórico completo. Para análisis de estabilidad deben revisarse tamaño de muestra, orden cronológico y pertinencia de recalcular límites en cada período.

## Cómo analizar TU Excel

- Archivos admitidos: **`.xlsx` y `.csv`**. Convertí archivos antiguos `.xls` a `.xlsx` con Excel.
- Cada hoja debe tener encabezados de columnas y mediciones numéricas por fila. Acepta varios nombres de hojas y filas con encabezado.
- Elegí **hoja**, **columna de medición**, **etiqueta/fecha opcional**, **límite inferior** y **límite superior**.
- Presioná **Generar cartas de control**. Podés subir otro archivo después sin reiniciar el servidor; cada análisis se agrega al historial.
- Cada análisis nuevo conserva **una copia exacta del Excel o CSV original y los resultados** en la base de datos. La biblioteca permite descargar esa copia sin alterarla. El historial conserva las cartas individuales.
- Hasta **12 MB por archivo** y **10.000 filas de datos por análisis** en esta primera versión.

## Biblioteca de Excel y compilado PowerPoint (VERSIÓN 2)

**Al hacer clic en “Generar cartas de control” o “Generar PowerPoint completo (G16)”**, Calibra crea automáticamente un registro nuevo en **Biblioteca Excel**. Solo se archivan cargas con análisis exitoso; inspeccionar un Excel sin analizarlo no lo archiva. Dos archivos con el mismo nombre y fechas distintas se guardan como registros separados.

Cada registro muestra **nombre y fecha**, tamaño, cantidad de ensayos, mediciones y el **semáforo de resultados**. Desde la biblioteca podés:

1. **▣ PowerPoint** — volver a descargar el PowerPoint completo de ese archivo, sin volver a subir el Excel.
2. **↓ Excel original** — recuperar exactamente los bytes que se analizaron, sin editar fórmulas ni formato.
3. **Ver ensayos** — abrir el análisis individual y sus cartas de control.
4. **Eliminar** — borrar voluntariamente esa copia archivada y sus análisis (con confirmación; irreversible).
5. **Elegir archivos + Descargar compilado** — marcar entre 1 y 10 Excel (hasta 60 análisis combinados) y crear un solo PPTX, con portada y semáforo general, seguido de separadores, semáforo individual y cartas de cada archivo. El compilado no mezcla ni recalcula la estadística entre los proyectos.

### PowerPoint automático

- El botón **Generar PowerPoint completo (G16)** analiza los nueve ensayos configurados para el Excel CORASCO–LLANSA y los registra bajo **un único archivo** en la biblioteca; además sigue permitiendo descarga inmediata del PPTX.
- En **Resultados del análisis**, usá **▣ Exportar PowerPoint** para la medición visible.
- En **Historial**, **▣ PowerPoint de todos los análisis** continúa funcionando como antes, incluso con registros antiguos que no estuvieran archivados en la nueva biblioteca.
- **Nuevo semáforo en PowerPoint:** panel con VERDE, AMARILLO, ROJO y NO EVALUABLE, número de parámetros en cada estado y tabla con lecturas / incumplimientos / estado. Se calcula con los resultados reales guardados, no con números fijos de una plantilla.
- El resto del informe **conserva las cuatro gráficas** (I, histograma, MR, Q–Q), estadísticas, detalle de incidencias y conclusiones técnicas.
- La **tolerancia poblacional 90/90** de las presentaciones de referencia **no se calcula todavía**. Las reglas del semáforo Calibra no equivalen al análisis completo 90/90 ni certifican contractualmente el proyecto.

### ¿Se conserva todo en Render?

**Solo si configurás una base de datos persistente mediante `DATABASE_URL`** (por ejemplo PostgreSQL externo). Tanto el **Excel original** como el **manifiesto de la biblioteca** y las **cartas calculadas** se guardan en la base de datos. El PowerPoint no ocupa almacenamiento adicional: se regenera a partir de los análisis conservados. Esto consume espacio en PostgreSQL a medida que analizás más archivos.

- En tu PC, el archivo `data/analyses.db` (SQLite) queda guardado mientras no borres esa carpeta. **Si venís de una versión anterior, conservá tu carpeta `data/`** al actualizar el proyecto para mantener el historial previo.
- En Render Free, el disco del servicio es efímero: si no configurás PostgreSQL, **podés perder el archivo histórico completo en reinicios o nuevos despliegues**. Tampoco lo guardés en una cuenta pública sin contraseña.
- Establecé `APP_PASSWORD` y un `SECRET_KEY` seguro en las variables de entorno: se almacenan documentos que pueden ser confidenciales. Implementá copias de seguridad de la base de datos para archivos importantes.

## Qué resultados interpreta

**Cumplimiento de especificaciones:** compara cada valor medido con los límites que ingresás, incluyéndolos como válidos (`LSL ≤ valor ≤ USL`). Señala cuántos están dentro, por debajo o por encima. Si no ingresás límites, no afirma que el lote cumpla. Si encuentra celdas de medición vacías o no numéricas, las reporta como omitidas; **no las cuenta como aprobadas**.

**Estabilidad del proceso:** usa una carta de individuales **I** y una de rango móvil **MR** (dos observaciones consecutivas) con constantes tradicionales del método:

- I: centro = promedio; límites = promedio ± 3·MR̅/1,128.
- MR: centro = MR̅; límite superior = 3,267·MR̅; límite inferior = 0.

Los **límites de control** calculados estadísticamente **no son los límites de especificación** que define el cliente o la norma. Puede haber valores dentro de especificación con señales de inestabilidad, y viceversa. Esta versión detecta puntos fuera de límites; no implementa todas las reglas complementarias de Nelson/Western Electric. Lo ideal es disponer de 20–25 o más mediciones consecutivas para estimar límites preliminares. No es un sistema certificado ni reemplaza un procedimiento de validación de calidad.

## Publicar en Render

1. Subí los archivos de **esta carpeta, sin cambiar su estructura**, a un repositorio de GitHub.
2. En Render elegí **New → Web Service → GitHub** y conectá el repositorio.
3. Configurá:

   | Campo | Valor |
   |---|---|
   | Build Command | `pip install -r requirements.txt` |
   | Start Command | `gunicorn --workers 1 --timeout 180 app:app` |
   | Plan | `Free` (si aparece disponible) |

4. En **Environment** agregá `SECRET_KEY` (clave aleatoria larga) y **`APP_PASSWORD`** (una contraseña segura). Si desplegás una aplicación sin `APP_PASSWORD`, el historial y los análisis podrán verse públicamente desde su enlace. **No cargués mediciones confidenciales en una instancia pública sin protección**.
5. Publicá y abrí la URL que te entregue Render.

### IMPORTANTE: conservar los datos entre despliegues

En tu computadora el historial se almacena por defecto en `data/analyses.db` (SQLite). En los servicios gratuitos de Render, **el almacenamiento local es efímero**, por lo que ese historial puede desaparecer al reiniciar o desplegar de nuevo. Para uso real, utilizá **PostgreSQL externo** y configurá `DATABASE_URL` en Render (por ejemplo una base PostgreSQL alojada en Neon). `psycopg2-binary` y Flask-SQLAlchemy ya están incluidos. Si todavía no necesitás guardar un historial duradero, podés probar sin conectar PostgreSQL.

**Seguridad:** la contraseña es compartida por el proyecto, no crea cuentas distintas. Recomendamos separar cuentas, roles y respaldos en una fase posterior para equipos grandes.

## Archivos principales

```
calibra_control_calidad/
├── app.py                    # Rutas web, biblioteca, reportes, login
├── excel_reader.py           # Lectura de XLSX/CSV y hojas
├── spc.py                    # Cálculos de las cartas I-MR
├── ppt_report.py             # Presentaciones PowerPoint basadas en resultados
├── batch_report.py           # Selección 9 ensayos del Excel G16
├── requirements.txt          # Dependencias Python
├── render.yaml               # Alternativa de despliegue Render Blueprint
├── ejemplo_mediciones.xlsx   # Archivo para probar
├── templates/
│   ├── index.html            # Panel principal
│   └── login.html            # Pantalla de acceso opcional
├── static/
│   ├── styles.css            # Estilos y diseño responsive
│   └── main.js               # Carga, gráficos interactivos y filtros
└── tests/
    ├── test_spc.py           # Pruebas de cálculos y lectura
    ├── test_ppt.py           # Pruebas de diapositivas
    └── test_library.py       # Pruebas integración de biblioteca Flask
```

## Observaciones técnicas pendientes

El análisis completo G16 se limita a los nueve ensayos preconfigurados y a los límites citados. La biblioteca puede archivar archivos Excel/CSV de otros proyectos **mediante análisis individual**, pero el botón G16 no es un clasificador universal de cualquier plantilla. Falta implementar reglas estadísticas adicionales, tolerancia 90/90 y autenticación multiusuario antes de uso crítico. No se modifican los Excel cargados.
