# Sistema para gestión de información masiva — Tiendas Ancla

## Qué hace

La aplicación recibe el archivo Excel base y lee la hoja **Consolidado**. Detecta las tiendas con fotografías en las tres columnas:

- **S:** Febrero Fotos
- **T:** Marzo 2026
- **U:** Abril 2026

En el Excel base validado existen **856 tiendas**, cada una repetida en **3 filas**, para un total de **2.568 filas** a clasificar.

Por cada tienda, la aplicación:

1. Muestra sus tres filas asociadas.
2. Presenta las tres fotos directamente en pantalla.
3. Permite seleccionar una clasificación:
   - GANA — verde
   - NO GANA — rojo
   - NESTUM EN LA EXHIBICION — rojo
   - NESTOGENO EN LA EXHIBICION — rojo
   - NESTUM Y NESTOGENO EN LA EXHIBICION — rojo
   - NO TIENE 3 O MAS CATEGORIAS — rojo
   - POCOS PRODUCTOS — amarillo
   - NO ES LA MISMA TIENDA — amarillo
4. Descarga una copia del Excel con la decisión en **W — ANOTACIONES** y con las **tres filas del IDCliente** coloreadas de **A a W**.

La columna **V — OBSERVACIÓN** no se reemplaza y queda disponible como antecedente de revisión.

## Instalación en Windows

### Requisito previo
Instala Python 3.11 o superior desde la página oficial de Python y activa la opción **Add Python to PATH** durante la instalación.

### Ejecución rápida
1. Descomprime la carpeta del sistema.
2. Haz doble clic en `EJECUTAR_SISTEMA.bat`.
3. Espera a que se abra la interfaz en el navegador.
4. Sube el Excel base.
5. Revisa las tiendas, guarda las clasificaciones y descarga el Excel actualizado.

## Ejecución manual alternativa

Abre una terminal en esta carpeta y ejecuta:

```bash
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Privacidad y funcionamiento

- El Excel se procesa localmente en el equipo donde se ejecuta la aplicación.
- La aplicación no envía el Excel a bases de datos ni servicios externos.
- Para mostrar las fotografías, consulta las URL que ya están incluidas en el archivo Excel, alojadas en Google Cloud Storage.
- El archivo original no se sobrescribe; se genera un nuevo `.xlsx` al descargar.

## Archivos del sistema

- `app.py`: interfaz de revisión.
- `motor_excel.py`: lectura del archivo, agrupación por IDCliente y exportación conservando estructura.
- `requirements.txt`: dependencias necesarias.
- `EJECUTAR_SISTEMA.bat`: inicio rápido para Windows.


## Versión 3 — Exclusión de tiendas ya validadas

Al cargar el Excel, una tienda se omite automáticamente si cualquiera de sus tres filas ya contiene información en:

- **V — OBSERVACIÓN**
- **W — ANOTACIONES**

Así, un `IDCliente` previamente trabajado no vuelve a aparecer en la cola de revisión.
La interfaz también incluye el botón **Guardar y siguiente pendiente**.


## Versión 4 — Corrección de exportación de Excel

Se corrigió la generación del archivo `.xlsx` para conservar las declaraciones internas
de compatibilidad del libro original. Esta corrección evita que Excel solicite reparar
el archivo descargado al abrirlo.


## Versión 5 — Recuperación de archivos exportados previamente

Además de excluir tiendas ya validadas, el exportador restaura los prefijos internos estándar de Excel
si se carga un archivo generado por una versión anterior. Esto permite continuar desde un respaldo
sin mantener el aviso de reparación.


## Versión 6 — Memoria persistente y rendimiento

Esta versión agrega mejoras importantes de seguridad y velocidad:

### Memoria persistente local

Cada vez que se guarda una clasificación, el sistema escribe automáticamente un respaldo en:

```text
data/sesiones
```

Esto permite recuperar el avance si se cierra el navegador, se reinicia Streamlit o se apaga el computador antes de descargar el Excel final.

Al volver a cargar el mismo archivo, la aplicación recupera las decisiones guardadas y muestra un aviso con la cantidad restaurada.

### Respaldo JSON

En la barra lateral aparece el botón **Descargar respaldo JSON**. Este archivo contiene únicamente las decisiones guardadas y sirve como respaldo liviano del avance.

### Caché local de imágenes

Las imágenes descargadas desde las URL se guardan en:

```text
data/cache_imagenes
```

Si vuelves a una tienda o reinicias la aplicación, las fotos ya descargadas se cargan desde disco, lo que mejora la velocidad.

### Descarga optimizada del Excel

Para mejorar el rendimiento, el Excel actualizado ya no se genera automáticamente en cada cambio de tienda. Ahora se genera solamente cuando presionas:

```text
Preparar Excel actualizado
```

Después aparece el botón:

```text
Descargar Excel actualizado
```

### Colores exactos en Excel

Los colores quedan configurados en formato ARGB correcto para Excel:

- Verde: `FF92D050`
- Rojo: `FFFF0000`
- Amarillo: `FFFFFF00`


## Versión 8 — Preparada para Streamlit Community Cloud

Cambios principales:

- Caché de imágenes limitado para evitar consumo excesivo de memoria en Streamlit Cloud.
- `st.cache_data` para imágenes limitado a `max_entries=30` y `ttl=1800`.
- Precarga de imágenes desactivada por defecto y limitada a pocas URL.
- Filtro inicial para elegir qué revisar al cargar el Excel:
  - Pendientes nuevas.
  - Filtrar por color del Excel: rojo, verde, amarillo, mixto o sin color.
  - Todas las tiendas con fotos.
- Detección de colores ya aplicados en el Excel entre columnas A y W.
- Opción para cargar un respaldo JSON y recuperar decisiones en Streamlit Cloud.
- Archivo `.streamlit/config.toml` incluido para despliegue.

### Recomendación para Streamlit Cloud

La memoria local en `data/sesiones` es útil mientras la app esté activa, pero en despliegues gratuitos puede reiniciarse. Para máxima seguridad, descarga el respaldo JSON cada cierto número de validaciones y vuelve a cargarlo si necesitas continuar.
