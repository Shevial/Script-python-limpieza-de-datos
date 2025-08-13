# Limpieza y Extracción de Correos Rebotados

Este repositorio contiene dos herramientas en Python diseñadas para **limpiar listas de contactos** eliminando direcciones de correo electrónico que han rebotado.

---

## 1. `extraer_correos.py`

### Propósito
Escanea una carpeta o un archivo `.zip` con mensajes de rebote y **extrae todas las direcciones** que aparecen después del texto:
Final-Recipient: rfc822;
Genera un archivo `.xlsx` o `.csv` con la lista de correos rebotados, listo para usar con `limpiar_contactos.py`.

---

### Uso
```bash
# Desde una carpeta con los mensajes
py extraer_correos.py <ruta_carpeta> -o correos_rebotados.xlsx

# Desde un archivo ZIP con mensajes
py extraer_correos.py <archivo.zip> -o correos_rebotados.xlsx
````

#### Parámetros:
 <ruta_carpeta> o <archivo.zip>: Origen de los mensajes de rebote.

-o o --output: Archivo de salida (por defecto: correos_rebotados.xlsx).
## 2. `limpiar_contactos.py`

###  Propósito
Usa un archivo de contactos (`.xlsx` o `.csv`) y elimina todas las direcciones que se encuentren en la lista de rebotes generada por `extraer_correos.py`.

---

### Uso
```bash
py limpiar_contactos.py contactos.xlsx correos_rebotados.xlsx -o contactos_limpios.xlsx
````
#### Parámetros:
contacts (obligatorio): Archivo de contactos (.xlsx o .csv).

rebotes (opcional): Archivo de rebotes (por defecto: correos_rebotados.xlsx).

-c o --email-column: Nombre exacto de la columna que contiene los emails (si no se detecta automáticamente).

-o o --output: Nombre del archivo de salida (por defecto: <nombre_original>_limpios.xlsx o .csv).

 ## Flujo de trabajo recomendado
1. Extraer rebotes

```bash
py extraer_correos.py carpeta_rebotes -o correos_rebotados.xlsx
```` 
2. Limpiar contactos

```bash
py limpiar_contactos.py contactos.xlsx correos_rebotados.xlsx -o contactos_limpios.xlsx
````
### Personalización
En limpiar_contactos.py, la lista de posibles nombres de columna para los correos está en la función detect_email_column().
Si tu archivo usa un nombre distinto para la columna de email, puedes:

Añadirlo ahí, o

Usar el parámetro --email-column "NombreColumna".

En extraer_correos.py, el patrón de búsqueda de correos está en la función find_emails_in_text().
Si tus mensajes usan otro formato, puedes modificar la expresión regular.
Dependencias
Estos scripts requieren Python 3 y :
```bash
pip install pandas openpyxl
````
