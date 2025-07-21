
verificar que la cadena que se va enviar sea igual a esto : &set_field.InvoiceNr=992-11159061&set_field.VECode=10197&set_field.InvDate=2025-07-18&set_field.TransDate=2025-07-18&set_field.PayDeal=CC&set_field.OKPersons=MAC&set_field.Objects=202506&set_field.OKFlag=0&set_field.PrelBook=0&set_field.PayVal=80&set_row_field.0.STP=1&set_row_field.0.AccNumber=511202&set_row_field.0.Objects=10561,FF,CAER,AI-8564-2025,HAWB-21420-2025&set_row_field.0.Sum=4.27&set_row_field.0.VATCode=E&set_row_field.0.Item=C101&set_row_field.0.qty=1&set_row_field.0.PeriodCode=&set_row_field.1.STP=1&set_row_field.1.AccNumber=511202&set_row_field.1.Objects=10561,FF,CAER,AI-8564-2025,HAWB-21419-2025&set_row_field.1.Sum=75.73&set_row_field.1.VATCode=E&set_row_field.1.Item=C101&set_row_field.1.qty=1&set_row_field.1.PeriodCode=




# VIVc_Guardar
Rutina para crear ordene de compras 
# VIVc_Guardar
Este script de Python está diseñado para leer datos de facturas desde un archivo de Excel, procesarlos y luego enviarlos a una API a través de solicitudes HTTP POST. Utiliza las bibliotecas `asyncio` y `aiohttp` para operaciones asíncronas, `pandas` para la manipulación de datos y `openpyxl` para leer archivos de Excel.

## Características

- Lee datos de facturas desde una hoja específica en un archivo de Excel.
- Procesa y agrupa los datos de las facturas por encabezados y detalles.
- Construye y envía solicitudes POST a una API con los datos de la factura.
- Maneja la autenticación básica para las solicitudes a la API.
- Realiza operaciones de red de forma asíncrona para mejorar el rendimiento.

## Requisitos

- Python 3.7+
- Las bibliotecas de Python se listan en el archivo `requirements.txt`.

## Instalación

1.  Clona este repositorio:
    ```bash
    git clone https://github.com/alexisvivas/VIVc_Guardar.git
    cd VIVc_Guardar
    ```

2.  Crea un entorno virtual (recomendado):
    ```bash
    python -m venv venv
    source venv/bin/activate  # En Windows usa `venv\Scripts\activate`
    ```

3.  Instala las dependencias:
    ```bash
    pip install -r requirements.txt
    ```

## Uso

1.  Asegúrate de que el archivo de Excel con los datos de la factura esté en la ruta especificada en la variable `file_path` dentro del script `VIVc_Guardar.py`.
2.  Actualiza las credenciales de `username` y `password`, y la `base_url` en el script si es necesario.
3.  Ejecuta el script:
    ```bash
    python VIVc_Guardar.py
    ```

El script procesará el archivo de Excel, imprimirá información sobre el primer y último número de factura, y luego enviará los datos a la API.

## Estructura del Código

-   **`VIVc_Detalle`**: Una clase de datos para representar las líneas de detalle de una factura.
-   **`Encabezado`**: Una clase de datos para representar el encabezado de una factura, que contiene una lista de objetos `VIVc_Detalle`.
-   **`HttpClass`**: Una clase de utilidad con un método estático `consulta` para realizar solicitudes GET a la API.
-   **`vivc_guardar()`**: La función principal que orquesta la lectura del archivo de Excel, el procesamiento de datos y el envío de los mismos a la API.
-   **`buscar_y_registrar()`**: Una función auxiliar para verificar si un registro de factura ya existe antes de crearlo.
-   **`main()`**: El punto de entrada asíncrono para ejecutar el script.
