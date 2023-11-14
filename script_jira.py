import requests
from requests.auth import HTTPBasicAuth
import json
import openpyxl
from datetime import datetime

token = 'ATATT3xFfGF0aFchKiPhIxSmbT2cjonIYV8FY2Euj-S8rQIiJMsHs_GoA6AC8w8FSEYljS18NQh5Z8KXjNwMrIotcTC05J6DRqbNhq08R8jT7I0Lj8qKkQv2TeTfmxkkYM0zhi2NyaVwuzSw1DeKlxdfZlPjo8cSjQMs7ffso83RSLi1N6ja59s=34B360DE'

url = "https://proyectosthese.atlassian.net/rest/api/3/search"

auth = HTTPBasicAuth("cecilia.gomez@these.com.uy", token)

headers_request = {
    "Accept": "application/json"
}

query = {
    'jql': 'project = GDP AND issuetype = "Actualización Semáforo" AND updated >= startOfDay(-2M) AND updated <= endOfDay()'
}

try:
    response = requests.get(url, headers=headers_request, params=query, auth=auth)
    response.raise_for_status() 
    data = response.json()
except json.JSONDecodeError:
    print("Error al decodificar JSON. Respuesta del servidor:")
    print(response.text)
    data = None

if data:
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    headers_sheet = ["Proyecto", "Creado", "Alcance", "Ambientes de trabajo", "Calidad", "Clima interno", "Compras y contrataciones", "Cronograma", "Interacción con el cliente"]  

    sheet.append(headers_sheet)

    for issue in data.get('issues', []):
        proyecto = issue['fields']['customfield_10043']['value']

        creado_datetime = datetime.strptime(issue['fields']['created'], "%Y-%m-%dT%H:%M:%S.%f%z")
        creado = creado_datetime.strftime("%d-%m-%Y")  # Formato deseado: 'YYYY-MM-DD'


        alcance = issue['fields']['customfield_10057']
        ambiente = issue['fields']['customfield_10064']
        calidad = issue['fields']['customfield_10059']
        clima = issue['fields']['customfield_10089']
        compras = issue['fields']['customfield_10063']
        cronograma = issue['fields']['customfield_10058']
        interaccion = issue['fields']['customfield_10056']

        if alcance and isinstance(alcance, list):
            alcance_values = [value.get('value', '').strip() for value in alcance]
            alcance = ', '.join(alcance_values)
        else:
            alcance = ''

        if ambiente and isinstance(ambiente, list):
            ambiente_values = [value.get('value', '').strip() for value in ambiente]
            ambiente = ', '.join(ambiente_values)
        else:
            ambiente = ''

        if calidad and isinstance(calidad, list):
            calidad_values = [value.get('value', '').strip() for value in calidad]
            calidad = ', '.join(calidad_values)
        else:
            calidad = ''

        if clima and isinstance(clima, list):
            clima_values = [value.get('value', '').strip() for value in clima]
            clima = ', '.join(clima_values)
        else:
            clima = ''

        if compras and isinstance(compras, list):
            compras_values = [value.get('value', '').strip() for value in compras]
            compras = ', '.join(compras_values)
        else:
            compras = ''

        if cronograma and isinstance(cronograma, list):
            cronograma_values = [value.get('value', '').strip() for value in cronograma]
            cronograma = ', '.join(cronograma_values)
        else:
            cronograma = ''

        if interaccion and isinstance(interaccion, list):
            interaccion_values = [value.get('value', '').strip() for value in interaccion]
            interaccion = ', '.join(interaccion_values)
        else:
            interaccion = ''
        
        row_data = [proyecto, creado, alcance, ambiente, calidad, clima, compras, cronograma, interaccion]
        sheet.append(row_data)

    workbook.save('resultado_script.xlsx')
    print("Datos guardados correctamente en salida.xlsx")
else:
    print("No se pudieron procesar los datos JSON.")
