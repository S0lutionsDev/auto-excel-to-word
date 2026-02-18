# Auto Excel a Word

Script en Python para generar documentos de Word (y opcionalmente PDF) de forma automática a partir de datos en un archivo Excel y una plantilla `.docx`.

Ideal para automatizar la creación masiva de certificados, reportes, constancias, contratos o cualquier documento personalizado.

---

## Tecnologías

- Python 3
- pandas
- python-docx
- openpyxl
- docx2pdf (opcional)

---

## Cómo funciona

1. El archivo Excel contiene los datos (una fila por registro).
2. La plantilla Word usa variables con el formato:
3. El script:
- Lee el Excel  
- Reemplaza las variables  
- Genera un documento por cada fila  

---

## Instalación

Clonar el repositorio:

```bash
git clone https://github.com/SantiCrudo/auto-excel-to-word.git
cd auto-excel-to-word
```
Instalar dependencias:
```
pip install pandas python-docx openpyxl
```
Opcional (exportar a PDF):
```
pip install docx2pdf
```

#Uso

Editar Students.xlsx con tus datos.

Asegurarse de que los nombres de las columnas coincidan con las variables del .docx.

Ejecutar:
```
python autoWord.py
```
