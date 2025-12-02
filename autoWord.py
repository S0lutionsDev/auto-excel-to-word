#  pip install docxtpl -q
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate

doc = DocxTemplate("plantilla.docx")
nombre = "Santiago"
telefono = "(123) 345-653"
correo = "hugo@gmail.com"
fecha = datetime.today().strftime("%d/%m/%Y")

const = {
    'nombre':nombre,
    'correo': correo,
    'telefono': telefono,
    'fecha': fecha
}


df = pd.read_excel('Students.xlsx')

for indice, fila in df.iterrows():
    #print(indice) # 0 al 9 indice de los alumnos
    # print(fila)
    contenido = {
        'nombre_alumno': fila["Nombre del Alumno"],
        'nota_mat': fila['Mat'],
        'nota_fis': fila['Fis'],
        'nota_qui': fila['Qui']
    }
    contenido.update(const)
    doc.render(contenido)
    doc.save(f"notas_de_{fila['Nombre del Alumno']}.docx")
    print(contenido)

# doc.render(const)
# doc.save(f"prueba.docx") # Con esto se lleno las casillas de variables con los datos que proporcionamos antes.