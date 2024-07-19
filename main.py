from docxtpl import DocxTemplate
from datetime import datetime
import pandas as pd
import os

def create_dictionary(subject, student):
    constantes = {
                'name': subject["Nombre"], 
                'email': subject["Correo"], 
                'phoneNumber': subject["Telefono"], 
                'date': datetime.today().strftime("%d/%m/%Y"), 
                'studentName': student["Nombre"], 
                'subject': subject["Materia"],
                'test': student["Examen"],
                'attendance': student["Asistencia"],
                'participation': student["Participacion"]
                }
    return constantes

def create_file(constantes, subject, student):
    doc.render(constantes)
    doc.save(f"Calificaciones/{subject["Materia"]}/Calificaiones_{student["Nombre"]}.docx")

def create_folders(project_file_path):
    os.mkdir(project_file_path)

def create_scores_folder(scores_folder_path):
    if not os.path.exists(scores_folder_path):
        os.mkdir(scores_folder_path)

if __name__ == "__main__":
    project_file_path = os.getcwd()
    project_file_path = os.path.join(project_file_path, "Calificaciones")
    create_scores_folder(project_file_path)
    doc = DocxTemplate("scores_template.docx")
    maestro_sheet = pd.read_excel(io="students_scores.xlsx", sheet_name="Maestro")

    for index, subject in maestro_sheet.iterrows():
        create_folders(os.path.join(project_file_path, subject["Materia"]))
        
        students_info = pd.read_excel(io="students_scores.xlsx", sheet_name=subject["Materia"])

        for index, student in students_info.iterrows():
            constantes = create_dictionary(subject, student)
            create_file(constantes, subject, student)
