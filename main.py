from docxtpl import DocxTemplate
from datetime import datetime
import pandas as pd
import os
import smtplib
from email.message import EmailMessage

def create_dictionary(subject, student):
    '''Creates a dictionary based on a list of students and its respective teacher information
    
    args:
        - subject: information for the subject and its respective teacher.
        - student: information for the current student.
    '''

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
    '''Creates the docx file with the information obtained from the Excel file.
    
    args:
        - constantes: Dictionary with the information to be passed to the docx file.
        - subject: Information to rename the new file (subject name).
        - student: Information to rename the new file (student name).

    '''
    doc.render(constantes)
    doc.save(f"Calificaciones/{subject["Materia"]}/Calificaiones_{student["Nombre"]}.docx")

def create_folders(project_file_path):
    '''Creates the folder for the current subject if not exists
    
    args:
        - project_file_path: Current folder path to be created.
    
    '''
    if not os.path.exists(project_file_path):
        os.mkdir(project_file_path)

def send_email(text, sender, receiver):

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(sender, "khhy drjh swgz hkwn")
    s.sendmail(sender, receiver, "This is a test email")

if __name__ == "__main__":
    email_text_path = os.path.join(os.getcwd(), "email_text.txt")
    subject = "Calificaciones | Python automation"
    sender = "djjuanluis14@gmail.com"
    receiver = "juanluis.reyes014@gmail.com"
    project_file_path = os.getcwd()
    project_file_path = os.path.join(project_file_path, "Calificaciones")
    create_folders(project_file_path)
    doc = DocxTemplate("scores_template.docx")
    maestro_sheet = pd.read_excel(io="students_scores.xlsx", sheet_name="Maestro")

    for index, subject in maestro_sheet.iterrows():
        create_folders(os.path.join(project_file_path, subject["Materia"]))
        
        students_info = pd.read_excel(io="students_scores.xlsx", sheet_name=subject["Materia"])

        for index, student in students_info.iterrows():
            constantes = create_dictionary(subject, student)
            create_file(constantes, subject, student)
            send_email(email_text_path, sender, receiver)

