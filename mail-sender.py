import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from collections import Counter

# Dane serwera SMTP
smtp_server = "smtp.server.com"
smtp_port = 587
smtp_user = "mail@server.com"
smtp_password = "password"

# Ścieżki do folderów i plików
input_folder = "/Users/filipchmielecki/Downloads/automation/students"  # Folder z podfolderami i plikami PDF
excel_file = "/Users/filipchmielecki/Downloads/automation/mails.xlsx"  # Plik Excel z adresami email
log_file = "/Users/filipchmielecki/Downloads/automation/log.txt"  # Plik logów

# Wczytywanie adresów email z pliku Excel
try:
    df = pd.read_excel(excel_file)
    if 'Imie i Nazwisko' not in df.columns or 'Email' not in df.columns:
        raise ValueError("Plik Excel musi zawierać kolumny 'Imie i Nazwisko' oraz 'Email'.")
except Exception as e:
    print(f"❌ Błąd przy wczytywaniu pliku Excel: {e}")
    exit(1)

# Sprawdzanie duplikatów w kolumnie 'Imie i Nazwisko'
duplicates = df['Imie i Nazwisko'][df.duplicated('Imie i Nazwisko', keep=False)]
if not duplicates.empty:
    with open(log_file, 'w') as log:
        log.write("❌ Znaleziono duplikaty w pliku Excel:\n")
        for name in duplicates.unique():
            log.write(f"- {name}\n")
    print("❌ Znaleziono duplikaty w pliku Excel. Sprawdź log.txt.")
    exit(1)

# Tworzenie pliku log
with open(log_file, 'w') as log:
    log.write("📋 Log wysyłania emaili:\n")

# Funkcja do wysyłania maila
def send_email(to_email, attachment_path, student_name):
    try:
        subject = f"Raport po I semestrze - {student_name}"
        body = f"""
        <html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; background-color: #eaf4fb; padding: 20px;">
    <div style="max-width: 600px; margin: 0 auto; background-color: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);">
        <h2 style="color: #007bff; text-align: center;">Dziękujemy za wspólny I semestr nauki języka!</h2>
        <p style="color: #333;">Szanowni Państwo,</p>
        <p style="color: #333;">Serdecznie dziękujemy za zaufanie oraz za wspólnie spędzony I semestr zajęć. Cieszymy się z postępów, jakie poczynił uczeń <strong>{student_name}</strong> i mamy nadzieję na kontynuację nauki w kolejnym semestrze.</p>
        <p style="color: #333;">W załączniku znajdą Państwo <b style="color: #007bff;">raport z I semestru zajęć</b>, który zawiera szczegółowe informacje na temat osiągnięć ucznia.</p>
        <p style="color: #333;">W razie jakichkolwiek pytań lub wątpliwości, prosimy o kontakt. Jesteśmy do Państwa dyspozycji.</p>
        <p style="color: #333;">Życzymy miłego dnia!</p>
        <br>
        <p style="text-align: right; color: #007bff;">Z poważaniem,<br><strong>Zespół Sówka Języki Obce</strong></p>
        <hr style="border: none; border-top: 1px solid #007bff; margin: 20px 0;">
        <p style="font-size: 12px; color: #888; text-align: center;">Ten raport został wygenerowany automatycznie na podstawie danych uzupełnianych przez naszych lektorów w dzienniku EduSky. W razie jakichkolwiek wątpliwości, zapraszamy do kontaktu.</p>
    </div>
</body>
</html>
        """

        message = MIMEMultipart()
        message["From"] = smtp_user
        message["To"] = to_email
        message["Subject"] = subject
        message.attach(MIMEText(body, "HTML"))

        with open(attachment_path, "rb") as attachment_file:
            attachment = MIMEBase("application", "pdf")
            attachment.set_payload(attachment_file.read())
            encoders.encode_base64(attachment)
            attachment.add_header("Content-Disposition", f"attachment; filename=Raport.pdf")
            message.attach(attachment)

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(message)
        server.quit()

        with open(log_file, 'a') as log:
            log.write(f"✅ Wysłano: {student_name} -> {to_email} ({os.path.basename(attachment_path)})\n")
        print(f"✅ Wysłano: {student_name} -> {to_email}")
    except Exception as e:
        with open(log_file, 'a') as log:
            log.write(f"❌ Błąd przy wysyłaniu do {student_name} ({to_email}): {e}\n")
        print(f"❌ Błąd przy wysyłaniu do {student_name}: {e}")

# Przetwarzanie plików PDF z podfolderów
for root, _, files in os.walk(input_folder):
    for file in files:
        if file.endswith(".pdf"):
            full_path = os.path.join(root, file)
            student_name = os.path.splitext(file)[0]  # Nazwa pliku bez rozszerzenia
            student_email = df[df['Imie i Nazwisko'] == student_name]['Email'].values

            if student_email.size > 0:
                send_email(student_email[0], full_path, student_name)
            else:
                with open(log_file, 'a') as log:
                    log.write(f"⚠️ Brak adresu email dla: {student_name} ({file})\n")
                print(f"⚠️ Brak adresu email dla: {student_name}")
