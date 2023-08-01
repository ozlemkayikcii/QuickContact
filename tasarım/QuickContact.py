import os
from flask import Flask, render_template, request, send_file
import pandas as pd
import openpyxl

app = Flask(__name__)

def add_prefix_to_column(input_file, column_name):
    try:
        # Excel dosyasını oku
        df = pd.read_excel(input_file, engine='openpyxl', header=1)

        # Belirli sütunun tüm hücrelerine "Gezi " ön eki ekle
        df[column_name] = "Gezi " + df[column_name].astype(str)

        # Değişiklikleri kaydet
        df.to_excel(input_file, index=False, engine='openpyxl')

        print(f"{column_name} sütunundaki isimlere başarıyla 'Gezi ' ön eki eklendi.")
    except Exception as e:
        print("Hata oluştu:", str(e))

def create_vcf_file(name, phone):
    vcf_data = f"BEGIN:VCARD\nVERSION:2.1\nN:{name}\nFN:{name}\nTEL;CELL:{phone}\nEND:VCARD\n"
    return vcf_data

def excel_to_vcf(input_file, name_column, phone_column, output_file):
    try:
        # Excel dosyasını oku
        df = pd.read_excel(input_file, engine='openpyxl')
        vcf_contents = ""

        # Belirli sütunlardaki verileri döngüyle gez ve VCF dosyasına dönüştür
        for index, row in df.iterrows():

            name = row[name_column]
            phone = row[phone_column]
            vcf_contents += create_vcf_file(name, phone)

        # VCF dosyasını oluştur
        with open(output_file, 'w', encoding='utf-8') as vcf_file:
            vcf_file.write(vcf_contents)

        print(f"{output_file} dosyası başarıyla oluşturuldu.")
    except Exception as e:
        print("Hata oluştu:", str(e))

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    try:
        # Get the uploaded file from the request
        file = request.files['file']

        # Check if the file is present and has the correct extension
        if file and file.filename.lower().endswith('.xlsx'):
            # Save the uploaded file temporarily
            uploaded_file_path = "uploaded_file.xlsx"
            file.save(uploaded_file_path)

            # Process the uploaded file and convert it to VCF
            name_column = "AD SOYAD"
            phone_column = "TELEFON"
            vcf_file_path = "converted_rehber.vcf"
            excel_to_vcf(uploaded_file_path, name_column, phone_column, vcf_file_path)

            # Remove the temporary uploaded file
            os.remove(uploaded_file_path)

            # Provide the VCF file for download  , attachment_filename='rehber.vcf' *****   http://localhost:5000/
            return send_file(vcf_file_path, as_attachment=True)

        else:
            return "Invalid file format. Please upload an xlsx file."

    except Exception as e:
        return f"An error occurred: {str(e)}"


if __name__ == '__main__':
    app.run(debug=True)

