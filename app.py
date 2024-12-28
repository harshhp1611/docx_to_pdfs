import os
import time
import pythoncom
import comtypes.client
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

# Path to the folder where you will store the uploaded and converted files
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'converted_pdfs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

def convert_doc_to_pdf(doc_path, pdf_path):
    # Initialize COM
    pythoncom.CoInitialize()  # Initialize COM before using any COM objects
    
    try:
        # Debug: Print the file paths
        print(f"Attempting to open document at: {doc_path}")
        print(f"Attempting to save PDF at: {pdf_path}")
        
        # Initialize COM object for Word (requires Microsoft Word to be installed)
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False  # Keep Word in the background

        # Small delay to ensure Word is fully initialized
        time.sleep(1)
        
        # Open the document
        doc = word.Documents.Open(doc_path)

        # Save as PDF
        doc.SaveAs(pdf_path, FileFormat=17)  # FileFormat=17 for PDF
        doc.Close()
        word.Quit()
    except Exception as e:
        print(f"Error: {e}")
    finally:
        # Uninitialize COM after using it
        pythoncom.CoUninitialize()

@app.route('/', methods=['GET', 'POST'])
def index():
    pdf_file = None
    error_message = None

    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            error_message = 'No file part'
            return render_template('index.html', error_message=error_message)

        file = request.files['file']
        
        if file.filename == '':
            error_message = 'No selected file'
            return render_template('index.html', error_message=error_message)
        
        if file and file.filename.endswith('.docx'):
            filename = file.filename
            doc_path = os.path.join(os.getcwd(), app.config['UPLOAD_FOLDER'], filename)  # Use absolute path
            file.save(doc_path)
            
            # Convert .docx to .pdf
            pdf_filename = filename.replace('.docx', '.pdf')
            pdf_path = os.path.join(os.getcwd(), app.config['OUTPUT_FOLDER'], pdf_filename)  # Use absolute path
            convert_doc_to_pdf(doc_path, pdf_path)

            # Provide the PDF file for download
            pdf_file = pdf_path  # Full path to the PDF file
            
            return render_template('index.html', pdf_file=pdf_file)

    return render_template('index.html', pdf_file=pdf_file, error_message=error_message)

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(os.getcwd(), app.config['OUTPUT_FOLDER'], filename)
    
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "File not found!", 404

if __name__ == '__main__':
    app.run(debug=True)
