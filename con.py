from flask import Flask, request, send_file, render_template_string
import os
from spire.doc import Document
from spire.doc.common import FileFormat
import tempfile

app = Flask(__name__)

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>File Converter</title>
<script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="flex items-center justify-center min-h-screen bg-gray-50">
  <div class="bg-white p-6 rounded shadow-lg max-w-md w-full space-y-4">
    <h2 class="text-xl font-semibold text-center">Upload File to Convert</h2>
    <form method="post" action="/convert" enctype="multipart/form-data" class="flex flex-col space-y-3">
      <input type="file" name="file" class="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-pink-50 file:text-pink-700 hover:file:bg-pink-100" accept=".doc,.docx,.pdf" required />
      <button type="submit" class="bg-blue-600 text-white py-2 px-4 rounded hover:bg-blue-700 transition">Convert & Download</button>
    </form>
  </div>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files.get('file')
    if not file:
        return "No file uploaded.", 400

    input_path = tempfile.mktemp(suffix=os.path.splitext(file.filename)[1])
    file.save(input_path)

    # Decide output format and path
    if input_path.lower().endswith('.pdf'):
        output_path = tempfile.mktemp(suffix='.docx')
        # PDF to Word - placeholder: in practice use Aspose or 3rd party
        # Here just copy input to output to demo
        os.system(f'cp "{input_path}" "{output_path}"')
    else:
        output_path = tempfile.mktemp(suffix='.pdf')
        # Word to PDF using Spire.Doc
        doc = Document()
        doc.LoadFromFile(input_path)
        doc.SaveToFile(output_path, FileFormat.PDF)

    return send_file(output_path, as_attachment=True, download_name=os.path.basename(output_path))


if __name__ == "__main__":
    app.run(debug=False)
