from flask import Flask, request, render_template, send_file
import os
import tempfile
import subprocess

app = Flask(__name__)  # <- Denna rad måste finnas


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        xml_file = request.files.get('xml')

        if not xml_file:
            return " Ingen XML-fil bifogad.", 400

        with tempfile.TemporaryDirectory() as tmpdir:
            xml_path = os.path.join(tmpdir, 'input.xml')

            # Skapa output-mapp i projektet
            output_dir = os.path.join(os.getcwd(), 'output')
            os.makedirs(output_dir, exist_ok=True)

            xml_file.save(xml_path)

            result = subprocess.run(
                ['python', 'RFQ_GIT.py', xml_path, output_dir],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )

            if result.returncode != 0:
                return f" RFQ_GIT.py misslyckades:\n\n{result.stderr}", 500

            output_path = os.path.join(output_dir, 'komplett_rfqdokument.docx')

            if not os.path.exists(output_path):
                return " Dokumentet kunde inte genereras.", 500

            return send_file(output_path, as_attachment=True)

    return render_template('form.html')

if __name__ == '__main__':
    app.run(debug=True)
