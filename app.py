import os
import shutil
import tempfile
from flask import Flask, request, send_file, jsonify, send_from_directory
from generator import parse_cpi_iflow_zip, build_visteon_ts_docx

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FRONTEND_DIR = os.path.join(BASE_DIR, 'frontend')
app = Flask(__name__, static_folder=FRONTEND_DIR, template_folder=FRONTEND_DIR)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # Limit upload size to 32MB


@app.errorhandler(413)
def handle_413(_error):
    return jsonify({'error': 'File too large. Please upload a smaller .zip file.'}), 413


@app.route('/')
def index():
    return send_from_directory(FRONTEND_DIR, 'index.html')


@app.route('/health')
def health():
    return jsonify({'status': 'ok'})


@app.route('/generate', methods=['POST'])
def generate():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in request'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not file.filename.lower().endswith('.zip'):
        return jsonify({'error': 'Please upload a valid SAP CPI iFlow .zip export.'}), 400

    # Create temporary isolated working directory for processing
    temp_dir = tempfile.mkdtemp()
    try:
        zip_path = os.path.join(temp_dir, file.filename)
        file.save(zip_path)

        # Parse iFlow ZIP data
        parsed_data = parse_cpi_iflow_zip(zip_path)

        # Generate DOCX document
        generated_docx_path = build_visteon_ts_docx(parsed_data, output_dir=temp_dir)

        download_name = f"UID_U057_{parsed_data['iflow_name']}.docx"

        # Send generated file back to user
        response = send_file(
            generated_docx_path,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        response.call_on_close(lambda: shutil.rmtree(temp_dir, ignore_errors=True))
        return response

    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({'error': f"Failed to generate Technical Specification: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)), debug=True)