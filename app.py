from flask import Flask, request, jsonify, render_template, send_file
from flasgger import Swagger
from utils.simple_functions import *
import os
import shutil
import traceback

app = Flask(__name__)


@app.route('/upload', methods=['POST'])
def upload_file():
        
        try:
            data = request.form
            user_id = data["user_id"]
            #if not user_id:
            #   return jsonify({"error": "Missing user_id"}), 400

            if 'file' not in request.files:
                return jsonify({"error": "No file uploaded"}), 400

            file = request.files['file']
            if file.filename == '':
                return jsonify({"error": "No selected file"}), 400

            user_path = get_user_path(user_id, "input")
            os.makedirs(user_path, exist_ok=True)
            for filename in os.listdir(user_path):
                file_path = os.path.join(user_path, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.chmod(file_path, stat.S_IWRITE) 
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path, onerror=lambda func, path, excinfo: os.chmod(path, stat.S_IWRITE) or func(path))
                except Exception as e:
                    print(f"Warning: impossibile cancellare {file_path}. Motivo: {e}")
                    traceback.print_exc()


        except Exception as e:
                traceback.print_exc()
                return jsonify({"error": str(e)}), 500





