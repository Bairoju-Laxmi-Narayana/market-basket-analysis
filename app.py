from flask import Flask, request
from flask_cors import CORS
import subprocess

app = Flask(__name__)

@app.route('/execute_program', methods=['GET'])
def execute_program():
    path = request.args.get('path')
    try:
        result = subprocess.run(['python', path], capture_output=True, text=True)
        return result.stdout
    except Exception as e:
        return str(e), 500

if __name__ == '__main__':
    app.run(debug=True)
