from flask import Flask, render_template, request, jsonify, Response, send_file
import os
import threading
import queue
import time
import json
import logging
import pandas as pd

# Ensure local imports work
import sys
sys.path.append(os.getcwd())

import cz
import hy

app = Flask(__name__)
app.secret_key = 'fuel_management_secret'

# Configuration
UPLOAD_FOLDERS = {
    'hy': '无人值守化验月报',
    'cz': '无人值守称重月报'
}
RESULT_FILES = {
    'hy': '化验月报汇总.xlsx',
    'cz': '称重月报汇总.xlsx'
}

# Ensure directories exist
for folder in UPLOAD_FOLDERS.values():
    if not os.path.exists(folder):
        os.makedirs(folder)

# Global queue for log streaming
log_queue = queue.Queue()

def log_callback(message):
    """Callback function to push logs to the queue."""
    log_queue.put(message)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    file_type = request.form.get('type') # 'hy' or 'cz'
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
        
    if file and file_type in UPLOAD_FOLDERS:
        filename = file.filename
        if not (filename.endswith('.xls') or filename.endswith('.xlsx')):
             return jsonify({'error': 'Invalid file type. Only Excel files allowed.'}), 400

        save_path = os.path.join(UPLOAD_FOLDERS[file_type], filename)
        file.save(save_path)
        log_callback(f"文件上传成功: {filename} -> {UPLOAD_FOLDERS[file_type]}")
        return jsonify({'message': f'File uploaded to {UPLOAD_FOLDERS[file_type]}'})
    
    return jsonify({'error': 'Invalid request'}), 400

@app.route('/api/run', methods=['POST'])
def run_task():
    task_type = request.json.get('type')
    
    if task_type == 'hy':
        thread = threading.Thread(target=run_hy_task)
        thread.start()
        return jsonify({'message': '化验月报汇总任务已启动'})
    elif task_type == 'cz':
        thread = threading.Thread(target=run_cz_task)
        thread.start()
        return jsonify({'message': '称重月报汇总任务已启动'})
    else:
        return jsonify({'error': 'Unknown task type'}), 400

def run_hy_task():
    log_callback(">>> 开始执行化验月报汇总...")
    try:
        hy.run_analysis(log_callback=log_callback)
        log_callback("<<< 化验汇总任务完成。")
    except Exception as e:
        log_callback(f"!!! 任务出错: {e}")

def run_cz_task():
    log_callback(">>> 开始执行称重月报汇总...")
    try:
        cz.run_weight_processing(log_callback=log_callback)
        log_callback("<<< 称重汇总任务完成。")
    except Exception as e:
        log_callback(f"!!! 任务出错: {e}")

@app.route('/api/logs')
def stream_logs():
    def event_stream():
        while True:
            try:
                # Wait for new log message
                message = log_queue.get(timeout=20) 
                yield f"data: {json.dumps({'message': message})}\n\n"
            except queue.Empty:
                # Send keep-alive
                yield f": keep-alive\n\n"
    
    return Response(event_stream(), mimetype="text/event-stream")

@app.route('/download/<type>')
def download_result(type):
    if type in RESULT_FILES:
        filename = RESULT_FILES[type]
        if os.path.exists(filename):
            return send_file(filename, as_attachment=True)
        else:
            return f"文件 {filename} 尚未生成，请先运行任务。", 404
    return "Invalid file type", 400

@app.route('/api/preview/<type>')
def preview_result(type):
    data = {}
    
    def format_numeric(x):
        try:
            if isinstance(x, (int, float)):
                return "{:,.2f}".format(x)
            return x
        except:
            return x

    try:
        if type == 'hy':
            filename = '化验月报汇总.xlsx'
        elif type == 'cz':
            filename = '称重月报汇总分类.xlsx'
        else:
            return jsonify({'error': 'Invalid type'}), 400

        if not os.path.exists(filename):
            return jsonify({'error': 'Results not generated yet'}), 404
        
        try:
            xls = pd.ExcelFile(filename)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                
                # Apply formatting to all numeric columns
                for col in df.select_dtypes(include=['number']).columns:
                    if col == '序号':
                        df[col] = df[col].apply(lambda x: str(int(x)) if pd.notnull(x) else "")
                    else:
                        df[col] = df[col].apply(format_numeric)
                
                # Convert to HTML table with styling classes
                data[sheet_name] = df.to_html(classes='result-table', index=False, na_rep='', border=0)
        except Exception as e:
            return jsonify({'error': f"Error reading Excel file: {str(e)}"}), 500

        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    # Clean console logged by flask
    log = logging.getLogger('werkzeug')
    log.setLevel(logging.ERROR)
    
    print("Web Server Starting on http://127.0.0.1:5000")
    app.run(debug=True, port=5000)
