from flask import Flask, request, render_template, send_file
from datetime import datetime, timedelta
from telegram_parser import run_parser

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        channel = request.form.get('channel')
        start = request.form.get('start') or (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        end = request.form.get('end') or datetime.now().strftime('%Y-%m-%d')
        start_date = datetime.strptime(start, '%Y-%m-%d')
        end_date = datetime.strptime(end, '%Y-%m-%d')
        filename = run_parser(channel, start_date, end_date)
        return send_file(filename, as_attachment=True)
    return render_template('index.html')