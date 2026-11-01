"""
============================================================================
MICROSOFT EXCEL WEB - PYTHON BACKEND & SERVER
Runs a local web server with SQLite / MySQL database persistence
============================================================================
Usage:
    python server.py
Then open in browser:
    http://localhost:8000
============================================================================
"""

import http.server
import socketserver
import json
import sqlite3
import os
import urllib.parse

PORT = 8000
DB_FILE = 'excel.db'

# Initialize SQLite database
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS workbooks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            data TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

init_db()

class ExcelRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        
        # API: List Workbooks
        if parsed.path == '/api/list':
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute('SELECT name, updated_at FROM workbooks ORDER BY updated_at DESC')
            rows = c.fetchall()
            conn.close()
            
            res = [{'name': r[0], 'updated_at': r[1]} for r in rows]
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'status': 'success', 'workbooks': res}).encode('utf-8'))
            return

        # API: Load Workbook
        if parsed.path == '/api/load':
            params = urllib.parse.parse_qs(parsed.query)
            name = params.get('name', [''])[0]
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute('SELECT data FROM workbooks WHERE name = ?', (name,))
            row = c.fetchone()
            conn.close()

            if row:
                cells = json.loads(row[0])
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'success', 'name': name, 'data': cells}).encode('utf-8'))
            else:
                self.send_response(404)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'error', 'message': 'Topilmadi'}).encode('utf-8'))
            return

        # Default static file handler
        return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)

        # API: Save Workbook
        if parsed.path == '/api/save' or 'action=save' in parsed.query:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length).decode('utf-8')
            
            try:
                payload = json.loads(body)
                name = payload.get('name', 'Книга1')
                data_json = json.dumps(payload.get('data', {}))

                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute('''
                    INSERT INTO workbooks (name, data, updated_at) 
                    VALUES (?, ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(name) DO UPDATE SET 
                        data = excluded.data, 
                        updated_at = CURRENT_TIMESTAMP
                ''', (name, data_json))
                conn.commit()
                conn.close()

                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'status': 'success',
                    'message': f"'{name}' muvaffaqiyatli saqlandi!"
                }).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'error', 'message': str(e)}).encode('utf-8'))
            return

        self.send_response(404)
        self.end_headers()

if __name__ == '__main__':
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    print(f"Microsoft Excel Web Server running at http://localhost:{PORT}")
    print("Press Ctrl+C to stop.")
    with socketserver.TCPServer(("", PORT), ExcelRequestHandler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped.")
