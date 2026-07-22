"""
server.py -- Multi-Threaded KRA Cloud Compiler Server
============================================================================
Runs a multi-threaded streaming HTTP server for Render and local environments.

Endpoints:
    GET /compile?fy=2026-27&quarter=Q1
    GET /status
"""

import sys
import os
import json
import io
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from docx import Document

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fetch_data import fetch_all_data
from populate_template import populate_all_tables
from appreciation_note import generate_appreciation_note, insert_appreciation_note

PORT = int(os.environ.get("PORT", 5005))


def find_template_file():
    candidates = [
        os.path.join(os.path.dirname(__file__), "Sectionwise Blank KRA Report for June 2026 end Quarter.docx"),
        os.path.join(os.path.dirname(os.path.dirname(__file__)), "Sectionwise Blank KRA Report for June 2026 end Quarter.docx"),
        "Sectionwise Blank KRA Report for June 2026 end Quarter.docx",
        os.path.join("consolidation", "Sectionwise Blank KRA Report for June 2026 end Quarter.docx")
    ]
    for c in candidates:
        if os.path.exists(c):
            return os.path.abspath(c)
    return candidates[0]


class CompilerHandler(BaseHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS, HEAD')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        parsed_url = urlparse(self.path)
        
        if parsed_url.path == "/compile":
            params = parse_qs(parsed_url.query)
            fy = params.get("fy", [""])[0]
            quarter = params.get("quarter", [""])[0]

            if not fy or not quarter:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"error": "Missing 'fy' or 'quarter' parameter"}).encode())
                return

            print(f"Cloud compile triggered for {quarter} FY {fy}...")
            
            try:
                data = {}
                template_path = find_template_file()
                doc = Document(template_path)
                populate_all_tables(doc, data)
                
                note_paragraphs = generate_appreciation_note(data, fy, quarter)
                if note_paragraphs:
                    insert_appreciation_note(doc, note_paragraphs)
                
                file_stream = io.BytesIO()
                doc.save(file_stream)
                file_stream.seek(0)
                file_bytes = file_stream.read()
                
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                self.send_header('Content-Disposition', f'attachment; filename="Consolidated_KRA_{quarter}_{fy}.docx"')
                self.send_header('Content-Length', str(len(file_bytes)))
                self.end_headers()
                self.wfile.write(file_bytes)
                print("Compilation successful! File streamed to client.")
                
            except Exception as e:
                import traceback
                err_msg = f"Error during compilation: {e}\n{traceback.format_exc()}"
                print(err_msg)
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"status": "failed", "message": str(e)}).encode())
            return

        elif parsed_url.path == "/status" or parsed_url.path == "/" or parsed_url.path == "/health":
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "online", "message": "KRA Compiler Server is ready"}).encode())
            return

        self.send_response(404)
        self.end_headers()


def run_server():
    server = ThreadingHTTPServer(('0.0.0.0', PORT), CompilerHandler)
    print(f"KRA Threading Cloud Compiler Server running on port {PORT}...")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.server_close()


if __name__ == "__main__":
    run_server()
