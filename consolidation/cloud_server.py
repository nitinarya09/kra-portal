"""
cloud_server.py -- Cloud-ready Web Compiler Server for Render
============================================================================
Listens for GET /compile?fy=2026-27&quarter=Q1.
Downloads data from Google Sheets, compiles the report in-memory, and
returns the generated .docx file directly to the browser as a download.
"""

import sys
import os
import json
import io
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from docx import Document

# Add path so imports work on Render
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fetch_data import get_credentials, fetch_all_data
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
            print(f"Found template file at: {c}")
            return os.path.abspath(c)
    print(f"WARNING: Template file not found in candidates, trying fallback path...")
    return candidates[0]


class CloudCompilerHandler(BaseHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_HEAD(self):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()

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
                print("Step 1: Fetching data...")
                data = fetch_all_data(fy, quarter)
                print(f"Step 1 Complete! Data worksheets: {len(data)}")
                
                print("Step 2: Finding template file...")
                template_path = find_template_file()
                print(f"Step 2 Complete! Opening doc: {template_path}")
                doc = Document(template_path)
                
                print("Step 3: Populating tables...")
                populate_all_tables(doc, data)
                print("Step 3 Complete!")
                
                print("Step 4: Generating appreciation note...")
                note_paragraphs = generate_appreciation_note(data, fy, quarter)
                if note_paragraphs:
                    insert_appreciation_note(doc, note_paragraphs)
                print("Step 4 Complete!")
                
                print("Step 5: Saving to memory buffer...")
                file_stream = io.BytesIO()
                doc.save(file_stream)
                file_stream.seek(0)
                file_bytes = file_stream.read()
                print(f"Step 5 Complete! Buffer size: {len(file_bytes)} bytes")
                
                # 6. Send file back to browser as attachment download
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                self.send_header('Content-Disposition', f'attachment; filename="Consolidated_KRA_{quarter}_{fy}.docx"')
                self.send_header('Content-Length', str(len(file_bytes)))
                self.end_headers()
                self.wfile.write(file_bytes)
                print("Compilation successful! File streamed to client.")
                
            except Exception as e:
                import traceback
                err_msg = f"Error during cloud compilation: {e}\n{traceback.format_exc()}"
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
            self.wfile.write(json.dumps({"status": "online", "message": "KRA Cloud Compiler Server is ready"}).encode())
            return

        self.send_response(404)
        self.end_headers()


def run_server():
    server = ThreadingHTTPServer(('0.0.0.0', PORT), CloudCompilerHandler)
    print(f"KRA Threading Cloud Compiler Server running on port {PORT}...")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.server_close()


if __name__ == "__main__":
    run_server()
