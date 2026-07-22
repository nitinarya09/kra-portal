"""
server.py -- Lightweight Background Compiler Trigger Server
============================================================================
Runs a local server on port 5000 that listens for requests from the browser
portal to compile the KRA report dynamically.

Endpoints:
    GET /compile?fy=2026-27&quarter=Q1
"""

import sys
import os
import json
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

# Import consolidation script
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from consolidate import run_compilation

PORT = 5005


class CompilerHandler(BaseHTTPRequestHandler):
    def end_headers(self):
        # Enable CORS for browser access
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed_url = urlparse(self.path)
        
        # 1. Endpoint: /compile
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

            print(f"Triggering compilation via API for {quarter} FY {fy}...")
            
            # Execute compilation
            success, msg = run_compilation(fy, quarter)

            if success:
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"status": "success", "message": msg}).encode())
            else:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"status": "failed", "message": msg}).encode())
            return

        # 2. Endpoint: /status (health check)
        elif parsed_url.path == "/status" or parsed_url.path == "/":
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "online", "message": "KRA Compiler Server is ready"}).encode())
            return

        # 404 Not Found
        self.send_response(404)
        self.end_headers()


def run_server():
    server = HTTPServer(('localhost', PORT), CompilerHandler)
    print(f"============================================================")
    print(f"  KRA Background Compiler Server running on: http://localhost:{PORT}")
    print(f"  Close this window to stop the server.")
    print(f"============================================================")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down compiler server...")
        server.server_close()


if __name__ == "__main__":
    run_server()
