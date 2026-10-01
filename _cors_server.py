import http.server
import socketserver

class CORSRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        super().end_headers()

PORT = 8123
with socketserver.TCPServer(("127.0.0.1", PORT), CORSRequestHandler) as httpd:
    httpd.serve_forever()
