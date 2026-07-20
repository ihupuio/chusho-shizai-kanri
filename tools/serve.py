import http.server
import os
import socketserver
import sys

os.chdir(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'docs'))
PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8461


class Handler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Cache-Control', 'no-store')
        super().end_headers()

    def log_message(self, *args):
        pass


socketserver.TCPServer.allow_reuse_address = True
with socketserver.ThreadingTCPServer(('0.0.0.0', PORT), Handler) as httpd:
    print(f'serving on http://0.0.0.0:{PORT}', flush=True)
    httpd.serve_forever()
