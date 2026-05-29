from http.server import HTTPServer, BaseHTTPRequestHandler
import os
import re

class GASHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/':
            filename = 'Index.html'
        else:
            filename = self.path.lstrip('/')

        if os.path.exists(filename):
            with open(filename, 'r') as f:
                content = f.read()

            # Simple include replacement
            def replace_include(match):
                inc_file = match.group(1) + '.html'
                if os.path.exists(inc_file):
                    with open(inc_file, 'r') as inc_f:
                        return inc_f.read()
                return f"<!-- {inc_file} not found -->"

            content = re.sub(r"<\?!= include\('(.+?)'\); \?>", replace_include, content)

            # Remove other GAS tags
            content = re.sub(r"<\?.+?\?>", "", content)

            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(content.encode())
        else:
            self.send_error(404)

httpd = HTTPServer(('localhost', 8000), GASHandler)
print("Serving on port 8000...")
httpd.serve_forever()
