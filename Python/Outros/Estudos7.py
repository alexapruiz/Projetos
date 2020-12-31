import SimpleHTTPServer
import SocketServer

PORT = 8201

Handler = SimpleHTTPServer.SimpleHTTPRequestHandler
httpd = SocketServer.TCPServer(("", PORT), Handler)

print("servidor web na porta ", PORT)
httpd.serve_forever()