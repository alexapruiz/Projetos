from sockets.python3.server import Server

class MyServer(Server):
    def act_on(self, data, addr):
        # Do something with data (in bytes) and return a string.
        return data.decode()

server = MyServer(listening_address=('127.0.0.1', 11113))