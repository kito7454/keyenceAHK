import Pyro5.api
import Pyro5.server
from ahk_routine import ahk_routine, check_for_variable  # import your existing code

# Decorate the class to expose it as a Pyro5 object
@Pyro5.api.expose
class RemoteAhkRoutine(ahk_routine):
    """
    A Pyro5-exposed wrapper around your ahk_routine class.
    Inherits everything from ahk_routine, just adds the @expose decorator.
    We override __init__ to accept arguments from the remote caller.
    """
    def __init__(self, file_name=r"C:\Users\USER1\PycharmProjects\keyenceAHK\VK_automation.xlsx", variableDictionary={}):
        super().__init__(file_name=file_name, variableDictionary=variableDictionary)

    def ping(self):
        print("pinged")
        return "pinged"


def start_server():
    # Create the Pyro5 daemon (the server that listens for remote calls)
    # host="0.0.0.0" means it accepts connections from any IP on your network
    # Change the port if needed
    daemon = Pyro5.server.Daemon(host="0.0.0.0", port=9090)

    # Register your class with the daemon, giving it a simple name
    uri = daemon.register(RemoteAhkRoutine, objectId="ahk.routine")

    print(f"Server is running. URI: {uri}")
    print("Waiting for remote calls...")

    # This blocks and keeps the server alive, listening for connections
    daemon.requestLoop()


if __name__ == "__main__":
    start_server()