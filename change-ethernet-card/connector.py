from pyvim.connect import SmartConnect, Disconnect
import ssl
import atexit


def vcenter_connector(**kwargs):
    """ Acceptable Values are:
     server_name = serve name with domain suffix. example: vcsa-HQ.alibaba.local
     user = the username with vCenter SSO domain. example: hamed@vsphere.local
     password = SSO password

       """
    for key, value in kwargs.items():
        server = kwargs["server"]
        user = kwargs["user"]
        password = kwargs["password"]
    context = None
    if hasattr(ssl, '_create_unverified_context'):
        context = ssl._create_unverified_context()
    connection = SmartConnect(host=f'{server}', user=user, pwd=password, sslContext=context)
    if not connection:
        print(
            "Could not connect to the specified host using specified ",
            "username and password"
        )
    # atexit.register(Disconnect, connection)
    return connection
