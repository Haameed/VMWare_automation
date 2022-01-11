from pyvim.connect import SmartConnect, Disconnect
import ssl
import atexit


def vcenter_connector(**kwargs):
    """ Acceptable Values are:
     server_name = serve name without domain suffix. example: vcsa-HQ
     domain = the domain name. example: 'somecompany.com'
     user = the username with vCenter SSO domain. example: hamed@vsphere.local
     password = SSO password

       """
    for key, value in kwargs.items():
        server_name = kwargs["server"]
        domain = kwargs["domain"]
        user = kwargs["user"]
        password = kwargs["password"]
    context = None
    if hasattr(ssl, '_create_unverified_context'):
        context = ssl._create_unverified_context()
    connection = SmartConnect(
        host=f'{server_name}.{domain}',
        user=user,
        pwd=password,
        sslContext=context)
    if not connection:
        print(
            "Could not connect to the specified host using specified ",
            "username and password"
        )
    atexit.register(Disconnect, connection)
    return connection.content
