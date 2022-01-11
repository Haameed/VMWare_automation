import connector
import secret
from pyVmomi import vim
import threading

server = 'vcsa-gorgan.alibaba.local'
username = secret.user
password = secret.pwd
connection = connector.vcenter_connector(server=server, user=username, password=password)
content = connection.content
view = content.viewManager.CreateContainerView(content.rootFolder, [vim.VirtualMachine], True)
change_list = []
threads = []
for vm in view.view:
    device_change = []
    vm_name = vm.name
    power_status = vm.summary.runtime.powerState
    if power_status == "poweredOn":
        for device in vm.config.hardware.device:
            if isinstance(device, vim.vm.device.VirtualEthernetCard):
                if not device.connectable.startConnected:
                    nicspec = vim.vm.device.VirtualDeviceSpec()
                    nicspec.operation = vim.vm.device.VirtualDeviceSpec.Operation.edit
                    nicspec.device = device
                    nicspec.device.connectable.startConnected = True
                    device_change.append(nicspec)
                    config_spec = vim.vm.ConfigSpec(deviceChange=device_change)
                    t = threading.Thread(target=vm.ReconfigVM_Task, args=(config_spec,))
                    t.start()
                    threads.append(t)

