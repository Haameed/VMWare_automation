import connector
import secret
import pandas as pd
import smtplib
from pyVmomi import vim
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

from datetime import datetime


file_name = f'vm_list_{datetime.now().strftime("%Y-%m-%d")}.xlsx'
def vm_list():
    vcenters = ["vcsa-1", "vcsa-2", "vcsa-3", "vcsa-4"]
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    dataframe = {}
    for vcenter in vcenters:
        # dataframe[vcenter] = {"DataCenter": [], "Cluster": [], "Host": [], "Name": [], "ID": [],
        #     "Power_status": [], "Is_template": [], "Number of CPUs": [],
        #     "CPU_Cores": [], "Total CPU": [], "CPU Hot_add": [], "Memory GB": [], "Memory Hot_add": [],
        #     "VM_Tools Version": [], "virtual_disks": [], "datastore_name": [], "VM Path": [],
        #     "Network Count": [], "network_name": [], "tools_installType":[], "tools_runningStatus": [],
        #     "tools_status": [], "tools_version": [], "tools_versionStatus": [], "tools_versionStatus2": [],
        #     "vm_guestFamilly": [], "vm_guestHostname": [], "vm_hwVersion": [], "vm_ip": [], "vm_ips": []}
        dataframe[vcenter] = {"DataCenter": [], "Cluster": [], "Host": [], "Name": [], "ID": [],
                              "Power_status": [], "Is_template": [], "Number of CPUs": [],
                              "CPU_Cores": [], "Total CPU": [], "CPU Hot_add": [], "Memory GB": [],
                              "Memory Hot_add": [],
                              "VM_Tools Version": [], "virtual_disks": [], "datastore list": [], "Ethernet list": [],
                              "Total allocated disk": [],
                              "Network Count": [], "networks_name": [], "tools_installType": [],
                              "tools_runningStatus": [],
                              "tools_status": [], "tools_version": [], "tools_versionStatus": [],
                              "tools_versionStatus2": [],
                              "vm_guestFamilly": [], "vm_guestHostname": [], "vm_hwVersion": [], "vm_ip": [],
                              "vm_ips": []}

        connection = connector.vcenter_connector(
            server=vcenter, domain=secret.domain, user=secret.user, password=secret.pwd
        )
        content = connection
        children = content.rootFolder.childEntity
        for child in children:
            dc = child
            dc_name = dc.name
            clusters = dc.hostFolder.childEntity
            for cluster in clusters:
                cluster_name = cluster.name
                hosts = cluster.host
                for host in hosts:
                    host_name = host.name
                    for vm in host.vm:
                        vm_name = vm.name
                        cpu_num = vm.config.hardware.numCPU
                        cpu_cores = vm.config.hardware.numCoresPerSocket
                        total_cpu = (cpu_cores * cpu_num)
                        memory_gb = vm.summary.config.memorySizeMB / 1024
                        vm_tools = vm.config.tools.toolsVersion
                        cpu_hotadd = vm.config.cpuHotRemoveEnabled
                        memory_hotadd = vm.config.memoryHotAddEnabled
                        power_status = vm.summary.runtime.powerState
                        is_template = vm.summary.config.template
                        # vm_path_name = vm.summary.config.vmPathName
                        ethernet_cards = vm.summary.config.numEthernetCards
                        virtual_disks = vm.summary.config.numVirtualDisks
                        vm_id = str(vm).split(":")[1].replace("'", "")
                        datastores_list = ""
                        ethernets_list = ""
                        totalAllocatedDisk= 0
                        for device in vm.config.hardware.device:
                            if isinstance(device, vim.vm.device.VirtualDisk):
                                datastores_list += f"({device.deviceInfo.label}, {device.backing.datastore.name}, {round(device.capacityInKB / 1024 / 1024)}),"
                                totalAllocatedDisk = totalAllocatedDisk + device.capacityInKB
                            elif isinstance(device, vim.vm.device.VirtualVmxnet3):
                                ethernets_list += f"({device.deviceInfo.label}, {device.macAddress}, connected = {device.connectable.connected} , wakeOnLanEnabled = {device.wakeOnLanEnabled}),"
                        totalAllocatedDisk = round(totalAllocatedDisk / 1024 / 1024)
                        datastores_list = datastores_list.rstrip(",")
                        ethernets_list = ethernets_list.rstrip(",")
                        vm_network = vm.summary.vm.network
                        networks_name = []
                        for net in vm_network:
                            networks_name.append(net.name)
                        tools_installType = vm.guest.toolsInstallType
                        tools_runningStatus = vm.guest.toolsRunningStatus
                        tools_status = vm.guest.toolsStatus
                        tools_version = vm.guest.toolsVersion
                        tools_versionStatus = vm.guest.toolsVersionStatus
                        tools_versionStatus2 = vm.guest.toolsVersionStatus2
                        vm_guestFamilly = vm.guest.guestFamily
                        vm_guestHostname = vm.guest.hostName
                        vm_hwVersion = vm.guest.hwVersion
                        vm_ip = vm.guest.ipAddress
                        vm_ips = []
                        if tools_status == 'toolsOk' or tools_status == 'toolsOld':
                            for nic in vm.guest.net:
                                if hasattr(nic, 'ipConfig') and hasattr(nic.ipConfig, 'ipAddress'):
                                    for addr in nic.ipConfig.ipAddress:
                                        vm_ips.append(addr.ipAddress)
                        else:
                            vm_ips = "none"
                        dataframe[vcenter]["DataCenter"].append(dc_name)
                        dataframe[vcenter]["Cluster"].append(cluster_name)
                        dataframe[vcenter]["Host"].append(host_name)
                        dataframe[vcenter]["Name"].append(vm_name)
                        dataframe[vcenter]["ID"].append(vm_id)
                        dataframe[vcenter]["Power_status"].append(power_status)
                        dataframe[vcenter]["Is_template"].append(is_template)
                        dataframe[vcenter]["Number of CPUs"].append(cpu_num)
                        dataframe[vcenter]["CPU_Cores"].append(cpu_cores)
                        dataframe[vcenter]["Total CPU"].append(total_cpu)
                        dataframe[vcenter]["CPU Hot_add"].append(cpu_hotadd)
                        dataframe[vcenter]["Memory GB"].append(memory_gb)
                        dataframe[vcenter]["Memory Hot_add"].append(memory_hotadd)
                        dataframe[vcenter]["VM_Tools Version"].append(vm_tools)
                        dataframe[vcenter]["virtual_disks"].append(virtual_disks)
                        dataframe[vcenter]["datastore list"].append(datastores_list)
                        dataframe[vcenter]["Total allocated disk"].append(totalAllocatedDisk)
                        # dataframe[vcenter]["VM Path"].append(vm_path_name)
                        dataframe[vcenter]["Network Count"].append(ethernet_cards)
                        dataframe[vcenter]["Ethernet list"].append(ethernets_list)
                        dataframe[vcenter]["networks_name"].append(networks_name)
                        dataframe[vcenter]["tools_installType"].append(tools_installType)
                        dataframe[vcenter]["tools_runningStatus"].append(tools_runningStatus)
                        dataframe[vcenter]["tools_status"].append(tools_status)
                        dataframe[vcenter]["tools_version"].append(tools_version)
                        dataframe[vcenter]["tools_versionStatus"].append(tools_versionStatus)
                        dataframe[vcenter]["tools_versionStatus2"].append(tools_versionStatus2)
                        dataframe[vcenter]["vm_guestFamilly"].append(vm_guestFamilly)
                        dataframe[vcenter]["vm_guestHostname"].append(vm_guestHostname)
                        dataframe[vcenter]["vm_hwVersion"].append(vm_hwVersion)
                        dataframe[vcenter]["vm_ip"].append(vm_ip)
                        dataframe[vcenter]["vm_ips"].append(vm_ips)


        df = pd.DataFrame(dataframe[vcenter])
        df.to_excel(writer, sheet_name=vcenter, index=False, encoding='utf-8')
        workbook = writer.book
        worksheet = writer.sheets[vcenter]
        header_format = workbook.add_format({
            'bold': True,
            'fg_color': '#6a7ed9',
            'align': 'center',
            'border': 1})

        for idx, col in enumerate(df):
            series = df[col]
            max_len = max((
                series.astype(str).map(len).max(),
                len(str(series.name))
            )) + 1
            worksheet.set_column(idx, idx, 20)
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

    writer.save()

def send_mail(contact , body):
    sender_email = "automation@mycompany.com"
    receiver_email = contact

    msg = MIMEMultipart()
    msg['Subject'] = 'VM List(all)'
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg.attach(MIMEText(body))
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(file_name, "rb").read())
    encoders.encode_base64(part)
    part.add_header('content-disposition', 'attachment', filename=file_name)
    msg.attach(part)
    with smtplib.SMTP('webmail.mycompany.ir', 25) as smtpObj:
        smtpObj.ehlo()
        smtpObj.starttls()
        smtpObj.sendmail(sender_email, receiver_email, msg.as_string())

if __name__ == "__main__":
    vm_list()
    send_mail("virtualization@mycompany.com", "Dear colleagues,\n Please find th attachment for VM list from all vCenters")
