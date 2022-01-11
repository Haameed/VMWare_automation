import connector
import secret
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime



file_name = f'vm_snapshot_{datetime.now().strftime("%Y-%m-%d")}.xlsx'
def vm_list():
    vcenters = ["vcsa-hq", "vcsa-gorgan", "vcsa-datacenter", "vcsa-yara"]
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    dataframe = {}
    for vcenter in vcenters:
        dataframe[vcenter] = {"DataCenter": [], "Cluster": [], "Host": [], "Name": [], "ID": [],
                     "current_snapshot_id": [], "snapshot_name": [], "snapshot_id": [],
                     "snapshot_age": [], "snapshot_state": [], "snapshot_desc": [], "snapshot_has_child": [], "snapshot_childs": []}
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
                        if vm.snapshot:
                            current_snapshot_id = vm.snapshot.currentSnapshot._moId
                            for snapshot in vm.snapshot.rootSnapshotList:
                                snapshot_name = snapshot.name
                                snapshot_id = snapshot.id
                                snapshot_creation_date = snapshot.createTime.replace(tzinfo=None)
                                snapshot_desc = snapshot.description
                                snapshot_state = snapshot.state
                            if snapshot.childSnapshotList:
                                snapshot_childs = []
                                snapshot_has_child = True
                                for i in snapshot.childSnapshotList:
                                    snapshot_childs.append(i.name)
                            else:
                                snapshot_has_child = False
                            snapshot_age = str(datetime.now() - snapshot_creation_date)
                            vm_name = vm.name
                            vm_id = str(vm).split(":")[1].replace("'", "")
                            dataframe[vcenter]["DataCenter"].append(dc_name)
                            dataframe[vcenter]["Cluster"].append(cluster_name)
                            dataframe[vcenter]["Host"].append(host_name)
                            dataframe[vcenter]["Name"].append(vm_name)
                            dataframe[vcenter]["ID"].append(vm_id)
                            dataframe[vcenter]["current_snapshot_id"].append(current_snapshot_id)
                            dataframe[vcenter]["snapshot_name"].append(snapshot_name)
                            dataframe[vcenter]["snapshot_id"].append(snapshot_id)
                            dataframe[vcenter]["snapshot_age"].append(snapshot_age)
                            dataframe[vcenter]["snapshot_state"].append(snapshot_state)
                            dataframe[vcenter]["snapshot_desc"].append(snapshot_desc)
                            dataframe[vcenter]["snapshot_has_child"].append(snapshot_has_child)
                            dataframe[vcenter]["snapshot_childs"].append(snapshot_childs)
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
            worksheet.set_column(idx, idx, max_len)
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

    writer.save()

def send_mail(contact , body):
    sender_email = "automation@mycompany.com"
    receiver_email = contact

    msg = MIMEMultipart()
    msg['Subject'] = 'snapshot list(all)'
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg.attach(MIMEText(body))
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(file_name, "rb").read())
    encoders.encode_base64(part)
    part.add_header('content-disposition', 'attachment', filename=file_name)
    msg.attach(part)
    with smtplib.SMTP('webmail.mycompany.com', 25) as smtpObj:
        smtpObj.ehlo()
        smtpObj.starttls()
        smtpObj.sendmail(sender_email, receiver_email, msg.as_string())

if __name__ == "__main__":
    vm_list()
    send_mail("virtualization@mycompany.com", "Dear colleagues,\n Please find th attachment for Snapshot list from all vCenters")


