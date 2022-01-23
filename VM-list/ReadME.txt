The script connects to multiple Vcenter servers to gather data about virtual machines in a single Excel file and email it to the user.
Each vcenter's information will be stored on a separate sheet. Here is the data for each machine:
1- Datacenter that hosts VM
2- the Cluster which the VM is a part of 
3- The pysical server which hosts the VM
4- Name of Virtual machine 
5- VM ID
6- VM power status 
7- Show if The VM is marked as template
8- Number of CPUs 
9- Number of CPU cores 
10- Total CPU count 
11- Cpu hot add Status
12- Alocated Memory to the VM
13- Memory Hot add status 
15- VM virtual disks and datastores and capacity (GB)
16- Total alocated Disk 
17- Data store names 
18- VM path
19- Number of network interfaces 
20- Port group names 
21- VMWare-Tools install type 
22- VMWare-tools running status 
23- VMWare-tools Status ( OK, Old, Not Installed,Not Running)
24- VMWare-tools version
25- VMWare-tools version status 
26- VM guest familly
27- VM guest hostname 
28- VM hardware version
29- VM main IP
30- Other VM IPs (virtual IPs or Other interfaces IP with binded mac address)

There may be a problem with the coding style or a better way to gather this information. Any comments to improve performance, coding, or coding style would be appreciated
