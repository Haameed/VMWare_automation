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
15- VM virtual disk 
16- Data store names 
17- VM path
18- Number of network interfaces 
19- Port group names 
20- VMWare-Tools install type 
21- VMWare-tools running status 
22- VMWare-tools Status ( OK, Old, Not Installed,Not Running)
23- VMWare-tools version
24- VMWare-tools version status 
25- VM guest familly
26- VM guest hostname 
27- VM hardware version
28- VM main IP
29- Other VM IPs (virtual IPs or Other interfaces IP with binded mac address)

There may be a problem with the coding style or a better way to gather this information. Any comments to improve performance, coding, or coding style would be appreciated
