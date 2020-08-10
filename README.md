# Creating an Excel report of Cisco device's Inventory

  - Use openpyxl moduel to create an Excel report
  - Read a tab-delimited file which contains the list of devices as a input file
  - Use sh inventory command to gather data

## Report preview
|Hostname|IP Address|Name|Description|PID|Serial Number|
|:------:|:-----------:|:---------:|:---------:|:---------:|:---------:|
|Switch1|10.1.1.1|Switch1 Supervisor 1|Something|Something|ABCD1234

## Tab-delimited file
Switch1	IOS	10.1.1.1	SSH   
Switch2	IOS	10.1.22.1	SSH   
Switch3	IOS	10.3.1.1	SSH
