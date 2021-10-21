# nsx-dfw-export

PowerShell script created to export NSX-T Distributed Firewall rules.

The DFW can be exported directly from the UI of NSX-T; however, it does not convert the Path/UUID of the objects to the display name of the groups.  This script does that using hash tables.

The PS1 file generates CSV outputs from the Invoke-RestMethod PowerShell command and combines them into a single Excel workbook.

Currently working towards a more optimal way of displaying VM and IP-set group members (in progress).
