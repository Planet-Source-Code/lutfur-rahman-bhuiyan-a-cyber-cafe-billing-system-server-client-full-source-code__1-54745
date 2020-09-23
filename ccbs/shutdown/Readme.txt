''''''''''''''''''''''''''''''''''''''''''''
'Written by Marty Forde                    '
''''''''''''''''''''''''''''''''''''''''''''
'About Me:                                 '
'Occupation: Sophmore HighSchool Student   '
'    and VB consultant                     '
'Age: 16                                   '
'Expertise Areas: VB Api programming and   '
'    multimedia programming                '
'E-mail: Marty1149@aol.com                 '
''''''''''''''''''''''''''''''''''''''''''''

this program can shutdown/reboot/logoff a 9x/nt/2000 and remote shutdown a 9x/nt/2000 

Shuttingdown your system:

To shutdown, reboot, or logoff a machine in VB, you use the ExitWindowEx API function. 
In addition, if you are running NT or 2000, you need to call the AdjustTokenProvileges 
function first. This function encapsulates the necessary tasks (i.e., checking the OS and 
calling the AdjustTokenPrivileges function if necessary).

One of the Enum Values in the declarations, passed to the function, determines 
whether you log off, restart, or shutdown, and whether the shutdown is graceful 
(running applications are given a chance to to exit and possibly cancel the shutdown) or 
forced. See the comments within the code for more detail.

Remote shuttingdown systems:

Uses the iniatesystemshutdown api in order to remote shutdown a system