Many of examples in this directory makes reference to generic computer and domain names. 
Because of this it is unlikely they will run unmodified in your enviroment.

The QADNA.mdb is written in Access 2000, but can be converted to later versions of Access with 
no problems.

The QADNA interface uses ADSI to catalogue computer objects and user/group resources. It will
work on NT4 or later. ADSI is included with Windows 2000 and later but is a separate install
for NT4. 

To catalogue resource objects, select the Configure QADNA option from the main QADNA menu. 
from the configuration form enter the name of your domain and select the Update Domain Objects
option. This will update any users, groups and computer objects from the specified domain.

One this operation is completed you can select the Update Computer Objects button to catalogue
file shares and print queues. 

The configuration option only needs to be done once and whenever new users/groups/shares/print 
queues are added to your system.