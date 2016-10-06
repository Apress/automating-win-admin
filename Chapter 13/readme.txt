A general issue is the location of sample files. In most cases you will have
to modify code to take into consideration the different locations of data files.

A number of the solution scripts use the OLE DB driver 3.51, which isn't included in the most
recent MDAC/OLE DB install. Use OLE DB driver version 4.0, or whatever is the most recent version
installed on your computer.

Many of the solution scripts use the Access Northwind samples database, which is available
from your Microsoft Office CD. It might not have been installed by default.

A number of the solutions output the contents of database tables to the screen using Wscript.Echo. 
Use cscript to execute these scripts instead of Wscript to avoid interactive message popup for 
each record.

Solution 13-1 defines a DSN used in Solution 13-8.

Solution 13-2 defines a File DSN used in Solution 13-9.

Solution 13-9 uses the UserData.xls file located in this directory. 
This must be manually updated with users, and the UserList range must be expanded
to included any new users. This solution uses a File DSN, ExcelUserData.dsn, which is also
included in this directory. This file might require editing to point to the correct path. Solution
13-2 details the steps for creating this File DSN.


The URLs in solutions 13-13 and 13-16, www.acme.com, is not a valid address.

The WSHENT.CopyTable component described in solution 13-17 can be found in the CopyTable.wsc 
file. It must be registered using regsvr32 or Explorer before it can be used.

Solution 13-19 requires the following 

PARAMETERS category Text, company Text;
SELECT Products.*
FROM Suppliers INNER JOIN (Categories INNER JOIN Products
 ON Categories.CategoryID = Products.CategoryID) 
ON Suppliers.SupplierID = Products.SupplierID
WHERE (((Categories.CategoryName)=[category]) 
AND ((Suppliers.CompanyName)=[company]));


Solution 13-22 defines the following stored procedure
CREATE PROCEDURE AddNewStore @StoreName varChar(40), @Address varChar(40), 
@City varChar(20),@StoreID integer OUTPUT
AS

DECLARE @newid int
Select @newid=max(stor_ID) from stores
Set @newid = @newid + 1

Insert into stores (stor_id,stor_name,stor_address,city)
Values(@newid,@StoreName,@Address,@city)

Select @StoreID =@newid

Solution 13-23 assumes the tables Order Details History and OrderHist exists in the 
Northwinds database. 
These tables are used to archive the order history. Order Details History is a copy
of Order Details and OrderHist is a copy of Orders. Both tables should be empty. 



