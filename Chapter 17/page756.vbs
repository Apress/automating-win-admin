'set the security descriptor for the file using the ADsSecurityUtility object
objSecurity.SetSecurityDescriptor "c:\data", 
                    ADS_PATH_FILE, objSD, ADS_SD_FORMAT_IID objSecurity.SetSecurityDescriptor objSD
