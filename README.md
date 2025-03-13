# Office-2024-Offline-Install
Included instructions can be used to install Microsoft Office LTSC Professional Plus 2024 inside an air gapped environment If changes are required, for example Office version, the XML configuration file will need to be modified. Office Deployment Tool is required to 
to generate the required files locally. 
### Pulling down the desired installation files from Microsoft
Install the Office Deployment Tool. https://www.microsoft.com/en-us/download/details.aspx?id=49117. Once installed we can focus on the XML config file. 
Here is an example for Microsoft Office LTSC Professional Plus 2024. You will need the correct GVLK key for the desired version. 
```
<Configuration>
  <Info Description="Office Professional Plus 2024 (64-bit)" />
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024" SourcePath="C:\OfficeSetup">
    <Product ID="ProPlus2024Volume" PIDKEY="XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" UpdatePath="\\<netpathhere>" />
  <RemoveMSI />
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
```
Once you have the correct XML execute the following command via PowerShell or CMD.
```
setup.exe /download configuration-office-professional-plus-2024-64-bit.xml
```
This will take sometime depending on your network connection. Once complete a folder with dat and CAB files should be present. This is referenced in **SourcePath** of the XML. **ODT** and **OfficeSetup** folders/files can now be moved to target machines. 
This process is how setup files are generated. It is a onetime step unless other versions are needed. 
### Installing via ODT
Copy **ODT** and **OffcieSetup** to the target machine and execute the follwoing command:
```
setup.exe /configure configuration-office-professional-plus-2024-64-bit.xml
```
This will intiate the install and should produce Office install window. 

### Troubleshooting
1. If command issues check ODT directory you created per Microsoft installation instructions for ODT.
2. If activation issues make sure the correct GVLK is in the XML. Recreating install files will be needed for XML changes. 
3. Cross reference XML file location names and directory location. This can be custom. 
