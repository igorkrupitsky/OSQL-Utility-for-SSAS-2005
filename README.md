# "OSQL Utility" for SSAS 2005

Originally posted here:
https://www.codeproject.com/Articles/27081/-OSQL-Utility-for-SSAS-2005

A script that lets you run many XMLA files against a SSAS 2005 database.

## Introduction
One cool thing about Microsoft Analysis Services 2005 is that it lets you script many administrative tasks (such as role creation and cube processing). The script files are saved with an XMLA extension and are written in Analysis Services Scripting Language (ASSL) format.

The problem arises when you need to run a lot (100+) of these files against a SSAS database. Unfortunately, I could not find a utility to run these files programmatically. This article describes how to create a VBS script file to run many XMLA files against SSAS.

## Background
The trick is to use HTTP access to send the Execute XMLA request to SSAS. The execute request, along with letting you send MDX commands, lets you send ASSL commands.

Deployment
The most difficult part is the configuration. Here are the steps:

1. Setup HTTP access to SQL Server 2005 Analysis Services: http://www.microsoft.com/technet/prodtechnol/sql/2005/httpasws.mspx.
2. Make sure that XMLA virtual directory is set to basic authentication only. Note that you can also use Anonymous access (with a user that has admin access to your SSAS server), but this option is less secure.
3. Change the configuration file (xmlaConfig.xml) to point to your XMLA provider (like http://MyServer/xmla/msmdpump.dll). Set the Windows user name (domain\username) and password that has admin access to your SSAS server.
4. Drop your XMLA files on top of the script file (xmla.vbs)
5. Optionally, you can create a subfolder called XMLA in the same folder where the script file resides. The script file will look for the XMLA subfolder and run all XMLA files within it.
