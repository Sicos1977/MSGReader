msgreader
=========

- 2014-03-06 Version 1.1
      -  Added support for special characters like German umlauts, they are parsed out of the HTML text that is                 embedded inside the RTF
      -  The RTFBody was loaded 4 times instead of once
      -  A CC was always added even when there was no CC (the To was taken in that case)
      -  Fixed some minor issues and cleaned up the code 
      
- 2014-01-16 First release


SIDE NOTE:
==========

This project can also be used from a COM based language like VB script or VB6.
To use it first compile the code and register the com visible assembly with the command:

Regasm.exe /codebase DocumentServices.Modules.Readers.MsgReader.dll

After that you can call it like this:

dim msgreader
set msgreader = createobject("DocumentServices.Modules.Readers.MsgReader.Reader")
msgreader.ExtractToFolder "<some input msg file>", "<some output folder>"
