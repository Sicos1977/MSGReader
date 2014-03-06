msgreader
=========

This project can also be used from a COM based language like VB script or VB6.

To use it first compile the code and register the com visible assembly with the command:

Regasm.exe /codebase DocumentServices.Modules.Readers.MsgReader.dll

After that you can call it like this:

dim msgreader

set msgreader = createobject("DocumentServices.Modules.Readers.MsgReader.Reader")
msgreader.ExtractToFolder "<some input msg file>", "<some output folder>"
