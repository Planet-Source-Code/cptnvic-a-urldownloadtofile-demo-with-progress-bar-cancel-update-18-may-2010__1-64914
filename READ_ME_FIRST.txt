IF YOU THINK YOU CAN JUST UN-ZIP THIS CODE AND HIT F5... YOU ARE NUTZ!!!
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Here's how to make it work:

1) Download the FREE type library at:
    http://www.mvps.org/emorcillo/download/vb6/tl_ole.zip

2) Un-zip the type library to your SYSTEM or SYSTEM32 folder (depending on your OS)

3) Register the type library:
    From the task bar:  START>RUN
    Type the following in the dialog box that appears:
    regtlib <Full path of .tlb file>    
    
    WHERE <Full path of .tlb file> is the location of your type library file.
    SO, if you unzip tl_ole.zip to the SYSTEM32 folder... you would type:

              regtlib c:\system32\olelib.tlb        ... and click the ok button

4) Start VB and Load this project

5) From the VB IDE, Select: PROJECT>REFERENCES
    In the references list that appears, find:
    Edamo's OLE interfaces & functions v1.81     in that list, and check the check box.
    Click the OK button.

NOW You can hit F5 and run the code!
