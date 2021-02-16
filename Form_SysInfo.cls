VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


'Created by:             Marco Caruso, Marketplace Enrollment - carusomj@upmc.edu
'Modified Date:          7/9/2019
'Version Go Live:        4/5/2019
'Purpose:                To log, quantify and display employee quality and production data
'
'                        Current Version = version 5.0
'
'                            Change Log
'                            5.0: New interface
'
'                            5.1: removed db and rs objects as global vars and made them local. With each sub, the objects are created and killed.
'                                 Was getting errors 3043 and 3040, most likely caused by having multiple db objects open at once. This was causing backend to corrupt.
'                                 See here for info on the errors: http://www.accessrepairnrecovery.com/blog/how-to-fix-ms-access-error-3043
'
'                                 Also updated the LogError sub to delimit each section with | character




Private Sub Form_Load()
    Dim test As String: test = 5
End Sub
