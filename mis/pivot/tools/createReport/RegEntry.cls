Attribute VB_Name = "clsRegEntry"
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'
' WhatString: mis/pivot/tools/createReport/RegEntry.cls 1.0 10-JUN-2008 10:32:38 MBA
'
'
'
' Maintained by: 
'
' Description  : einzelner Registryeintrag /Document Property Eintrag
'
' Keywords     :
'
' Reference    :
'
' Copyright    : varetis COMMUNICATIONS GmbH, Grillparzer Str.10, 81675 Muenchen, Germany
'
'----------------------------------------------------------------------------------------
'
'Declarations


'Options
Option Explicit

'Declare variables
Dim mstrValue As String         'Registry Wert
Dim mstrKey As String           'Registry Schlüssel
Dim mstrName As String          'Registry Eintrag
Dim mstrPropName As String      'Document Property Name
Dim mintPropType As Integer     'Document Property Datentyp

'Declare constants
Const what = "@(#) mis/pivot/tools/createReport/RegEntry.cls 1.0 10-JUN-2008 10:32:38 MBA"

'-------------------------------------------------------------
' Description   :
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub Class_Initialize()

    'Properties initialisieren
    mstrValue = ""
    mstrKey = ""
    mstrName = ""
    mstrPropName = ""
    mintPropType = msoPropertyTypeString    'defaultmäßig Typ Text (msoPropertyTypeString)
End Sub


'-------------------------------------------------------------
' Description   :
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Private Sub Class_Terminate()

End Sub



'-------------------------------------------------------------
' Description   : SubKey
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get Key() As String

    Key = mstrKey
End Property

'-------------------------------------------------------------
' Description   : SubKey
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Let Key(ByVal pstrKey As String)

    mstrKey = pstrKey
End Property

'-------------------------------------------------------------
' Description   : Wert (hier können nur Stringeintäge verwendet werden)
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get Value() As String

    Value = mstrValue
End Property

'-------------------------------------------------------------
' Description   : Wert (hier können nur Stringeintäge verwendet werden)
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Let Value(ByVal pstrValue As String)

    mstrValue = pstrValue
End Property

'-------------------------------------------------------------
' Description   : Name
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get Name() As String

    Name = mstrName
End Property

'-------------------------------------------------------------
' Description   : Name des Registryeintrags
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Let Name(ByVal pstrName As String)

    mstrName = pstrName
End Property

'-------------------------------------------------------------
' Description   : Name des Document Property Eintrags
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get PropName() As String

    PropName = mstrPropName
End Property

'-------------------------------------------------------------
' Description   : Name des Document Property Eintrags
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Let PropName(ByVal pstrPropName As String)

    mstrPropName = pstrPropName
End Property

'-------------------------------------------------------------
' Description   : Datentyp des Document Property Eintrags
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Get PropType() As Integer

    'defaultmäßig Typ Text (msoPropertyTypeString)
    PropType = mintPropType
End Property

'-------------------------------------------------------------
' Description   : Datentyp des Document Property Eintrags
'
' Reference     :
'
' Parameter     :
'
' Exception     :
'-------------------------------------------------------------
'
Public Property Let PropType(ByVal pintPropType As Integer)
    
    mintPropType = pintPropType
End Property
