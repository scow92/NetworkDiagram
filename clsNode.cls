VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================================================================
' clsNode - one device/equipment endpoint referenced by a circuit.
'===========================================================================
Public Key       As String    ' unique node key (room|rack|equip)
Public Room      As String    ' room / location name
Public Rack      As String    ' rack / cabinet name
Public Equip     As String    ' equipment / device name
Public IsPassive As Boolean   ' True for ODF / patch panels / cassettes
