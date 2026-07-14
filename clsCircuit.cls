VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCircuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================================================================
' clsCircuit - one physical cable/circuit segment between two device ends.
'===========================================================================
Public Cable      As String   ' cable reference / label
Public CableType  As String   ' "Fibre" | "DAC" | "Copper"
Public FromKey    As String   ' End-A node key (room|rack|equip)
Public ToKey      As String   ' End-B node key (room|rack|equip)
Public FromPort   As String   ' End-A port
Public ToPort     As String   ' End-B port
Public Cal        As String   ' CAL / continuity id (fibre + DAC)
Public Length     As String   ' cable length
Public ConnectorA As String   ' End-A connector type
Public TxA        As String   ' End-A transmit / fibre detail
Public ConnectorB As String   ' End-B connector type
Public TxB        As String   ' End-B transmit / fibre detail
