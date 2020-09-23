Attribute VB_Name = "Firewall"
Option Explicit
  'winsock1 - to obtain local ip

  Private Type PF_FILTER_DESCRIPTOR
    dwFilterFlags As Long
    dwRule        As Long
    pfatType      As Long 'ENUM_PFADDRESSTYPE
    SrcAddr       As Long
    SrcMask       As Long
    DstAddr       As Long
    DstMask       As Long
    dwProtocol    As Long
    lfLateBound   As Long
    wSrcPort      As Integer
    wDstPort      As Integer
    wSrcPortHighRange As Integer
    wDstPortHighRange As Integer
  End Type

  Private Const FILTER_PROTO_ANY = 0
  Private Const FILTER_PROTO_ICMP = 1
  Private Const FILTER_PROTO_TCP = 6
  Private Const FILTER_PROTO_UDP = 17 '11
  
  Private Const FILTER_TCPUDP_PORT_ANY = 0
  
  Private Const FILTER_ICMP_TYPE_ANY = 255
  Private Const FILTER_ICMP_CODE_ANY = 255

  Private Const FD_FLAGS_NOSYN = 1
  Private Const FD_FLAGS_ALLFLAGS = FD_FLAGS_NOSYN

  Private Const IP_LOCALHOST = "localhost"
  Private Const IP_MASKALL = "0.0.0.0"
  Private Const IP_MASKNONE = "255.255.255.255"
  
  Private Const ERROR_BASE As Integer = 23000
  Private Const PFERROR_NO_PF_INTERFACE As Integer = (ERROR_BASE + 0)
  Private Const PFERROR_NO_FILTERS_GIVEN As Integer = (ERROR_BASE + 1)
  Private Const PFERROR_BUFFER_TOO_SMALL As Integer = (ERROR_BASE + 2)
  Private Const ERROR_IPV6_NOT_IMPLEMENTED As Integer = (ERROR_BASE + 3)
  Private Const NO_ERROR As Long = 0

  Private Enum ENUM_PFADDRESSTYPE
    PF_IPV4 = 0
    PF_IPV6 = 1
  End Enum

  Private Enum ENUM_PF_ACTION
    PF_ACTION_FORWARD = 0
    PF_ACTION_DROP = 1
  End Enum
  
  Private Declare Function PfCreateInterface Lib "Iphlpapi.dll" Alias "_PfCreateInterface@24" _
    (ByVal dwName As Long, ByVal inAction As ENUM_PF_ACTION, ByVal outAction As ENUM_PF_ACTION, _
    ByVal bUseLog As Boolean, ByVal bMustBeUnique As Boolean, ByRef ppInterface As Long) As Long
    
  Private Declare Function PfBindInterfaceToIPAddress Lib "Iphlpapi.dll" Alias "_PfBindInterfaceToIPAddress@12" _
    (ByVal pInterface As Long, ByVal PFADDRESSTYPE As ENUM_PFADDRESSTYPE, IPAddress As Any) As Long
  
  Private Declare Function PfAddFiltersToInterface Lib "Iphlpapi.dll" Alias "_PfAddFiltersToInterface@24" _
    (ByVal ih As Long, ByVal cInFilters As Long, ByRef pfiltIn As Any, ByVal cOutFilters As Long, ByRef pfiltOut As Any, _
    ByRef pfHandle As Any) As Long

  Private Declare Function PfRemoveFilterHandles Lib "Iphlpapi.dll" Alias "_PfRemoveFiltersFromInterface@20" _
    (ByVal ih As Long, ByVal cInFilters As Long, ByRef pfiltIn As Any, ByVal cOutFilters As Long, _
    ByRef pfiltOut As Any) As Long

  Private Declare Function PfDeleteInterface Lib "Iphlpapi.dll" Alias "_PfDeleteInterface@4" _
    (ByVal ppInterface As Long) As Long
  
  Private Declare Function PfUnBindInterface Lib "Iphlpapi.dll" Alias "_PfUnBindInterface@4" _
    (ByVal ppInterface As Long) As Long
    
  '*****************************************************************************************************************
  '*****************************************************************************************************************
'  Private InFilter As PF_FILTER_DESCRIPTOR
'  Private OutFilter As PF_FILTER_DESCRIPTOR
  Private hInterface As Long
  Private Handle3    As Long
  Private Result     As Long
  
  Public RulesFirewallDate As Date 'to refresh rutes array
  
  Dim ArrIP(3) As Byte, IpStr$ 'to pass local ip to the respective function
Public Sub StartFilters() 'allow all less the rules
  On Error GoTo StartFilters_Error
  Dim Registro() As String

  IpStr = Form1.Winsock1.LocalIP
  Registro = Split(IpStr, ".")
  
  ArrIP(0) = CByte(Registro(0)) '172 ' 192
  ArrIP(1) = CByte(Registro(1)) '26 '168
  ArrIP(2) = CByte(Registro(2)) '0 '5
  ArrIP(3) = CByte(Registro(3)) '2 '136
    
  Result = PfCreateInterface(0&, PF_ACTION_FORWARD, PF_ACTION_FORWARD, False, True, hInterface)
  Result = PfBindInterfaceToIPAddress(hInterface, ByVal PF_IPV4, ArrIP(0))
  
  AddFilter

  Exit Sub
StartFilters_Error:
  MsgBox "MbInfo: Error en StartFilters: " & Err.Description, vbCritical, Now
     
End Sub
Public Sub StopFilters()
  On Error GoTo StopFilters_Error

  If hInterface <> 0 Then
    PfUnBindInterface hInterface                            '// Un-Bind
    PfDeleteInterface hInterface                            '// Delete
    hInterface = 0
  End If

  Exit Sub
StopFilters_Error:
  MsgBox "MbInfo: Error en StopFilters: " & Err.Description, vbCritical, Now

End Sub
Public Sub BlockAll()
  On Error GoTo BlockAll_Error
  
  'Result = PfCreateInterface(0&, PF_ACTION_FORWARD, PF_ACTION_FORWARD, False, True, hInterface)
  Result = PfCreateInterface(0&, PF_ACTION_DROP, PF_ACTION_DROP, False, True, hInterface)
  Result = PfBindInterfaceToIPAddress(hInterface, ByVal PF_IPV4, ArrIP(0))
  
  Exit Sub
BlockAll_Error:
  MsgBox "MbInfo: Error en BlockAll: " & Err.Description, vbCritical, Now
     
End Sub
Public Sub ResetFilters() 'used in Form1 Time
  On Error GoTo ResetFilters_Error

  StopFilters
  StartFilters

  Exit Sub
ResetFilters_Error:
  MsgBox "MbInfo: Error en ResetFilters: " & Err.Description, vbCritical, Now

End Sub
Private Sub AddFilter() 'use the IsIp function to verify
  On Error GoTo AddFilter_Error
  
  If Form1.ListView1.ListItems.Count = 0 Then Exit Sub

  Dim i&
  Dim aIP(3) As Byte, aIpAny(3) As Byte, sIP$, Registro() As String 'ip arrays
  Dim SrcAddr&, SrcMask&, DstAddr&, DstMask&
  
  aIpAny(0) = 0
  aIpAny(1) = 0
  aIpAny(2) = 0
  aIpAny(3) = 0
  
  For i = 1 To Form1.ListView1.ListItems.Count
    Dim InFilter  As PF_FILTER_DESCRIPTOR
    
    'commom
    InFilter.dwFilterFlags = FD_FLAGS_NOSYN
    InFilter.pfatType = PF_IPV4
    InFilter.dwRule = 0
    InFilter.lfLateBound = 0
    
    'ip source
    sIP = Form1.ListView1.ListItems(i).Text
    If Val(sIP) = 0 Then
      'InFilter.SrcAddr = aIpAny(0)
      SrcAddr = aIpAny(0)
    Else
      Registro = Split(sIP, ".")
      aIP(0) = CByte(Registro(0))
      aIP(1) = CByte(Registro(1))
      aIP(2) = CByte(Registro(2))
      aIP(3) = CByte(Registro(3))
      'InFilter.SrcAddr = aIP(0)
      SrcAddr = aIP(0)
    End If
    
    'mask source
    sIP = Form1.ListView1.ListItems(i).SubItems(1)
    If sIP = "" Then
      SrcMask = vbNull
    ElseIf Val(sIP) = 0 Then
      'InFilter.SrcMask = aIpAny(0)
      SrcMask = aIpAny(0)
    Else
      Registro = Split(sIP, ".")
      aIP(0) = CByte(Registro(0))
      aIP(1) = CByte(Registro(1))
      aIP(2) = CByte(Registro(2))
      aIP(3) = CByte(Registro(3))
      'InFilter.SrcMask = aIP(0)
      SrcMask = aIP(0)
    End If
    
    'port source
    InFilter.wSrcPort = Val(Form1.ListView1.ListItems(i).SubItems(2))
    InFilter.wSrcPortHighRange = Val(Form1.ListView1.ListItems(i).SubItems(2))
    
    'destin. ip
    sIP = Form1.ListView1.ListItems(i).SubItems(3)
    If Val(sIP) = 0 Then
      'InFilter.DstAddr = aIpAny(0)
      DstAddr = aIpAny(0)
    Else
      Registro = Split(sIP, ".")
      aIP(0) = CByte(Registro(0))
      aIP(1) = CByte(Registro(1))
      aIP(2) = CByte(Registro(2))
      aIP(3) = CByte(Registro(3))
      'InFilter.DstAddr = aIP(0)
      DstAddr = aIP(0)
    End If
    
    'destin. mask
    sIP = Form1.ListView1.ListItems(i).SubItems(4)
    If sIP = "" Then
      DstMask = vbNull
    ElseIf Val(sIP) = 0 Then
      'InFilter.DstMask = aIpAny(0)
      DstMask = aIpAny(0)
    Else
      Registro = Split(sIP, ".")
      aIP(0) = CByte(Registro(0))
      aIP(1) = CByte(Registro(1))
      aIP(2) = CByte(Registro(2))
      aIP(3) = CByte(Registro(3))
      'InFilter.DstMask = aIP(0)
      DstMask = aIP(0)
    End If
    
    'destin. Port
    InFilter.wDstPort = Val(Form1.ListView1.ListItems(i).SubItems(5))
    InFilter.wDstPortHighRange = Val(Form1.ListView1.ListItems(i).SubItems(5))
    
    'protocol
    InFilter.dwProtocol = GetProtocol(Form1.ListView1.ListItems(i).SubItems(6))
    
    InFilter.SrcAddr = VarPtr(SrcAddr)
    InFilter.SrcMask = VarPtr(SrcMask)
    InFilter.DstAddr = VarPtr(DstAddr)
    InFilter.DstMask = VarPtr(DstMask)
    
    If Form1.ListView1.ListItems(i).SubItems(7) = "OUT" Then 'out
      Result = PfAddFiltersToInterface(hInterface, 0, ByVal 0&, 1, InFilter, Handle3)
    Else 'in
      Result = PfAddFiltersToInterface(hInterface, 1, InFilter, 0, ByVal 0&, Handle3)
    End If
    
    If Result = PFERROR_NO_FILTERS_GIVEN Then MsgBox "Error: PFERROR_NO_FILTERS_GIVEN"
  Next 'i
  
  Exit Sub
AddFilter_Error:
  MsgBox "MbInfo: Error en AddFilter: " & Err.Description, vbCritical, Now
     
End Sub
Public Sub LoadRulesFirewall() 'load firewall rules file
  On Error GoTo LoadRulesFirewall_Error
  
  Dim Canal%, x&, sLine As String, a$, Registro() As String
  
  'if don´t exists the file, create it
  If Dir$(App.Path & "\FirewallRules.txt") = vbNullString Then 'si no existe, lo crea
    Canal = FreeFile
    Open App.Path & "\FirewallRules.txt" For Output As #Canal
    Close #Canal
  End If
  
  If FileLen(App.Path & "\FirewallRules.txt") = 0 Then
    RulesFirewallDate = FileDateTime(App.Path & "\FirewallRules.txt")
    Exit Sub
  End If
   
  Form1.ListView1.ListItems.Clear   'lo borra
  Canal = FreeFile
  'read the file and load it to the lv1
  Open App.Path & "\FirewallRules.txt" For Input As #Canal
  While Not EOF(Canal)
    Line Input #Canal, sLine
    If sLine <> vbNullString Then
      If Left$(sLine, 1) <> "#" Then
        If InStr(sLine, "#") > 0 Then
          a = Trim$(Left$(sLine, InStr(sLine, "#") - 1)) '# = for comments
          a = Replace$(a, vbTab, "")
          Registro = Split(a, ",") '8 registers
          If UBound(Registro) <> 7 Then
            MsgBox "bad register", vbExclamation, UBound(Registro)
            Close #Canal
            Exit Sub
          End If
          Form1.ListView1.ListItems.Add , , Registro(0)
        Else
          Registro = Split(sLine, ",")
          If UBound(Registro) <> 7 Then
            MsgBox "bad register", vbExclamation, UBound(Registro)
            Close #Canal
            Exit Sub
          End If
          Form1.ListView1.ListItems.Add , , Registro(0)
        End If
          
        For x = 1 To 7
          Form1.ListView1.ListItems(Form1.ListView1.ListItems.Count).SubItems(x) = cn(Registro(x))
        Next 'x
        Form1.ListView1.ListItems(Form1.ListView1.ListItems.Count).SubItems(8) = "Block"
      End If
    End If
  Wend
  Close #Canal
    
  RulesFirewallDate = FileDateTime(App.Path & "\FirewallRules.txt")

  Exit Sub
LoadRulesFirewall_Error:
  MsgBox "MbInfo: Error en LoadRulesFirewall: " & Err.Description, vbExclamation, Now

End Sub
Public Function GetProtocol(Datos As String) As Long
  On Error GoTo GetProtocol_Error

  Datos = LCase$(Datos)
  Select Case Datos
    Case "all"
     GetProtocol = 0
    Case "icmp"
      GetProtocol = 1
    Case "tcp"
      GetProtocol = 6
    Case "udp"
      GetProtocol = 17
    Case Else
      GetProtocol = 0
  End Select

  Exit Function
GetProtocol_Error:
  MsgBox "MbInfo: Error en GetProtocol del Firewall: " & Err.Description, vbCritical, Now

End Function
Public Function IsIp(Cadena As String) As Boolean 'to verify the string
  On Error GoTo IsIp_Error

  Dim Temp() As String 'base 0
  Dim x&
  
  If Cadena = "" Then 'if null
    IsIp = False
    Exit Function
  End If
  
  If Not IsNumeric(Cadena) Then 'if not numbers
    IsIp = False
    Exit Function
  End If
    
  Temp = Split(Cadena, ".")
  If UBound(Temp) <> 3 Then 'if not 4 fields
    IsIp = False
    Exit Function
  End If
  
  If Val(Temp(0)) = 0 Then 'if 0 in the 1º field
    IsIp = False
    Exit Function
  End If
  
  For x = LBound(Temp) To UBound(Temp)
    If (Temp(x)) = "" Then 'if field is null
      IsIp = False
      Exit Function
    End If
    
    If Val(Temp(x)) >= 0 And Val(Temp(x)) < 256 Then 'the range
      IsIp = True
    Else
      IsIp = False
      Exit Function
    End If
  Next 'x
  
  Exit Function
IsIp_Error:
  MsgBox "MbCap: Error en IsIp: " & Err.Description, vbExclamation, Now

End Function
Public Function cn(pVal As String) As String  'to control the nulls
  On Error GoTo Cn_Error

  If IsMissing(pVal) Then
    cn = "Missing"
  ElseIf IsNull(pVal) Then
    cn = "Null"
  ElseIf IsEmpty(pVal) Then
    cn = "Null"
  Else
    cn = pVal
  End If
  
  Exit Function
Cn_Error:
  MsgBox "MbInfo: Error en cn: " & Err.Description, vbExclamation, Now

End Function
