VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComplexDataConsumer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Ҫ��: ����һ�� recordset ��������Ϊ�������Դ������:

'Private WithEvents RS As ADODB.RecordSet

Public Property Get DataSource() As DataSource
    ' Ҫ��: �����������Դ��������:
   
    'Set DataSource = RS.DataSource
End Property
Public Property Set DataSource(ByVal objDataSource As DataSource)
    ' Ҫ��: �����������Դ��������:
    
    'Set RS.DataSource = objDataSource
    'PropertyChanged "DataSource"
End Property
Public Property Get DataMember() As DataMember
    ' Ҫ��: �������ݳ�Ա������:
    
    'Set DataMember = RS.DataMember
End Property
Public Property Let DataMember(ByVal DataMember As DataMember)
    ' Ҫ��: ָ��һ���µ����ݳ�Ա������:
    
    'Set RS.DataMember = DataMember
    'PropertyChanged "DataMember"
End Property
