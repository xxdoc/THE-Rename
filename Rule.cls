VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contient UNE règle de renommage"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarRuleIndex As Integer
Private mvarRuleName As String
Private mvarRuleActive As Boolean
Private mvarRuleType As Integer
Private mvarRuleCondition As Integer
Private mvarRuleTestValue As Variant
Private mvarRuleDescription As String
Public Property Let RuleIndex(ByVal Index As Integer)
    mvarRuleIndex = Index
End Property
Public Property Get RuleIndex() As Integer
    RuleIndex = mvarRuleIndex
End Property
Public Property Let RuleDescription(ByVal vData As String)
    mvarRuleDescription = vData
End Property
Public Property Get RuleDescription() As String
Attribute RuleDescription.VB_Description = "Description pour l'utilisateur (générée par le programme)"
    RuleDescription = mvarRuleDescription
End Property
Public Property Let RuleTestValue(ByVal vData As Variant)
Attribute RuleTestValue.VB_Description = "Valeur de test"
    mvarRuleTestValue = vData
End Property
Public Property Set RuleTestValue(ByVal vData As Variant)
    Set mvarRuleTestValue = vData
End Property
Public Property Get RuleTestValue() As Variant
    If IsObject(mvarRuleTestValue) Then
        Set RuleTestValue = mvarRuleTestValue
    Else
        RuleTestValue = mvarRuleTestValue
    End If
End Property
Public Property Let RuleCondition(ByVal vData As Integer)
Attribute RuleCondition.VB_Description = "Par exemple 1 pour >, 2 pour <,3 pour >= ..."
    mvarRuleCondition = vData
End Property
Public Property Get RuleCondition() As Integer
    RuleCondition = mvarRuleCondition
End Property
Public Property Let RuleType(ByVal vData As Integer)
Attribute RuleType.VB_Description = "Indique le type de règle, (par exemple 1=taille, 2=Attributs)"
    mvarRuleType = vData
End Property
Public Property Get RuleType() As Integer
    RuleType = mvarRuleType
End Property
Public Property Let RuleActive(ByVal vData As Boolean)
Attribute RuleActive.VB_Description = "Renvoie vrai si cette règle est active"
    mvarRuleActive = vData
End Property
Public Property Get RuleActive() As Boolean
    RuleActive = mvarRuleActive
End Property
Public Property Let RuleName(ByVal vData As String)
    mvarRuleName = vData
End Property
Public Property Get RuleName() As String
Attribute RuleName.VB_Description = "Nom de la règle donné par l'utilisateur"
Attribute RuleName.VB_UserMemId = 0
    RuleName = mvarRuleName
End Property

