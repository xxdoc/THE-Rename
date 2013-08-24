VERSION 5.00
Begin VB.UserControl LabelText 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "LabelText.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "LabelText.ctx":003F
   Begin VB.TextBox txt 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   495
   End
End
Attribute VB_Name = "LabelText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Zone de texte avec étiquette"
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

' Variables de propriétés
Private m_LabelWidth As Long    ' single
Private m_LabelPosition As LabelPositionConstants

' Valeurs de propriétés par défaut
Private Const m_def_LabelWidth = 0
Private Const m_def_LabelPosition = 0

' Variables pour LabelWidth tracking
Private fTracking As Boolean
Private oldLabelWidth As Integer
Private oldX As Integer

' Enumérations
Public Enum BackStyleConstants
    ltBSTransparent
    ltBSOpaque
End Enum

Public Enum LabelPositionConstants
    ltPositionLeft
    ltPositionRight
    ltPositionTop
    ltPositionBottom
End Enum

' Déclarations d'événements
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Event Click() 'MappingInfo=txt,txt,-1,Click
Attribute Click.VB_Description = "Survient lorsque l'utilisateur enfonce puis relâche un bouton de la souris sur un objet."
Attribute Click.VB_UserMemId = -600
Event DblClick() 'MappingInfo=txt,txt,-1,DblClick
Attribute DblClick.VB_Description = "Survient lorsque l'utilisateur enfonce puis relâche un bouton de la souris puis l'enfonce et le relâche de nouveau sur un objet."
Attribute DblClick.VB_UserMemId = -601
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txt,txt,-1,KeyDown
Attribute KeyDown.VB_Description = "Survient lorsque l'utilisateur enfonce une touche alors qu'un objet a le focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txt,txt,-1,KeyPress
Attribute KeyPress.VB_Description = "Survient lorsque l'utilisateur enfonce puis relâche une touche ANSI."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txt,txt,-1,KeyUp
Attribute KeyUp.VB_Description = "Survient lorsque l'utilisateur relâche une touche alors qu'un objet a le focus."
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txt,txt,-1,MouseDown
Attribute MouseDown.VB_Description = "Survient lorsque l'utilisateur enfonce le bouton de la souris alors qu'un objet a le focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txt,txt,-1,MouseMove
Attribute MouseMove.VB_Description = "Se déclenche lorsque l'utilisateur déplace la souris."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txt,txt,-1,MouseUp
Attribute MouseUp.VB_Description = "Survient lorsque l'utilisateur relâche le bouton de la souris alors qu'un objet a le focus."
Attribute MouseUp.VB_UserMemId = -607
Event Change() 'MappingInfo=txt,txt,-1,Change
Attribute Change.VB_Description = "Survient lorsque le contenu d'un contrôle a été modifié."
Attribute Change.VB_MemberFlags = "200"
Event OLECompleteDrag(Effect As Long) 'MappingInfo=txt,txt,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Survient au contrôle source d'une opération de glisser-déplacer OLE après la fin ou l'annulation d'un glisser-déplacer manuel ou automatique."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txt,txt,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Survient lorsque des données sont lâchées sur un contrôle via une opération de glisser-déplacer OLE, alors que OLEDropMode est défini sur manuel."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=txt,txt,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Survient lorsque la souris est déplacée sur le contrôle au cours d'une opération de glisser-déplacer OLE, si sa propriété OLEDropMode est définie sur manuelle."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=txt,txt,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Survient au contrôle source d'une opération de glisser-déplacer OLE lorsque le curseur de la souris doit être modifié."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=txt,txt,-1,OLESetData
Attribute OLESetData.VB_Description = "Survient au contrôle source d'une opération de glisser-déplacer OLE lorsque la cible demande des données qui n'étaient pas fournies au DataObject durant l'événement OLEDragStart."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=txt,txt,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Survient lorsqu'une opération de glisser-déplacer OLE est initialisée manuellement ou automatiquement."
Event LabelWidthChanged()
Attribute LabelWidthChanged.VB_Description = "Modification de la largeur du libellé"
'Default Property Values:
Const m_def_SelectAll = True
'Property Variables:
Dim m_SelectAll As Boolean



'=================== Propriétés ===================

'-------- Alignment --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=lbl,lbl,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
    Alignment = lbl.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    If New_Alignment < 0 Or New_Alignment > 2 Then
      Exit Property
    End If
    lbl.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'-------- Appearance --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_UserMemId = -520
    Appearance = txt.Appearance
End Property

'-------- AutoSize --------
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_UserMemId = -500
    AutoSize = lbl.AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    lbl.AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    ' Redimensionne le contrôle
    If New_AutoSize Then UserControl_Resize
End Property

'-------- BackColor --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=lbl,lbl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Renvoie ou définit la couleur d'arrière-plan utilisée pour afficher le texte et les graphiques d'un objet."
Attribute BackColor.VB_UserMemId = -501
    BackColor = lbl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lbl.BackColor() = New_BackColor
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'-------- BackStyle --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=lbl,lbl,-1,BackStyle
Public Property Get BackStyle() As BackStyleConstants
Attribute BackStyle.VB_UserMemId = -502
    BackStyle = lbl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BackStyleConstants)
    If New_BackStyle < 0 Or New_BackStyle > 1 Then
      Exit Property
    End If
    UserControl.BackStyle() = New_BackStyle
    lbl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'-------- BorderStyle --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = txt.BorderStyle
End Property

'-------- Caption --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=lbl,lbl,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lbl.Caption() = New_Caption
    PropertyChanged "Caption"
    ' Redimensionne
    UserControl_Resize
End Property

'-------- Enabled --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Renvoie ou définit une valeur qui détermine si un objet peut répondre à des événements générés par l'utilisateur."
Attribute Enabled.VB_UserMemId = -514
    Enabled = txt.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txt.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'-------- Font --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=lbl,lbl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = lbl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lbl.Font = New_Font
    Set UserControl.Font = New_Font ' Pour le calcul de la hauteur dans resize
    PropertyChanged "Font"
    ' Redimensionne le contrôle
    If AutoSize Then UserControl_Resize
End Property

'-------- ForeColor --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=lbl,lbl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Renvoie ou définit la couleur de premier plan utilisée pour afficher le texte et les graphiques d'un objet."
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = lbl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lbl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'-------- HideSelection --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,HideSelection
Public Property Get HideSelection() As Boolean
    HideSelection = txt.HideSelection
End Property

'-------- LabelPosition --------
Public Property Get LabelPosition() As LabelPositionConstants
    LabelPosition = m_LabelPosition
End Property

Public Property Let LabelPosition(ByVal New_LabelPosition As LabelPositionConstants)
    ' N'accepte une modification qu'en création
    If Ambient.UserMode Then
        Err.Raise Number:=33001, Description:="Propriété non modifiable en exécution"
        Exit Property
    End If
    
    m_LabelPosition = New_LabelPosition
    PropertyChanged "LabelPosition"
    ' Redimensionne
    If AutoSize Then lbl.AutoSize = True
    UserControl_Resize
End Property

'-------- LabelWidth --------
Public Property Get LabelWidth() As Single
Attribute LabelWidth.VB_ProcData.VB_Invoke_Property = "General"
    LabelWidth = m_LabelWidth
End Property

Public Property Let LabelWidth(ByVal New_LabelWidth As Single)
    m_LabelWidth = New_LabelWidth
    If m_LabelWidth < 0 Then m_LabelWidth = 0
    PropertyChanged "LabelWidth"
    ' Retire AutoSize
    If lbl.AutoSize = True Then
        AutoSize = False
    End If
    ' Transmet au label
    lbl.width = m_LabelWidth
    ' et redimensionne
    UserControl_Resize
    ' Gérère l'événement
    RaiseEvent LabelWidthChanged
End Property

'-------- Locked --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,Locked
Public Property Get Locked() As Boolean
    Locked = txt.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txt.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'-------- MaxLength --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,MaxLength
Public Property Get MaxLength() As Long
    MaxLength = txt.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txt.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'-------- MousePointer --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = txt.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    txt.MousePointer() = New_MousePointer
    lbl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'-------- MouseIcon --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = txt.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set txt.MouseIcon = New_MouseIcon
    Set lbl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'-------- MultiLine --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,MultiLine
Public Property Get MultiLine() As Boolean
    MultiLine = txt.MultiLine
End Property

'-------- OLEDragMode --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,OLEDragMode
Public Property Get OLEDragMode() As OLEDragConstants
    OLEDragMode = txt.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As OLEDragConstants)
    txt.OLEDragMode() = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

'-------- OLEDropMode --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
    OLEDropMode = txt.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    txt.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'-------- PasswordChar --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,PasswordChar
Public Property Get PasswordChar() As String
    PasswordChar = txt.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txt.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'-------- ScrollBars --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,ScrollBars
Public Property Get ScrollBars() As Integer
    ScrollBars = txt.ScrollBars
End Property

'-------- SelLength --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,SelLength
Public Property Get SelLength() As Long
    SelLength = txt.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txt.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'-------- SelStart --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,SelStart
Public Property Get SelStart() As Long
    SelStart = txt.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txt.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'-------- SelText --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,SelText
Public Property Get SelText() As String
    SelText = txt.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txt.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'-------- Text --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,Text
Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = txt.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txt.Text() = New_Text
    PropertyChanged "Text"
End Property

'-------- TextBackColor --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,BackColor
Public Property Get TextBackColor() As OLE_COLOR
    TextBackColor = txt.BackColor
End Property

Public Property Let TextBackColor(ByVal New_TextBackColor As OLE_COLOR)
    txt.BackColor() = New_TextBackColor
    PropertyChanged "TextBackColor"
End Property

'-------- TextFont --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,Font
Public Property Get TextFont() As Font
    Set TextFont = txt.Font
End Property

Public Property Set TextFont(ByVal New_TextFont As Font)
    Set txt.Font = New_TextFont
    PropertyChanged "TextFont"
End Property

'-------- TextForeColor --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,ForeColor
Public Property Get TextForeColor() As OLE_COLOR
    TextForeColor = txt.ForeColor
End Property

Public Property Let TextForeColor(ByVal New_TextForeColor As OLE_COLOR)
    txt.ForeColor() = New_TextForeColor
    PropertyChanged "TextForeColor"
End Property

'-------- ToolTipText --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,ToolTipText
Public Property Get ToolTipText() As String
    ToolTipText = txt.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    txt.ToolTipText() = New_ToolTipText
    lbl.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'-------- UseMnemonic --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=lbl,lbl,-1,UseMnemonic
Public Property Get UseMnemonic() As Boolean
    UseMnemonic = lbl.UseMnemonic
End Property

Public Property Let UseMnemonic(ByVal New_UseMnemonic As Boolean)
    lbl.UseMnemonic() = New_UseMnemonic
    PropertyChanged "UseMnemonic"
End Property

'-------- WhatsThisHelpID --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,WhatsThisHelpID
Public Property Get WhatsThisHelpID() As Long
    WhatsThisHelpID = txt.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
    txt.WhatsThisHelpID() = New_WhatsThisHelpID
    lbl.WhatsThisHelpID() = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property

'-------- WordWrap --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=lbl,lbl,-1,WordWrap
Public Property Get WordWrap() As Boolean
    WordWrap = lbl.WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    lbl.WordWrap() = New_WordWrap
    PropertyChanged "WordWrap"
End Property

Private Sub txt_GotFocus()
    If m_SelectAll = True Then
        On Error Resume Next
        txt.SelStart = 0
        txt.SelLength = 300
    End If
End Sub
'-------- InitProperties --------
Private Sub UserControl_InitProperties()
    m_LabelWidth = lbl.width
    m_LabelPosition = ltPositionLeft
    ' Police par défaut
    Set Font = Ambient.Font
    Set TextFont = Ambient.Font
    m_SelectAll = m_def_SelectAll
End Sub

'-------- ReadProperties --------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lbl.Alignment = PropBag.ReadProperty("Alignment", vbLeftJustify)
    lbl.AutoSize = PropBag.ReadProperty("AutoSize", False)
    lbl.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    lbl.BackStyle = PropBag.ReadProperty("BackStyle", ltBSOpaque)
    lbl.Caption = PropBag.ReadProperty("Caption", "Label1")
    lbl.ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
    txt.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_LabelPosition = PropBag.ReadProperty("LabelPosition", m_def_LabelPosition)
    m_LabelWidth = PropBag.ReadProperty("LabelWidth", m_def_LabelWidth)
    lbl.width = m_LabelWidth
    txt.Locked = PropBag.ReadProperty("Locked", False)
    txt.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txt.MousePointer = PropBag.ReadProperty("MousePointer", vbArrow)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txt.OLEDragMode = PropBag.ReadProperty("OLEDragMode", vbOLEDragManual)
    txt.OLEDropMode = PropBag.ReadProperty("OLEDropMode", vbOLEDropNone)
    txt.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    txt.SelLength = PropBag.ReadProperty("SelLength", 0)
    txt.SelStart = PropBag.ReadProperty("SelStart", 0)
    txt.SelText = PropBag.ReadProperty("SelText", "")
    txt.Text = PropBag.ReadProperty("Text", "Text1")
    txt.BackColor = PropBag.ReadProperty("TextBackColor", vbWindowBackground)
    Set txt.Font = PropBag.ReadProperty("TextFont", Ambient.Font)
    txt.ForeColor = PropBag.ReadProperty("TextForeColor", vbWindowText)
    txt.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    lbl.UseMnemonic = PropBag.ReadProperty("UseMnemonic", True)
    txt.WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
    lbl.WordWrap = PropBag.ReadProperty("WordWrap", False)
    m_SelectAll = PropBag.ReadProperty("SelectAll", m_def_SelectAll)
    txt.Alignment = PropBag.ReadProperty("TextAlignement", 0)
End Sub

'-------- WriteProperties --------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", lbl.Alignment, vbLeftJustify)
    Call PropBag.WriteProperty("AutoSize", lbl.AutoSize, False)
    Call PropBag.WriteProperty("BackColor", lbl.BackColor, vbButtonFace)
    Call PropBag.WriteProperty("BackStyle", lbl.BackStyle, ltBSOpaque)
    Call PropBag.WriteProperty("Caption", lbl.Caption, "Label1")
    Call PropBag.WriteProperty("ForeColor", lbl.ForeColor, vbButtonText)
    Call PropBag.WriteProperty("Enabled", txt.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("LabelPosition", m_LabelPosition, m_def_LabelPosition)
    Call PropBag.WriteProperty("LabelWidth", m_LabelWidth, m_def_LabelWidth)
    Call PropBag.WriteProperty("Locked", txt.Locked, False)
    Call PropBag.WriteProperty("MaxLength", txt.MaxLength, 0)
    Call PropBag.WriteProperty("MousePointer", txt.MousePointer, vbArrow)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("OLEDragMode", txt.OLEDragMode, vbOLEDragManual)
    Call PropBag.WriteProperty("OLEDropMode", txt.OLEDropMode, vbOLEDropNone)
    Call PropBag.WriteProperty("PasswordChar", txt.PasswordChar, "")
    Call PropBag.WriteProperty("SelLength", txt.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txt.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txt.SelText, "")
    Call PropBag.WriteProperty("Text", txt.Text, "Text1")
    Call PropBag.WriteProperty("TextBackColor", txt.BackColor, vbWindowBackground)
    Call PropBag.WriteProperty("TextFont", Font, Ambient.Font)
    Call PropBag.WriteProperty("TextForeColor", txt.ForeColor, vbWindowText)
    Call PropBag.WriteProperty("ToolTipText", txt.ToolTipText, "")
    Call PropBag.WriteProperty("UseMnemonic", lbl.UseMnemonic, True)
    Call PropBag.WriteProperty("WhatsThisHelpID", txt.WhatsThisHelpID, 0)
    Call PropBag.WriteProperty("WordWrap", lbl.WordWrap, False)
    Call PropBag.WriteProperty("SelectAll", m_SelectAll, m_def_SelectAll)
    Call PropBag.WriteProperty("TextAlignement", txt.Alignment, 0)
End Sub

'=================== Evénements ===================

'-------- Change --------
Private Sub txt_Change()
    RaiseEvent Change
End Sub

'-------- Click --------
Private Sub txt_Click()
    RaiseEvent Click
End Sub

'-------- DblClick --------
Private Sub txt_DblClick()
    RaiseEvent DblClick
End Sub

'-------- KeyDown --------
Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'-------- KeyPress --------
Private Sub txt_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'-------- KeyUp --------
Private Sub txt_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'-------- MouseDown --------
Private Sub txt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, _
        ScaleX(X + txt.left, vbTwips, vbContainerPosition), _
        ScaleY(Y + txt.top, vbTwips, vbContainerPosition))
End Sub

'-------- MouseMove --------
Private Sub txt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, _
        ScaleX(X + txt.left, vbTwips, vbContainerPosition), _
        ScaleY(Y + txt.top, vbTwips, vbContainerPosition))
End Sub

'-------- MouseUp --------
Private Sub txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, _
        ScaleX(X + txt.left, vbTwips, vbContainerPosition), _
        ScaleY(Y + txt.top, vbTwips, vbContainerPosition))
End Sub

'-------- OLECompleteDrag --------
Private Sub txt_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'-------- OLEDragDrop --------
Private Sub txt_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

'-------- OLEDragOver --------
Private Sub txt_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

'-------- OLEGiveFeedback --------
Private Sub txt_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

'-------- OLESetData --------
Private Sub txt_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

'-------- OLEStartDrag --------
Private Sub txt_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

'=================== Méthodes ===================

'-------- OLEDrag --------
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=txt,txt,-1,OLEDrag
Public Sub OLEDrag()
    txt.OLEDrag
End Sub

'=================== Autres ===================

'---- Changement de taille du contrôle
Private Sub UserControl_Resize()
    ' Largeur du libellé
    Dim w As Integer
    w = lbl.width
    ' Hauteur du texte du libellé
    Dim h As Integer
    h = TextHeight(lbl.Caption)
    
    ' Positionne selon le cas
    Select Case LabelPosition
        Case ltPositionLeft
            ' Positionne libellé
            lbl.Move 0, Screen.TwipsPerPixelY * 3, w, height
            ' et le texte
            If width < w Then
                ' Trop petit, cache le texte
                txt.Visible = False
            Else
                ' OK, positionne le texte
                txt.Visible = True
                txt.Move w, 0, width - w, height
            End If
            
        Case ltPositionRight
            ' Positionne libellé
            lbl.Move width - w, Screen.TwipsPerPixelY * 3, w, height
            ' et le texte
            If width < w Then
                ' Trop petit, cache le texte
                txt.Visible = False
            Else
                ' OK, positionne le texte
                txt.Visible = True
                txt.Move 0, 0, width - w, height
            End If
    
        Case ltPositionTop
            ' Positionne libellé
            lbl.Move Screen.TwipsPerPixelX * 3, 0, width, h
            ' et le texte
            If height < h Then
                ' Trop petit, cache le texte
                txt.Visible = False
            Else
                ' OK, positionne le texte
                txt.Visible = True
                txt.Move 0, h, width, height - h
            End If
    
        Case ltPositionBottom
            ' Positionne libellé
            lbl.Move Screen.TwipsPerPixelX * 3, height - h, width, h
            ' et le texte
            If height < h Then
                ' Trop petit, cache le texte
                txt.Visible = False
            Else
                ' OK, positionne le texte
                txt.Visible = True
                txt.Move 0, 0, width, height - h
            End If
    End Select
    
    ' Modifie LabelWidth
    m_LabelWidth = lbl.width
    
    ' Génère l'événement
    RaiseEvent Resize
End Sub

'---- Traitement du label en mode modification
Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Seulement si en mode création
    If Ambient.UserMode = False Then
        ' Si click près de la limite entre label et texte
        If lbl.width - X < 5 * Screen.TwipsPerPixelX Then
            ' Initialise tracking
            fTracking = True
            oldLabelWidth = LabelWidth
            oldX = X
        End If
    End If
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Seulement si en mode création
    If Ambient.UserMode = False Then
        If fTracking Or lbl.width - X < 5 * Screen.TwipsPerPixelX Then
            Screen.MousePointer = vbSizeWE
        Else
            Screen.MousePointer = vbArrow
        End If
        
        If fTracking Then
            ' Nouvelle largeur
            LabelWidth = oldLabelWidth + X - oldX
        End If
    End If
End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Seulement si en mode création
    If Ambient.UserMode = False Then
        ' Termine tracking
        fTracking = False
        Screen.MousePointer = vbArrow
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get SelectAll() As Boolean
    SelectAll = m_SelectAll
End Property

Public Property Let SelectAll(ByVal New_SelectAll As Boolean)
    m_SelectAll = New_SelectAll
    PropertyChanged "SelectAll"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txt,txt,-1,Alignment
Public Property Get TextAlignement() As AlignmentConstants
    TextAlignement = txt.Alignment
End Property

Public Property Let TextAlignement(ByVal New_TextAlignement As AlignmentConstants)
    txt.Alignment() = New_TextAlignement
    PropertyChanged "TextAlignement"
End Property

