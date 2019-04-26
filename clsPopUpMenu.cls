'---------------------------------------------------------------------------------------
' Class Module  : clsPopUpMenu
' Author        : paweld
' Purpose       : Klasa pozwala generowac swoje menu kontekstowe
' Licence       : Free
' Version       : 1.0 beta
'
'---------------------------------------------------------------------------------------
' Call          :
'---------------------------------------------------------------------------------------
'             prawy przycisk myszki na kontrolce dla ktorej zostal zainicjowany
'
'---------------------------------------------------------------------------------------
' Components    :
'---------------------------------------------------------------------------------------
' clsPopUpMenu          - modul klasy
' clsPopUpMenu_Def      - klasa definicja menu
' //dowolny moduł//     - inicjowanie klasy
'
'---------------------------------------------------------------------------------------
' Changes       :
'---------------------------------------------------------------------------------------
' 2019-04-24        , PD : Utworzenie klasy.
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private Const clsPopUpMenuIsON = True
Private Const CurrentModeName = "clsPopUpMenu"

Private Enum enControlTyp
    Buttona = 1
   'ComboBox = 2    'UNDER CONSTRUCTION
   'ListBox = 3     'UNDER CONSTRUCTION
End Enum

Private bBeforNextCtontolAddLine As Boolean

Private Menus As New clsPopUpMenu_Definition

'Private Enum vbaProceduryZlecenYY
'    vbAktualizacjaPriorytetu = 1
'    vbAktualizacjaCiaiaguTechnologicznego = 2
'    vbDodanieCechyCharakterystycznej = 3
'    vbKlonowanieMaszyn = 4
'    vbAktualizacjaRejonu = 5
'End Enum

Private Sub Class_Initialize()
    
    Debug.Print "Konstruktor klasy PopUpMenu (" & VBA.Now & ")"
        'Dim Lst As clsAktualizacjeLista
   
    'Call Lst.UtworzTabliceDanychoAktualizacjach
End Sub

Public Sub INIT()
    Debug.Print "INIT"
End Sub

'          Przypisz Do kontrolki typu przycisk
'Public Sub AssignToTheButton(ctrButton As CommandButton, sNazwaMenu As String, Optional AddShortcutMenuBar As Boolean = True)
'                                                                                   "OnClick"
Public Sub AssignToTheButton(ctrButton As CommandButton, sNazwaMenu As String)
    'Call AssignToControl(
    Debug.Print "UNDER CONSTRUCTION"
End Sub

'Public Sub AssignToThePicture(ctrButton As Image, sNazwaMenu As String, Optional AddShortcutMenuBar As Boolean = True, Optional CallOnClick As Boolean = True)
Public Sub AssignToThePicture(ctrButton As Image, sNazwaMenu As String)
    'Call AssignToControl(
    Debug.Print "UNDER CONSTRUCTION"
End Sub

Public Function AssignToControl(objCtr As Object, sNazwaMenu As String) As Boolean
    Debug.Print "UNDER CONSTRUCTION"
    AssignToControl = False
   'AssignToControl = True
End Function

'Public Function AssignToControl(objCtr As Object, objTyp As enControlTyp, NazwaMenu, OnClick) As Boolean
'    Debug.Print "UNDER CONSTRUCTION"
'End Function


Public Function CreatPopupMenu(sCbName As String, cb As CommandBar) As Boolean
    
    Call UtworzPopupMenu(sCbName, cb)
    
    Exit Function
    '------ BUDOWANIE MENU POPUP ------ nie dziala w klasie ???
    'Sprawdź czy istnieje w kolekcji
    Set cb = Application.CommandBars(sCbName)
    'Jeśli istnieje, usuń, by utworzyć na nowo
    If Not cb Is Nothing Then DeleteCommandBars (sCbName)

    'W razie błędu idź do podprogramu obsługi błędów
    'Utwórz pasek menu popup na nowo
    Set cb = Application.CommandBars.Add(Name:=sCbName, Position:=msoBarPopup)
End Function

Public Function AddButon(cb As CommandBar, sCaption As String, sOnAction As String, Optional iFaceId As Integer, Optional sDescription As String, Optional sTooltipText As String, Optional bVisible As Boolean = True, Optional bTranslate As Boolean = True)
    Dim ctr As CommandBarControl
    
    If bTranslate Then sCaption = Sur.tr(sCaption)
    
    If Len(sDescription) > 0 And bTranslate Then sDescription = Sur.tr(sDescription)
    If Len(sDescription) = 0 Then sDescription = sCaption
    
    If Len(sTooltipText) > 0 And bTranslate Then If bTranslate Then sTooltipText = Sur.tr(sTooltipText)
    If Len(sTooltipText) = 0 Then sTooltipText = sCaption
    
    Set ctr = cb.Controls.Add(msoControlButton)
    With ctr
                                                    ' obsluga dodawani lini
        If bBeforNextCtontolAddLine Then .BeginGroup = True
        If bBeforNextCtontolAddLine Then bBeforNextCtontolAddLine = False

        .Caption = sCaption                         ' tytuł przycisku
        .DescriptionText = sDescription             ' opis
        .TooltipText = sTooltipText                 ' flaga/chmurka/podpowiedz (przy innych typach menu widoczny)
        .Visible = bVisible                         ' widocznosc przycisku
        If iFaceId > 0 Then .FaceId = iFaceId       ' FaceId = obrazek
        .OnAction = sOnAction                       'akcja (procedura) wykonywana przy naciśnięciu na przycisk
        
        .Style = msoButtonIconAndCaption            ' wygląd przycisku: etykieta + obrazek
        
    End With
End Function

Public Function AddLine()
    bBeforNextCtontolAddLine = True
End Function

Public Function AddLabel(cb As CommandBar, sCaption As String, Optional iFaceId As Integer, Optional bTranslate As Boolean = True)
    'tak naprawde to nieaktywny przcisk
    Dim ctr As CommandBarControl
    
    If bTranslate Then sCaption = Sur.tr(sCaption)
   'If Len(sDescription) = 0 Then sDescription = sCaption
   'If Len(sTooltipText) = 0 Then sTooltipText = sCaption
    
    Set ctr = cb.Controls.Add(msoControlButton)
    With ctr
                                                    ' obsluga dodawani lini
        If bBeforNextCtontolAddLine Then .BeginGroup = True
        If bBeforNextCtontolAddLine Then bBeforNextCtontolAddLine = False
        .Enabled = False                            ' deaktywowanie ikony i opisu zeby imitowalo naglowek
        .Caption = sCaption                         ' tresc etykiety
        .TooltipText = .Caption                     ' flaga/chmurka/podpowiedz
        If iFaceId > 0 Then .FaceId = iFaceId       ' FaceId = obrazek
        
        .Style = msoButtonIconAndCaption            ' wygląd przycisku: etykieta + obrazek
    End With
End Function

Private Sub DeleteCommandBars(sCbToDelete As String)
    On Error Resume Next
    Application.CommandBars(sCbToDelete).DELETE
End Sub
