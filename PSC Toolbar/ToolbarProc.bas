Attribute VB_Name = "mdlBandProc"
Option Explicit

Function GetText(ByVal hwnd As Long) As String
   
   GetText = Space$(GetWindowTextLength(hwnd))
   GetWindowText hwnd, GetText, Len(GetText) + 1

End Function
Public Function IInputObject_HasFocusIO(ByVal This As IInputObject) As Long
Dim oCallback As IInputObjectCallback

   Set oCallback = This
   
   IInputObject_HasFocusIO = oCallback.HasFocus()

End Function


Private Function MessageProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lPrevProc As Long
Dim oObj As IInputObjectCallback
   
   ' Get the previous window procedure
   lPrevProc = GetProp(hwnd, PROP_PREVPROC)
   
   Select Case wMsg
      
      Case WM_COMMAND
      
         ' Get the callback
         Set oObj = PtrToObj(GetProp(hwnd, PROP_OBJECT))
         
         Select Case wParam \ &H10000
            Case EN_KILLFOCUS
               oObj.OnFocus False
            Case EN_SETFOCUS
               oObj.OnFocus True
            Case Is < &H100
               oObj.ButtonClicked wParam
         End Select
         
         
      Case Else
   
         MessageProc = CallWindowProc(lPrevProc, hwnd, wMsg, wParam, lParam)
         
   End Select
   
End Function

Function IInputObject_TranslateAcceleratorIO(ByVal This As IInputObject, lpmsg As olelib.MSG) As Long
Dim oCallback As IInputObjectCallback

   Set oCallback = This
   IInputObject_TranslateAcceleratorIO = oCallback.TranslateAccelerator(lpmsg)
   
End Function

Private Function PtrToObj(ByVal lPtr As Long) As IInputObjectCallback
Dim oUnk As IInputObjectCallback

   MoveMemory oUnk, lPtr, 4&
   Set PtrToObj = oUnk
   MoveMemory oUnk, 0&, 4&
            
End Function


Public Sub SubClass(ByVal hwnd As Long, ByVal Obj As IInputObjectCallback)

   ' Set the properties
   SetProp hwnd, PROP_OBJECT, ObjPtr(Obj)
   SetProp hwnd, PROP_PREVPROC, GetWindowLong(hwnd, GWL_WNDPROC)
   
   ' Subclass the windows
   SetWindowLong hwnd, GWL_WNDPROC, AddressOf MessageProc
   
End Sub


Public Sub UnsubClass(ByVal hwnd As Long)
Dim lProc As Long

   ' Get the window procedure
   lProc = GetProp(hwnd, PROP_PREVPROC)
   
   ' Unsubclass the window
   SetWindowLong hwnd, GWL_WNDPROC, lProc
   
   ' Remove the properties
   RemoveProp hwnd, PROP_OBJECT
   RemoveProp hwnd, PROP_PREVPROC

End Sub

'
' ReplaceVTableEntry
' ==================
'
' Replaces an entry in a object v-table
'
Public Function ReplaceVTableEntry(ByVal oObject As Long, ByVal nEntry As Integer, ByVal pFunc As Long) As Long
Dim pFuncOld As Long, pVTableHead As Long
Dim pFuncTmp As Long, lOldProtect As Long
     
    ' Object pointer contains a pointer to v-table--copy it to temporary
    ' pVTableHead = *oObject;
    MoveMemory pVTableHead, ByVal oObject, 4

    ' Calculate pointer to specified entry
    pFuncTmp = pVTableHead + (nEntry - 1) * 4
    
    ' Save address of previous method for return
    ' pFuncOld = *pFuncTmp;
    MoveMemory pFuncOld, ByVal pFuncTmp, 4
    
    ' Ignore if they're already the same
    If pFuncOld <> pFunc Then
        ' Need to change page protection to write to code
        VirtualProtect ByVal pFuncTmp, 4, PAGE_EXECUTE_READWRITE, lOldProtect
        
        ' Write the new function address into the v-table
        MoveMemory ByVal pFuncTmp, pFunc, 4     ' *pFuncTmp = pfunc;
        
        ' Restore the previous page protection
        VirtualProtect ByVal pFuncTmp, 4, lOldProtect, lOldProtect 'Optional
        
    End If
    
    'return address of original proc
    ReplaceVTableEntry = pFuncOld
    
End Function


