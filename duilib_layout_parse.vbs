'------------------------------------------------
' duilib_layout_parse.vbs
' 用于生成布局文件的按钮响应，拷贝到 cpp 源文件
' https://www.douban.com/people/wyrover/
' wyrover@gmail.com
Const COM_FSO           = "Scripting.FileSystemObject"

'------------------------------------------------
' FSO
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
'------------------------------------------------
' CScriptRun
Sub CScriptRun 
    Dim Args
    Dim Arg
    If LCase(Right(WScript.FullName,11)) = "wscript.exe" Then
        Args = Array("cmd.exe /k CScript.exe", """" & WScript.ScriptFullName & """" )
            For Each Arg In WScript.Arguments
            ReDim Preserve Args(UBound(Args)+1)
            Args(UBound(Args)) = """" & Arg & """"
        Next
        WScript.Quit CreateObject("WScript.Shell").Run(Join(Args), 1, True)
    End If
End Sub

'------------------------------------------------
' 打印字符串
Sub Echo(message)
    WScript.Echo message
End Sub


'------------------------------------------------
' 路径末尾添加\
Function DisposePath(sPath)
    On Error Resume Next
    
    If Right(sPath, 1) = "\" Then
        DisposePath = sPath
    Else
        DisposePath = sPath & "\"
    End If
    
    DisposePath = Trim(DisposePath)
End Function 

'------------------------------------------------
' 获取文件路径
Function GetFilePath(filename)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    GetFilePath = DisposePath(FSO.GetParentFolderName(filename))
End Function 


'------------------------------------------------
' 获取文件绝对路径
Function GetAbsolutePathName(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    GetAbsolutePathName = FSO.GetAbsolutePathName(file)
End Function

'------------------------------------------------
' 获取文件名
Function GetFileName(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    GetFileName = FSO.GetFileName(file)
End Function

'------------------------------------------------
' 获取基本文件名
Function GetBaseName(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    GetBaseName = FSO.GetBaseName(file)
End Function


CScriptRun


Call Main

Sub Main()
    Dim obj_dui_layout_doc : Set obj_dui_layout_doc = New dui_layout_doc
    obj_vc_project_doc.filename = WScript.Arguments(0)
    'obj_dui_layout_doc.filename = "main_frm.xml"
    obj_dui_layout_doc.parse()    
    Set obj_dui_layout_doc = Nothing
End Sub
	

Class dui_layout_doc
    Private filename_
    Private doc_
    Private root_       
    Private outfile_
    Private fso_

    Private Sub Class_Initialize()          
        Set fso_ = CreateObject(COM_FSO)
    End Sub  
  
    Private Sub Class_Terminate()     
        If Not (outfile_ is Nothing) Then
            outfile_.Close
        End If
        Set doc_ = Nothing
        Set fso_ = Nothing
    End Sub  

    Public Property Let filename(value)   
        Dim fso, basename, dir
        Set fso = CreateObject("Scripting.FileSystemObject")
        filename_ = fso.GetAbsolutePathName(value)
        Set fso = Nothing        
        
        Echo filename_
        
        dir = GetFilePath(filename_)
        basename = GetBaseName(filename_)        


        Set doc_ = CreateObject("MSXML2.DOMDocument")
        doc_.async = False
        doc_.load(filename_)
        If doc_.parseError.errorCode = 0 Then            
            Set root_ = doc_.documentElement     
            Set outfile_ = fso_.OpenTextFile(dir & basename & "_out.cpp", ForWriting, True)
        End If
    End Property
    
    Public Property Get filename()
        filename = filename_
    End Property      

    Private Function WriteLine(ByVal data)
        Echo data
        outfile_.WriteLine data
    End Function

    

    Public Function parse()
        write_define
        write_find_control
        write_action
        write_action_impl
    End Function

    Private Function write_define()
        Dim node, buttons, options
        If doc_.parseError.errorCode = 0 Then                    

            WriteLine "//--声明成员变量--"
            '------------------------------------------------
            ' buttons
            Set buttons = root_.selectNodes("//Button")
            If Not (buttons is Nothing) Then		
                For I = 0 To buttons.length-1
                    WriteLine "CButtonUI* " & buttons(I).getAttribute("name") & "_"	
                Next	             
            End If 
            
            '------------------------------------------------
            ' options
            Set options = root_.selectNodes("//Option")
            If Not (options is Nothing) Then		
                For I = 0 To options.length-1
                    WriteLine "COptionUI* " & options(I).getAttribute("name") & "_"	
                Next	
            End If                

        End If
    End Function

    Private Function write_find_control()
        Dim node, buttons, options
        If doc_.parseError.errorCode = 0 Then                    

            WriteLine "//--绑定委托--"
            '------------------------------------------------
            ' buttons
            Set buttons = root_.selectNodes("//Button")
            If Not (buttons is Nothing) Then                
                For I = 0 To buttons.length-1
                    WriteLine buttons(I).getAttribute("name") & "_ = static_cast<CButtonUI*>(m_pm.FindControl(_T(""" & buttons(I).getAttribute("name") & """)));"	
                    WriteLine "if (" & buttons(I).getAttribute("name") & "_) " & buttons(I).getAttribute("name") & "_->OnNotify += MakeDelegate(this, &CXXXXXXDlg::on_" & buttons(I).getAttribute("name") & "_click);"
                    WriteLine ""
                Next
            End If      
            

        End If
    End Function

    Private Function write_action()
        Dim node, buttons, options
        If doc_.parseError.errorCode = 0 Then                    

            WriteLine "//--响应函数--"
            '------------------------------------------------
            ' buttons
            Set buttons = root_.selectNodes("//Button")
            If Not (buttons is Nothing) Then                
                For I = 0 To buttons.length-1
                    WriteLine "bool on_" & buttons(I).getAttribute("name") & "_click(void* param);"	                    
                Next
            End If       
            

        End If
    End Function

    Private Function write_action_impl()
        Dim node, buttons, options
        If doc_.parseError.errorCode = 0 Then                    

            WriteLine "//--响应函数实现--"
            '------------------------------------------------
            ' buttons
            Set buttons = root_.selectNodes("//Button")
            If Not (buttons is Nothing) Then                
                For I = 0 To buttons.length-1
                    WriteLine "bool CXXXXXXDlg::on_" & buttons(I).getAttribute("name") & "_click(void* param)"	                    
                    WriteLine "{"
                    WriteLine "    DuiLib::TNotifyUI* pMsg = (DuiLib::TNotifyUI*)param;"
                    WriteLine "    if (pMsg->sType == _T(""click"")) {"
                    WriteLine "    "
                    WriteLine "    }"
                    WriteLine ""
                    WriteLine "    return true;"
                    WriteLine "}"
                    WriteLine ""
                Next
            End If       
            

        End If
    End Function

    

    
End Class


