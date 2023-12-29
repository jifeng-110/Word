Sub AddCustomButton()
    ' 在 Word 中添加自定义按钮
    Dim ribbonXml As String
    ribbonXml = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>" & _
                "  <ribbon>" & _
                "    <qat>" & _
                "      <sharedControls>" & _
                "        <button id='mso:HomeTab' label='Home' visible='false' enabled='false' />" & _
                "      </sharedControls>" & _
                "    </qat>" & _
                "    <tabs>" & _
                "      <tab id='mso:TabAddIns'>" & _
                "        <group id='mso:Group1' label='Custom Group'>" & _
                "          <button id='mso:ButtonID' label='Click Me' imageMso='HappyFace' onAction='ShowMessageBox' />" & _
                "        </group>" & _
                "      </tab>" & _
                "    </tabs>" & _
                "  </ribbon>" & _
                "</customUI>"

    ' 加载自定义 UI
    Application.AddCustomUI ("MyRibbonUI"), ribbonXml
End Sub

Sub ShowMessageBox(control As IRibbonControl)
    ' 按钮点击时显示消息框
    MsgBox "Hello, Word!"
End Sub
