# Word
在 Word 中添加自定义按钮的 VBA 代码

代码详解
ribbonXml 变量：此变量包含自定义界面的 XML 代码。
Application.AddCustomUI 方法：此方法将自定义界面加载到 Word 中。第一个参数是自定义 UI 的名称，第二个参数是自定义界面的 XML 代码。
ShowMessageBox 子过程：此子过程在自定义按钮被点击时显示一个消息框。
使用说明
将上述代码复制并粘贴到 Word 的 Visual Basic 编辑器中。
保存并关闭编辑器。
现在，您应该在 Word 的功能区上看到一个名为“自定义组”的新选项卡。该选项卡包含一个名为“单击我”的新按钮。
单击“单击我”按钮，将显示一个消息框。
注意
在 Word 中，您无法隐藏功能区上的某些选项卡，如“开始”选项卡。因此，我们使用了一个名为“HomeTab”的共享控件来禁用“开始”选项卡。这将防止用户在单击自定义按钮时切换到“开始”选项卡。
自定义按钮只能在当前文档中使用。如果您想在所有文档中使用自定义按钮，则需要使用加载项。
