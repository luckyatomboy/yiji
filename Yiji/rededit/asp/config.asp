<%
Dim sUsername, sPassword
sUsername = "admin"
sPassword = "admin"

Dim aStyle()
Redim aStyle(11)
aStyle(1) = "gray|||gray|||office|||uploadfile/|||550|||350|||rar|zip|exe|doc|xls|txt|chm|hlp|||swf|||gif|jpg|jpeg|bmp|||rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov|||gif|jpg|jpeg|bmp|||500|||100|||10000|||100|||100|||1|||1|||EDIT|||1|||0|||0|||||||||1|||0|||Office标准风格，部分常用按钮，标准适合界面宽度|||1|||zh-cn|||0|||500|||300|||1|||版权所有|||FF0000|||12|||宋体||||||0|||jpg|jpeg|||300|||FFFFFF|||1"
aStyle(2) = "full|||blue|||coolblue|||uploadfile/|||550|||400|||rar|zip|exe|doc|xls|txt|chm|hlp|||swf|||gif|jpg|jpeg|bmp|||rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov|||gif|jpg|jpeg|bmp|||500|||100|||100|||100|||100|||1|||1|||EDIT|||1|||0|||0|||||||||1|||0|||酷蓝样式，全部功能按钮|||1|||zh-cn|||0|||500|||300|||0|||版权所有...|||FF0000|||12|||宋体||||||0|||jpg|jpeg|||300|||FFFFFF|||1"
aStyle(3) = "light|||gray|||light|||uploadfile/|||550|||350|||rar|zip|exe|doc|xls|txt|chm|hlp|||swf|||gif|jpg|jpeg|bmp|||rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov|||gif|jpg|jpeg|bmp|||500|||100|||100|||100|||100|||1|||1|||EDIT|||1|||0|||0|||||||||1|||0|||Office标准风格按钮+淡色，部分常用按钮，标准适合界面宽度|||1|||zh-cn|||0|||500|||300|||0|||版权所有...|||FF0000|||12|||宋体||||||0|||jpg|jpeg|||300|||FFFFFF|||1"
aStyle(4) = "blue|||blue|||blue|||../uploadfile/|||550|||350|||rar|zip|exe|doc|txt|xls|chm|hlp|||swf|||gif|jpg|jpeg|bmp|||rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov|||gif|jpg|jpeg|bmp|||204800|||100|||102400|||100|||100|||1|||1|||EDIT|||1|||0|||3|||||||||1|||1|||Office标准风格按钮+蓝色，部分常用按钮，标准适合界面宽度|||1|||zh-cn|||1|||500|||300|||1|||des.yi.org|||FF0000|||12|||宋体||||||0|||jpg|jpeg|||300|||FFFFFF|||1"
aStyle(5) = "green|||gray|||green|||uploadfile/|||550|||350|||rar|zip|exe|doc|txt|xls|chm|hlp|||swf|||gif|jpg|jpeg|bmp|||rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov|||gif|jpg|jpeg|bmp|||500|||100|||100|||100|||100|||1|||1|||EDIT|||1|||0|||0|||||||||1|||0|||Office标准风格按钮+绿色，部分常用按钮，标准适合界面宽度|||1|||zh-cn|||0|||500|||300|||0|||版权所有...|||FF0000|||12|||宋体||||||0|||jpg|jpeg|||300|||FFFFFF|||1"
aStyle(6) = "red|||gray|||red|||uploadfile/|||550|||350|||rar|zip|exe|doc|xls|txt|chm|hlp|||swf|||gif|jpg|jpeg|bmp|||rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov|||gif|jpg|jpeg|bmp|||500|||100|||100|||100|||100|||1|||1|||EDIT|||1|||0|||0|||||||||1|||0|||Office标准风格按钮+红色，部分常用按钮，标准适合界面宽度|||1|||zh-cn|||0|||500|||300|||0|||版权所有...|||FF0000|||12|||宋体||||||0|||jpg|jpeg|||300|||FFFFFF|||1"
aStyle(7) = "yellow|||gray|||yellow|||uploadfile/|||550|||350|||rar|zip|exe|doc|xls|txt|chm|hlp|||swf|||gif|jpg|jpeg|bmp|||rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov|||gif|jpg|jpeg|bmp|||500|||100|||100|||100|||100|||1|||1|||EDIT|||1|||0|||0|||||||||1|||0|||Office标准风格按钮+黄色，部分常用按钮，标准适合界面宽度|||1|||zh-cn|||0|||500|||300|||0|||版权所有...|||FF0000|||12|||宋体||||||0|||jpg|jpeg|||300|||FFFFFF|||1"
aStyle(8) = "3d|||gray|||office3d|||uploadfile/|||550|||350|||rar|zip|exe|doc|xls|txt|chm|hlp|||swf|||gif|jpg|jpeg|bmp|||rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov|||gif|jpg|jpeg|bmp|||500|||100|||100|||100|||100|||1|||1|||EDIT|||1|||0|||0|||||||||1|||0|||Office标准风格3D凹凸按钮，部分常用按钮，标准适合界面宽度|||1|||zh-cn|||0|||500|||300|||0|||版权所有...|||FF0000|||12|||宋体||||||0|||jpg|jpeg|||300|||FFFFFF|||1"
aStyle(9) = "coolblue|||blue|||coolblue|||uploadfile/|||550|||350|||rar|zip|exe|doc|xls|txt|chm|hlp|||swf|||gif|jpg|jpeg|bmp|||rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov|||gif|jpg|jpeg|bmp|||500|||100|||100|||100|||100|||1|||1|||EDIT|||1|||0|||0|||||||||1|||0|||COOL界面，蓝色主调，标准风格，部分常用按钮，标准适合界面宽度，默认样式|||1|||zh-cn|||0|||500|||300|||0|||版权所有...|||000000|||12|||宋体||||||0|||jpg|jpeg|||300|||FFFFFF|||1"
aStyle(10) = "mini|||blue|||coolblue|||uploadfile/|||550|||350|||rar|zip|exe|doc|xls|txt|chm|hlp|||swf|||gif|jpg|jpeg|bmp|||rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov|||gif|jpg|jpeg|bmp|||500|||100|||100|||100|||100|||1|||1|||EDIT|||1|||0|||0|||||||||1|||0|||mini全菜单风格，全部功能按钮，工具栏占位小，标准界面宽度|||1|||zh-cn|||0|||500|||300|||0|||版权所有...|||FF0000|||12|||宋体||||||0|||jpg|jpeg|||300|||FFFFFF|||1"
aStyle(11) = "popup|||blue|||coolblue|||uploadfile/|||550|||350|||rar|zip|exe|doc|xls|txt|chm|hlp|||swf|||gif|jpg|jpeg|bmp|||rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov|||gif|jpg|jpeg|bmp|||500|||100|||100|||100|||100|||1|||1|||EDIT|||1|||0|||0|||||||||1|||0|||酷蓝样式，主要用于标准弹窗，增加弹窗返回按钮。|||1|||zh-cn|||0|||500|||300|||0|||版权所有...|||FF0000|||12|||宋体||||||0|||jpg|jpeg|||300|||FFFFFF|||1"

Dim aToolbar()
Redim aToolbar(33)
aToolbar(1) = "1|||TBHandle|FormatBlock|FontName|FontSize|ZoomSelect|Bold|Italic|UnderLine|StrikeThrough|SuperScript|SubScript|TBSep|JustifyLeft|JustifyCenter|JustifyRight|JustifyFull|||格式工具栏|||1"
aToolbar(2) = "1|||TBHandle|Cut|Copy|Paste|PasteText|PasteWord|FindReplace|Delete|RemoveFormat|TBSep|UnDo|ReDo|SelectAll|UnSelect|TBSep|OrderedList|UnOrderedList|Indent|Outdent|ParagraphAttr|TBSep|ForeColor|BgColor|BackColor|BackImage|TBSep|absolutePosition|zIndexForward|zIndexBackward|||常用工具栏|||2"
aToolbar(3) = "1|||TBHandle|Image|Flash|Media|File|TBSep|TableMenu|FormMenu|ShowBorders|TBSep|Fieldset|Iframe|HorizontalRule|Marquee|CreateLink|Anchor|Map|Unlink|TBSep|Symbol|Emot|Excel|Art|NowDate|NowTime|Quote|TBSep|Maximize|About|||高级工具栏|||3"
aToolbar(4) = "2|||TBHandle|FormatBlock|FontName|FontSize|ZoomSelect|Bold|Italic|UnderLine|StrikeThrough|SuperScript|SubScript|TBSep|JustifyLeft|JustifyCenter|JustifyRight|JustifyFull|||格式工具栏|||1"
aToolbar(5) = "2|||TBHandle|Cut|Copy|Paste|PasteText|PasteWord|FindReplace|Delete|RemoveFormat|TBSep|UnDo|ReDo|SelectAll|UnSelect|TBSep|OrderedList|UnOrderedList|Indent|Outdent|ParagraphAttr|TBSep|ForeColor|BgColor|BackColor|BackImage|TBSep|absolutePosition|zIndexForward|zIndexBackward|||常用工具栏|||2"
aToolbar(6) = "2|||TBHandle|Image|Flash|Media|File|RemoteUpload|TBSep|TableMenu|FormMenu|ShowBorders|TBSep|Fieldset|Iframe|HorizontalRule|Marquee|CreateLink|Anchor|Map|Unlink|TBSep|Symbol|Emot|Excel|Art|NowDate|NowTime|TBSep|Quote|Code|||高级工具栏|||3"
aToolbar(7) = "2|||TBHandle|Paragraph|BR|TBSep|TableInsert|TableProp|TBSep|TableCellProp|TableCellSplit|TBSep|TableRowProp|TableRowInsertAbove|TableRowInsertBelow|TableRowMerge|TableRowSplit|TableRowDelete|TBSep|TableColInsertLeft|TableColInsertRight|TableColMerge|TableColSplit|TableColDelete|TBSep|FormText|FormTextArea|FormRadio|FormCheckbox|FormDropdown|FormButton|||表格与表单|||4"
aToolbar(8) = "2|||TBHandle|EditMenu|FontMenu|ParagraphMenu|ComponentMenu|ObjectMenu|ToolMenu|TableMenu|FormMenu|FileMenu|ZoomMenu|TBSep|Refresh|Save|TBSep|Print|TBSep|ModeCode|ModeEdit|ModeText|ModeView|TBSep|SizePlus|SizeMinus|TBSep|Maximize|TBSep|EQ|TBSep|About|Site|||菜单与状态|||5"
aToolbar(9) = "3|||TBHandle|FormatBlock|FontName|FontSize|ZoomSelect|Bold|Italic|UnderLine|StrikeThrough|SuperScript|SubScript|TBSep|JustifyLeft|JustifyCenter|JustifyRight|JustifyFull|||格式工具栏|||1"
aToolbar(10) = "3|||TBHandle|Cut|Copy|Paste|PasteText|PasteWord|FindReplace|Delete|RemoveFormat|TBSep|UnDo|ReDo|SelectAll|UnSelect|TBSep|OrderedList|UnOrderedList|Indent|Outdent|ParagraphAttr|TBSep|ForeColor|BgColor|BackColor|BackImage|TBSep|absolutePosition|zIndexForward|zIndexBackward|||常用工具栏|||2"
aToolbar(11) = "3|||TBHandle|Image|Flash|Media|File|TBSep|TableMenu|FormMenu|ShowBorders|TBSep|Fieldset|Iframe|HorizontalRule|Marquee|CreateLink|Anchor|Map|Unlink|TBSep|Symbol|Emot|Excel|Art|NowDate|NowTime|Quote|TBSep|Maximize|About|||高级工具栏|||3"
aToolbar(12) = "4|||TBHandle|FormatBlock|FontName|FontSize|ZoomSelect|Bold|Italic|UnderLine|StrikeThrough|SuperScript|SubScript|TBSep|JustifyLeft|JustifyCenter|JustifyRight|JustifyFull|||格式工具栏|||1"
aToolbar(13) = "4|||TBHandle|Cut|Copy|Paste|PasteText|PasteWord|FindReplace|Delete|RemoveFormat|TBSep|UnDo|ReDo|SelectAll|UnSelect|TBSep|OrderedList|UnOrderedList|Indent|Outdent|ParagraphAttr|TBSep|ForeColor|BgColor|BackColor|BackImage|TBSep|absolutePosition|zIndexForward|zIndexBackward|||常用工具栏|||2"
aToolbar(14) = "4|||TBHandle|Image|Flash|Media|File|TBSep|TableMenu|FormMenu|ShowBorders|TBSep|Fieldset|Iframe|HorizontalRule|Marquee|CreateLink|Anchor|Map|Unlink|TBSep|Symbol|Emot|Art|NowDate|NowTime|Quote|TBSep|Maximize|LocalUpload|ImportWord|RemoteUpload|||高级工具栏|||3"
aToolbar(15) = "5|||TBHandle|FormatBlock|FontName|FontSize|ZoomSelect|Bold|Italic|UnderLine|StrikeThrough|SuperScript|SubScript|TBSep|JustifyLeft|JustifyCenter|JustifyRight|JustifyFull|||格式工具栏|||1"
aToolbar(16) = "5|||TBHandle|Cut|Copy|Paste|PasteText|PasteWord|FindReplace|Delete|RemoveFormat|TBSep|UnDo|ReDo|SelectAll|UnSelect|TBSep|OrderedList|UnOrderedList|Indent|Outdent|ParagraphAttr|TBSep|ForeColor|BgColor|BackColor|BackImage|TBSep|absolutePosition|zIndexForward|zIndexBackward|||常用工具栏|||2"
aToolbar(17) = "5|||TBHandle|Image|Flash|Media|File|TBSep|TableMenu|FormMenu|ShowBorders|TBSep|Fieldset|Iframe|HorizontalRule|Marquee|CreateLink|Anchor|Map|Unlink|TBSep|Symbol|Emot|Excel|Art|NowDate|NowTime|Quote|TBSep|Maximize|About|||高级工具栏|||3"
aToolbar(18) = "6|||TBHandle|FormatBlock|FontName|FontSize|ZoomSelect|Bold|Italic|UnderLine|StrikeThrough|SuperScript|SubScript|TBSep|JustifyLeft|JustifyCenter|JustifyRight|JustifyFull|||格式工具栏|||1"
aToolbar(19) = "6|||TBHandle|Cut|Copy|Paste|PasteText|PasteWord|FindReplace|Delete|RemoveFormat|TBSep|UnDo|ReDo|SelectAll|UnSelect|TBSep|OrderedList|UnOrderedList|Indent|Outdent|ParagraphAttr|TBSep|ForeColor|BgColor|BackColor|BackImage|TBSep|absolutePosition|zIndexForward|zIndexBackward|||常用工具栏|||2"
aToolbar(20) = "6|||TBHandle|Image|Flash|Media|File|TBSep|TableMenu|FormMenu|ShowBorders|TBSep|Fieldset|Iframe|HorizontalRule|Marquee|CreateLink|Anchor|Map|Unlink|TBSep|Symbol|Emot|Excel|Art|NowDate|NowTime|Quote|TBSep|Maximize|About|||高级工具栏|||3"
aToolbar(21) = "7|||TBHandle|FormatBlock|FontName|FontSize|ZoomSelect|Bold|Italic|UnderLine|StrikeThrough|SuperScript|SubScript|TBSep|JustifyLeft|JustifyCenter|JustifyRight|JustifyFull|||格式工具栏|||1"
aToolbar(22) = "7|||TBHandle|Cut|Copy|Paste|PasteText|PasteWord|FindReplace|Delete|RemoveFormat|TBSep|UnDo|ReDo|SelectAll|UnSelect|TBSep|OrderedList|UnOrderedList|Indent|Outdent|ParagraphAttr|TBSep|ForeColor|BgColor|BackColor|BackImage|TBSep|absolutePosition|zIndexForward|zIndexBackward|||常用工具栏|||2"
aToolbar(23) = "7|||TBHandle|Image|Flash|Media|File|TBSep|TableMenu|FormMenu|ShowBorders|TBSep|Fieldset|Iframe|HorizontalRule|Marquee|CreateLink|Anchor|Map|Unlink|TBSep|Symbol|Emot|Excel|Art|NowDate|NowTime|Quote|TBSep|Maximize|About|||高级工具栏|||3"
aToolbar(24) = "8|||TBHandle|FormatBlock|FontName|FontSize|ZoomSelect|Bold|Italic|UnderLine|StrikeThrough|SuperScript|SubScript|TBSep|JustifyLeft|JustifyCenter|JustifyRight|JustifyFull|||格式工具栏|||1"
aToolbar(25) = "8|||TBHandle|Cut|Copy|Paste|PasteText|PasteWord|FindReplace|Delete|RemoveFormat|TBSep|UnDo|ReDo|SelectAll|UnSelect|TBSep|OrderedList|UnOrderedList|Indent|Outdent|ParagraphAttr|TBSep|ForeColor|BgColor|BackColor|BackImage|TBSep|absolutePosition|zIndexForward|zIndexBackward|||常用工具栏|||2"
aToolbar(26) = "8|||TBHandle|Image|Flash|Media|File|TBSep|TableMenu|FormMenu|ShowBorders|TBSep|Fieldset|Iframe|HorizontalRule|Marquee|CreateLink|Anchor|Map|Unlink|TBSep|Symbol|Emot|Excel|Art|NowDate|NowTime|Quote|TBSep|Maximize|About|||高级工具栏|||3"
aToolbar(27) = "9|||TBHandle|FormatBlock|FontName|FontSize|ZoomSelect|Bold|Italic|UnderLine|StrikeThrough|SuperScript|SubScript|TBSep|JustifyLeft|JustifyCenter|JustifyRight|JustifyFull|||格式工具栏|||1"
aToolbar(28) = "9|||TBHandle|Cut|Copy|Paste|PasteText|PasteWord|FindReplace|Delete|RemoveFormat|TBSep|UnDo|ReDo|SelectAll|UnSelect|TBSep|OrderedList|UnOrderedList|Indent|Outdent|ParagraphAttr|TBSep|ForeColor|BgColor|BackColor|BackImage|TBSep|absolutePosition|zIndexForward|zIndexBackward|||常用工具栏|||2"
aToolbar(29) = "9|||TBHandle|Image|Flash|Media|File|TBSep|TableMenu|FormMenu|ShowBorders|TBSep|Fieldset|Iframe|HorizontalRule|Marquee|CreateLink|Anchor|Map|Unlink|TBSep|Symbol|Emot|Excel|Art|NowDate|NowTime|Quote|TBSep|Maximize|About|||高级工具栏|||3"
aToolbar(30) = "10|||TBHandle|FormatBlock|FontName|FontSize|ZoomSelect|TBSep|EditMenu|FontMenu|ParagraphMenu|ComponentMenu|ObjectMenu|ToolMenu|FormMenu|TableMenu|FileMenu|Maximize|||mini工具栏|||1"
aToolbar(31) = "11|||TBHandle|FormatBlock|FontName|FontSize|ZoomSelect|Bold|Italic|UnderLine|StrikeThrough|SuperScript|SubScript|TBSep|JustifyLeft|JustifyCenter|JustifyRight|JustifyFull|||格式工具栏|||1"
aToolbar(32) = "11|||TBHandle|Cut|Copy|Paste|PasteText|PasteWord|FindReplace|Delete|RemoveFormat|TBSep|UnDo|ReDo|SelectAll|UnSelect|TBSep|OrderedList|UnOrderedList|Indent|Outdent|ParagraphAttr|TBSep|ForeColor|BgColor|BackColor|BackImage|TBSep|absolutePosition|zIndexForward|zIndexBackward|||常用工具栏|||2"
aToolbar(33) = "11|||TBHandle|Image|Flash|Media|File|TBSep|TableMenu|FormMenu|ShowBorders|TBSep|Fieldset|Iframe|HorizontalRule|Marquee|CreateLink|Anchor|Map|Unlink|TBSep|Symbol|Emot|Excel|Art|NowDate|NowTime|Quote|TBSep|Save|About|||高级工具栏|||3"
%>