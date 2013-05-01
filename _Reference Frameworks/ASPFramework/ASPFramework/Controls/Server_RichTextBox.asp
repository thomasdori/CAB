<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Rich TextBox
'::By Raymond Irving (xwisdom at GotDotNet)

	'Helper function.
	Public Function New_ServerRichTextBox(name) 
		Set New_ServerRichTextBox = New ServerRichTextBox
			New_ServerRichTextBox.Control.Name = name
	End Function

	Public Function New_ServerRichTextBoxEX(name,width,height)	
		Set New_ServerRichTextBoxEX = New ServerTextBox
		New_ServerRichTextBoxEX.Control.Name = name		
		New_ServerRichTextBoxEX.Width=width
		New_ServerRichTextBoxEX.Height=height
	End Function

	'Flag to only allow one richtext script to be loaded
	Public AlreadyLoadRichTextScript : AlreadyLoadRichTextScript= False

	Page.RegisterLibrary "ServerRichTextBox"	
			
	Class ServerRichTextBox
		
		Public Control
		Public Text
		Public Mode '1 Standard (not implemented)
		Public Width
		Public Height
		Public ReadOnly
		Public TextChanged
		Public ToolBarTemplate
		public FooterTemplate
		Public HideToolBar
		Public RaiseOnChanged
		Public DisableImagePreloading
				
		Public BackColor
		Public ButtonColor
		Public DocumentStyle
		Public ImageGalleryURL		'displays customized image gallery page in a popup window
		
		Private mRichState
			
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
				Control.ImplementsOnLoad = True
				Control.ImplementsProcessPostBack = True
				Control.SupportsClientSideEvent   = True
				
			Text    = ""
			Mode    = 1
			ReadOnly = False
			Width 	= 520
			Height 	= 200
			
			BackColor 			= "#d2b48c"
			ButtonColor			= "#efedde"
			DocumentStyle		= ""
			ImageGalleryURL		= ""
			
			Control.CssClass	= ""
			ToolBarTemplate 	= ""
			FooterTemplate 		= ""
			
			TextChanged=False			
			HideToolBar = False
			RaiseOnChanged = False
			DisableImagePreloading = False
			mRichState = (Browser.IsMSIE And Browser.Version>=5.5)
			
		End Sub
		
		Public Function ProcessPostBack()
		
			If  Request.Form.Key(Control.Name+"_Editor_Value") <> "" Then
				Text  = Request.Form(Control.Name+"_Editor_Value")
				If Not Control.ViewState Is Nothing Then
					TextChanged = (Control.ViewState.Read("T")<>Text)
				End If
			Else
				If Not Control.ViewState Is Nothing Then
					Text  = Control.ViewState.Read("T")
				End If
			End If
		
			If RaiseOnChanged  Then				
				If TextChanged Then
					Call Page.RegisterPostBackEventHandler(Me,"OnChanged","")
				End If
			End If
					
		End Function

		Public Function OnLoad()
			If AlreadyLoadRichTextScript Then Exit Function
			Response.Write "<script language=""JavaScript"" src=""" + SCRIPT_LIBRARY_PATH + "richtext/richtextbox.js""></script>"
			'preload toolbar images
			Call PreloadToolBarImages()	
			%>
			<style type="text/css">
				.RichBtn 	{border: #efedde 1px solid;}
				.RichSelect	{font-family: 'Microsoft Sans Serif', Verdana, sans-serif;font-size: 8pt;color: #000000;}
			</style>
			<%
			AlreadyLoadRichTextScript = True
		End Function

		Public Function ReadProperties(bag)			
			Mode	 		= CInt(bag.Read("M"))
			Width	 		= CInt(bag.Read("W"))
			Height	 		= CInt(bag.Read("H"))
			ReadOnly 		= CBool(bag.Read("RO"))
			HideToolbar 	= CBool(bag.Read("TB"))
			DocumentStyle 	= bag.Read("DS")
			ToolBarTemplate = bag.Read("TT")
			FooterTemplate 	= bag.Read("FT")
			DisableImagePreloading = bag.Read("DI")
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "M",Mode
			bag.Write "W",Width
			bag.Write "H",Height
			bag.Write "RO",ReadOnly
			bag.Write "T",Text
			bag.Write "TB",HideToolbar
			bag.Write "DS",DocumentStyle
			bag.Write "TT",ToolBarTemplate
			bag.Write "FT",FooterTemplate
			bag.Write "DI",DisableImagePreloading
		End Function
	   
		Public Function HandleClientEvent(e)
			HandleClientEvent = False	
		End Function					
		
		Public Function SetValueFromDataSource(value)
			Text = value
		End Function

		Public Function SetValue(value) 
			Text = value
		End Function
	   	   
		Public Default Function Render()
			
			Dim varStart	 

			If Control.IsVisible = False Then
				Exit Function
			End If

			varStart = Now

			If Control.TabIndex = 0 Then  'If Not assigned, then autoassign
				Control.TabIndex = Page.GetNextTabIndex()
			End If

			If mRichState Then
				'pure richtextbox
				RenderRichTextBox			
			Else
				'downgraded textbox
				RenderDownGraded
			End If

			Page.TraceRender varStart,Now,Control.Name
		 	 
		End Function
		
		Private Function RenderDownGraded()
		%>
			<table bgcolor="<%=BackColor%>" border="0" cellspacing="0" cellpadding="1"><tr><td>
			<table 
				<% 
					If ButtonColor<>"" Then Response.Write "bgcolor="""+ ButtonColor +""""
					If Control.CssClass<>"" Then Response.Write " Class='" & Control.CssClass  + "' " 
					If Control.Style<>"" Then Response.Write " Style='" & Control.Style + "' "
					If Control.Attributes <> "" Then Response.Write Control.Attributes   & " "
				%> 
				border="0" cellpadding="1" cellspacing="0">
			<% If Not HideToolBar Then %>
			<tr><td>
			<a href="javascript:;" onclick="RichTextBox.Preview(document.forms.frmForm['<%=Control.Name%>_Editor_Value'].value)"><img class="RichBtn" border="0" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/preview.gif" alt="Preview" onmouseover="RichTextBox.ButtonOver(this.style);" onmouseout="RichTextBox.ButtonOut(this.style);" onmousedown="RichTextBox.ButtonDown(this.style);" onmouseup="RichTextBox.ButtonUp(this.style);" width="23" height="22"></a>
			</td></tr>
			<% End If %>
			<tr>
				<td><textarea 
						name="<%=Control.Name%>_Editor_Value"  
						cols = "<%=Width/10%>" 
						rows = "<%=Height/10%>" 
						style = "width:<%=Width%>px; height:<%=Height%>px;" 
						<%=IIF(ReadOnly,"readonly","")%> <%=IIF(Not Control.Enabled,"disabled","")%>><%=Server.HTMLEncode(Text)%></textarea>
				</td>
			</tr>
			</table>
			</td></tr></table>
		<%
		End Function

		Private Function RenderRichTextBox()
			Dim fnc
		%>
			
			<table bgcolor="<%=BackColor%>" border="0" cellspacing="0" cellpadding="1"><tr><td>
			<table 
				<% 
					If ButtonColor<>"" Then Response.Write "bgcolor="""+ ButtonColor +""""
					If Control.CssClass<>"" Then Response.Write " Class='" & Control.CssClass  + "' " 
					If Control.Style<>"" Then Response.Write " Style='" & Control.Style + "' "
					If Control.Attributes <> "" Then Response.Write Control.Attributes   & " "
				%> 
				border="0" cellpadding="1" cellspacing="0">
			<tr>
				<td>
				<% 
					If Not HideToolBar Then 
						If ToolBarTemplate = "" Then  
							RenderToolBar
						Else
							Set fnc = GetFunctionReference(ToolBarTemplate)	
							If Not fnc Is Nothing Then
								Call fnc()
							End If			
						End If
					End If
				%>
				</td>
			  </tr>
			  <tr>
				<td>
				<iframe id="<%=Control.Name%>_Editor" width="<%=Width%>" height="<%=Height%>"></iframe>
				<input name="<%=Control.Name%>_Editor_Value" type="hidden" value="">
				<script language="JavaScript">
					var htm = '<%=CJString(Text)%>'; 									
					var ro = <%=IIF(ReadOnly,"true","false")%>;
					var en = <%=IIF(Not Control.Enabled,"false","true")%>;
					var style = '<%=DocumentStyle%>';
					var igurl = '<%=ImageGalleryURL%>';
					var <%=Control.Name%> = new RichTextBox("<%=Control.Name%>_Editor",htm,ro,en,style,igurl);
				</script>
				</td>
			  </tr>
			  <% If Not HideToolBar And FooterTemplate <>"" Then %>
			  <tr>
				<td bgcolor="#efedde">
					<%
						Set fnc = GetFunctionReference(FooterTemplate)	
						If Not fnc Is Nothing Then
							Call fnc()
						End If			
					%>
				</td>
			  </tr>
			  <% End If %>
			  </table>
			  </td></tr></table>
		<%
		End Function	   

		Public Function RenderToolBar()
		%>	
			<!-- Toolbar start -->
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
			<td align="top">
				<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td>&nbsp;
					<select class="RichSelect" onchange="<%=Control.Name%>.cmdExec('formatBlock',this[this.selectedIndex].value);this.selectedIndex=0;"<%=IIF(Not Control.Enabled," Disabled", "")%>>
						<option selected>Style</option>
						<option value="Normal">Normal</option>
						<option value="Heading 1">Heading 1</option>
						<option value="Heading 2">Heading 2</option>
						<option value="Heading 3">Heading 3</option>
						<option value="Heading 4">Heading 4</option>
						<option value="Heading 5">Heading 5</option>
						<option value="Address">Address</option>
						<option value="Formatted">Formatted</option>
						<option value="Definition Term">Definition Term</option>
					</select>
					<select class="RichSelect" onchange="<%=Control.Name%>.cmdExec('fontname',this[this.selectedIndex].value);this.selectedIndex=0;"<%=IIF(Not Control.Enabled," Disabled", "")%>>
						<option selected>Font</option>
						<option value="Arial">Arial</option>
						<option value="Arial Black">Arial Black</option>
						<option value="Arial Narrow">Arial Narrow</option>
						<option value="Comic Sans MS">Comic Sans MS</option>
						<option value="Courier New">Courier New</option>
						<option value="System">System</option>
						<option value="Tahoma">Tahoma</option>
						<option value="Times New Roman">Times New Roman</option>
						<option value="Verdana">Verdana</option>
						<option value="Wingdings">Wingdings</option>
					</select>
					<select class="RichSelect" onchange="<%=Control.Name%>.cmdExec('fontsize',this[this.selectedIndex].value);this.selectedIndex=0;"<%=IIF(Not Control.Enabled," Disabled", "")%>>
						<option selected>Size</option>
						<option value="1">1</option>
						<option value="2">2</option>
						<option value="3">3</option>
						<option value="4">4</option>
						<option value="5">5</option>
						<option value="6">6</option>
						<option value="7">7</option>
						<option value="8">8</option>
						<option value="10">10</option>
						<option value="12">12</option>
						<option value="14">14</option>
					</select>
					<img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/fgcolor.gif" alt="Fore Color" onClick="<%=Control.Name%>.showColorPicker()" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22">
					</td>
					<td nowrap height="22">&nbsp;<img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/hr.gif" alt="Insert Horizontal Rule" onClick="<%=Control.Name%>.cmdExec('InsertHorizontalRule')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"></td>
					<td nowrap height="22"><img src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/vertical_line.gif" width="2" height="22"></td>
					<td nowrap height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/link.gif" alt="Insert Link" onClick="<%=Control.Name%>.cmdExec('createLink')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/imagelink.gif" alt="Insert Image" onClick="<%=Control.Name%>.showImageDialog()" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><% If Len(ImageGalleryURL) > 0 Then %><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/image.gif" alt="Image Gallery" onClick="<%=Control.Name%>.showImageGallery()" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><% End If %><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/table.gif" alt="Insert Table" onClick="<%=Control.Name%>.showTableDialog()" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img name = "<%=Control.Name%>_Editor_TBvhtml" class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/htag.gif" alt="View HTML" onClick="<%=Control.Name%>.toggleHTMLView()" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/preview.gif" alt="Preview" onClick="<%=Control.Name%>.showPreview()" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22">&nbsp;</td>
				</tr>
				</table>
				<table border="0" cellpadding="0" cellspacing="0" ID="Table1">
				<tr>
					<td nowrap height="22">&nbsp;<img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/Cut.gif" alt="Cut" onClick="<%=Control.Name%>.cmdExec('cut')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/Copy.gif" alt="Copy" onClick="<%=Control.Name%>.cmdExec('copy')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/Paste.gif" alt="Paste" onClick="<%=Control.Name%>.cmdExec('paste')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/undo.gif" alt="Undo" onClick="<%=Control.Name%>.cmdExec('Undo')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/redo.gif" alt="Redo" onClick="<%=Control.Name%>.cmdExec('Redo')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"></td>
					<td width="2"><img src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/vertical_line.gif" width="2" height="22"></td>
					<td nowrap><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/Bold.gif" alt="Bold" onClick="<%=Control.Name%>.cmdExec('bold')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/Italic.gif" alt="Italics" onClick="<%=Control.Name%>.cmdExec('italic')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/under.gif" alt="Underline" onClick="<%=Control.Name%>.cmdExec('underline')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/strikethrough.gif" alt="Strike-through" onClick="<%=Control.Name%>.cmdExec('StrikeThrough')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/Superscript.gif" alt="Superscript" onClick="<%=Control.Name%>.cmdExec('SuperScript')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/Subscript.gif" alt="Subscript" onClick="<%=Control.Name%>.cmdExec('SubScript')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"></td>
					<td width="2"><img src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/vertical_line.gif" width="2" height="22"></td>
					<td nowrap><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/left.gif" alt="Justify Left" onClick="<%=Control.Name%>.cmdExec('justifyleft')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/Center.gif" alt="Center" onClick="<%=Control.Name%>.cmdExec('justifycenter')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/right.gif" alt="Justify Right" onClick="<%=Control.Name%>.cmdExec('justifyright')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"></td>
					<td width="2"><img src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/vertical_line.gif" width="2" height="22"></td>
					<td nowrap><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/numlist.gif" alt="Ordered List" onClick="<%=Control.Name%>.cmdExec('insertorderedlist')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/bullist.gif" alt="Unordered List" onClick="<%=Control.Name%>.cmdExec('insertunorderedlist')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/inindent.gif" alt="Increase Indent" onClick="<%=Control.Name%>.cmdExec('indent')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"><img class="RichBtn" hspace="1" vspace="1" align=absmiddle src="<%=SCRIPT_LIBRARY_PATH%>richtext/images/outdent.gif" alt="Decrease Indent" onClick="<%=Control.Name%>.cmdExec('outdent')" onmouseover="<%=Control.Name%>.buttonOver(this);" onmouseout="<%=Control.Name%>.buttonOut(this);" onmousedown="<%=Control.Name%>.buttonDown(this);" onmouseup="<%=Control.Name%>.buttonUp(this);" width="23" height="22"></td>
				</tr>
				</table>
			</td>
			</tr>
			</table>						
			<!-- Toolbar end -->
		<%
		End Function 
		
		'CJString - converts server-side string (multi-line) to javascript (signle-line) string
		Private Function CJString(text)
			if Len(text)=0 Then Exit Function
			text=Replace(text,"\","\\")			' replace \ with \\
			text=Replace(text,"'","\'")			' replace ' with \'
			text=Replace(text,vbCrLf,"\n")		' replace CrLf with \n
			text=Replace(text,vbLf,"\n")		' replace single Lf with \n
			text=Replace(text,vbCr,"\r")		' replace single Cr with \n
			text=Replace(text,"</script>","<\/script>",1,-1,1)	' replace js ending </script> tag with <\/script>
			CJString = text 
		End Function


		Private Function PreloadToolBarImages()
			Dim pth

			If DisableImagePreloading Then Exit Function

			pth = SCRIPT_LIBRARY_PATH +"richtext/images/"
			Page.RegisterClientImage "RT23",pth+"preview.gif"
			
			If mRichState Then
				Page.RegisterClientImage "RT01",pth+"Bold.gif"
				Page.RegisterClientImage "RT02",pth+"bullist.gif"
				Page.RegisterClientImage "RT03",pth+"Center.gif"
				Page.RegisterClientImage "RT04",pth+"Copy.gif"
				Page.RegisterClientImage "RT05",pth+"Cut.gif"
				Page.RegisterClientImage "RT06",pth+"DeIndent.gif"
				Page.RegisterClientImage "RT07",pth+"empty.gif"
				Page.RegisterClientImage "RT08",pth+"fgcolor.gif"
				Page.RegisterClientImage "RT09",pth+"HR.gif"
				Page.RegisterClientImage "RT10",pth+"htag.gif"
				Page.RegisterClientImage "RT11",pth+"image.gif"
				Page.RegisterClientImage "RT12",pth+"imageLink.gif"
				Page.RegisterClientImage "RT13",pth+"imageLocal.gif"
				Page.RegisterClientImage "RT14",pth+"imageUpload.gif"
				Page.RegisterClientImage "RT15",pth+"inindent.gif"
				Page.RegisterClientImage "RT16",pth+"Italic.gif"
				Page.RegisterClientImage "RT17",pth+"left.gif"
				Page.RegisterClientImage "RT18",pth+"Link.gif"
				Page.RegisterClientImage "RT19",pth+"list.txt"
				Page.RegisterClientImage "RT20",pth+"numlist.gif"
				Page.RegisterClientImage "RT21",pth+"outdent.gif"
				Page.RegisterClientImage "RT22",pth+"Paste.gif"
				Page.RegisterClientImage "RT24",pth+"redo.gif"
				Page.RegisterClientImage "RT25",pth+"right.gif"
				Page.RegisterClientImage "RT26",pth+"Save.gif"
				Page.RegisterClientImage "RT27",pth+"Strikethrough.gif"
				Page.RegisterClientImage "RT28",pth+"Subscript.gif"
				Page.RegisterClientImage "RT29",pth+"Superscript.gif"
				Page.RegisterClientImage "RT30",pth+"table.gif"
				Page.RegisterClientImage "RT31",pth+"under.gif"
				Page.RegisterClientImage "RT32",pth+"Undo.gif"
				Page.RegisterClientImage "RT33",pth+"vertical_line.gif"
			End If
		End Function

	End Class

%>
