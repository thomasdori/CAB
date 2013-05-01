<%
	If Application("CLASP_PH") = vbEmpty Then
		Application("CLASP_PH") = 0
	End If
	Application("CLASP_PH") = Clng(Application("CLASP_PH")) + 1
	Application("CLASP_LAST_HIT") = Now
%>
<script language="javascript">
	//clasp.launchApplication("CLASP_SAMPLES",800,600);
</script>
<table border="0" width="100%" ID="Table1">
   <tr>
		<td width="50%"><img src="images/clasp.gif"></td>
		<td width="50%" align="right">
			<table ID="Table2">
				<tr><td>
						<span Class="Caption">Samples List<br>
						</span><a href="javascript:window.top.location.href='help/help.htm'">Load Help...</a>
					</td>
				</tr>
			</table>
		</td>
  </tr>
</table>
<hr>
<TABLE width="100%">
	<TR>
		<TD><A HREF="Buttons.asp">Buttons.asp</A></TD>
		<TD><A HREF="CheckBox.asp">CheckBox.asp</A></TD>
		<TD><A HREF="TextBox.asp">TextBox.asp</A></TD>
		<TD><A HREF="Label.asp">Label.asp</A></TD>
		<TD><A HREF="RadioButton.asp">RadioButton.asp</A></TD>
	</TR>
	<TR>
		<TD><A HREF="DropDown.asp">DropDown.asp</A></TD>
		<TD><A HREF="HyperLink.asp">HyperLink.asp</A></TD>
		<TD><A HREF="GenericHTMLControl.asp">GenericHTML.asp</A></TD>
		<TD><A HREF="DataGrid.asp">DataGrid.asp</A></TD>
		<TD><A HREF="DataRepeater.asp">DataRepeater.asp</A></TD>
	</TR>
	<TR>
		<TD><A HREF="DataList.asp">DataList.asp</A></TD>
		<TD><A HREF="DrillDown.asp">DrillDown.asp</A></TD>
		<TD><A HREF="TextBoxAdvance.asp">Formated TextBoxes</A></TD>
		<TD><A HREF="RichTextBox.asp">RichText TextBox</A></TD>
		<TD><A HREF="ImageList.asp">ImageList</A></TD>
	</TR>
	<TR>
		<TD><A HREF="Image.asp">Image.asp</A></TD>
		<TD><A HREF="AutoFocus.asp">Page AutoFocus.asp</A></TD>
		<TD><A HREF="Events.asp">Events.asp</A></TD>
		<TD><A HREF="PerformanceTest.asp">PerformanceTest.asp</A></TD>
		<TD><A HREF="UserControls.asp">UserControls.asp</A></TD>
	</TR>
	<TR>
		<TD><A HREF="Tabs.asp">Tabs.asp</A></TD>
		<TD><A HREF="Panel.asp">Panel.asp</A></TD>
		<TD><A HREF="PageCache.asp">PageCache.asp</A></TD>
		<TD><A HREF="EmployeeList.asp">Employee List View</A></TD>
		<TD><A HREF="Marquee.asp">Marquee.asp</A></TD>
	</TR>
	<TR>
		<TD><a HREF="FileUploader.asp">FileUploader.asp</a></TD>
		<TD><a HREF="PlaceHolder.asp">PlaceHolder.asp</a></TD>
		<TD><a HREF="Browser.asp">Browser Detection</a></TD>
		<TD><A HREF="Timer.asp">Timer.asp</A></TD>
		<TD><A HREF="ServerValidation.asp">ServerValidation.asp</A></TD>
	</TR>
	<TR>
		<TD><A HREF="Tree.asp"><b>Tree.asp (beta)</b></A></TD>
		<TD><A HREF="Menu.asp"><b>Menu.asp (beta)</b></A></TD>
		<TD><A HREF="Graph.asp"><b>Graph.asp (beta)</b></A></TD>
		<TD><A HREF="Layer.asp"><b>Layer.asp (beta)</b></A></TD>
		<TD><A HREF="BarCode.asp"><b>BarCode.asp (beta)</b></A></TD>
	</TR>
	<TR>
		<TD><A HREF="Frame.asp"><B>Frame.asp</B></A></TD>
		<TD><A HREF="ToolBar.asp"><B>ToolBar.asp</B></A></TD>

	</TR>

	
</TABLE>
<HR>
