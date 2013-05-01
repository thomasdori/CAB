<%
	
	Page.IsTemplatePresent = True 'Notifies CLASP that there is a template...
	
	Public Function ServerTextBox_Skin(ctrl)
		ctrl.Style = "border:1px solid blue;background-color:#DDDDDD"
		if ctrl.owner.mode = 2 then
			ctrl.style = ctrl.style & ";color:red"
		end if
	End Function			 
				 
	Public Function ServerButton_Skin(ctrl)
		ctrl.Style = "color:green;"
	End Function			 
%>
