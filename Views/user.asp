<!-- #include File="../Helpers/ViewHelper.asp" -->

<%=view.GetHeader()%>
	<%=view.GetToken()%>
    <!-- Add your site or application content here -->
    <p>Hello world! This is Classic Asp Boilerplate.</p>

    <script type="text/javascript">
    	var userComm = new CommunicationHandler();

    	//send a NOT successfull request
    	userComm.post("../Controllers/UserController.asp", {'firstName':'thomas', 'lastName':'dori'});

    	//send a successfull request
    </script>
<%=view.GetFooter()%>