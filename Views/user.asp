<!-- #include File="Partials/HeaderPartial.asp" -->
<!-- #include File="Partials/FooterPartial.asp" -->

<%=header.getContent()%>
    <!-- Add your site or application content here -->
    <p>Hello world! This is HTML5 Boilerplate.</p>

	<%
		' Unfortunately does not work. Solution: do not use cookies ;)
		'Call cookie.write("test1","<>!@#$%^&*()_+")
        'Call cookie.write("test2","thomas")
        'Call cookie.write("test3","xyz")
'
		'output.writeLine("test1: " & cookie.read("test1"))
        'output.writeLine("test2: " & cookie.read("test2"))
        'output.writeLine("test3: " & cookie.read("test3"))
    %>

    <script type="text/javascript">
    	var userComm = new CommunicationHandler();

    	//send a NOT successfull request
    	userComm.post("../Controllers/UserController.asp", {data1:'', data2:''});

    	//send a successfull request
    </script>
<%=footer.getContent()%>