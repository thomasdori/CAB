<!-- #include File="Views/Partials/HeaderPartial.asp" -->
<!-- #include File="Views/Partials/FooterPartial.asp" -->

<%=header.getContent()%>
    <!-- Add your site or application content here -->
    <p>Hello world! This is HTML5 Boilerplate.</p>
    <div>
    	<%
    		Call cookie.write("test1","<>!@#$%^&*()_+")
            Call cookie.write("test2","thomas")
            Call cookie.write("test3","xyz")

    		output.writeLine("test1: " & cookie.read("test1"))
            output.writeLine("test2: " & cookie.read("test2"))
            output.writeLine("test3: " & cookie.read("test3"))
    	%>
    </div>

    <script type="text/javascript">
    	var userComm = new CommunicationHandler();

    	//send a request
    	//userComm.post("url", {data1:'', data2:''})
    </script>
<%=footer.getContent()%>