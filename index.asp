<!-- #include File="Views/Partials/HeaderPartial.asp" -->
<!-- #include File="Views/Partials/FooterPartial.asp" -->

<%=header.getContent()%>
    <!-- Add your site or application content here -->
    <p>Hello world! This is HTML5 Boilerplate.</p>
    <div>
    	<%
    		Call cookie.write("test","<>!@#$%^&*()_+")
            Call cookie.write("test2","thomas")
            Call cookie.write("test3","xyz")

    		output.write(cookie.read("test"))
    	%>
    </div>

    <script type="text/javascript">
    	var userComm = new CommunicationHandler();

    	//send a request
    	//userComm.post("url", {data1:'', data2:''})
    </script>
<%=footer.getContent()%>