<!-- #include File="Partials/HeaderPartial.aps" -->
<!-- #include File="Partials/FooterPartial.asp" -->

<%=header.getContent()%>
    <!-- Add your site or application content here -->
    <p>Hello world! This is HTML5 Boilerplate.</p>

    <script type="text/javascript">
    	var userComm = new CommunicationHandler();

    	//send a request
    	//userComm.post("url", {data1:'', data2:''})
    </script>
<%=footer.getContent()%>