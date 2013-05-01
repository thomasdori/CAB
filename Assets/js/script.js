function CommunicationHandler() {
	this.post = function(url, data) {
		data = $.extend(data, {'token': $("#token").val()});

		$.post(url, data)
		.done(function(data, textStatus, jqXHR) {
			if(data.token){
				alert("received new token:" + data.token);
				$("#token").val(data.token);
			} else {
				//todo: handle error
			}
		})
		.fail(function(jqXHR, textStatus, errorThrown) {
			//todo: handle error
		})
		.always(function() {
			//todo
		});
	}
}