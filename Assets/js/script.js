function CommunicationHandler() {
	this.post = function(url, data) {
		data = $.extend(data, {'token': $("#token").val()});

		$.post(url, data)
		.done(function(data) { console.log("second success"); })
		.fail(function(data) { console.log("error"); })
		.always(function(data) { console.log("finished"); });
	}
}