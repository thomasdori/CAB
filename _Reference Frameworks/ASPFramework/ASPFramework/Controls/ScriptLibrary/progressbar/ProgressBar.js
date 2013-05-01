var CLASP_ProgressBarObj = null;
var CLASP_ProgressBarReady = false;

function CLASP_ProgressBar_Start() {
	var url = SCRIPT_LIBRARY_PATH + 'progressbar/progressbar.html';
	CLASP_ProgressBarObj = window.open(url,"pbar","height=100,width=400,status=yes,toolbar=no,menubar=no,location=no");
}

function CLASP_Progress_Ready() {
	CLASP_ProgressBarReady = true;
}

function CLASP_ProgressBar_Update(v) {
	if(CLASP_ProgressBarReady)
		CLASP_ProgressBarObj.UpdateProgress(v);
	else
		setTimeout("CLASP_ProgressBarObj.UpdateProgress(" + v + ");",500);
}

function CLASP_ProgressBar_End(v) {
	if(CLASP_ProgressBarObj) {
		CLASP_ProgressBarObj.close();
	}
}