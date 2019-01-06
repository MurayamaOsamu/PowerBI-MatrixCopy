window.onload = function() {
	document.getElementById("popup-close").onclick = function() {window.close();}
	chrome.tabs.getSelected(null, function(tab) {
		if (!tab.url.match(/^https:\/\/app.powerbi.com\//)) {
			document.getElementById("popup-msg").innerHTML = "Not PowerBI site.";
			document.getElementById("popup-close").style.display = "block";
		} else {
			chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
				chrome.tabs.sendMessage(tabs[0].id, {},
				function(response) {
					if (response == undefined) {
						document.getElementById("popup-msg").innerHTML = "Unknown Error.";
						document.getElementById("popup-close").style.display = "block";
					} else {
						if (response.result) {
							var selection = document.getElementById("CellsText");
							selection.style.display = "block";
							selection.value = response.message;
							selection.select();
							document.execCommand('copy');
							document.getElementById("popup-msg").innerHTML = "Copy complete.<br>You can paste in Excel etc.";
							document.getElementById("popup-close").style.display = "block";
						} else {
							document.getElementById("popup-msg").innerHTML = response.message;
							document.getElementById("popup-close").style.display = "block";
						}
					}
				});
			});
		}
	
	});
};