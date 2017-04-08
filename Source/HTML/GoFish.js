<script type="text/javascript">

	//Finds y value of given object
	function findPos(obj) {
			var curtop = 0;
			if (obj.offsetParent) {
				   do {
						   curtop += obj.offsetTop;
				   } while (obj = obj.offsetParent);
			return [curtop];
			}
	}

	//Scroll to location of "matchdiv"
	function ScrollToMatch(){
		var MatchDiv = document.getElementById('matchline');
		var matchPosition = findPos(MatchDiv);
		var bodyHeight = document.body.offsetHeight;
		var scrollTo = parseInt(matchPosition) - (parseInt(bodyHeight) / 3);
		window.scroll(0, scrollTo);
	}

	window.onload=ScrollToMatch;
	
</script>