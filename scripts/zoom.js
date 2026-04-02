var w = 0;
var h = 0;
var w_original = 0;
var h_original = 0;
var MIN = 100;

function enableControls(enable) {
	if (enable) {	
		$('input[type="button"]').removeAttr("disabled");
		
		if (w < MIN || h < MIN || w_original < MIN || h_original < MIN) {
			$('#menosButton').attr("disabled", "disabled");
		}
	} else {
		$('input[type="button"]').attr("disabled", "disabled");
	}
}

function animate() {
	enableControls(false);
	
	$('#img').animate({
		width: w,
		height: h
	}, 400, function() {
		enableControls(true);
	});			
}

function autoAjuste(animated) {
	var width = $(window).width() - 20;
	var height = $(window).height() - 50;
	
	if (h_original < w_original) 
	{
		w = width;
		h = (h_original * width) / (w_original);
		if (h > height)
		{
			h = height;
			w = (w_original * height) / (h_original);
		}
	}
	else
	{
		h = height;
		w = (w_original * height) / (h_original);
		if (w > width)
		{
			w = width;
			h = (h_original * width) / (w_original);
		}
	}
	
	if (animated) {
		animate();	
	} else {
		$('#img').width(w);
		$('#img').height(h);
	}	
}

$(document).ready(function () {
	$('#img').one('load', function () {
		w_original = $("#img").width();
		h_original = $("#img").height();

		autoAjuste(false);
	});

	$("#maisButton").click(function () {
		w *= 1.2;
		h *= 1.2;

		animate();
	});

	$("#menosButton").click(function () {
		w *= 0.8;
		h *= 0.8;

		animate();
	});

	$("#originalButton").click(function () {
		w = w_original;
		h = h_original;

		animate();
	});

	$("#ajusteButton").click(function () {
		autoAjuste(true);
	});
});