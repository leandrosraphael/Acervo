/* Inicialización para empregar galego no plugin jQuery UI date picker. */
/* Traducido por OpenHost S.L. para o proxecto OpenNemas (http://www.openhost.es/es/opennemas) */
jQuery(function($){
	$.datepicker.regional['gl'] = {
		closeText: 'Pechar',
        prevText: '&#x3c;Prev',
        nextText: 'Seg&#x3e;',
		currentText: 'Hoxe',
		monthNames: ['Xaneiro','Febreiro','Marzo','Abril','Maio','Xu&ntilde;o',
		'Xullo','Agosto','Setembro','Outubro','Novembro','Decembro'],
		monthNamesShort: ['Xan', 'Feb', 'Mar', 'Abr', 'Mai', 'Xu&ntilde;',
		'Xul', 'Ago', 'Set', 'Out', 'Nov', 'Dec'],
        dayNames: ['Domingo', 'Luns', 'Martes', 'M&eacute;rcores', 'Xoves', 'Venres', 'S&aacute;bado'],
        dayNamesShort: ['Dom', 'Lun', 'Mar', 'M&eacute;r', 'Xov', 'Ven', 'S&aacute;b'],
		dayNamesMin: ['Do','Lu','Ma','M&eacute;','Xo','Ve','S&aacute;'],
		weekHeader: 'Sm',
		dateFormat: 'dd/mm/yy',
		firstDay: 1,
		isRTL: false,
		showMonthAfterYear: false,
		yearSuffix: ''};
	$.datepicker.setDefaults($.datepicker.regional['gl']);
});
