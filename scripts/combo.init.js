$(document).ready(function () {
	$(".styled_combo").each(function () {
		var id = $(this).attr("id");
		$("#" + id).multiselect({
			multiple: false,
			minWidth: 20,
			selectedList: 1,
			header: false,
			open: function () {
				var sel = $("#" + id).multiselect().multiselect("getChecked");
				var selLi = sel.parent().parent();
				var selUl = selLi.parent();
				var index = selUl.find("li").index(selLi);
				$(selUl).scrollTop($(selLi).height() * index);

			}
		});
	});
});