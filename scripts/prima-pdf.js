var PrimaPDF = function () {
    var self = this;
    var _UI = [];
    var _originalfileName;
    var _codigoMidia;
    self.totalPages = 0;
    self.codigoMidia = 0;

    self.init = function (originalFileName, codigoMidia) {
        _UI.txtPage = $("#txtPage");
        _UI.spanTotalPages = $("#spanTotalPages");
        _UI.container = $(".container");
        _UI.img = _UI.container.find("img");
        _UI.zoomButtons = $(".pdf-control-zoomIn, .pdf-control-zoomOut");
        _UI.expandButton = $(".pdf-control-expand");
        _UI.paginationButtonPrev = $(".pdf-control-prev");
        _UI.paginationButtonNext = $(".pdf-control-next");
        _UI.pdfBody = $(".pr-pdf .pr-pdf-body");
        _UI.pdfWrapper = $(".pr-pdf .pr-pdf-body .pdf-wrapper");
        _UI.pdfName = $(".pr-pdf .pr-pdf-header .pr-pdf-header-name");

        _resizePrimaPDF();
        $(window).resize(function () {
            _resizePrimaPDF();
        });

        _bindElements();
        
        _originalfileName = originalFileName;
        _codigoMidia = codigoMidia;

        _initPDF(codigoMidia);
    };

    self.goToPage = function (page) {
        page = (page === undefined) ? 1 : (page < 1) ? 1 : (page > self.totalPages) ? self.totalPages : page;
        _UI.txtPage.val(page);

        _getFileAtPage(page);

        _inactivePaginationButtons();
    };

    var _bindElements = function () {
        _UI.zoomButtons.bind("click", _zoomButtonClickHandler);
        _UI.expandButton.bind("click", _expandButtonClickHandler);
        _UI.paginationButtonPrev.bind("click", _paginationButtonClickHandler);
        _UI.paginationButtonNext.bind("click", _paginationButtonClickHandler);

        _UI.txtPage.keypress(function (event) {
            if (event.which == 13) {
                self.goToPage(_UI.txtPage.val());
            }
        });

        _UI.txtPage.change(function () {
            self.goToPage(_UI.txtPage.val());
        });

        _UI.img.load(function () {
            setTimeout(function() {
                _UI.img.on("contextmenu", function (e) {
                    return false;
                });
            }, 1000);
        });
    };

    var _getPdfInfo = function () {
        _UI.pdfName.html(_originalfileName);
        
        _UI.spanTotalPages.text(self.totalPages);
        _UI.txtPage.attr("max", self.totalPages);

        _inactivePaginationButtons();
        self.goToPage(1);
        _UI.img.smartZoom();
    };

    var _zoomButtonClickHandler = function (e) {
        var scaleToAdd = 0.5;
        if (e.target.className.indexOf("pdf-control-zoomOut") != -1)
            scaleToAdd = -scaleToAdd;
        _UI.img.smartZoom("zoom", scaleToAdd, undefined, 300);
    };

    var _expandButtonClickHandler = function () {
        _UI.img.smartZoom("zoom", -5000, undefined, 100);
    };

    var _paginationButtonClickHandler = function (e) {
        if (e.target.className.indexOf("inactive") == -1) {
            var page = parseInt(_UI.txtPage.val());
            page = (e.target.className.indexOf("pdf-control-prev") == -1) ? ++page : --page;
            self.goToPage(page);
        }
    }

    var _inactivePaginationButtons = function () {
        var page = parseInt(_UI.txtPage.val());
        if (page == 1) {
            _UI.paginationButtonPrev.addClass("inactive");
        } else {
            _UI.paginationButtonPrev.removeClass("inactive");
        }

        if (page == self.totalPages) {
            _UI.paginationButtonNext.addClass("inactive");
        } else {
            _UI.paginationButtonNext.removeClass("inactive");
        }
    }

    var _resizePrimaPDF = function () {
        var w = $(window).width();
        _UI.pdfWrapper.width(w - 140);
        var h = $(window).height();
        _UI.pdfBody.height(h - 36);
    };

    var _initPDF = function (codigoMidia) {
        var url = "prima-arquivo-pdf.asp?codigoMidia=" + codigoMidia;
        $.ajax({
            type: "GET",
            url: url,
            cache: false,
            success: function (data) {
                self.totalPages = parseInt(data);
                _getPdfInfo();
            },
            error: function () {
                alert('Erro ao carregar arquivo.');
            }
        });
    }

    var _getFileAtPage = function (page) {
        var ts = (new Date()).getTime();
        _UI.img[0].src = "prima-pagina-pdf.asp?tstmp=" + ts + "&pagina=" + page + "&codigoMidia=" + _codigoMidia;
    }
};