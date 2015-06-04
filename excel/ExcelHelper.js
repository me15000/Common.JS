var ExcelHelper = {};
ExcelHelper.getBrowserName = function () {

    var ua = window.navigator.userAgent;
    //ie 
    if (ua.indexOf("MSIE") >= 0) {
        return 'ie';
    }
        //firefox 
    else if (ua.indexOf("Firefox") >= 0) {
        return 'Firefox';
    }
        //Chrome
    else if (ua.indexOf("Chrome") >= 0) {
        return 'Chrome';
    }
        //Opera
    else if (ua.indexOf("Opera") >= 0) {
        return 'Opera';
    }
        //Safari
    else if (ua.indexOf("Safari") >= 0) {
        return 'Safari';
    }
}

ExcelHelper.tableToExcel = (function () {
    var uri = 'data:application/vnd.ms-excel;base64,',
    template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>',
      base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) },
      format = function (s, c) {
          return s.replace(/{(\w+)}/g,
          function (m, p) { return c[p]; })
      }

    return function (table, name) {
        var ctx = { "worksheet": name || 'Worksheet', table: table.innerHTML }
        window.location.href = uri + base64(format(template, ctx))
    }
})();


ExcelHelper.CreateExcelByTable = function (table) {
    var bn = this.getBrowserName();
    if (bn == "ie") {
        var ax = new ActiveXObject("Excel.Application");
        var wb = ax.Workbooks.Add();
        var sheet = wb.Worksheets(1);

        var tr = document.body.createTextRange();
        tr.moveToElementText(table);
        tr.select();
        tr.execCommand("Copy");
        sheet.Paste();

        ax.Visible = true;

        var si = null;

        var cleanup = function () {
            if (si) {
                window.clearInterval(si);
            }
        }

        try {
            var fname = ax.Application.GetSaveAsFilename("Excel.xls", "Excel Spreadsheets (*.xls), *.xls");
        } catch (e) {
            print("Nested catch caught " + e);
        } finally {
            wb.SaveAs(fname);
            var savechanges = false;
            wb.Close(savechanges);

            ax.Quit();
            ax = null;

            si = window.setInterval(cleanup, 1);
        }
    } else {
        this.tableToExcel(table);
    }
}
