var url = "./documents/Olympia%202018_startovka.xlsx";
var poc = 0;

function readFile(e) {
    fetch(url).then(function (res) {
        /* get the data as a Blob */
        if (!res.ok) throw new Error("fetch failed");
        return res.arrayBuffer();
    }).then(function (ab) {
        /* parse the data when it is received */
        var data = new Uint8Array(ab);
        var workbook = XLSX.read(data, {
            type: "array"
        });

        /* DO SOMETHING WITH workbook HERE */
        ProcessExcel(workbook);
        poc++;
        if (poc > 600) {
            location.reload();
        }
          setTimeout(readFile, 1500);
    });
}

function ProcessExcel(workbook) {
    /* DO SOMETHING WITH workbook HERE */
    var first_sheet_name = workbook.SheetNames[0];

    /* Var array by ROW */
    var excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[first_sheet_name]);    

    if ($('#table').length) {

        var find = false;
        var row = 1;
        while (!find) {
            for (var i = 0; i < excelRows.length; i++) {
                if ((i + 1) < excelRows.length) {
                    if (!isNaN(excelRows[i][row + ".kolo"]) && isNaN(excelRows[i + 1][row + ".kolo"])) {
                        find = reloadTable(excelRows, i, row);
                        break;
                    }
                } else if ((i + 1) == excelRows.length) {
                    if (!isNaN(excelRows[i][row + ".kolo"]) && isNaN(excelRows[0][(row + 1) + ".kolo"])) {
                        find = reloadTable(excelRows, i, row);
                        break;
                    }
                }
            }
            row++;

            if (row > 4) {
                find = true;
            }
        }
    } else {
        $('#table_here').append($("<table>"));
        $('table').attr("class", "table table-striped");
        $('table').attr("id", "table");
        $('table').append($("<thead>"));
        $('thead').append($("<tr>"));
        $('tr').append($("<th>").text("Jmeno"));
        $('tr').append($("<th>").text("Oddil"));
        $('tr').append($("<th>").text("rok nar"));
        $('tr').append($("<th>").text("nejlepší čas"));
        $('table').append($("<tbody>"));

        $('#name').text(excelRows[0]["jméno"]);


        if (isNaN(excelRows[0]["1.kolo"])) {
            $('#time').text("0.00 s");
        } else {
            $('#time').text(excelRows[0]["1.kolo"] + " s");
        }

        for (var i = 0; i < excelRows.length; i++) {
            $('tbody').append($("<tr id = \"row-" + i + "\">"));
            $('#row-' + i).append($("<td>").text(excelRows[i]["jméno"]));
            $('#row-' + i).append($("<td>").text(excelRows[i]["oddíl"]));
            $('#row-' + i).append($("<td>").text(excelRows[i]["rok nar."]));
            var reg = new RegExp('^[0-9]+.*');
            if (!(reg.test(excelRows[i]["nejlepší čas"]))) {
                $('#row-' + i).append($("<td>").text("0.00 s"));
            } else {
                $('#row-' + i).append($("<td>").text(excelRows[i]["nejlepší čas"]  + " s"));
            }
        }

        var max = 8;
        if (excelRows.length < 8) {
            max = excelRows.length;
        }

        for (max; max < excelRows.length; max++) {
            $('#row-' + max).addClass("hidden");
        }
        $('#row-0').addClass("hidden");
    }
}

function reloadTable(excelRows, i, row) {
    $('#name').empty().text(excelRows[i]["jméno"]);

    if (excelRows[i][row + ".kolo"] === "999.99") {
        $('#time').empty().text("NaN");
    } else {
        $('#time').empty().text(excelRows[i][row + ".kolo"] + " s");
    }

    var max = 8;
    if (excelRows.length < 8) {
        max = excelRows.length;
    }

    for (var j = i; j > 0; j--) {
        $('#row-' + j).addClass("hidden");
    }

    for (var j = (i + max - 1); j > i; j--) {
        $('#row-' + j).removeClass("hidden");
    }

    return true;
}
