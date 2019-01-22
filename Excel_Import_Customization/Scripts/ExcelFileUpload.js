
function getExcelColumnsList() {
    var orderedExcelColumnDef = [];

    $('#Excel_droppable').each(function () {
        // this is inner scope, in reference to the .phrase element
        var phrase = '';
        let i = 0;
        $(this).find('li').each(function () {
            var current = $(this);
            let text = current.text();
            let value = current.attr('value');
            orderedExcelColumnDef.push({
                columnName: text,
                columnIndex: value,
                columnOrder: ++i,
            });
        });
    });
    console.log(orderedExcelColumnDef);
    return orderedExcelColumnDef;
}

$(document).ready(function () {

    let DBMetaData =
        [
            {
                Index: 0,
                Name: 'A',
            },
            {
                Index: 1,
                Name: 'B',
            },
            {
                Index: 2,
                Name: 'C',
            },
            {
                Index: 3,
                Name: 'D',
            },
            {
                Index: 4,
                Name: 'E',
            },
            {
                Index: 5,
                Name: 'F',
            }
        ];

    function init() {
        $("#Excel_droppable.droppable-area1, .droppable-area2").sortable({
            // connectWith: ".connected-sortable",
            stack: '.connected-sortable ul'
        }).disableSelection();
    }

    function generateExcelMetadaDropDown(MetaData) {
        let selectHtml = '<select class="form-control">';
        $.each(MetaData, function (i,item) {
            selectHtml += '<option value=' + item.Index + '>' + item.Name + '</option>';
        });
        selectHtml += '</select>';
        return selectHtml;
    }

    function generateExcelMetadaDraggableList(MetaData) {
        let ulList = '<ul id="Excel_droppable" class="connected-sortable droppable-area1">';
        $.each(MetaData, function (i, item) {
            ulList += '<li class="draggable-item" value=' + item.Index + '>' + item.Name + '<span class="close"><i class="fa fa-angle-double-right" aria-hidden="true"></i></span> </li>';
        });
        ulList += '</ul>';
        return ulList;
    }

    function generateDatabaseMetadaDraggableList(MetaData) {
        let ulList = '<ul class="connected-sortable droppable-area1">';
        $.each(MetaData, function (i, item) {
            ulList += '<li class="draggable-item" value=' + item.Index + '>' + item.Name + '<span class="close"><i class="fa fa-angle-double-right" aria-hidden="true"></i></span> </li>';
        });
        ulList += '</ul>';
        return ulList;
    }

    function getTabularData(responseData) {
        let htmlStr = '';

        if (responseData.length === 0)
            return 'No data found in excel';
         
        let arrayData = responseData[0];
        let ExcelColumnHeads = [];

        let headerText = '';
        let bodyText = '';
        headerText += '<tr>';

        let index = 0;

        for (let propName in arrayData) {
            headerText += '<th>';
            headerText += propName;
            headerText += '</th>';
            ExcelColumnHeads.push({
                Index: ++index,
                Name: propName,
            });
        }
        headerText += '</tr>';

        for (let Data in responseData) {
            let thisData = responseData[Data];
            bodyText += '<tr>';
            for (let propName in thisData) {
                if (thisData.hasOwnProperty(propName)) {
                    var propValue = thisData[propName];
                    bodyText += '<td>';
                    bodyText += propValue;
                    bodyText += '</td>';
                }
            }
            bodyText += '</tr>';
        }

        return {
            tableHead: headerText,
            tableBody: bodyText,
            ExcelMetaData: ExcelColumnHeads
        };
    }

    $('#formReportSearch').submit(function (e) {
        e.preventDefault();
        //   let VendorId = ($("#cmbSupplier").val() == "-1" || $("#cmbSupplier").val() == null || $("#cmbSupplier").val() == "") ? 0 : $("#cmbSupplier").val();

        var isRequiredFieldMissing = false;
        let MissingFileds = [];
        
        $.ajax({
            url: "/ExcelImport/UploadExcel",
            type: "POST",
            data: new FormData(this),
            cache: false,
            contentType: false,
            processData: false,
            success: function (data) {
                console.log(JSON.parse(data.Data));
                let resultArray = JSON.parse(data.Data);
                let tblHtml = '';
                let tabularData= getTabularData(resultArray);

                $("#tblExcelFile > thead").html(tabularData.tableHead);
                $("#tblExcelFile > tbody").html(tabularData.tableBody);

            
               let excelMetadataHtml = generateExcelMetadaDraggableList(tabularData.ExcelMetaData);
                let databaseMetadataHtml = generateDatabaseMetadaDraggableList(DBMetaData);



                $("#choiceConfiguration").html('<div class="col-md-6"><h6>DB Columns</h6> ' + databaseMetadataHtml + '</div>' + '<div class="col-md-6"><h6>Excel Columns</h6>' + excelMetadataHtml  +'</div>');

                init();
            },
            error: function (xhr, error, status) {
                
                console.log(error, status);
            }
        });
    });



});