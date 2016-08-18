/// <reference path="../App.js" />
/// <reference path="D:\Edward\Project\MSDNProject\MSDNProject\TaskPaneWeb\Scripts/jquery-1.9.1.js" />

(function () {
    "use strict";
    
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
                ['TextBindings', 'TextCoercion', 'TableCoercion', 'TableBindings', 'Settings',
             'MatrixBindings', 'MatrixCoercion', 'ImageCoercion', 'DocumentEvents',
             'BindingEvents', 'ExcelApi'].forEach(function (d) {
                 console.log(d + ':' + Office.context.requirements.isSetSupported(d, '1.1'));
             });
            $('#get-data-from-selection').click(getDataFromSelection);
            $('#Bindings - addFromSelectionAsync').click(addFromSelectionAsync);
            $('#addRows').click(addRows);
            $('#DocumentRefer').click(DocumentRefer);
            $('#getUrl').click(getUrl);
            $('#createTableAtSelection').click(createTableAtSelection);
            $('#createObject').click(createObject);
            $('#insertImageWithStream').click(insertImageWithStream);
            $('#getIP').click(getIP);
        });
    };

    function getUrl() {
        app.showNotification(document.URL+" ; "+window.location.href+" ; "+window.document.referrer);
    }

    function DocumentRefer() {
        var x = document.referrer;
        app.showNotification(x);
    }

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            { valueFormat: "formatted" },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

})();


function OpenExcel() {
    var Excel = new ActiveXObject("Excel.Application");
    Excel.visible = true;
    Excel.Workbooks.open("C:\Users\v-tazho\Desktop\OpMark.xlsx");
}

function getIP()
{
    $(document).ready(function () {
        $.getJSON("http://jsonip.com/?callback=?", function (data) {
            console.log(data);
            app.showNotification(data.ip);
        });
    });
}
var base64Img;
function insertImageWithStream()
{
    //var can = document.getElementById("imgCanvas");
    //var img = document.getElementById("imageid");
    //var ctx = can.getContext("2d");
    //ctx.drawImage(img, 10, 10);   
    base64Img.replace("data:image/png;base64,","");
    Office.context.document.setSelectedDataAsync(base64Img, {
        coercionType: Office.CoercionType.Image,
        imageLeft: 50,
        imageTop: 50,
        imageWidth: 100,
        imageHeight: 100
    },
      function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.log("Action failed with error: " + asyncResult.error.message);
          }
      });
    
}


function toDataUrl(url, callback, outputFormat) {
    var img = new Image();
    img.crossOrigin = 'Anonymous';
    img.onload = function () {
        var canvas = document.createElement('CANVAS');
        var ctx = canvas.getContext('2d');
        var dataURL;
        canvas.height = this.height;
        canvas.width = this.width;
        ctx.drawImage(this, 0, 0);
        dataURL = canvas.toDataURL(outputFormat);
        callback(dataURL);
        canvas = null;
    };
    img.src = url;
}

function previewFile() {
  var preview = document.querySelector('img');
  var file    = document.querySelector('input[type=file]').files[0];
  var reader  = new FileReader();

  reader.addEventListener("load", function () {
    preview.src = reader.result;
  }, false);

  if (file) {
      reader.readAsDataURL(file);
      base64Img = reader.result;
      base64Img.replace("data:image/png;base64,", "");
      Office.context.document.setSelectedDataAsync(base64Img, {
          coercionType: Office.CoercionType.Image,
          imageLeft: 50,
          imageTop: 50,
          imageWidth: 100,
          imageHeight: 100
      },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log("Action failed with error: " + asyncResult.error.message);
            }
        });
  }
}
//addFromSelectionAsync
function addFromSelectionAsync() {
    Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Table, { id: 'addFromSelection' },
       function (asyncResult) {
           app.showNotification('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
       });
}

//addBindingFromPrompt
function addBindingFromPrompt() {

    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Table, { id: 'MyBinding', promptText: 'Select text to bind to.' }, function (asyncResult) {
        app.showNotification('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    });
}

//add rows
function addRows() {
    Office.context.document.bindings.getByIdAsync("addFromSelection",
        function (asyncResult) {
            var binding = asyncResult.value;
            binding.addRowsAsync([["hello", "word"]], function (asyncResult) {
                app.showNotification(asyncResult.value);
            });
        })
}

//add tabledata
function addTableData() {
    // Build table.
    var myTable = new Office.TableData();
    //myTable.headers = [["Cities"]];
    //myTable.rows = [['Berlin'], ['Roma'], ['Tokyo'], ['Seattle']];
    myTable.headers = [["C1", "C2"]];
    //one row
    //myTable.rows = [['Berlin1', 'Berlin2']];
    //multiple rows
    myTable.rows = [['Berlin1',''  ], ['Roma1', '' ], ['Tokyo1', '']]; //this is right 
    //myTable.rows = [['Berlin1', ], ['Roma1', ], ['Tokyo1', ]];//this is wrong
    // Write table.
    //Office.context.document.setSelectedDataAsync(myTable, { coercionType: "table" },
    //    function (result) {
    //        var error = result.error
    //        if (result.status === "failed") {
    //            app.showNotification(error.name + ": " + error.message);
    //        }
    //    });
    //Write text to the current user selection
    //Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'MyBinding' },
    //    function (asyncResult) {
    //        app.showNotification('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    //    }
    //);
    //Office.context.document.goToByIdAsync("MyBinding", Office.GoToType.Binding, function (asyncResult) {
    //    if (asyncResult.status == "failed") {
    //        app.showNotification("Action failed with error: " + asyncResult.error.message);
    //    }
    //    else {
    //        app.showNotification("Navigation successful");
    //    }
    //});

    Office.context.document.setSelectedDataAsync(
      "=123+Sheet1!A1",
      function (asyncResult) {
          if (asyncResult.status == "failed") {
              app.showNotification("Action failed with error: " + asyncResult.error.message);
          } else {
              app.showNotification("Success! Click the Next button to move on.");
          }
      });

}

function addLinkWord() {
    //Office.context.document.setSelectedDataAsync("This is <a href='http://www.bing.com/'>Bing</a>", { coercionType: "html" },
    //    function (result)
    //    {
    //        var error = result.error
    //        if (result.status === "failed") {
    //            app.showNotification(error.name + ": " + error.message);
    //        }
    //    });
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
        { valueFormat: "unformatted", filterType: "all" },
         function (asyncResult) {
             if (asyncResult.status == 'succeeded') {
                 tempData = asyncResult.value;
                 var res = /(https?:\/\/[^\s]+)/g.test(tempData);
                 if (res == false) {
                     newdata = tempData.replace(/Bing/gm, "<a target='_self' href='https://bing.com'>Bing</a>");
                 }
                 else {
                     newdata = tempData.replace(/<a\b[^>]*>(.*?)<\/a>/gm, "Bing");

                 }
                 writeHtmlData(newdata);
             }
         });
}

function writeHtmlData(newData) {

    Office.context.document.setSelectedDataAsync(newData, { coercionType: "html" }, function (asyncResult) {
        if (asyncResult.status == "failed") {
            WriteinPage('Error: ' + asyncResult.error.message);
        }
    });
}

//add event handler
function addEventHandlerToDocument() {
    //DocumentSelectionChanged
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function (asyncResult)
    {
        //get selected value
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Html, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                tempData = result.value;
                var value = getLinks(tempData); //get link from the returned html
                $(window.parent.document).find("#Frm_Main").attr("src", "https://www.bing.com");//set the src of iframe with the link
                app.showNotification('The selected text is:', value + '"');
            } else {
                app.showNotification('Error:', result.error.message);
            }
        });
    });
}
//get links
function getLinks(text) {
    var urlRegex = /(https?:\/\/[^\s]+)/g;
    return text.replace(urlRegex, function (url) {
        return '<a href="' + url + '">' + url + '</a>';
    })

}

//add table wtih formulas
function addTablewithFormula() {
    var data = [['numbers', 'formulas'], [1, '=1'], [2, '=2'], [3, '=3']];

    data.headers = [['numbers', 'formulas']];
    data.rows = [
      [1, '=1'],
      [2, '=2'],
      [3, '=3'],
    ];
    var data1 = new Office.TableData();

    data1.headers = [['numbers', 'formulas']];
    data1.rows = [
      [1, '=[@numbers]'],
      [2, '=[@numbers]'],
      [3, '=[@numbers]'],
    ];

    var opts = { coercionType: Office.CoercionType.Matrix };
    var callback = function (res) {
        if (res.status === Office.AsyncResultStatus.Failed)
            return console.log("Err: setting data: " + res.error.message);
    };
    var callback1 = function (res) {
        if (res.status === Office.AsyncResultStatus.Failed)
            return console.log("Err: setting data: " + res.error.message);
    };
    Office.context.document.setSelectedDataAsync([['numbers', 'formulas'], [1, '=1'], [1, '=2'], [1, '=3']], { coercionType: "matrix" },
    function (asyncResult) {
        var error = asyncResult.error;
        if (asyncResult.status === "failed"){
            write(error.name + ": " + error.message);
        }
    });

}
function getDocumentAsCompressed() {
    Office.context.document.getFileAsync("text",
       { sliceSize: 100000 },
       function (result) {

           if (result.status == Office.AsyncResultStatus.Succeeded) {

               // Get the File object from the result.
               var myFile = result.value;
               var state = {
                   file: myFile,
                   counter: 0,
                   sliceCount: myFile.sliceCount
               };

               updateStatus("Getting file of " + myFile.size +
                   " bytes");

               getSlice(state);
           }
           else {
               updateStatus(result.status);
           }
       });
}
// Get a slice from the file and then call sendSlice.
function getSlice(state) {

    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {

            updateStatus("Sending piece " + (state.counter + 1) +
                " of " + state.sliceCount+" "+ result.data);

            getDocumentAsCompressed(result.value, state);
        }
        else {
            updateStatus(result.status);
        }
    });
}
// Create a function for writing to the status div. 
function updateStatus(message) {
    var statusInfo = document.getElementById("status");
    statusInfo.innerHTML += message + "<br/>";
}
function callUrl()
{
    //var local = "https://10.168.196.230/testWebService/WebService1.asmx/HelloWorld"; 
    var local = "https://www.baidu.com/";
    window.open(local,"_self");
    //$.ajax({
    //    url: local ,
    //    type: "POST",
    //    async: false,
    //    cache: false,
    //    contentType: "application/json; charset=utf-8",
       
    //    dataType: "json",
    //    processData: false,
    //    success: function (data, status) {
    //        app.showNotification(data.d);
    //    },
    //    error: function (ajaxResult) {
    //        app.showNotification(ajaxResult.status + ": " + ajaxResult.statusText);
    //    }
    //});
}

function AjaxCall(type, path, params, values, success, error) {
    params = params || [];
    values = values || [];

    var result = null;
    var data = '{';
    for (var i in params) {
        if (i > 0)
            data += ',';
        data += '"' + params[i] + '":"' + values[i] + '"'
    }
    data += '}';

    var local = "https://xx.xx.xxx.xx/OfficeQuotingServices/OfficeQuoting.svc/";
    $.ajax({
        url: local + path,
        type: type,
        async: false,
        cache: false,
        contentType: "application/json; charset=utf-8",
        data: data,
        dataType: "jsonp",
        processData: false,
        success: success || function (ajaxResult) {
            result = ajaxResult;
        },
        error: error || function (ajaxResult) {
            result = ajaxResult.status + ": " + ajaxResult.statusText;
        }
    });
    return result;
}

function addHandlerAsync()
{
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function (eventArgs) {
        app.showNotification(eventArgs);
    }, function (asyncResult) {
        app.showNotification(asyncResult);
    });

}
var r;
function customXmlPart()
{
    Office.context.document.customXmlParts.addAsync("<suppliers><supplier ID='1'>Contoso</supplier><supplier ID='2'>Wingtip Toys</supplier></suppliers>",
        function (result) {
             r = result.value.id;
    });

}
function customXmlPartEvent() {
    Office.context.document.customXmlParts.getByIdAsync(r, function (result) {
        var xmlPart = result.value;
        xmlPart.addHandlerAsync(Office.EventType.NodeDeleted, function (eventArgs) {
            app.showNotification("A node has been deleted.");
        });
    });
}
// Function that writes to a div with id='message' on the page.
function write(message) {
    document.getElementById('message').innerText += message;
}

function customXmlPartEvent1()
{
    Office.context.document.customXmlParts.getByIdAsync(r, function (asyncResult) {
        var byIdXmlPart = asyncResult.value;

        byIdXmlPart.addHandlerAsync(Office.EventType.NodeInserted, function (eventArgs) { // listen to node Inserted event.
            app.showNotification("A node has been Inserted.");
        });

        byIdXmlPart.addHandlerAsync(Office.EventType.NodeDeleted, function (eventArgs) { // listen to node deleted event.
            app.showNotification("A node has been deleted.");
        });

        byIdXmlPart.addHandlerAsync(Office.EventType.NodeReplaced, function (eventArgs) { // listen to node replaced event.
            app.showNotification("A node has been replaced.");
        });
    });
}

function deleteXMLPart() {
    Office.context.document.customXmlParts.getByIdAsync(r, function (result) {
        var xmlPart = result.value;
        xmlPart.getNodesAsync('*/*', function (nodeResults) {
            for (i = 0; i < nodeResults.value.length; i++) {
                var node = nodeResults.value[i];
                node.setXmlAsync('<root categoryId="1" xmlns="http://tempuri.org"><item name="Cheap" price="$193.95"/><item name="Expensive" price="$931.88"/></root>');
            }
        })
        //xmlPart.deleteAsync(function (eventArgs) { // listen to node replaced event.
        //    app.showNotification("delete");
        //});

    });
}
// Function that writes to a div with id='message' on the page.
function write(message) {
    document.getElementById('message').innerText += message;
}

function writeMatrix() {
    Office.context.document.setSelectedDataAsync([["Red", "Rojo"], ["Green", "Verde"], ["Blue", "Azul"]], {coercionType: "matrix"},
    function(asyncResult) {
        var error = asyncResult.error;
        if (asyncResult.status === "failed"){
            write(error.name + ": " + error.message);
        }
    });


}

function EmptyMatrix() {
    Office.context.document.setSelectedDataAsync([["", ""], ["", ""], ["", ""]], { coercionType: "matrix" },
    function (asyncResult) {
        var error = asyncResult.error;
        if (asyncResult.status === "failed") {
            write(error.name + ": " + error.message);
        }
    });
}
var xmlString;
function getWordOOXML()
{
    Office.context.document.getSelectedDataAsync("ooxml", function (asyncResult) {
        var error = asyncResult.error;
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            app.showNotification(error.name + ": " + error.message);
        }
        else {
            // Get selected data.
            var dataValue = asyncResult.value;
            xmlString = asyncResult.value;
            app.showNotification('Selected data is ' + dataValue);
        }
    });
}
    
function setWordOOXML() {
    var tarXml=xmlString.replace('w:right="1800"', 'w:right="720"');
    var tarXml=xmlString.replace('Test', 'Hello');
    Office.context.document.setSelectedDataAsync(tarXml, { coercionType: "ooxml" }, function (asyncResult) {
        var error = asyncResult.error;
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            app.showNotification(error.name + ": " + error.message);
        }
        else {
            // Get selected data.
            var dataValue = asyncResult.value;
            app.showNotification('Selected data is ' + dataValue);
        }
    })
}

function getExcelRangeAddress() {
    Excel.run(function (ctx) {
        var names = ctx.workbook.names;
        var range = names.getItem('MyRange').getRange();
        range.load('address');
        return ctx.sync().then(function () {
            console.log(range.address);
        });
    }).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}

function getWordSelectionHtml() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.name + ": " + error.message);
            }
            else {
                // Get selected data.
                var dataValue = asyncResult.value;
                console.log('Selected data is ' + dataValue);
            }
        });

}

function getWordSelectionOoxml()
{
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Ooxml,
          
           function (asyncResult) {
               var error = asyncResult.error;
               if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                   console.log(error.name + ": " + error.message);
               }
               else {
                   // Get selected data.
                   var dataValue = asyncResult.value;
                   console.log('Selected data is ' + dataValue);
               }
           });

}

function createCookies() {
    document.cookie = "username=John Doe;expires=Thu, 18 Dec 2016 12:00:00 UTC";
}

function addHandler() {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged,MyHandler);
}
function MyHandler(eventArgs) {
    app.showNotification('Evnet raised');
}

function insertImage() { var htmlString = "<img " +
                   "src='https://i.ytimg.com/vi/qk51u8-4uo4/hqdefault.jpg'"
                   + " alt ='apps for Office image' img/>";
    var htmlString1 = "<b>Hello</b> World!";
    Office.context.document.setSelectedDataAsync(htmlString1, { CoercionType: Office.CoercionType.Html },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                app.showNotification('Error: ' + asyncResult.error.message);
            }

           
        })
}

function BindingDataChangedEvent() {
    Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: "BindingId" }, function (asyncResult) {
        asyncResult.value.addHandlerAsync(Office.EventType.BindingDataChanged, function (eventArgs) {
            app.showNotification('binding changed '+ eventArgs.binding.id);
        });
    });
}
var tableData;
var options;
function createObject() {
    //tableData.headers = [['header1']];
    tableData = new Office.TableData();
    var rows = [], cells = [];

    var formatColor = { fontColor: "blue" };
    var formats = [];
    var formatCells = {};

    for (var r = 0; r < 10000; r++) {
        cells = [];

        for (var c = 0; c < 20; c++) {
            cells.push("data");

            formatCells = {};
            formatCells.cells = {};
            formatCells.format = {};

            formatCells.cells.row = r;
            formatCells.cells.column = c;
            formatCells.format = formatColor;
            formats.push(formatCells);
        }

        rows.push(cells);
    }


    tableOptions = { headerRow: true, filterButton: true, bandedRows: false, style: "none" };

     options = {
        coercionType: Office.CoercionType.Table,
        tableOptions: tableOptions,
        cellFormat: formats
    };

    tableData.rows = rows;
    //app.showNotification("create object successfully");
}

function createTableAtSelection() {   
    
    Office.context.document.setSelectedDataAsync(tableData, options,
    function (asyncResult) {
        app.showNotification(asyncResult.status);
        
    });
}






    