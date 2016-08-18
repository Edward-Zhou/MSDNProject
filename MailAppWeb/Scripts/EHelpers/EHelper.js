function findItem(id) {
    // Return a GetItem operation request for the subject of the specified item. 
    var result =
            '<?xml version="1.0" encoding="utf-8"?>'+
            '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" '+
            '    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> '+
            '<soap:Body> '+
            '  <FindItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" '+
            '    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"'+
            '    Traversal="Shallow"> '+
            '<ItemShape> '+
            '<t:BaseShape>IdOnly</t:BaseShape> '+
            '</ItemShape> '+
            '<ParentFolderIds> '+
            '<t:DistinguishedFolderId Id="inbox"/> ' +
            '</ParentFolderIds> '+
            '</FindItem> '+
            '</soap:Body> '+
            '</soap:Envelope> '
    return result;
}


function sendRequest() {
    // Create a local variable that contains the mailbox.
    var mailbox = Office.context.mailbox;

    mailbox.makeEwsRequestAsync(findItem(mailbox.item.itemId), callback);
}


function getItem()
{
    // Return a GetItem operation request for the subject of the specified item. 
    var result =
                '<?xml version="1.0" encoding="utf-8"?> ' +
                '<soap:Envelope ' +
                '  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
                '  xmlns:xsd="http://www.w3.org/2001/XMLSchema" ' +
                '  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
                '  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
                '  <soap:Body>' +
                '    <GetItem ' +
                '      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
                '      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
                '      <ItemShape> ' +
                '        <t:BaseShape>Default</t:BaseShape> ' +
                '        <t:IncludeMimeContent>true</t:IncludeMimeContent> ' +
                '      </ItemShape> ' +
                '      <ItemIds> ' +
                ' <t:ItemId Id="AAAmAHYtdGF6aG9AT2ZmaWNlRGV2R3JvdXAub25taWNyb3NvZnQuY29tAEYAAAAAAGBlZGGzbjhItPx41cZC0dwHANhk3Ku+0ZlIl9ziTnbjc78AAAAAAQwAANhk3Ku+0ZlIl9ziTnbjc78AAJE4/OgAAA==" />' +

                '      </ItemIds> ' +
                '    </GetItem> ' +
                '  </soap:Body> ' +
                '</soap:Envelope>'
    return result;
}
function getItem(id)
{
    // Return a GetItem operation request for the subject of the specified item. 
    var result =
                '<?xml version="1.0" encoding="utf-8"?> ' +
                '<soap:Envelope ' +
                '  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
                '  xmlns:xsd="http://www.w3.org/2001/XMLSchema" ' +
                '  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
                '  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
                '  <soap:Header> '+
                '<t:RequestServerVersion Version="Exchange2013" />'+
                ' </soap:Header> '+
                '  <soap:Body>' +
                '    <GetItem ' +
                '      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
                '      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
                '      <ItemShape> ' +
                '        <t:BaseShape>Default</t:BaseShape> ' +
                '        <t:IncludeMimeContent>true</t:IncludeMimeContent> ' +
                '      </ItemShape> ' +
                '      <ItemIds> ' +
                ' <t:ItemId Id="'+id+'" />' +
                '      </ItemIds> ' +
                '    </GetItem> ' +
                '  </soap:Body> ' +
                '</soap:Envelope>'
    return result;
}
function GetItem()
{
    // Create a local variable that contains the mailbox.
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(getItem(mailbox.item.itemId), callback);
}

//GetItem operation (contact)
function getItemContact(id) {
    var result=
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"  '+
'    xmlns:xsd="http://www.w3.org/2001/XMLSchema"   '+
'    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"   '+
'    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">  ' +
                '  <soap:Header> ' +
                '<t:RequestServerVersion Version="Exchange2013" />' +
                ' </soap:Header> ' +
'<soap:Body>  '+
'  <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">  '+
        '    <ItemShape>  '+
'      <t:BaseShape>AllProperties</t:BaseShape>  '+
'    </ItemShape>  '+
'    <ItemIds>  '+
'      <t:ItemId Id="'+id+'" />  '+
'    </ItemIds>  '+
'  </GetItem>  '+
'</soap:Body>  '+
'</soap:Envelope>'

      return result;
}

function GetItemContact()
{
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(getItemContact(mailbox.item.itemId),callback);
}
function callback(asyncResult) {
    var result = asyncResult.value;
    var context = asyncResult.context;

}