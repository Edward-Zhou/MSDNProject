function getItemSoap(id) {
    // Return a GetItem operation request for the subject of the specified item. 
    var result =
            '<?xml version="1.0" encoding="utf-8"?> ' +
            '<soap:Envelope ' +
            '  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
            '  xmlns:xsd="http://www.w3.org/2001/XMLSchema" ' +
            '  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
            '  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
            '  <soap:Body> ' +
            '    <GetItem ' +
            '      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
            '      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
            '      <ItemShape> ' +
            '        <t:BaseShape>Default</t:BaseShape> ' +
            '        <t:IncludeMimeContent>true</t:IncludeMimeContent> ' +
            '      </ItemShape> ' +
            '      <ItemIds> ' +
            '        <t:ItemId Id="' + id + '" ChangeKey="CQAAAB" /> ' +
            '      </ItemIds> ' +
            '    </GetItem> ' +
            '  </soap:Body> ' +
            '</soap:Envelope>'
    return result;
}


function getItem() {
    // Create a local variable that contains the mailbox.
    var mailbox = Office.context.mailbox;

    mailbox.makeEwsRequestAsync(getItemSoap(mailbox.item.itemId), callback);
}

function callback(asyncResult) {
    var result = asyncResult.value;
    var context = asyncResult.context;

}