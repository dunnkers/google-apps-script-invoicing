export const EmailPreview = email => `
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
    <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">
</head>

<body>
    <textarea id="message" rows="4" style="width: 100%;" placeholder="Dear client..."></textarea>
    <iframe 
        style="width: 99%; height: 430px;"
        src="${'data:text/html,' + encodeURIComponent(email)}">
        
    </iframe>
    <input type="checkbox" id="checkbox1" checked>
    <label for="checkbox1">Attach hour declaration</label>
    <div style="float: right;">
        <button class="action">Send</button>
        <button>Save to drafts</button>
        <input type="button" value="Cancel" onclick="google.script.host.close()" />
    </div>
    
</body>

</html>
`;