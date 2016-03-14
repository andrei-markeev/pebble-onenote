var UI = require('ui');
var ajax = require('ajax');
var Settings = require('settings');
var Vector2 = require('vector2');

var appUrl = 'http://markeev.com/pebble/onenote.html';

var auth_url = '';
var baseApiUrl = '';
var clientId='';
var clientSecret='';

Settings.config({ url: appUrl }, function() {
    localStorage.removeItem('refresh_token');
    console.log('Settings closed!');
    retrieveNotebooks();
});

if (!Settings.option('auth_code'))
{
    showText('Account is not configured. Please visit app settings on phone.');
    Pebble.openURL(appUrl);
    return;
}
else
    retrieveNotebooks();

function retrieveNotebooks()
{
  var account_type = Settings.option('account_type');
    
  if (account_type == 'onenote.com') {
      clientId = '000000004818A9A6'; // Live app id (https://account.live.com/developers/applications/index)
      clientSecret = ''; // Live app secret
      auth_url = 'https://login.live.com/oauth20_token.srf';
      baseApiUrl = 'https://www.onenote.com/api/v1.0/';
  } else {
      clientId='127482e0-2e1a-4e80-9b81-8a13f9fb822c'; // Azure app id (https://manage.windowsazure.com/)
      clientSecret=''; // Azure app secret
      auth_url = 'https://login.microsoftonline.com/common/oauth2/token?api-version=beta';
      baseApiUrl = 'https://graph.microsoft.com/beta/';
  }

  if (localStorage.getItem('refresh_token') !== null)
    authWithRefreshToken();
  else
    authWithAccessCode();
   
}

function authWithRefreshToken()
{
    console.log('Auth with refresh token');


    ajax({
            url: auth_url,
            method: 'POST',
            data: {
                grant_type: 'refresh_token',
                refresh_token: localStorage.getItem('refresh_token'),
                redirect_uri: appUrl,
                client_id: clientId,
                client_secret: clientSecret
            }
        },
        authorizedCallback,
        showError
    );
}

function authWithAccessCode()
{
    console.log('Auth with access code');

    ajax({
            url: auth_url,
            method: 'POST',
            data: {
                grant_type: 'authorization_code',
                code: Settings.option('auth_code'),
                redirect_uri: appUrl,
                client_id: clientId,
                client_secret: clientSecret
            }
        },
        authorizedCallback,
        showError
    );
}

var access_token = '';
function authorizedCallback(data)
{
    console.log("returned: " + data);
    var dataObj = JSON.parse(data);
    access_token = dataObj.access_token;
    var refresh_token = dataObj.refresh_token;
    localStorage.setItem('refresh_token', refresh_token);
    
    // get list of notebooks
    ajax(
      {
        url: baseApiUrl + '/me/notes/notebooks',
        headers: { "Authorization": "Bearer " + access_token }
      },
      function(data) {
        showMenu('Notebooks', data, 'name', notebookSelected);
      },
      showError
    );
}

var notebook_id = '';
var notebook_title = '';
function notebookSelected(e, id)
{
    notebook_id = id;
    notebook_title = e.item.title;

    // get sections of the selected notebook
    ajax(
        {
            url: baseApiUrl + '/me/notes/notebooks/' + notebook_id + '/sections',
            headers: { "Authorization": "Bearer " + access_token }
        },
        function(data) {
          showMenu('Sections', data, 'name', sectionSelected);
        },
        showError
    );
}

var section_id = '';
function sectionSelected(e, id)
{
    section_id = id;

    // get pages of the section
    ajax(
        {
            url: baseApiUrl + '/me/notes/sections/' + section_id + '/pages',
            headers: { "Authorization": "Bearer " + access_token }
        },
        function(data) {
          showMenu('Pages', data, 'title', pageSelected);
        },
        showError
    );
    
}

var page_id = '';
function pageSelected(e, id)
{
    page_id = id;

    // get contents of the page
    ajax(
        {
            url: baseApiUrl + '/me/notes/pages/' + page_id + '/content',
            headers: { "Authorization": "Bearer " + access_token }
        },
        pageContentRequestSuccess,
        showError
    );
    
}

function pageContentRequestSuccess(data)
{
    var pageFontSize = Settings.option('font_size') || 24;
    
    console.log(data);
    var title;
    var text = data.replace(/<title>([^<]*)<\/title>/, function(m) { title = m.replace(/<(?:.|\n)*?>/gm, ''); return ''; });
    
    var listItems = [];
    text = text.replace(/<li>(.*?)<\/li>/g, function(m) {
        listItems.push({ 
            text: decodeHTMLEntities(m.replace(/<(?:.|\n)*?>/gm, '').replace(/\s+/g, ' ')),
            tag: m.match(/data\-tag="([A-Za-z,:\-]+)"/)[1]
        });
        return '[listItem]';
    });

    console.log(JSON.stringify(listItems));
    
    text = text.replace(/<(br \/|p|\/p|li)>/gm, '[br]');
    text = text.replace(/<(?:.|\n)*?>/gm, '').replace(/\s+/g, ' ').replace(/\[br\]/g,'\n');
    text = decodeHTMLEntities(text);

    var page = new UI.Window({
      backgroundColor: 'white',
      scrollable: true
    });
    
    var yPos = 0;
    var titleSize = sizeVector(title, pageFontSize);
    var titleText = new UI.Text({ position: new Vector2(0, yPos), size: titleSize, text: title, textAlign: 'center', textOverflow: 'wrap', color: 'black', font: 'gothic-' + pageFontSize + '-bold' });
    page.add(titleText);
    yPos += titleSize.y + 5;

    var texts = text.split('[listItem]');
    while (texts.length > 0)
    {
        var fragmentText = texts.shift();
        
        if (fragmentText.replace(/[\s\n]+/g,'') !== '') {
            var fragmentSize = sizeVector(fragmentText, pageFontSize);
            var fragmentElement = new UI.Text({ position: new Vector2(4, yPos), size: fragmentSize, text: fragmentText, textAlign: 'left', textOverflow: 'wrap', color: 'black', font: 'gothic-' + pageFontSize });
            page.add(fragmentElement);
            yPos += fragmentSize.y + 2;
        }

        if (listItems.length > 0) {
            var listItem = listItems.shift();
            var isTodo = listItem.tag.indexOf('to-do') != -1;
            if (isTodo) {
                page.add(new UI.Rect({ position: new Vector2(2, yPos + 4), size: new Vector2(16, 16), borderColor: 'black' }));
                page.add(new UI.Rect({ position: new Vector2(3, yPos + 5), size: new Vector2(14, 14), borderColor: 'black' }));
                if (listItem.tag.indexOf('to-do:completed') != -1) {
                    console.log('adding the checkmark image');
                    page.add(new UI.Image({ position: new Vector2(5, yPos + 7), size: new Vector2(11, 11), image: 'images/check.bmp' }));
                }
            }
            var listItemSize = sizeVector(listItem.text, pageFontSize, isTodo ? 20 : 0);
            var listItemElement = new UI.Text({ position: new Vector2(isTodo ? 24 : 4, yPos), size: listItemSize, text: listItem.text, textAlign: 'left', textOverflow: 'wrap', color: 'black', font: 'gothic-' + pageFontSize });
            page.add(listItemElement);
            yPos += listItemSize.y + 2;
        }
    }
    
    page.show();
}

function decodeHTMLEntities(text) {
    var entities = [
        ['apos', '\''],
        ['amp', '&'],
        ['lt', '<'],
        ['gt', '>']
    ];

    for (var i = 0, max = entities.length; i < max; ++i) 
        text = text.replace(new RegExp('&'+entities[i][0]+';', 'g'), entities[i][1]);

    text = text.replace(/&#([0-9]{1,5});/gi, function(match, numStr) {
        var num = parseInt(numStr, 10); // read num as normal number
        return String.fromCharCode(num);
    });
    
    return text;
}

function showMenu(title, data, title_property, select_callback)
{
    console.log("data received: " + data);
    var dataObj = JSON.parse(data);
    
    if (dataObj.value.length === 0) {
        showText(title + ' not found.');
        return;
    }

    var menuitems = [];
    var ids = [];
    for (var i=0;i<dataObj.value.length;i++)
    {
        menuitems.push({
            title: dataObj.value[i][title_property] || '(untitled)'
        });
        
        ids.push(dataObj.value[i].id);
    }
    var menu = new UI.Menu({
        sections: [
            {
                title: title,
                items: menuitems
            }
        ]
    });
    
    menu.on('select', function(e) { select_callback(e, ids[e.itemIndex]); });
    menu.show();
}

function showError()
{
    var main = new UI.Card({
        title: 'Error',
        body: JSON.stringify(arguments),
        scrollable: true
    });
    main.show();
}

function showText(text)
{
    var main = new UI.Card({
        body: text,
        scrollable: true
    });
    main.show();
    
}

function strTruncate(string, width) {
    if (string.length >= width) {
        var result = string[width - 1] === ' ' ? string.substr(0, width - 1) : string.substr(0, string.substr(0, width).lastIndexOf(' '));
        if (result.length === 0)
            result = string.substr(0, width - 1);
        return result;
    }
    return string;
}

function strTruncateWhole(string, width) {
    var arr = [];
    var b = 0;
    while (b < string.length) {
        arr.push(strTruncate(string.substring(b), width));
        b += arr[arr.length - 1].length;
    }
    return arr;
}

function sizeVector(string, fontSize, substractWidth) {
    var width = 136 - (substractWidth || 0);
    var charsPerLine;
    if (fontSize==14)
        charsPerLine = width / 5;
    else if (fontSize==18)
        charsPerLine = width / 6;
    else if (fontSize==24)
        charsPerLine = width / 7;
    var lines = string.split('\n');
    var height = 0;
    while (lines.length)
    {
        var split = strTruncateWhole(lines.shift(), charsPerLine);
        height += split.length * fontSize;
    }
    console.log('return vector: ' + width + 'x' + height);
    return new Vector2(width, height);
}
