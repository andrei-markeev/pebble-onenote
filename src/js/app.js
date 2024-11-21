var UI = require('ui');
var ajax = require('ajax');
var Settings = require('settings');
var Vector2 = require('vector2');
var Feature = require('platform/feature');

var appUrl = 'https://markeev.com/pebble/onenote_v3';
var scope = 'offline_access https://graph.microsoft.com/.default';
var clientId='c8581455-5588-4516-88cf-b983d157df0e';
var clientSecret='';
var auth_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
var baseApiUrl = 'https://graph.microsoft.com/v1.0';

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
                scope: scope,
                refresh_token: localStorage.getItem('refresh_token'),
                redirect_uri: appUrl,
                client_id: clientId,
                client_secret: clientSecret
            }
        },
        authorizedCallback,
        showError('authWithRefreshToken')
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
                scope: scope,
                code: Settings.option('auth_code'),
                redirect_uri: appUrl,
                client_id: clientId,
                client_secret: clientSecret
            }
        },
        authorizedCallback,
        showError('authWithAccessCode')
    );
}

var access_token = '';
function authorizedCallback(data)
{
    var dataObj = JSON.parse(data);
    access_token = dataObj.access_token;
    var refresh_token = dataObj.refresh_token;
    localStorage.setItem('refresh_token', refresh_token);
  
    // get list of notebooks
    ajax(
      {
        url: baseApiUrl + '/me/onenote/notebooks',
        headers: { "Authorization": "Bearer " + access_token }
      },
      function(data) {
        showMenu('Notebooks', data, 'displayName', notebookSelected);
      },
      showError('authorizedCallback')
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
            url: baseApiUrl + '/me/onenote/notebooks/' + notebook_id + '/sections',
            headers: { "Authorization": "Bearer " + access_token }
        },
        function(data) {
          showMenu('Sections', data, 'displayName', sectionSelected);
        },
        showError('notebookSelected')
    );
}

var section_id = '';
function sectionSelected(e, id)
{
    section_id = id;

    // get pages of the section
    ajax(
        {
            url: baseApiUrl + '/me/onenote/sections/' + section_id + '/pages',
            headers: { "Authorization": "Bearer " + access_token }
        },
        function(data) {
          showMenu('Pages', data, 'title', pageSelected);
        },
        showError('sectionSelected')
    );
    
}

var page_id = '';
function pageSelected(e, id)
{
    page_id = id;

    // get contents of the page
    ajax(
        {
            url: baseApiUrl + '/me/onenote/pages/' + page_id + '/content',
            headers: { "Authorization": "Bearer " + access_token }
        },
        pageContentRequestSuccess,
        showError('pageSelected')
    );
    
}

function pageContentRequestSuccess(data)
{
    var pageFontSize = Settings.option('font_size') || 24;
    
    console.log(data);
    
    var title;
    var text = data.replace(/<title>([^<]*)<\/title>/, function(m) { title = m.replace(/<(?:.|\n)*?>/gm, ''); return ''; });
    
    if (Feature.round()) {
        // simplified support for Pebble Round
        new UI.Card({
            title: title,
            body: htmlToText(text),
            scrollable: true
        }).show();
    }
    else {
        var page = new UI.Window({
          backgroundColor: 'white',
          scrollable: true,
          
        });
        
        var yPos = addTextToWindow(page, 0, 0, title, pageFontSize) + 2;
    
        var items = parseHtmlTree(text);
        console.log(JSON.stringify(items[0], null, 4));
        showItem(items[0], page, pageFontSize, yPos);
        
        page.show();
    }
}


function showMenu(title, data, title_property, select_callback)
{
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

function showError(funcName)
{
  return function(response, statusCode, request)
  {
      var main = new UI.Card({
          title: 'Error @' + funcName,
          body: response,
          scrollable: true
      });
      main.show();
  };
}
  
function showText(text)
{
    var main = new UI.Card({
        body: text,
        scrollable: true
    });
    main.show();
    
}


function parseHtmlTree(text) {
    var items = [];
    var listTypes = [];
    var listCounters = [];
    var openedListItems = [];
    var counter = 0;
    text = '<li>' + text + '</li>';
    text = text.replace(/<\/?[uo]l>|<li[^>]*>(?:<span[^>]*>)?|<\/li>/gm, function(m, offset) {
        if (m == '<ul>' || m == '<ol>') {
            listTypes.push(m);
            listCounters.push(counter);
            counter = 0;
            return m;
        }
        if (m == '</ul>' || m == '</ol>') {
            listTypes.pop();
            counter = listCounters.pop();
            return m;
        }
        if (m == '</li>') {
            var closingLi = openedListItems.pop();
            closingLi.endOffset = offset;
            closingLi.text = text.slice(closingLi.startOffset, offset);
            
            for (var i=items.length-1;i>=0;i--) {
                var li = items[i];
                if (li.level == closingLi.level+1 && li.startOffset > closingLi.startOffset)
                {
                    closingLi.children.unshift(li);
                    var start = li.startOffset - closingLi.startOffset;
                    var end = li.endOffset - closingLi.startOffset;
                    console.log(closingLi.text.slice(start,end));
                    closingLi.text = closingLi.text.substr(0, start) + '[---childItem---]' + closingLi.text.substr(end);
                }
            }
            
            closingLi.text = htmlToText(closingLi.text);
            return m;
        }

        counter++;
        var counterMatch = m.match(/^<li value="([A-Za-z0-9]+)"/);
        if (counterMatch && counterMatch.length > 1 && +counterMatch[1])
            counter = +counterMatch[1];

        var tagMatch = m.match(/data\-tag="([A-Za-z,:\-]+)"/);
        var tag = tagMatch && tagMatch.length > 1 ? tagMatch[1] : '';
        if (tag === '' && listTypes.length > 0)
            tag = listTypes[listTypes.length - 1];
        var newListItem = {
            tag: tag,
            level: listTypes.length,
            counter: counter,
            startOffset: offset + m.indexOf('>') + 1,
            children: []
        };
        items.push(newListItem);
        openedListItems.push(newListItem);
        return m;
    });
    
    return items;
}

function showItem(item, page, pageFontSize, yPos)
{
    var texts = item.text.split('[---childItem---]');
    while (texts.length > 0)
    {
        var padding = item.level * 4 + 2;
        var fragmentText = texts.shift();
        
        if (fragmentText.replace(/[\s\n]+/g,'') !== '') {
            yPos = addTextToWindow(page, (item.level === 0 ? 0 : 20 ) + padding, yPos, fragmentText, pageFontSize);
        }

        if (item.children.length > 0) {
            var childItem = item.children.shift();
            if (childItem.tag.indexOf('to-do') != -1) {
                page.add(new UI.Rect({ position: new Vector2(padding, yPos + 4 + (pageFontSize - 18)), size: new Vector2(16, 16), borderColor: 'black' }));
                page.add(new UI.Rect({ position: new Vector2(padding + 1, yPos + 5 + (pageFontSize - 18)), size: new Vector2(14, 14), borderColor: 'black' }));
                if (childItem.tag.indexOf('to-do:completed') != -1)
                    page.add(new UI.Image({ position: new Vector2(padding + 3, yPos + 7 + (pageFontSize - 18)), size: new Vector2(11, 11), image: 'images/check.bmp' }));
            }
            else if (childItem.tag == '<ul>')
                page.add(new UI.Circle({ position: new Vector2(padding + 6, yPos + 10 + (pageFontSize - 18)), radius: 4, borderColor: 'black', backgroundColor: 'black' }));
            else if (childItem.tag == '<ol>')
                page.add(new UI.Text({ position: new Vector2(padding, yPos + (pageFontSize - 18)), size: new Vector2(18, 18), text: childItem.counter + '.', color: 'black', font: 'gothic-18' }));
            
            yPos = showItem(childItem, page, pageFontSize, yPos);
        }
    }
    
    return yPos;
    
}

function htmlToText(text)
{
    text = text.replace(/<!\-\- OutlineGroupNode is not supported \-\->/g, '*** Error: fragment missing (not supported by API). Try adding some text before this node. *** [---br---]');
    text = text.replace(/<(br \/|p[^>]*|\/p|td[^>]*|tr|li[^>]*)>/g, '[---br---]');
    text = text.replace(/<[\s\S]*?>/g, ' ').replace(/\s+/g, ' ').replace(/\[\-\-\-br\-\-\-\]/g,'\n');
    text = text.replace(/[\s\r]*\n[\s\r]*\n[\s\r\n]*/, '\n\n').replace(/[\s\r]*\n[\s\r\n]*/, '\n');
    text = text.replace(/^[\s\r\n]+/m,'').replace(/[\s\r\n]+$/m,'').replace(/[\s\r\n]+\[\-\-\-childItem\-\-\-\]/m,'');
    text = decodeHTMLEntities(text);
    return text;
}

function decodeHTMLEntities(text) {
    var entities = [
        ['quot', '"'],
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

function addTextToWindow(window, xPos, yPos, text, pageFontSize)
{
    do {
        var result = truncateAndCalculateSize(text, pageFontSize, xPos);
        var fragmentSize = result[0];
        var fragmentText = result[1];
        text = result[2];
        var fragmentElement = new UI.Text({ position: new Vector2(xPos, yPos), size: fragmentSize, text: fragmentText, textAlign: 'left', textOverflow: 'wrap', color: 'black', font: 'gothic-' + pageFontSize });
        window.add(fragmentElement);
        console.log('--- text element added at ' + xPos + ':' + yPos + ' ---');
        yPos += fragmentSize.y;
    } while (text);
    return yPos;
}

function truncateAndCalculateSize(string, fontSize, substractWidth) {
    var width = 144 - (substractWidth || 0);
    var lines = string.split('\n');
    var rowsLeft = 30;
    var height = 0;
    var result = '';
    var restOfText = '';
    while (lines.length)
    {
        var paragraph = lines.shift();
        var split = strTruncateWhole(paragraph, fontSize, width);
        if (rowsLeft <= split.length)
        {
            result += split.slice(0,rowsLeft).join('');
            restOfText = split.slice(rowsLeft).join('') + '\n' + lines.join('\n');
            return [new Vector2(width, height + 3), result, restOfText];
        }
        
        rowsLeft -= split.length;
        height += split.length * fontSize;
        result += paragraph + '\n';
    }

    return [new Vector2(width, height + 3), string, ''];
}


function strTruncate(string, fontSize, width) {
    var len = 0;
    var i = 0;
    while (len <= width && i < string.length - 1) {
        len += getCharWidth(string.substr(i, 1), fontSize);
        i++;
    }
    if (len > width) {
        if (len - width > 2)
            i--;
        var result = string.substr(0, i);
        var wrappablePos = Math.max(result.lastIndexOf(' '), result.lastIndexOf('-'));
        if (string.length > i && string.substr(i, 1) != ' ' && wrappablePos > 0)
            result = string.substr(0, wrappablePos);

        console.log('strTruncate: ' + result + ' -> ' + len + ' : ' + i);
        return result;
    }
    return string;
}

function strTruncateWhole(string, fontSize, width) {
    var arr = [];
    var b = 0;
    while (b < string.length) {
        arr.push(strTruncate(string.substring(b), fontSize, width));
        b += arr[arr.length - 1].length;
    }
    return arr;
}

function getCharWidth(ch, fontSize) {
    
    if (fontSize == 18) {
        if (/[\(\)t]/.test(ch))
            return 5;
        else if (/[jiIJl1:;\.,]/.test(ch))
            return 4;
        else if (/[ABCDEFHGKLNOPQRSTUVXYZАБВГДЕЁЗИЙКЛНОПРСТУХЦЧЭЯ]/.test(ch))
            return 8;
        else if (/[Ww]/.test(ch))
            return 10;
        else if (/[MmмшщюМШЩЮ]/.test(ch))
            return 11;
        else if (ch == 'r' || ch == '-')
            return 5;
        else if (ch == ' ')
            return 3;
        else
            return 7;
    }
    else if (fontSize == 24)
    {
        if (/[\(\)t]/.test(ch))
            return 6;
        else if (/[jiIJl1:;\.,]/.test(ch))
            return 4;
        else if (/[ABCDEFHGKLNOPQRSTUVXYZАБВГДЕЁЗИЙКЛНОПРСТУХЦЧЭЯ]/.test(ch))
            return 9;
        else if (/[wW]/.test(ch))
            return 11;
        else if (/[mMмшщюМШЩЮ]/.test(ch))
            return 12;
        else if (ch == 'r' || ch == '-')
            return 5;
        else if (ch == ' ')
            return 4;
        else
            return 8;
    }
    else if (fontSize == 28)
    {
        if (/[\(\)t]/.test(ch))
            return 8;
        else if (/[jiIJl1:;\.,]/.test(ch))
            return 5;
        else if (/[ABCDEFHGKLNOPQRSTUVXYZАБВГДЕЁЗИЙКЛНОПРСТУХЦЧЭЯ]/.test(ch))
            return 13;
        else if (/[wW]/.test(ch))
            return 14;
        else if (/[mMмшщюМШЩЮ]/.test(ch))
            return 15;
        else if (ch == 'r' || ch == '-')
            return 7;
        else if (ch == ' ')
            return 4;
        else
            return 11;
    }
    
    console.log('Error in getCharWidth: fontSize = ' + fontSize + ', char = ' + ch);
    return 8;
}
