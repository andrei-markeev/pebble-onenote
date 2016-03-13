var UI = require('ui');
var ajax = require('ajax');
var Settings = require('settings');

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
    showText('Account is not configured. Please visit settings.');
    Pebble.openURL(appUrl);
    return;
}
else
    retrieveNotebooks();

function retrieveNotebooks()
{
  var account_type = Settings.option('account_type');
    
  if (account_type == 'onenote.com') {
      clientId = '000000004818A9A6';
      clientSecret = '';
      auth_url = 'https://login.live.com/oauth20_token.srf';
      baseApiUrl = 'https://www.onenote.com/api/v1.0/';
  } else {
      clientId='127482e0-2e1a-4e80-9b81-8a13f9fb822c'; // Azure app id
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
    var titleMatch = data.match(/<title>([^<]*)<\/title>/);
    var title = titleMatch.length > 1 ? titleMatch[1] : '';
    console.log(data);
    var text = data.replace(/<title>([^<]*)<\/title>/, '').replace(/<(?:.|\n)*?>/gm, '').replace(/\s+/g, ' ');
    
    var card = new UI.Card({
        title: title || 'Untitled page',
        body: text,
        scrollable: true
    });
    card.show();

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