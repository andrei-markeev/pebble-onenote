var UI = require('ui');
var ajax = require('ajax');

var clientId=''; // put the Azure app id here
var clientSecret=''; // put the Azure app secret here
var auth_code = ''; // put the auth_code here

// auth_code was retrieved manually from the following URL:
// https://login.microsoftonline.com/common/oauth2/authorize?api-version=beta&response_type=code&client_id=your-client-id-goes-here&redirect_uri=your-redirect-url-goes-here&resource=https://graph.microsoft.com&prompt=login
// this is because you cannot authenticate directly on the watch as it doesn't have browser as such
// there is a way to do it through configuration page, but we decided to skip this step because of time shortage

// get access token
ajax({
        url: 'https://login.microsoftonline.com/common/oauth2/token?api-version=beta',
        method: 'POST',
        data: {
            grant_type: 'authorization_code',
            code: auth_code,
            redirect_uri: 'http://markeev.com/pebble/onenote.html',
            client_id: clientId,
            client_secret: clientSecret
        }
    },
    authorizedCallback,
    errorCallback
);

var access_token = '';
function authorizedCallback(data)
{
    var dataObj = JSON.parse(data);
    access_token = dataObj.access_token;
    console.log("returned: " + data);
    
    // get list of notebooks
    ajax(
      {
        url: 'https://graph.microsoft.com/beta/me/notes/notebooks',
        headers: { 
          "Authorization": "Bearer " + access_token
        }
      },
      notebooksRequestSuccess,
      errorCallback
    );
}

var notebook_id = '';
var notebook_title = '';
function notebooksRequestSuccess(data)
{
    console.log("notebooks data: " + data);
    var dataObj = JSON.parse(data);

    var notebooks_menu = [];
    var notebooks_ids = [];
    for (var i=0;i<dataObj.value.length;i++)
    {
        notebooks_menu.push({
            title: dataObj.value[i].name
        });
        
        notebooks_ids.push(dataObj.value[i].id);
    }
    var menu = new UI.Menu({
        sections: [
            {
                title: 'Notebooks',
                items: notebooks_menu
            }
        ]
    });
    
    menu.on('select', function(e) {
        notebook_id = notebooks_ids[e.itemIndex];
        notebook_title = e.item.title;
        
        // get sections of the selected notebook
        ajax(
          {
            url: 'https://graph.microsoft.com/beta/me/notes/notebooks/' + notebook_id + '/sections',
            headers: { 
              "Authorization": "Bearer " + access_token
            }
          },
          sectionsRequestSuccess,
          errorCallback
        );
    });
        
    menu.show();
}

var section_id = '';
function sectionsRequestSuccess(data)
{
    console.log("sections data: " + data);
    var dataObj = JSON.parse(data);
    var sections_menu = [];
    var sections_ids = [];
    for (var i=0;i<dataObj.value.length;i++)
    {
        sections_menu.push({
            title: dataObj.value[i].name
        });
        
        sections_ids.push(dataObj.value[i].id);
    }
    var menu = new UI.Menu({
        sections: [
            {
                title: 'Sections of ' + notebook_title,
                items: sections_menu
            }
        ]
    });
    
    menu.on('select', function(e) {
        section_id = sections_ids[e.itemIndex];
        
        // get pages of the section
        ajax(
          {
            url: 'https://graph.microsoft.com/beta/me/notes/sections/' + section_id + '/pages',
            headers: { 
              "Authorization": "Bearer " + access_token
            }
          },
          pagesRequestSuccess,
          errorCallback
        );
    });
        
    menu.show();
    
}

var page_id = '';
function pagesRequestSuccess(data)
{
    console.log("pages data: " + data);
    var dataObj = JSON.parse(data);
    var pages_menu = [];
    var pages_ids = [];
    for (var i=0;i<dataObj.value.length;i++)
    {
        if (!dataObj.value[i].title)
            continue;
        
        pages_menu.push({
            title: dataObj.value[i].title
        });
        
        pages_ids.push(dataObj.value[i].id);
    }
    var menu2 = new UI.Menu({
        sections: [
            {
                title: 'Pages',
                items: pages_menu
            }
        ]
    });

    menu2.on('select', function(e) {
        page_id = pages_ids[e.itemIndex];
        
        // get contents of the page
        ajax(
          {
            url: 'https://graph.microsoft.com/beta/me/notes/pages/' + page_id + '/content',
            headers: { 
              "Authorization": "Bearer " + access_token
            }
          },
          pageContentRequestSuccess,
          errorCallback
        );
    });
        
    menu2.show();
    
}

function pageContentRequestSuccess(data)
{
    var text = data.replace(/<(?:.|\n)*?>/gm, '');
    text = text.replace(/ [ ]+/g, '');
    console.log(text);
    
    var card = new UI.Card({
        title: 'Page content',
        body: text
    });
    card.show();

}

function errorCallback()
{
    var main = new UI.Card({
        title: 'Error',
        body: JSON.stringify(arguments),
        scrollable: true
    });
    main.show();
}