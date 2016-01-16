var UI = require('ui');
var ajax = require('ajax');

var clientId='';
var clientSecret='';
var auth_code = '';

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
        console.log('selected section!');
        console.log(e);
        section_id = sections_ids[e.itemIndex];
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
        pages_menu.push({
            title: dataObj.value[i].title
        });
        
        pages_ids.push(dataObj.value[i].id);
    }
    var menu = new UI.Menu({
        sections: [
            {
                title: 'Pages',
                items: pages_menu
            }
        ]
    });

    
    menu.on('select', function(e) {
        console.log('selected section!');
        console.log(e);
        page_id = pages_ids[e.itemIndex];
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
        
    menu.show();
    
}

function pageContentRequestSuccess(data)
{
    console.log('page content callback!!!');
    var text = data.replace(/<(?:.|\n)*?>/gm, '');
    console.log(text);
    
    var main = new UI.Card({
        title: 'Page content',
        icon: 'images/menu_icon.png',
        body: text,
        scrollable: true
    });
    main.show();

}

function errorCallback()
{
    var main = new UI.Card({
        title: 'Error',
        icon: 'images/menu_icon.png',
        body: JSON.stringify(arguments),
        scrollable: true
    });
    main.show();
}