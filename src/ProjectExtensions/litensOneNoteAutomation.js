String.format = function() {
    var s = arguments[0];
    for (var i = 0; i < arguments.length - 1; i++) {       
        var reg = new RegExp("\\{" + i + "\\}", "gm");             
        s = s.replace(reg, arguments[i + 1]);
    }
    return s;
};

Function.registerNamespace('litens');

litens.notes = function () {
    const SITE_COLLECTION_ID = 'f5739ea0-7c8b-4cc6-a896-abe099ac980b';
    const ONENOTE_API_ROOT_URL = 'https://www.onenote.com/api/v1.0/myOrganization/siteCollections';
    var _self = this;
    this.token = null;


    var _log = function (message, logToScreen) {
        if (!message) {
            console.log('_log(): the message is undefined.');
            return;
        }
        console.log(message);
        if (logToScreen) {
            var status = $('#updateStatus');
            if (status) {
                status.text(message);
            } else {
                console.log('_log(): cannot log the message on the screen, as the container element is not found.');
            }
        }
    }

    var _init = function() {
         CDNManager.getScript('jquery-1.11.2.min.js', function() {
            CDNManager.getScript('adal.js', function() {
                _load();
            });
        });
     };

    var _load = function () {
        var template = _getFormTemplate();
        $('#oneNoteHook').html(template);
        $('#btnSave').click(_save).prop('disabled', true);
        $('#btnCancel').click(_cancel);

        _acquireToken().done(function (token) {
            _self.token = token;
            $('#btnSave').prop('disabled', false);
            _log('OpenID Connect access token for the OneNote was acquired successfully from the Azure AD', true);
        }).fail(function () {
            _log('Failed to acquire OpenID Connect token required for the OneNote access', true);
        });
    };

    var _getFormTemplate = function () {
        var html =
            '<h1>OneNote Automation Form</h1>' +
            '<style>' +
            '  .sbstbl {border: 1px solid #CCCCCC; text-align: left; border-collapse:collapse;padding-left:10px;padding-top:4px; padding-bottom:4px;}' +
            '  .btntbl {border-width: 0px; text-align: center; border-collapse:collapse;padding-left:10px;padding-top:8px; padding-bottom:8px;}' +
            '  #webidtxt, #sectionstxt {width: 360px;}' +
            '</style>' +
            '<table id="noteTable" class="sbstbl" width="70%">' +
            '  <tr>' +
            '    <td class="sbstbl">Project website id:</td>' +
            '    <td class="sbstbl"><input type="text" id="webidtxt" value="04d9514b-40ab-41c7-bdcf-5f90669a1d64"/></td>' +
            '  </tr>' +
            '  <tr>' +
            '    <td class="sbstbl">OneNote section names:</td>' +
            '    <td class="sbstbl"><textarea id="sectionstxt" cols="50" rows="5">Strategy Planning and Concept Phase, Preliminary Design and Quotation Phase, Design and Development Phase, Production Readiness Phase, Product and Process Validation Phase, Closure Phase</textarea></td>' +
            '  </tr>' +
            '  <tr>' +
            '    <td class="sbstbl">Status:</td>' +
            '    <td class="sbstbl"><span id="updateStatus"></span></td>' +
            '  </tr>' +
            '  <tr>' +
            '    <td class="btntbl" colspan="2">' +
            '      <input type="button" value="Save" id="btnSave"/>' +
            '      <input type="button" id="btnCancel" value="Cancel"/>' +
            '    </td>' +
            '  </tr>' +
            '</table>';
        return html;
    }

    var _save = function () {
        if (_self.token) {
            var targetWebId = $('#webidtxt').val();
            if (!targetWebId || 0 == targetWebId.length) {
                _log('No target web Id was specified. Please enter a valid web Id.', true);
                return;
            }
            var sectionNamesText = $('#sectionstxt').val();
            if (!sectionNamesText || 0 == sectionNamesText.length) {
                _log('No sections were specified. Please enter one or more section names separated by commas', true);
                return;
            }
            var sectionNames = sectionNamesText.trim().split(/,\s*/);
            _updateNotebook(_self.token, targetWebId, sectionNames);
        } else {
            _log('Invalid access token: cannot continue.', true);
            return;
        }
    }

    var _cancel = function () {
        alert("Cancel!");
    }

    var _acquireToken = function () {
        _log('Loading...', true);
        var deferred = $.Deferred(); 

        _self.config = {
            tenant: 'softforte.onmicrosoft.com',
            clientId: 'c2c82c39-3cbb-4b0a-8f8e-d2a2c8617123',
            postLogoutRedirectUri: window.location.origin,
            endpoints: {
                notesApiUri: 'https://onenote.com/'
            }
        };
        
        var authContext = new AuthenticationContext(_self.config);
        var isCallback = authContext.isCallback(window.location.hash);
        authContext.handleWindowCallback();
        if(isCallback && !authContext.getLoginError()) {
            window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
        }
        
        var user = authContext.getCachedUser();
        if(!user) {
            authContext.login();
        }
        
        authContext.acquireToken(_self.config.endpoints.notesApiUri, function(error, token) {
            if(error || !token) {
                _log('ADAL error: ' + error);
                deferred.reject();
            } else {
                _log('ADAL: token successfully acquired.');
                deferred.resolve(token);
            }
        });
        
        return deferred.promise();
    }

    var _updateNotebook = function (token, webId, sections) {
        _log('Locating project site notebook...', true);
        var notesUrl = String.format(
            '{0}/{1}/sites/{2}/notes/notebooks',
            ONENOTE_API_ROOT_URL,
            SITE_COLLECTION_ID, 
            webId);
        _get(notesUrl, token).done(function (response) {
            var notebooks = response.value;
            if (notebooks && notebooks.length && notebooks.length > 0) {
                var notebook = notebooks[0];                
                _log('Found notebook ID=' + notebook.id);
                _checkSections(token, notesUrl, notebook.id, sections);
            } else {
                _log('No notebooks were found on the project site with webId=' + webId, true);
            }
        }).fail(function (err) {
            _log('Fetching of project site notes has failed. Error: ' + err, true);
        });
    };
    
    var _checkSections = function (token, notesUrl, notebookId, sections) {
        _log('Checking if specified sections already exist...', true);
        var requestUrl = String.format('{0}/{1}/sections', notesUrl, notebookId);
        _get(requestUrl, token).done(function (response) {
            var existingSections = response.value;
            var newSections = [];
            for (var i = 0; i < sections.length; ++i) {
                if (!_contains(existingSections, sections[i])) {
                    newSections.push(sections[i]);
                }
            }
            _addSections(token, requestUrl, newSections);
        }).fail(function () {
            _log('Failed to get existing sections.', true);
        });
    };

    var _addSections = function (token, requestUrl, sections) {
        _log('Adding sections...', true);
        var updateRequests = [];
        for (var i = 0; i < sections.length; ++i) {
            var body = '{"name":"' + sections[i] + '"}';
            var request = _post(requestUrl, body, token);
            updateRequests.push(request);
        }

        $.when.apply($, updateRequests).done(function () {
            _log('Sections were created.', true);
        }).fail(function () {
            _log('Some or all of section creation requests have failed.', true);
        });
    };

    var _contains = function (array, stringValue) {
        for (var i = 0; i < array.length; ++i) {
            if (array[i].name.toLowerCase() == stringValue.toLowerCase()) {
                return true;
            }
        }
        return false;
    };

    var _get = function (url, token) {
        _log('About to send a "GET" to URL: ' + url);
        return $.ajax({
            type: 'GET',
            url: url,
            headers: {
                'Authorization': 'Bearer ' + token,
                'Accept': 'application/json'
            }
        });
    };

    var _post = function (url, body, token) {
        _log('About to send a "POST" to URL: ' + url);
        return $.ajax({
            type: 'POST',
            url: url,
            data: body,
            dataType: 'json',
            contentType: 'application/json',
            headers: {
                'Authorization': 'Bearer ' + token,
                'Accept': 'application/json'
            }
        });
    };

    var _instance = {
        init: _init 
    };
    return _instance;
}();

litens.notes.init();
