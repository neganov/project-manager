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

    var _htmlEnhancers = [
        {
            // TODO Checkbox unchecked
            regex: /(<.+?data-tag="to-do".+?>)/gmi,
            enhancement: '<img class="NoteTagImage" role="checkbox" aria-checked="false" src="https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/3_16_N.16x16x32.png" alt="To Do" title="To Do" style="left: -32px;">'
        },
        {
            // TODO Checkbox checked
            regex: /(<.+?data-tag="to-do:completed".+?>)/gmi,
            enhancement: '<img class="NoteTagImage" role="checkbox" aria-checked="true" src="https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/3_16_M.16x16x32.png" alt="To Do" title="To Do" style="left: -32px;">'
        },
        {
            // Attendence checkbox unchecked
            regex: /(<.+?data-tag="discuss-with-person-a".+?>)/gmi,
            enhancement: '<img class="NoteTagImage" role="checkbox" aria-checked="false" src="https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/94_16_N.16x16x32.png" alt="In Attendance" title="In Attendance" style="left: -32px;">'
        },
        {
            // Attendance checkbox checked
            regex: /(<.+?data-tag="discuss-with-person-a:completed".+?>)/gmi,
            enhancement: '<img class="NoteTagImage" role="checkbox" aria-checked="true" src="https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/94_16_M.16x16x32.png" alt="In Attendance" title="In Attendance" style="left: -32px;">'
        }
    ];

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

    var _init = function () {
        CDNManager.getScript('jquery-1.11.2.min.js', function () {
            CDNManager.getScript('adal.js', function () {
                CDNManager.getScript('jspdf.debug.js', function () {
                        _load();
                });
            });
        });
    };

    var _load = function () {
        var template = _getFormTemplate();
        $('#oneNoteHook').html(template);
        $('#btnSave').click(_save).prop('disabled', true);
        $('#btnPDF').click(_convertPage);
        $('#btnImg').click(_convertPage);
        $('#btnNew').click(_createNewPage);
        $('#htmContent').empty();
        $('#imgContent').empty();

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
            '  #webidtxt, #sectionstxt, #pagetitletxt, #newpagenametxt, #sectionnametxt, #htmltxt {width: 355px;padding-right:5px;}' +
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
            '    <td class="sbstbl">Page name for approval:</td>' +
            '    <td class="sbstbl"><input type="text" id="pagetitletxt" value="RFQ & Risk Assessment Review Meeting (APQP #2)"/></td>' +
            '  </tr>' +
            '  <tr>' +
            '    <td class="sbstbl">New page name:</td>' +
            '    <td class="sbstbl"><input type="text" id="newpagenametxt" value=""/></td>' +
            '  </tr>' +
            '  <tr>' +
            '    <td class="sbstbl">New page section name:</td>' +
            '    <td class="sbstbl"><input type="text" id="sectionnametxt" value="Product and Process Validation Phase"/></td>' +
            '  </tr>' +
            '  <tr>' +
            '    <td class="sbstbl">New page HTML:</td>' +
            '    <td class="sbstbl"><textarea id="htmltxt" cols="50" rows="5"></textarea></td>' +
            '  </tr>' +
            '  <tr>' +
            '    <td class="sbstbl">Status:</td>' +
            '    <td class="sbstbl"><span id="updateStatus"></span></td>' +
            '  </tr>' +
            '  <tr>' +
            '    <td class="btntbl" colspan="2">' +
            '      <input type="button" value="Save" id="btnSave"/>' +
            '      <input type="button" id="btnPDF" value="Generate PDF from the page"/>' +
            '      <input type="button" id="btnImg" value="Approve existing page"/>' +
            '      <input type="button" id="btnNew" value="Create new page"/>' +
            '    </td>' +
            '  </tr>' +
            '</table>';
        return html;
    }

    var _getNewImagePageTemplate = function () {
        var html = 
            '<!DOCTYPE html>' +
            '<html>' +
            '  <head>' +
            '    <title>Title Placeholder</title>' +
		    '    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />' +
		    '    <meta name="created" content="ISO date placeholder" />' +
            '  </head>' +
            '  <body>' +
            '    <img data-render-src="name:EmbeddedHtmlId" width="100%" />' +
            '  </body>' +
            '</html>';
        return html;
    };

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
        }
    }

    var _generatePdf = function (pageContent, token, webId, title, pageId, sectionId) {
        _log('Generating PDF from page content...', true);
        var pdf = new jsPDF('l', 'pt', 'letter');
        var margins = {
            top: 50,
            left: 60,
            width: 720
        };
        pdf.fromHTML(
            pageContent,
            margins.left,
            margins.top,
            {
                'width': margins.width
            },
            function (dispose) {
                pdf.save('page.pdf');
            }
        );
    };

    var _prepareHtml = function (htmlTemplate, pageName, dateString) {
        var html = htmlTemplate.replace(/<title>.+?<\/title>/i, '<title>' + pageName + '</title>');
        html = html.replace(/<meta name="created".+?>/i, '<meta name="created" content="' + dateString + '" />');
        return html;
    }

    var _enhanceHtml = function (html, enhancers, index) {
        if (!index) {
            index = 0;
        }

        if(index < enhancers.length) {
            var enhancedHtml = html.replace(enhancers[index].regex, '$&' + enhancers[index].enhancement);
            return _enhanceHtml(enhancedHtml, enhancers, ++index);
        } else {
            return html;
        }
    }

    var _generateImage = function (htmlContent, token, targetWebId, title, pageId, sectionId) {
        _log('Generating image from page content...', true);
        
        if (token) {
            var htmlPageTemplate = _getNewImagePageTemplate();
            var pageName = 'Read-Only: ' + title;
            var dateString = new Date().toISOString().substring(0, 19) + '-05:00';
            var htmlPage = _prepareHtml(htmlPageTemplate, pageName, dateString);
            var enhancedContent = _enhanceHtml(htmlContent, _htmlEnhancers);
            //var sectionName = $('#sectionnametxt').val();
            //var encodedSectionName = encodeURIComponent(sectionName);
            //var sectionUrl = String.format(
            //    '{0}/{1}/sites/{2}/notes/sections?filter=name eq \'{3}\'',
            //    ONENOTE_API_ROOT_URL,
            //    SITE_COLLECTION_ID,
            //    targetWebId,
            //    encodedSectionName);
            //_get(sectionUrl, token).then(function (result) {
            //    var sections = result.value;
            //    if (sections.length < 1) {
            //        _log('No section with name \'' + sectionName + '\' was found.', true);
            //        return;
            //    } else {
            //        var sectionId = sections[0].id;
                    var newPageUrl = String.format(
                        '{0}/{1}/sites/{2}/notes/sections/{3}/pages',
                        ONENOTE_API_ROOT_URL,
                        SITE_COLLECTION_ID,
                        targetWebId,
                        sectionId);
                    _multipartPost(newPageUrl, htmlPage, enhancedContent, token).then(function () {
                        _log('Snapshot page created succesfully. Deleting original page now...', true);
                    }, function (err) {
                        _log('Failed to generate image from page: ' + err, true);
                    });
                //}
            //}, function (err) {
            //    console.log('Error executing request: ' + err);
            //});
        } else {
            _log('Invalid access token: cannot continue.', true);
        }
    };

    var _createNewPage = function () {
        if (_self.token) {
            var htmlTemplate = $('#htmltxt').val();
            var pageName = $('#newpagenametxt').val();
            var dateString = new Date().toISOString().substring(0, 19) + '-05:00';
            var html = _prepareHtml(htmlTemplate , pageName, dateString);
            var sectionName = $('#sectionnametxt').val();
            var encodedSectionName = encodeURIComponent(sectionName);
            var targetWebId = $('#webidtxt').val();
            var sectionUrl = String.format(
                '{0}/{1}/sites/{2}/notes/sections?filter=name eq \'{3}\'',
                ONENOTE_API_ROOT_URL,
                SITE_COLLECTION_ID,
                targetWebId,
                encodedSectionName);
            _get(sectionUrl, _self.token).then(function (result) {
                var sections = result.value;
                if (sections.length < 1) {
                    _log('No section with name \'' + sectionName + '\' was found.', true);
                    return;
                } else {
                    var sectionId = sections[0].id;

                    var newPageUrl = String.format(
                        '{0}/{1}/sites/{2}/notes/sections/{3}/pages',
                        ONENOTE_API_ROOT_URL,
                        SITE_COLLECTION_ID,
                        targetWebId,
                        sectionId);
                    var request = _post(newPageUrl, html, 'application/xhtml+xml', _self.token);
                    request.then(function () {
                        _log('Page created successfully.', true);
                    }, function (err) {
                        _log('Error executing request: ' + err);
                    });

                }
            }, function (err) {
                console.log('Error executing request: ' + err);
            });
        } else {
            _log('Invalid access token: cannot continue.', true);
        }
    };

    var _convertPage = function (event) {
        if (_self.token) {
            var targetWebId = $('#webidtxt').val();
            if (!targetWebId || 0 == targetWebId.length) {
                _log('No target web Id was specified. Please enter a valid web Id.', true);
                return;
            }
            var pageTitle = $('#pagetitletxt').val();
            if (!pageTitle || 0 == targetWebId.length) {
                _log('No target page title was specified. Please enter a valid page title.', true);
                return;
            }

            var converterFunc = ('btnPDF' == event.target.id)? _generatePdf : _generateImage;
            _extractPageContent(_self.token, targetWebId, pageTitle, converterFunc);
        } else {
            _log('Invalid access token: cannot continue.', true);
        }
    }
    
    var _extractPageContent = function (token, webId, title, converterFunc) {
        _log('Locating target notebook page...', true);
        var pageId = null;
        var sectionId = null;
        var pagesUrl = String.format(
            '{0}/{1}/sites/{2}/notes/pages?select=id,title&expand=parentSection(select=name,id),parentNotebook(select=name)',
            ONENOTE_API_ROOT_URL,
            SITE_COLLECTION_ID,
            webId);
        _get(pagesUrl, token).then(function (response) {
            var pages = response.value;
            var page = null;

            for (var i = 0; i < pages.length; ++i) {
                if (pages[i].title.toLowerCase() == title.toLowerCase()) {
                    page = pages[i];
                    break;
                }
            }

            if (null == page) {
                _log(String.format('Page with title \'{0}\' was not found. Is the web id correct?', title), true);
            } else {
                pageId = page.id;
                sectionId = page.parentSection.id;
            }
        }, function (err) {
            _log('Fetching of notebook pages has failed. Error: ' + err, true);
        }).then(function () {
            if (pageId == null) {
                return;
            }
            _log('About to fetch contents of a page with Id ' + pageId, true);
            var pageUrl = String.format(
                '{0}/{1}/sites/{2}/notes/pages/{3}/content?includeIDs=true,preAuthenticated=true',
                ONENOTE_API_ROOT_URL,
                SITE_COLLECTION_ID,
                webId,
                pageId);
            _get(pageUrl, token).then(function (content) {
                converterFunc(content, token, webId, title, pageId, sectionId);
            }, function (err) {
                _log('Failed to fetch page content. Error: ' + err, true);
            });
        }, function (err) {
            _log('Fetching of notbook page contents has failed. Error: ' + err, true);
        });
    };

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
            var request = _post(requestUrl, body, 'application/json', token);
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

    var _post = function (url, body, ct, token) {
        _log('About to send a "POST" to URL: ' + url);
        return $.ajax({
            type: 'POST',
            url: url,
            data: body,
            dataType: 'json',
            contentType: ct,
            headers: {
                'Authorization': 'Bearer ' + token,
                'Accept': 'application/json'
            }
        });
    };

    var _multipartPost = function (url, presentationHtml, contentHtml, token) {
        _log('About to send a "multipart POST" to URL: ' + url);
        var timeStamp = Math.floor(Date.now() / 1000);
        var boundary = 'bn' + timeStamp;

        var data = '';
        data += '--' + boundary + '\r\n';
        data += 'Content-Disposition:form-data; name="Presentation"\r\n';
        data += 'Content-Type:text/html\r\n';
        data += '\r\n'  // we must have a blank line between metadata and the data.
        data += presentationHtml
        data += '\r\n';
        data += '--' + boundary + '\r\n';
        data += 'Content-Disposition:form-data; name="EmbeddedHtmlId"\r\n';
        data += 'Content-Type:text/html\r\n';
        data += '\r\n'  // we must have a blank line between metadata and the data.
        data += contentHtml
        data += '\r\n';
        data += '--' + boundary + '--';

        var deferred = $.Deferred();
        var xhr = new XMLHttpRequest();
        xhr.addEventListener('load', function (event) {
            _log('Page creation succeeded', true);
            deferred.resolve();
        });
        xhr.addEventListener('error', function (event) {
            _log('Failed to create a page with binary content: ' + event.message, true);
            deferred.reject();
        })
        xhr.open('POST', url);
        xhr.setRequestHeader('Content-Type', 'multipart/form-data; boundary=' + boundary);
        xhr.setRequestHeader('Authorization', 'Bearer ' + token);
        xhr.send(data);
        return deferred.promise();
    };

    var _instance = {
        init: _init 
    };

    return _instance;
}();

litens.notes.init();
