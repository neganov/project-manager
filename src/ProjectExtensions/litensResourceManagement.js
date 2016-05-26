Function.registerNamespace('litens');

litens.pm = function () {
    var self = this;
    
    self._projects = null;
    self._project = null;
    self._projectDraft = null;
    self._enterpriseResources = null;
    self._context = null;
    self._projectResources = null;
    self._projectResourcesMap = {};
    self._ddlMarkup = '';
    self._fields = null;
    self._firstEnterpriseResource = null;
    self._roleFieldInternalName = null;
    self._rolesEntries = null;
    self._roleMap = {};
    self._suggestedSubstitutionMap = {};
    self._tasks = null;
    self._allAssignments = [];
    self._assignmentsToReplace = [];
    self._substitutionMap = {};
    self._enterpriseResourceInternalNameMap = {};
    self._replacementEnterpriseResources = {};
    self._replacementDraftResources = {};
    self._oldProjectResources = {};
    self._draftAssignments = [];
    self._draftCapacities = [];
    self._draftResourceToEnterpriseResourceIdMap = {};
        
    var _init = function () {
        // Use Scott Hillier's CDN Manager app to inject references. More info:
        // https://www.itunity.com/article/managing-javascript-references-office-365-823
        // https://github.com/OfficeDev/PnP/tree/master/Solutions/Core.CDNManager

        CDNManager.getScript('jquery-1.11.2.min.js', function () {
            CDNManager.getScript('adal.js', function () {
                SP.SOD.registerSod('ps.js', '/sites/pwa/_layouts/15/ps.debug.js');
                SP.SOD.registerSod('sp.js', '/sites/pwa/_layouts/15/sp.debug.js');
                SP.SOD.registerSodDep('ps.js', 'sp.js');
                SP.SOD.executeFunc('ps.js', 'PS.ProjectContext', _loadResourceForm);
            });
        });
    };

    var _loadResourceForm = function () {
        var template = _getFormTemplate();
        $('#substitutionForm').html(template);

        self._context = PS.ProjectContext.get_current();
        self._projects = self._context.get_projects();
        self._enterpriseResources = self._context.get_enterpriseResources();
        self._context.load(self._projects, 'Include(Name, CreatedDate, Id)');
        self._context.load(self._enterpriseResources);
        self._context.executeQueryAsync(
            function () {
                self._firstEnterpriseResource = self._enterpriseResources.itemAt(0);
                self._fields = self._firstEnterpriseResource.get_customFields();
                self._context.load(self._fields);
                self._context.executeQueryAsync(function () {
                    var field = self._fields.get_item(0);
                    self._roleFieldInternalName = field.get_internalName();
                    var rolesTable = field.get_lookupTable();
                    self._rolesEntries = rolesTable.get_entries();
                    self._context.load(self._rolesEntries);
                    self._context.executeQueryAsync(function () {
                        for (var i = 0; self._rolesEntries.itemAt(i) ; ++i) {
                            var entry = self._rolesEntries.itemAt(i)
                            self._roleMap[entry.get_internalName()] = {
                                'id': entry.get_id().toString(),
                                'value': entry.get_fullValue()
                            };
                        }

                        var enumerator = self._enterpriseResources.getEnumerator();
                        self._ddlMarkup += '<option value="">&nbsp;</option>';
                        while (enumerator.moveNext()) {
                            var resource = enumerator.get_current();
                            var internalName = resource.get_item(self._roleFieldInternalName);
                            self._enterpriseResourceInternalNameMap[internalName] = resource;
                            if (!(resource.get_isGeneric())) {
                                var role = self._roleMap[internalName];
                                self._ddlMarkup += ('<option value="' + internalName + '">' + resource.get_name() + '</option>');
                                self._suggestedSubstitutionMap[role.value] = internalName;
                            }
                        }

                        var projectId = _getQueryStringParameter('projuid');
                        self._project = self._projects.getByGuid(projectId);
                        self._projectDraft = self._project.get_draft();
                        self._projectResources = self._projectDraft.get_projectResources();
                        self._context.load(self._projectDraft);
                        self._context.load(self._projectResources);
                        self._context.executeQueryAsync(function () {
                            var enumerator = self._projectResources.getEnumerator();
                            while (enumerator.moveNext()) {
                                var resource = enumerator.get_current();
                                var isGeneric = resource.get_isGenericResource();

                                if (!isGeneric) {
                                    var id = resource.get_id().toString();
                                    self._draftResourceToEnterpriseResourceIdMap[id] = resource;
                                    continue;
                                }

                                var resourceId = resource.get_id().toString();
                                self._projectResourcesMap[resourceId] = resource;
                                var tr = $('<tr><td class="sbstbl">' + resource.get_name() + '</td><td class="sbstbl"><select id="' + resourceId + '">' + self._ddlMarkup + '</select></td></tr>').appendTo('#resourceTable');
                                var internalName = self._suggestedSubstitutionMap[resource.get_name()];
                                if (internalName) {
                                    tr.find('select').val(internalName);
                                }
                            }
                        },
                        function (response, err) {
                            console.log('Failed to load project.');
                        });
                        $('#btnSave').click(_save);
                        $('#btnCancel').click(_cancel);
                    }, function (a, b) {
                        console.log('Failed to load role field entries.');
                    });
                }, function (a, b) {
                    console.log('Failed to load fields for resource.');
                });
            },
            function () {
                console.log('Failed to load initial objects.');
            }
        );
    };

    var _getFormTemplate = function () {
        var html =
            '<h1>Resource Substitution Form</h1>' +
            '<p id="replacementStatus"></p>' +
            '<style>' +
            '  .sbstbl {border: 1px solid #CCCCCC; text-align: left; border-collapse:collapse;padding-left:10px;padding-top:4px; padding-bottom:4px;}' +
            '  .btntbl {border-width: 0px; text-align: right; border-collapse:collapse;padding-left:10px;padding-top:8px; padding-bottom:8px;}' +
            '</style>' +
            '<table id="resourceTable" class="sbstbl" width="70%">' +
            '  <tr>' +
            '    <th class="sbstbl">Generic Project Resource</th>' +
            '    <th class="sbstbl">Enterprise Resource</th>' +
            '  </tr>' +
            '</table>' +
            '<table id="buttonTable" class="btntbl" width="70%">' +
            '  <tr>' +
            '    <td class="btntbl" width="100%">' +
            '      <input type="button" value="Save" id="btnSave"/>' +
            '      <input type="button" id="btnCancel" value="Cancel"/>' +
            '    </td>' +
            '  </tr>' +
            '</table>';
        return html;
    };

    var _replaceResources = function () {
        $('#replacementStatus').text('Preparing to replace resources...');
        self._substitutionMap = {};
        $('select').each(function(index){
            var select = $(this);
            var id = select.attr('id');
            var value = select.val();
            if (value != '') {
                // key=generic project resource id; value=enterprise resource internal name.
                self._substitutionMap[id] = value;
            }
        });
        
        // Build collections of resources to add and to remove.
        for(oldResourceId in self._substitutionMap) {
            var replacementResourceInternalName = self._substitutionMap[oldResourceId];
            var replacementResource = self._enterpriseResourceInternalNameMap[replacementResourceInternalName];
            var oldProjectResource = self._projectResourcesMap[oldResourceId];
            self._replacementEnterpriseResources[oldResourceId] = replacementResource;
            self._oldProjectResources[oldResourceId] = oldProjectResource;
        }
        
        // Build a collection of assignments to replace.
        for(var i = 0; i < self._allAssignments.length; ++i) {
            var assignment = self._allAssignments[i];
            var assignmentResourceId = assignment.get_resource().get_id().toString();
            
            if(self._oldProjectResources[assignmentResourceId]) {
                self._assignmentsToReplace.push(assignment);
            }
        }
                
        // Do the replacements now:
        // 1. Add new resources.
        // 2. Substitute assignments.
        // 3. Remove old project resources.
        // 4. Reload form to only show not yet replaced resources.
        
        var isAddingNewResource = false;
        for(var key in  self._replacementEnterpriseResources) {
            var enterpriseResource = self._replacementEnterpriseResources[key];
            var id = enterpriseResource.get_id().toString();
            var draftResource = self._draftResourceToEnterpriseResourceIdMap[id];

            if (!draftResource) {
                isAddingNewResource = true;
                draftResource = self._projectResources.addEnterpriseResource(enterpriseResource);
                self._draftResourceToEnterpriseResourceIdMap[id] = draftResource;
                self._draftResourceToEnterpriseResourceIdMap[id]['isNew'] = true;
            }

            self._replacementDraftResources[key] = draftResource;
        }
        
        if (isAddingNewResource) {
            $('#replacementStatus').text('Adding enterprise resources to the project...');
            self._addResourcesJob = self._projectDraft.update();
            self._context.waitForQueueAsync(self._addResourcesJob, 30, function (a) {
                console.log('Resource addition result: ' + a);
                // Resources were added, lets hydrate their respective objects.
                for (var key in self._draftResourceToEnterpriseResourceIdMap) {
                    if (self._draftResourceToEnterpriseResourceIdMap[key].isNew) {
                        self._context.load(self._draftResourceToEnterpriseResourceIdMap[key]);
                        self._draftResourceToEnterpriseResourceIdMap[key].isNew = false;
                    }
                }

                self._context.executeQueryAsync(function () {
                    _replaceResources2();
                }, function (a, b) {
                    throw ('Failed to load draft resources returned by addition of enterprise resources to the project');
                });
            });
        } else {
            _replaceResources2();
        }
     }
    
    var _replaceResources2 = function () {
        $('#replacementStatus').text('Removing generic resource assignments from project tasks...');
        for (var i = 0; i < self._assignmentsToReplace.length; ++i) {
            var oldAssignment = self._assignmentsToReplace[i];

            var id = SP.Guid.newGuid().toString();
            var start = oldAssignment.get_start();
            var finish = oldAssignment.get_finish();
            var notes = oldAssignment.get_notes();
            var capacity = oldAssignment.get_resourceCapacity();

            var oldResource = oldAssignment.get_resource();
            var oldResourceId = oldResource.get_id().toString();
            var newDraftResource = self._replacementDraftResources[oldResourceId];
            var resourceId = newDraftResource.get_id().toString();

            var task = oldAssignment.get_task();
            var taskId = task.get_id().toString();

            var info = new PS.AssignmentCreationInformation();
            info.set_id(id);
            info.set_start(start);
            info.set_finish(finish);
            info.set_resourceId(resourceId);
            info.set_taskId(taskId);
            info.set_notes(notes);
            oldAssignment['replacementInfo'] = info;
            self._draftCapacities.push(capacity);

            var assignments = task.get_assignments();
            assignments.remove(oldAssignment);
        }

        self._removeAssignmentsJob = self._projectDraft.update();
        self._context.waitForQueueAsync(self._removeAssignmentsJob, 30, function (a) {
            console.log('Assignments removal result: ' + a);
            $('#replacementStatus').text('Adding enterprise resource assignments to tasks...');

            for (var i = 0; i < self._assignmentsToReplace.length; ++i) {
                var oldAssignment = self._assignmentsToReplace[i];
                var task = oldAssignment.get_task();
                var assignments = task.get_assignments();
                var info = oldAssignment.replacementInfo;
                var draftAssignment = assignments.add(info);
                self._draftAssignments.push(draftAssignment);
            }

            self._assignmenetAdditionJob = self._projectDraft.update();
            self._context.waitForQueueAsync(self._assignmenetAdditionJob, 30, function (a) {
                console.log('Assignments addition result: ' + a);
                $('#replacementStatus').text('Recalculating enterprise resource capacities on assignments...');

                for (var i = 0; i < self._draftAssignments.length; ++i) {
                    try {
                        self._context.load(self._draftAssignments[i]);
                    } catch (ex) {
                        console.log('Failed to load assignment for capacity recalculation, index=' + i + ', original error: ' + ex);
                    }
                }

                self._context.executeQueryAsync(function () {
                    console.log('Recalculating capacities...');
                    for (var i = 0; i < self._draftAssignments.length; ++i) {
                        self._draftAssignments[i].set_resourceCapacity(self._draftCapacities[i]);
                    }

                    self._capacityUpdateJob = self._projectDraft.update();
                    self._context.waitForQueueAsync(self._capacityUpdateJob, 30, function (a) {
                        console.log('Capacities update result: ' + a);
                        $('#replacementStatus').text('Removing generic resources from the project...');
                        for (var key in self._oldProjectResources) {
                            var oldResource = self._oldProjectResources[key];
                            self._projectResources.remove(oldResource);
                        }

                        self._resorceRemovalJob = self._projectDraft.update();
                        self._context.waitForQueueAsync(self._resorceRemovalJob, 30, function (a) {
                            console.log('Resource removal result: ' + a);
                            console.log('Now reloading the resources...');
                            $('#replacementStatus').text('Reloading form...');
                            $('#btnSave').val('Save');

                            _loadResourceForm();
                        });
                    });
                }, function (a, b) {
                    throw ('Failed to load newly created draft assignments');
                });
            });
        });
    };

    var _save = function () {
        $('#replacementStatus').text('Beginning...');
        $('#btnSave').val('Saving...');
        self._tasks = self._projectDraft.get_tasks();
        self._context.load(self._tasks);
        self._context.executeQueryAsync(function(){
            var enumerator = self._tasks.getEnumerator();
            while (enumerator.moveNext()) {
                var task = enumerator.get_current();
                var assignments= task.get_assignments();
                self._context.load(task);
                self._context.load(assignments);
            }            
            self._context.executeQueryAsync(function(){
                var enumerator = self._tasks.getEnumerator();
                while(enumerator.moveNext()) {
                    var task = enumerator.get_current();
                    var taskId = task.get_id().toString();
                    var enum2 = task.get_assignments().getEnumerator();
                    while(enum2.moveNext()){
                        var assignment = enum2.get_current();
                        self._context.load(assignment);
                        self._context.load(assignment.get_resource());
                        self._context.load(assignment.get_task());
                        self._allAssignments.push(assignment);
                    }
                }
                
                self._context.executeQueryAsync(
                    _replaceResources,
                    function (a, b) {
                        throw 'Failed to load assignment details';
                    });
            }, function(a, b){
                throw 'Failed to load task assignments';
            });
        }, function(a, b){
            throw 'Failed to load tasks';
        });
    };
    
    var _cancel = function() {
        alert('cancel!');
    };
       
    var _getQueryStringParameter = function(paramToRetrieve) {
        var params =
            document.URL.split('?')[1].split('&');
        var strParams = '';
        for (var i = 0; i < params.length; i = i + 1) {
            var singleParam = params[i].split('=');
            if (singleParam[0].toLowerCase() == paramToRetrieve.toLowerCase())
                return singleParam[1];
        }
    };
    
    var instance = {
        init: _init
    };
    return instance;
}();

litens.pm.init();
