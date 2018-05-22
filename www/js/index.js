/*
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 * 
 */

var useV2 = false; // use newer graph API instead of limited Azure API

var authority = "https://login.windows.net/common",
    redirectUri = "https://citizant.sharepoint.com",
    resourceUri = "https://graph.windows.net",
    clientId = "10a2d9f2-3571-4209-be4e-cf65ff348b36",
    graphApiVersion = "2013-11-08";

var app = {
    // Invoked when Cordova is fully loaded.
    onDeviceReady: function () {
        console.log("onDeviceReady");
        document.getElementById('search').addEventListener('click', app.search);
        document.getElementById('userlist').innerHTML = "ready";
        if (useV2) {
            authority = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize";
            redirectUri = "https://citizant.sharepoint.com";
            resourceUri = "https://graph.microsoft.com";
            clientId = "10a2d9f2-3571-4209-be4e-cf65ff348b36";
            graphApiVersion = "v1.0";
        }
    },
    // Implements search operations.
    search: function () {
        console.log("search");
        document.getElementById('userlist').innerHTML = "";

        app.authenticate(function (authresult) {
            var searchText = document.getElementById('searchfield').value;
            app.requestData(authresult, searchText);
        });
    },
    // Shows user authentication dialog if required.
    authenticate: function (authCompletedCallback) {
        console.log("Calling Auth");
        app.context = new Microsoft.ADAL.AuthenticationContext(authority);
        app.context.tokenCache.readItems().then(function (items) {
            if (items.length > 0) {
                authority = items[0].authority;
                app.context = new Microsoft.ADAL.AuthenticationContext(authority);
            }
            // Attempt to authorize user silently
            app.context.acquireTokenSilentAsync(resourceUri, clientId)
                .then(authCompletedCallback, function () {
                    // We require user cridentials so triggers authentication dialog
                    app.context.acquireTokenAsync(resourceUri, clientId, redirectUri)
                        .then(authCompletedCallback, function (err) {
                            app.error("Failed to authenticate: " + err);
                            console.log("Failed to authenticate: " + err);
                        });
                });
        });

    },
    // Makes Api call to receive user list.
    requestData: function (authResult, searchText) {
        var req = new XMLHttpRequest();
        var url = "";

        if (useV2) {
            // graph.windows.com
            var url = resourceUri + "/" + graphApiVersion + "/me/";
        } else {
            // graph.windows.net has no includes or other "like" op. See https://msdn.microsoft.com/en-us/library/azure/ad/graph/howto/azure-ad-graph-api-supported-queries-filters-and-paging-options
            // url = resourceUri + "/" + authResult.tenantId + "/me?api-version=" + graphApiVersion;
            // url = resourceUri + "/" + authResult.tenantId + "/tenantDetails?api-version=" + graphApiVersion;

            url = resourceUri + "/" + authResult.tenantId + "/users?api-version=" + graphApiVersion;

            url = searchText ? url + "&$filter=startswith(displayName,'" + searchText + "')" : url + "&$orderby=displayName&$top=10";
            // url = searchText ? url + "&$filter=mailNickname eq '" + searchText + "'" : url + "&$top=10";
        }

        console.log("C-Calling:" + url);
        req.open("GET", url, true);
        req.setRequestHeader('Authorization', 'Bearer ' + authResult.accessToken);

        req.onload = function (e) {
            if (e.target.status >= 200 && e.target.status < 300) {
                console.log("C-Response:" + e.target.response);
                var data = JSON.parse(e.target.response);
                console.log("C-data:" + Object.prototype.toString.call(data).slice(8, -1) + ":" + data);
                if (!data) {
                    app.error("Unable to parse:" + json);
                    return;
                }
                var dataType = data && Object.keys(data)[0];
                console.log("C-dataType:" + Object.prototype.toString.call(dataType).slice(8, -1) + ":" + dataType);        

                app.renderUserListData(data);
                return;
            }
            app.error('Data request failed: ' + e.target.response);
        };
        req.onerror = function (e) {
            app.error('Data request failed: ' + e.error);
        }

        req.send();
    },
    // Renders user list.
    renderUserListData: function (data) {
        var users = data && data.value;
        console.log("C-users:" + Object.prototype.toString.call(users).slice(8, -1) + ":" + users);
        if (!users || users.length === 0) {
            app.error("No users found");
            return;
        }

        var userlist = document.getElementById('userlist');
        userlist.innerHTML = "";

        // Helper function for generating HTML
        function $new(eltName, classlist, innerText, children, attributes) {
            var elt = document.createElement(eltName);
            classlist.forEach(function (className) {
                elt.classList.add(className);
            });

            if (innerText) {
                elt.innerText = innerText;
            }

            if (children && children.constructor === Array) {
                children.forEach(function (child) {
                    elt.appendChild(child);
                });
            } else if (children instanceof HTMLElement) {
                elt.appendChild(children);
            }

            if (attributes && attributes.constructor === Object) {
                for (var attrName in attributes) {
                    elt.setAttribute(attrName, attributes[attrName]);
                }
            }

            return elt;
        }

        users.map(function (userInfo) {
            return $new('li', ['topcoat-list__item'], null, [
                $new('div', [], null, [
                    $new('p', ['userinfo-label'], 'First name: '),
                    $new('input', ['topcoat-text-input', 'userinfo-data-field'], null, null, {
                        type: 'text',
                        readonly: '',
                        placeholder: '',
                        value: userInfo.givenName || ''
                    })
                ]),
                $new('div', [], null, [
                    $new('p', ['userinfo-label'], 'Last name: '),
                    $new('input', ['topcoat-text-input', 'userinfo-data-field'], null, null, {
                        type: 'text',
                        readonly: '',
                        placeholder: '',
                        value: userInfo.surname || ''
                    })
                ]),
                $new('div', [], null, [
                    $new('p', ['userinfo-label'], 'UPN: '),
                    $new('input', ['topcoat-text-input', 'userinfo-data-field'], null, null, {
                        type: 'text',
                        readonly: '',
                        placeholder: '',
                        value: userInfo.userPrincipalName || ''
                    })
                ]),
                $new('div', [], null, [
                    $new('p', ['userinfo-label'], 'Phone: '),
                    $new('input', ['topcoat-text-input', 'userinfo-data-field'], null, null, {
                        type: 'text',
                        readonly: '',
                        placeholder: '',
                        value: userInfo.telephoneNumber || ''
                    })
                ])
            ]);
        }).forEach(function (userListItem) {
            userlist.appendChild(userListItem);
        });
    },
    log: function (message, isError) {
        isError ? console.error(message) : console.log(message);
        var logItem = document.createElement('li');
        logItem.classList.add("topcoat-list__item");
        isError && logItem.classList.add("error-item");
        var timestamp = '<span class="timestamp">' + new Date().toLocaleTimeString() + ': </span>';
        logItem.innerHTML = (timestamp + message);
        app.logArea.insertBefore(logItem, app.logArea.firstChild);
    },
    // Renders application error.
    error: function (err) {
        var userlist = document.getElementById('userlist');
        //userlist.innerHTML = "Errors:";

        var errorItem = document.createElement('li');
        errorItem.classList.add('topcoat-list__item');
        errorItem.classList.add('error-item');
        errorItem.innerText = err;

        userlist.appendChild(errorItem);
    }
};

document.addEventListener('deviceready', app.onDeviceReady, false);
