/**
    app name pd-spserverjsom
    need to import ajax and use with getUserData
 */
import * as $ from 'jquery';
import {loadSPScript} from 'pd-sputil';

const waitForScriptsReady = function(scriptName) {
    var def = $.Deferred();

    ExecuteOrDelayUntilScriptLoaded(function() {
        return def.resolve('Ready');
    }, scriptName);

    return def.promise();
};
const clearRequestDigest = function() {
    //this function was to clear the web manager when a taxonomy field was on the dom and you couldnt use jsom across site collections
    //the issue seems to be fixed 7/26/16 and i am commenting out the places where it is call in this file
    var manager = Sys.Net.WebRequestManager;
    if (manager._events !== null &&
        manager._events._list !== null) { 
        var invokingRequests = manager._events._list.invokingRequest; 

        while( invokingRequests !== null && invokingRequests.length > 0) 
        { 
            manager.remove_invokingRequest(invokingRequests[0]); 
        } 
    }
};
const jsomToObj = function(spItemCollection) {
    var cleanArray = [],
        itemsToTranform;

    if (spItemCollection.context) {
        itemsToTranform = spItemCollection.listItems;
    } else {
        itemsToTranform = spItemCollection;
    }

    if (itemsToTranform.getEnumerator) {
        var enumerableResponse = itemsToTranform.getEnumerator();

        while (enumerableResponse.moveNext()) {
            cleanArray.push(
                enumerableResponse.get_current().get_fieldValues()
            );
        }

        return cleanArray;
    }

    itemsToTranform.forEach(function(item) {
        cleanArray.push(item.get_fieldValues());
    });
    return cleanArray;
};

export function jsomGetDataFromSearch(props, currentResults) {
    // props {
    //     url: ,
    //     properties: []
    //     query: "EmpPositionNumber=\""+ posNumber + "\"",
    //     sourceId: ,
    //     trimDuplicates: optional,
    //     rowLimit: optional,
    //     startRow: optional,
    // }

    let scriptCheck = null;
    let allResults = currentResults || [];
    let glob = Microsoft.SharePoint.Client;
    if(glob && glob.Search) {
        scriptCheck = $.Deferred().resolve();
    } else {
        scriptCheck = loadSPScript("SP.Search.js");
    }

    return scriptCheck.then(function() {

        let clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext.get_current(),
            keywordQuery = new Microsoft.SharePoint.Client.Search.Query.keywordQuery(clientContext);

        if(!props.startRow) {
			props.startRow = 0;
		}
        if(!props.rowLimit) {
            props.rowLimit = 250
        }

        keywordQuery.set_queryText(props.query);
        keywordQuery.set_sourceId(props.sourceId);
        keywordQuery.set_trimDuplicates(props.trimDuplicates || false);
        keywordQuery.set_startRow(props.startRow);
        keywordQuery.set_rowLimit(props.rowLimit);

        if(props.properties) {
            let propertiesObj = keywordQuery.get_selectProperties();
            props.properties.forEach(item => {propertiesObj.add(item);});
        }
        let searchExecute = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(clientContext);
        let results = searchExecute.executeQuery(keywordQuery);

        return jsomSendDataToServer({
            context: clientContext,
            results: results
        });
    }).then(function(response) {
        let tableData = response.results.get_value(),
		    requestProps = tableData.ResultTables[0],
		    results = requestProps.ResultRows;

        allResults = allResults.concat(results.ResultRows);

        if(requestProps.TotalRows > (props.startRow + requestProps.RowCount)) {
            props.startRow = props.startRow + requestProps.RowCount;
            return jsomGetDataFromSearch(props, allResults);
        } else {
            return allResults;
        }
    });
}
export function jsomListItemRequest(props) {
    //props is obj {url, listId, query, columnsToInclude}
    return waitForScriptsReady('SP.js')
    .then(function() {
        //clearRequestDigest();

        var clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext.get_current(),
            list = clientContext.get_web().get_lists().getById( props.listId ),
            camlQuery = new SP.CamlQuery(),
            pagingSetup,
            listItemCollection;

        if (props.position) {
            //position should be listItems.get_listItemCollectionPosition().get_pagingInfo()
            //to go forwards listItems.get_listItemCollectionPosition().get_pagingInfo()
            //to go backwards previousPagingInfo = "PagedPrev=TRUE&Paged=TRUE&p_ID=" + spItems.itemAt(0).get_item('ID'); 
            pagingSetup = new SP.ListItemCollectionPosition();
            pagingSetup.set_pagingInfo(props.position);
            camlQuery.set_listItemCollectionPosition(pagingSetup);
        }
        if (props.folderRelativeUrl) {
            //server relative url to scope the query, so it will only look in a certain folder
            camlQuery.set_folderServerRelativeUrl(props.folderRelativeUrl);
        }

        camlQuery.set_viewXml(props.query);
        listItemCollection = list.getItems(camlQuery);

        if (props.columnsToInclude) {
            clientContext.load(listItemCollection, 'Include('+ props.columnsToInclude.join(',') +')');
        }else { 
            clientContext.load(listItemCollection);
        }

        

        return jsomSendDataToServer({
            context: clientContext,
            listItems: listItemCollection
        });
    });
}
export function jsomEnsureUser(user, url) {
    //user can be an object or array
    var datatype = Object.prototype.toString.call(user),
        startStringCheck = /^i:0#\.f\|membership\|/,
        verifiedUsers = [],
        usersToVerify,
        def = $.Deferred(),
        context,
        userLogin,
        web,
        temp;

    if (datatype === '[object Object]') {
        usersToVerify = [user];
    }
    if (datatype === '[object Array]') {
        usersToVerify = user;
    }
    if (!usersToVerify) {
        // never got set so the wrong datatype was passed
        throw new Error('an object or array must be the parameter to jsomEnsureUser');
    }
    context = url ? new SP.ClientContext(url) : new SP.ClientContext.get_current();
    web = context.get_web();


    usersToVerify.forEach(function(userData, index) {
        //i:0#.f|membership|
        userLogin = userData.AccountName || userData.WorkEmail;

        if (!startStringCheck.test(userLogin)) {
            userData.AccountName = 'i:0#.f|membership|'+userLogin.toLowerCase();
            userLogin = userData.AccountName;
        }

        temp = web.ensureUser(userLogin);
        verifiedUsers[index] = temp;
        context.load(verifiedUsers[index]);
    });

    jsomSendDataToServer({
        context: context
    }).then(function() {
        var giveBackValue,
            userTemp;

        usersToVerify.forEach(function(user, index) {
            userTemp = verifiedUsers[index]; 
            user.id = userTemp.get_id();
            if (!user.WorkEmail) {
                user.WorkEmail = userTemp.get_email();
            }
            if (!user.PreferredName) {
                user.PreferredName = userTemp.get_title();
            }
        });
        giveBackValue = datatype === '[object Object]' ? usersToVerify[0] : usersToVerify;
        def.resolve(giveBackValue);
    }).fail(function() {
        def.reject();
    });

    return def.promise();
}
export function jsomGetItemsById(props) {
    //props is obj {url, listId || listName, arrayOfIDs, numberToStartAt columnsToInclude}

    return waitForScriptsReady('SP.js')
    .then(function() {
        //clearRequestDigest();

        var clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext.get_current(),
            arrayOfResults = props.previousResults || [],
            totalItemsPerTrip = props.maxPerTrip || 200,
            totalItemsToGet = props.arrayOfIDs.length,
            ii = props.numberToStartAt || 0,
            listItemCollection = [],
            list = clientContext.get_web().get_lists();

        if (props.listId) {
            list = list.getById( props.listId );
        } else {
            list = list.getByTitle( props.listName );
        }

        while (ii < totalItemsToGet) {
            var item = list.getItemById( props.arrayOfIDs[ii] );
            if (props.columnsToInclude) {
                //Include('properties') does not work here;
                clientContext.load (item, props.columnsToInclude);
            }else { 
                clientContext.load(item);
            }
            listItemCollection.push( item );
            
            if ( listItemCollection.length === totalItemsPerTrip ) {
                ii++;
                break;
            } else {
                ii++;
                continue;
            }
        }   

        return jsomSendDataToServer({
            context: clientContext,
            listItems: listItemCollection
        }).then(function(data) {
            var cleanedResults = jsomToObj(data.listItems),
                combinedArray = arrayOfResults.concat( cleanedResults );
                
            if ( ii < totalItemsToGet ) {
                props.numberToStartAt = ii;
                props.previousResults = combinedArray;
                return jsomGetItemsById(props);
            }

            return combinedArray;
        });
    });
}
export function jsomGetFilesByRelativeUrl(props) {
    //props is obj {url, fileRefs, numberToStartAt columnsToInclude}

    return waitForScriptsReady('SP.js')
    .then(function() {
        //clearRequestDigest();

        var clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext.get_current(),
            web = clientContext.get_web(),
            totalItemsToGet = props.fileRefs.length,
            ii = 0,
            fileObjCollection = [];

        while (ii < totalItemsToGet) {
            var file = web.getFileByServerRelativeUrl(props.fileRefs[ii]);
            if (props.columnsToInclude) {
                //Include('properties') does not work here;
                clientContext.load (file, props.columnsToInclude);
            }else { 
                clientContext.load(file);
            }
            fileObjCollection.push( file );
            ii++;
        }   

        return jsomSendDataToServer({
            context: clientContext,
            files: fileObjCollection
        }).then(function(data) {
            return data;
        });
    });
}
export function jsomTaxonomyRequest(termSetID) {
    //item.IsAvailableForTagging
    return loadSPScript('sp.taxonomy.js')
    .then(function() {
        //clearRequestDigest();

        var clientContext = new SP.ClientContext.get_current(),
            taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(clientContext),
            termStore = taxonomySession.get_termStores().getById("5b7c889a745c4087bccb796372e50d36"),
            termSet = termStore.getTermSet(termSetID),
            terms = termSet.getAllTerms();
            
            clientContext.load(terms, 'Include(CustomProperties, Id,'+
                'IsAvailableForTagging, LocalCustomProperties, Name, PathOfTerm)');

        return jsomSendDataToServer({
            context: clientContext,
            terms: terms
        });
    });
}
export function jsomSendDataToServer(serverData) {
    var def = $.Deferred();
            
    serverData.context.executeQueryAsync(
        function() {
            //success
            def.resolve(serverData);
        },
        function() {
            def.reject(arguments);
        }
    ); //end QueryAsync
    return def.promise();
}
export function jsomListItemDataExtractor(listItemCollection) {
    
    return jsomToObj(listItemCollection);
}
