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
api.jsomCUD = (function() {
	// v 1.1
	// it based off jsom
	//by Jed
	//08/16/16
	var privateAPI = {
		sPoint: SP,
		setItemValues: function(context, list, item, columnInfoObj) {
			var toServerValue,
				toServerType,
				columnName,
				columnValue,
				taxColObj,
				taxField;

			for (columnName in columnInfoObj) {
				if (columnInfoObj.hasOwnProperty(columnName)) {
					//default 
					toServerValue = columnInfoObj[columnName].value;
					toServerType = api.getDataType(toServerValue);

					if (toServerType === '[object Object]' && toServerValue !== undefined) {
						//for comparison and constructing toServer
						columnValue = columnInfoObj[columnName].value;

						if (columnValue.termGuid || columnValue.termGuid === null) {
							//single metadata, {termLabel: , termGuid:}
							var taxonomySingle;

							taxColObj = list.get_fields().getByInternalNameOrTitle(columnName);
							taxField = context.castTo(taxColObj, privateAPI.sPoint.Taxonomy.TaxonomyField);
							
							if (columnValue.termGuid === null) {
								//there is no value
								taxonomySingle = null;
								taxField.validateSetValue(item, taxonomySingle);
							} else {
								//there is a value
								taxonomySingle = new privateAPI.sPoint.Taxonomy.TaxonomyFieldValue();
								taxonomySingle.set_label(columnValue.termLabel);
								taxonomySingle.set_termGuid(columnValue.termGuid);
								taxonomySingle.set_wssId(-1);
								taxField.setFieldValueByValue(item, taxonomySingle);
							}
							continue;
						}
						else if (columnValue.multiTerms) {
							//multi metadata, {multiTerms: [{label: '', guid: ''}, {label: '', guid: ''}]}
							taxColObj = list.get_fields().getByInternalNameOrTitle(columnName);
							taxField = context.castTo(taxColObj, privateAPI.sPoint.Taxonomy.TaxonomyField);

							var termPrep = columnValue.multiTerms.map(privateAPI.multiTerms);

							var terms = new privateAPI.sPoint.Taxonomy.TaxonomyFieldValueCollection(termPrep.join(';#'),taxField);

							taxField.setFieldValueByValueCollection(item, terms);
							continue;
						}
						else if (columnValue.choices) {
							//multi choice, {choices: [1,2,3]}
							toServerValue = columnValue.choices;
						}
						else if (columnValue.itemId) {
							//single lookup, {itemId: number}
							toServerValue = new privateAPI.sPoint.FieldLookupValue();
							toServerValue.set_lookupId(columnValue.itemId);
						}
						else if (columnValue.idArray) {
							//multi lookup, {idArray: [1,2,3,4]}
							toServerValue = columnValue.idArray.map(privateAPI.multiLookup);
						}
						else if (columnValue.acct) {
							//person field single, {acct: }  acct can be email or account name
							toServerValue = privateAPI.sPoint.FieldUserValue.fromUser(columnValue.acct);
						}
						else if (columnValue.acctArray) {
							//multi person field, {acctArray: [acct, acct,acct]}  acct can be email or account name
							toServerValue = columnValue.acctArray.map(privateAPI.multiPerson);
						}
						else if (columnValue.url) {
							//picture of hyperlink, {url: , description: }
							toServerValue = new privateAPI.sPoint.FieldUrlValue();
							toServerValue.set_url(columnValue.url);
							toServerValue.set_description(columnValue.description);
						}
					}
					item.set_item(columnName, toServerValue);
				}
			}
		},
		multiLookup: function(item) {
			var lookupValue = new privateAPI.sPoint.FieldLookupValue();
			return lookupValue.set_lookupId(item);
		},
		multiTerms: function(termInfo) {
			//-1;#Mamo|10d05b55-6ae5-413b-9fe6-ff11b9b5767c
			return "-1;#" + termInfo.label + "|" +termInfo.guid;  
		},
		multiPerson: function(item) {

			return privateAPI.sPoint.FieldUserValue.fromUser(item);
		},
		itemLoad: function(context, item, serverArray, currentIndex) {
			item.update();
			serverArray[currentIndex] = item;
			context.load(serverArray[currentIndex]);
		},
		callByType: function(list, data, action) {
			var dataType = api.getDataType(data),
				currentId,
				total,
				ii;

			if (dataType === '[object Array]') {
				total = data.length;
				for (ii = 0; ii < total; ii++) {
					currentId = data[ii];
					if (action === 'delete') {
						privateAPI.deleteItems(list, currentId);
					}
					if (action === 'recycle') {
						privateAPI.recycleItems(list, currentId);
					}
				}
			}
			if (dataType === '[object Number]' && action === 'delete') {
				privateAPI.deleteItems(list, data);
			}
			if (dataType === '[object Number]' && action === 'recycle') {
				privateAPI.recycleItems(list, data);
			}
		},
		deleteItems: function(list, itemId) {
			var listItem = list.getItemById(itemId);  
			listItem.deleteObject();
		},
		recycleItems: function(list, itemId) {
			var listItem = list.getItemById(itemId);  
			listItem.recycle();
		},
	};
	return {
		columnType: {
			slt: 'Single line of text',
			mlt: 'Multiple lines of text',
			num: 'Number',
			currency: 'Currency',
			date: 'Date and Time',
			choice: 'Choice',
			metadata: 'Managed Metadata',
			person: 'Person or Group',
			contentType: 'Content Type',
			yesNo: 'Yes/No',
			lookup: 'Lookup'
		},
		//this sets the value of the column
		ValuePrep: function(type, value) {
			//this function ensures values in fields are what the SP server expects
			//validation happens before you get here
			// Single line of text
			// Multiple lines of text
			// Number
			// Currency
			// Date and Time
			// Choice
			// Managed Metadata
			// Person or Group
			// Content Type
			// Yes/No
			// Lookup
			var valueConst = api.jsomCUD.ValuePrep;
			if (!(this instanceof valueConst)) {
				return new valueConst(type, value);
			}
			this.type = type;
			this.value = value;
		},
		PrepClientData: function(action, info, itemId) {
			//action is create, delete or recycle, update
			// info is object {
			// 	columnName: instance of valueprep
			// }
			var clientConst = api.jsomCUD.PrepClientData;

			if (!(this instanceof clientConst)) {
				return new clientConst(action, info, itemId);
			}

			this.action = action;
			if (info) {
				this.columnInfo = info;
			}
			if (itemId) {
				this.itemId = itemId;
			}
		},
		prepServerData: function(listGUID, siteURL, serverRequest) {
			/*
				serverRequest should be an array of objects
				{
					action: 'update',
					itemId: 3,
					columnInfo: {
						columnName: Valueprep instance,
						columnName: object
					}
				}
			*/
			var requestType = api.getDataType(serverRequest),
				toServer,
				totalItemsForServer,
				ii,
				currentObj,
				listItem,
				itemInfo,
				list,
				action;
			
			if (requestType === '[object Array]') {
				toServer = {};
				toServer.itemArray = [];
				totalItemsForServer = serverRequest.length;

				if (siteURL) {
					toServer.context = new privateAPI.sPoint.ClientContext(siteURL);
				} else {
					toServer.context = new privateAPI.sPoint.ClientContext.get_current();
				}

				if (privateAPI.sPoint.Guid.isValid(listGUID)) {
					list = toServer.context.get_web().get_lists().getById(listGUID);
				} else {
					list = toServer.context.get_web().get_lists().getByTitle(listGUID);
				}

				// create update
				for (ii = 0; ii < totalItemsForServer; ii++) {
					currentObj = serverRequest[ii];

					action = currentObj.action;
					if (!action) {
						// if no action throw error
						api.issue("Server Request with no action!");
					}

					action = action.toLowerCase();

					if (action === 'delete' || action === 'recycle') {
						// delete Items
						// {
						//     action: 'delete' or 'recycle',
						//     itemId: [1,2,3,4] or 3
						// }

						if (!currentObj.itemId) {
							api.issue("You can not delete/remove items without an ID!");
						}
						privateAPI.callByType(list, currentObj.itemId, action);
						continue;
					}
					if (action === 'create') {
						//exp serverRequest 
						// [
						//     {
						//         action: 'create',
						//         columnInfo: {
						//             column1: 'a slt field',
						//             column2: {termLabel: 'florida', termguid: '123-122-3244-234235-3423'}
						//         }
						//     }
						// ]
						listItem = list.addItem(new privateAPI.sPoint.ListItemCreationInformation());
						itemInfo = currentObj.columnInfo;
					}
					if (action === 'update') {
						//for updat exp serverRequest should be an array of objs [{itemid: number, columnInfo: {},{itemid: number, columnInfo: {}] 
						// [
						//     {
						//         itemId: 2,
						//         action: 'update'
						//         columnInfo: {
						//             column1: 'a slt field',
						//             column2: {termLabel: 'florida', termguid: '123-122-3244-234235-3423'}
						//     }
						// ]

						if (!currentObj.itemId) {
							api.issue('You can not update a list item without an ID!');
						}

						listItem = list.getItemById(currentObj.itemId);
						itemInfo = currentObj.columnInfo;
					}
					privateAPI.setItemValues(toServer.context, list, listItem, itemInfo);
					privateAPI.itemLoad(toServer.context, listItem, toServer.itemArray, ii);
				}
			} else if (requestType === '[object Object]') {
				//if an object is passed  recurse with serverRequest correted
				return api.jsomCUD.prepServerData(listGUID, siteURL, [serverRequest]);
			} else {
				//error
				api.issue("Incorrect serverRequest data type.");
			}

			return toServer;    
		}
	};
})();
