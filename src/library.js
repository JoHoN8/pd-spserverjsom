/**
	app name pd-spserverjsom
	need to import ajax and use with getUserData
 */
import * as $ from 'jquery';
import {
	loadSPScript,
	validGuid,
	waitForScriptsReady
} from 'pd-sputil';

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
			props.rowLimit = 250;
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
			var cleanedResults = jsomListItemDataExtractor(data.listItems),
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
export function jsomListItemDataExtractor(spItemCollection) {
	
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
}
export class jsomCUD{
	constructor() {
		this.itemsToSend = [];
		this.userRequests = [];
	}
	_getContext(site) {

		if(site) {
			this.context = new this.sp.ClientContext(this.siteUrl);
		} else {
			this.context = new this.sp.ClientContext.get_current();
		}
	}
	_getList(listId) {
		var isGuid = validGuid(this.listId);

		if (isGuid) {
			this.list = this.context.get_web().get_lists().getById(listId);
		} else {
			this.list = this.context.get_web().get_lists().getByTitle(listId);
		}
	}
	_dataTranmitter() {
		//recursive send loop here, return promise

	}
	_addColumnData(listItem, colData) {

		colData.forEach((colObj) => {

			if (colObj.dataType === 'taxBlank') {

				let taxColObj = this.list.get_fields().getByInternalNameOrTitle(colObj.columnName),
					taxField = this.context.castTo(taxColObj, this.sp.Taxonomy.TaxonomyField);
				
				taxField.validateSetValue(listItem, null);
			}
			else if (colObj.dataType === 'taxSingle') {
				let taxColObj = this.list.get_fields().getByInternalNameOrTitle(colObj.columnName),
					taxField = this.context.castTo(taxColObj, this.sp.Taxonomy.TaxonomyField);

				let taxonomySingle = new this.sp.Taxonomy.TaxonomyFieldValue();
					taxonomySingle.set_label(colObj.termLabel);
					taxonomySingle.set_termGuid(colObj.termGuid);
					taxonomySingle.set_wssId(-1);
					taxField.setFieldValueByValue(listItem, taxonomySingle);
			}
			else if (colObj.dataType === 'taxMulti') {
				let taxColObj = this.list.get_fields().getByInternalNameOrTitle(colObj.columnName),
					taxField = this.context.castTo(taxColObj, this.sp.Taxonomy.TaxonomyField),
					termPrep = colObj.multiTerms.map((termInfo) => {
						//-1;#Mamo|10d05b55-6ae5-413b-9fe6-ff11b9b5767c
						return `-1;# ${termInfo.termLabel}|${termInfo.termGuid}`;  
					});

				let terms =this.sp.Taxonomy.TaxonomyFieldValueCollection(termPrep.join(';#'),taxField);

				taxField.setFieldValueByValueCollection(listItem, terms);
			}
			else if (colObj.dataType === 'choiceMulti') {
				listItem.set_item(colObj.columnName, colObj.choices);
			}
			else if (colObj.dataType === 'lookupSingle') {
				let lookupVal = new this.sp.FieldLookupValue();
				lookupVal.set_lookupId(colObj.itemId);
				listItem.set_item(colObj.columnName, lookupVal);
			}
			else if (colObj.dataType === 'lookupMulti') {
				let multiLookupVal = colObj.idArray.map((ppId) => {
					return this.sp.FieldUserValue.fromUser(ppId);
				});
				listItem.set_item(colObj.columnName, multiLookupVal);
			}
			else if (colObj.dataType === 'ppSingle') {
				let ppVal = this.sp.FieldUserValue.fromUser(colObj.acct);
				listItem.set_item(colObj.columnName, ppVal);
			}
			else if (colObj.dataType === 'ppMulti') {
				let multiPPVal = colObj.acctArray.map((ppId) => {
					let lookupValue = new this.sp.FieldLookupValue();
					return lookupValue.set_lookupId(ppId);
				});
				listItem.set_item(colObj.columnName, multiPPVal);
			}
			else if (colObj.dataType === 'hyperlink') {
				let hyperLink = new this.sp.FieldUrlValue();
				hyperLink.set_url(colObj.url);
				hyperLink.set_description(colObj.description);
				listItem.set_item(colObj.columnName, hyperLink);
			} else if (colObj.dataType === 'simple') {
				listItem.set_item(colObj.columnName, colObj.columnValue);
			}
		}, this);

	}
	_loadItem(listItem) {
		let nextIndex = this.itemsToSend.length;

		listItem.update();
		this.itemsToSend[nextIndex] = listItem;
		this.context.load(this.itemsToSend[nextIndex]);
	}
	_createListItems() {

		this.userRequests.forEach((obj) => {
			let listItem = null;
			if (obj.action === 'create') {
				listItem = new this.sp.ListItemCreationInformation();
				this._addColumnData()._loadItem();
			} else if (obj.action === 'update') {
				listItem = this.list.getItemById(obj.itemId);
				this._addColumnData()._loadItem();
			} else  if (obj.action === 'recycle') {
				listItem = this.list.getItemById(obj.itemId);  
				listItem.recycle();
			} else if (obj.action === 'delete') {
				listItem = this.list.getItemById(obj.itemId);  
				listItem.deleteObject();
			}
		}, this);
	}
	_determineDataType(value) {

		if (value.termGuid === null) {
			//blanks out the field
			//value.termGuid: null
			value.dataType = 'taxBlank';
		} else if (value.termGuid) {
			//adds single value to field
			//value.termLabel: string
			//value.termGuid: string guid
			value.dataType = 'taxSingle';
		} else if (value.multiTerms) {
			//add multiple terms to field
			//value.multiTerms: [{termLabel: '', termGuid: ''}, {termLabel: '', termGuid: ''}]
			value.dataType = 'taxMulti';
		} else if (value.choices) {
			//adds multiple choices to field
			//value.choices: ['one','two','three']
			value.dataType = 'choiceMulti';
		} else if (value.itemId) {
			//adds single value to field
			//value.itemId: number
			value.dataType = 'lookupSingle';
		} else if (value.idArray) {
			//adds multiple items to field
			//value.idArray: [1,2,3]
			value.dataType = 'lookupMulti';
		} else if (value.acct) {
			//adds single employee to field
			//value.acct: someone@onmicrosoft.com  acct can be email or account name
			value.dataType = 'ppSingle';
		} else if (value.acctArray) {
			//adds multiple employees to field
			//value.acctArray: [someone@onmicrosoft.com, someone2@onmicrosoft.com]  acct can be email or account name
			value.dataType = 'ppMulti';
		} else if (value.url) {
			//pictue or hyperlink field
			//value.url: string
			//value.description: string 
			value.dataType = 'hyperlink';
		} else {
			//all other types
			//value.columnValue
			value.dataType = 'simple';
		}
	}
	_addDataType(colInfo) {
		let fixedData = [];

		colInfo.forEach((item) => {
			let copy = Object.assign({}, item);
			this._determineDataType(copy);
			fixedData.push(copy);
		});

		return fixedData;
	}
	/**
	 * Add an item to be sent to a list.
	 * column info is an array of objects that contain the data to construct the list item
	 * if the item is not create then itemId is required
	 * if action is delete or recycle then no columnInfo is needed just the id
	 * 
	 * for create or update column data should be passed as follows
	 * every columnInfo object must contain columnName
	 * if single tax field add termLabel and termGuid to column object
	 * if multi tax field add multiTerms to column object, [{termLabel: '', termGuid: ''}, {termLabel: '', termGuid: ''}]
	 * if multi choice field add choices to column object, ['one','two','three']
	 * if single lookup field add itemId to column object, will contain id number
	 * if multi lookup field add idArray to column object, [1,2,3]
	 * if single person field add account to column object, account is email or account name
	 * if multi person field add accountArray to column object, [someone@onmicrosoft.com, someone2@onmicrosoft.com]
	 * if hyperlink field add url and description to column object
	 * if none of these match your column type then pass the data to be stored as columnValue
	 * @param {string} action create, update, recycle or delete 
	 * @param {number} itemId id of the item to update, recycle or delete 
	 * @param {object[]} columnInfo array of objs to send to the server, there must be a object for each column
	 */
	addItem(action, columnInfo, itemId) {

		let prepedObj = {};

		if (!action) {
			throw new Error('action must be provided to add an object to addItem function');
		}
		prepedObj.action = action.toLowerCase();

		if (action !== 'create' && !itemId) {
			throw new Error('a item id must be provided to update, delete or recycle an item');
		}
		prepedObj.itemId = itemId;

		//after you call this method you will have an array of objs
		//objs will be {columnName: , columnData: , dataType: }
		prepedObj.columnData = this._addDataType(columnInfo);

		this.userRequests.push(prepedObj);
	}
	/**
	 * Sends the data added with addItem method to the server
	 * @param {string} site relative site url 
	 * @param {string} listId guid or title of the list
	 */
	sendToSever(site, listId) {
		//make sure everything is in place before doing process
		let def = $.Deferred(),
			self = this;

		waitForScriptsReady('sp.js')
		.then(() => {
			self.sp = SP;

			self
			._getContext(site)
			._getList(listId)
			._createListItems();

			return self._dataTranmitter();
		}).then((response) => {
			def.resolve(response);
		}).fail((data) => {
			def.reject(data);
		});

		return def.promise();
	}
}
