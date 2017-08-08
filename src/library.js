/**
	app name pd-spserverjsom
 */
import * as $ from 'jquery';
import {
	loadSPScript,
	validGuid,
	waitForScriptsReady,
	getDataType
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
const jsomListItemDataExtractor = function(spItemCollection) {
	
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
const fromSearchWorker = function(props) {
	let scriptCheck = null;
	let currentResults = props.allResults || [];
	let glob = Microsoft.SharePoint.Client;
	if(glob && glob.Search) {
		scriptCheck = $.Deferred().resolve();
	} else {
		scriptCheck = loadSPScript("SP.Search.js");
	}

	return scriptCheck.then(function() {

		let clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext('/search'),
			keywordQuery = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(clientContext);

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

		if(results.length === 0) {
			props.allResults = [];
		} else {
			props.allResults = currentResults.concat(results);
		}


		if(requestProps.TotalRows > (props.startRow + requestProps.RowCount)) {
			props.startRow = props.startRow + requestProps.RowCount;
			return fromSearchWorker(props);
		} else {
			return props.allResults;
		}
	});
};

/**
 * Retrieves data from the SP search index
 * url is a relative url
 * trimDuplicates, rowLimit and startRow is optional
 * @param {{url:string, query:string, sourceId:string, trimDuplicates:boolean, rowLimit:number, startRow:number}} props
 * @returns {promise}
 */
export function jsomGetDataFromSearch(props) {
	return waitForScriptsReady('sp.js')
	.then(() => {
		return fromSearchWorker(props);
	});
}
/**
 * Retrieves list items based on caml query
 * url is a site relative url
 * either pass listGUID or listTitle not both
 * folderRelativeUrl is optional and is a relative url to scope results to a folder
 * query is a caml query
 * @param {{url:string, listGUID:string, listTitle:string, query:string, columnsToInclude:array, folderRelativeUrl:string}} props
 * @returns {promise}
 */
export function jsomListItemRequest(props) {
	//todo make this function recursive
	return waitForScriptsReady('SP.js')
	.then(function() {

		var clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext.get_current(),
			camlQuery = new SP.CamlQuery(),
			pagingSetup,
			listItemCollection,
			list;

		if (props.listGUID) {
			list = clientContext.get_web().get_lists().getById(props.listGUID);
		} else {
			list = clientContext.get_web().get_lists().getById(props.listTitle);
		}

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
/**
 * Ensures user is in the site collections user information list
 * user is an object or an array of objects that contains AccountName or WorkEmail
 * url is a site relative url
 * @param {string} url
 * @param {object|array} user 
 * @returns {promise}
 */
export function jsomEnsureUser(url, user) {

	var datatype = getDataType(user),
		startStringCheck = /^i:0#\.f\|membership\|/,
		verifiedUsers = [],
		usersToVerify,
		def = $.Deferred(),
		context,
		userLogin,
		web,
		temp;

	if (datatype === 'object') {
		usersToVerify = [user];
	}
	if (datatype === 'array') {
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
/**
 * Gets list items based on id
 * url is a site relative url
 * pass listGUID or listTitle not both
 * @param {{url:string, listGUID:string, listTitle:string, arrayOfIds:number[], columnsToInclude:string[]}} props
 * @returns {promise} 
 */
export function jsomGetItemsById(props) {

	return waitForScriptsReady('SP.js')
	.then(function() {

		var clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext.get_current(),
			currentResults = props.allResults || [],
			totalItemsPerTrip = props.maxPerTrip || 100,
			totalItemsToGet = props.arrayOfIds.length,
			ii = props.numberToStartAt || 0,
			listItemCollection = [],
			list;

		if (props.listGUID) {
			list = clientContext.get_web().get_lists().getById( props.listGUID );
		} else {
			list = clientContext.get_web().get_lists().getByTitle( props.listTitle );
		}

		while (ii < totalItemsToGet) {
			var item = list.getItemById( props.arrayOfIds[ii] );
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
			var cleanedResults = jsomListItemDataExtractor(data.listItems);
				
			props.allResults = currentResults.concat( cleanedResults );
				
			if ( ii < totalItemsToGet ) {
				props.numberToStartAt = ii;
				return jsomGetItemsById(props);
			}

			return props.allResults;
		});
	});
}
/**
 * Retrieves file data by relative url
 * url is a site relative url
 * fileRefs is an array of relative urls of the files
 * columnToInclude is an array of column names that you want to retrieve, optional
 * @param {{url:string, fileRefs:string[], columnsToInclude:string[]}} props
 * @returns {promise}
 */
export function jsomGetFilesByRelativeUrl(props) {

	return waitForScriptsReady('SP.js')
	.then(function() {

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
/**
 * Retrieves term store terms based on termSetId
 * @param {string} termStoreId 
 * @param {string} termSetId
 * @return {promise}
 */
export function jsomTaxonomyRequest(termStoreId, termSetId) {
	//item.IsAvailableForTagging
	return waitForScriptsReady('sp.js')
	.then(() => {
		let tax;

		if(SP.Taxonomy) {
			//already loaded
			tax = $.Deferred().resolve();
		} else {
			tax = loadSPScript('sp.taxonomy.js');
		}
		return tax;
	}).then(function() {
		// termStoreId ex - "5b7c889a745c4087bccb796372e50d36"

		var clientContext = new SP.ClientContext.get_current(),
			taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(clientContext),
			termStore = taxonomySession.get_termStores().getById(termStoreId),
			termSet = termStore.getTermSet(termSetId),
			terms = termSet.getAllTerms();
			
			clientContext.load(terms, 'Include(CustomProperties, Id,'+
				'IsAvailableForTagging, LocalCustomProperties, Name, PathOfTerm)');

		return jsomSendDataToServer({
			context: clientContext,
			terms: terms
		});
	});
}
/**
 * Sends data to server
 * must have a context key
 * there can be a secondary key, ex listItems
 * the object that is passed will be return on resolve
 * @param {{context:object}} serverData 
 * @returns {promise}
 */
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
/**
 * Class that allows for batch create, update, recycle, delete
 */
export class jsomCUD{
	constructor() {
		this.itemsToSend = [];
		this.userRequests = [];
	}
	_getContext(site) {

		if(site) {
			this.context = new this.sp.ClientContext(site);
		} else {
			this.context = new this.sp.ClientContext.get_current();
		}
		return this;
	}
	_getList(listId) {
		var isGuid = validGuid(listId);

		if (isGuid) {
			this.list = this.context.get_web().get_lists().getById(listId);
		} else {
			this.list = this.context.get_web().get_lists().getByTitle(listId);
		}
		return this;
	}
	_dataTranmitter() {
		
		return jsomSendDataToServer({
			context: this.context,
			listItems: this.itemsToSend
		});
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
				let ppVal = this.sp.FieldUserValue.fromUser(colObj.account);
				listItem.set_item(colObj.columnName, ppVal);
			}
			else if (colObj.dataType === 'ppMulti') {
				let multiPPVal = colObj.accountArray.map((ppId) => {
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
		return this;

	}
	_loadItem(listItem) {
		let nextIndex = this.itemsToSend.length;

		listItem.update();
		this.itemsToSend[nextIndex] = listItem;
		this.context.load(this.itemsToSend[nextIndex]);

		return this;
	}
	_createListItems() {

		this.userRequests.forEach((obj) => {
			let listItem = null;
			if (obj.action === 'create') {
				listItem = this.list.addItem(new this.sp.ListItemCreationInformation());
				this._addColumnData(listItem, obj.columnData)._loadItem(listItem);
			} else if (obj.action === 'update') {
				listItem = this.list.getItemById(obj.itemId);
				this._addColumnData(listItem, obj.columnData)._loadItem(listItem);
			} else  if (obj.action === 'recycle') {
				listItem = this.list.getItemById(obj.itemId);  
				listItem.recycle();
			} else if (obj.action === 'delete') {
				listItem = this.list.getItemById(obj.itemId);  
				listItem.deleteObject();
			}
		}, this);

		return this;
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
		} else if (value.account) {
			//adds single employee to field
			//value.account: someone@onmicrosoft.com  account can be email or account name
			value.dataType = 'ppSingle';
		} else if (value.accountArray) {
			//adds multiple employees to field
			//value.accountArray: [someone@onmicrosoft.com, someone2@onmicrosoft.com]  acct can be email or account name
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
	_addItem(action, columnInfo, itemId) {

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
		if (action === 'create' || action === 'update') {
			prepedObj.columnData = this._addDataType(columnInfo);
		}

		this.userRequests.push(prepedObj);
	}
	/**
	 * Adds a create operation to the queue
	 * columnInfo is an array of objects that contain the data to create the list item
	 * 
	 * column data should be passed as follows
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
	 * @param {number} itemId 
	 * @param {object[]} columnInfo
	 */
	createItem(columnInfo) {
		this._addItem('create', columnInfo);
	}
	/**
	 * Adds a update operation to the queue
	 * columnInfo is an array of objects that contain the data to updates the list item
	 * 
	 * column data should be passed as follows
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
	 * @param {number} itemId 
	 * @param {object[]} columnInfo
	 */
	updateItem(itemId, columnInfo) {
		this._addItem('update', columnInfo, itemId);
	}
	/**
	 * Adds a recycle operation to the queue
	 * itemId can be a number or an array of number to recycle
	 * @param {number|number[]} itemId 
	 */
	recycleItem(itemId) {

		let type = getDataType(itemId);

		if(type === 'number') {
			this._addItem('recycle', null, itemId);
		} else if (type === 'array') {
			itemId.forEach((id) => {
				this._addItem('recycle', null, id);
			}, this);
		} else {
			throw new Error('invalid datatype passed to recycle item function');
		}
	}
	/**
	 * Adds a delete operation to the queue
	 * itemId can be a number or an array of number to recycle
	 * WARNING deleting an item skips the recycle bin and is unrecoverable
	 * @param {number|number[]} itemId 
	 */
	deleteItem(itemId) {

		let type = getDataType(itemId);

		if(type === 'number') {
			this._addItem('delete', null, itemId);
		} else if (type === 'array') {
			itemId.forEach((id) => {
				this._addItem('delete', null, id);
			}, this);
		} else {
			throw new Error('invalid datatype passed to delete item function');
		}
	}
	totalRequests() {
		return this.userRequests.length;
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
/**
 * Create list items on a meter so you dont get throttled
 * url is a site relative url
 * pass listGUID or listTitle not both
 * columnInfo is an array of objects that contain the column data to create the list item
 *
 * column data should be passed as follows
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
 * @param {{url:string, listGUID:string, listTitle:string, columnInfo:object[]}} props
 * @returns {promise} 
 */
export function jsomCreateItemsMetered(props) {
	let processData = null;

	if (!props.configured) {
		let defaults = {
			totalPerTrip: 50,
			numberToStartAt: 0,
			totalItems: props.columnInfo.length,
			allItems: [],
			configured: true
		};
		processData = Object.assign({}, defaults, pros);
	} else {
		processData = props;
	}

	let itemCreator = new jsomCUD(),
		index = processData.numberToStartAt;

	for (index; index < processData.totalItems; index++) {
		
		itemCreator.createItem(processData.columnInfo[index]);

		if (itemCreator.totalRequests() === processData.totalPerTrip) {
			index++;
			break;
		}
	}

	return itemCreator.sendToSever(props.url, props.listGUID)
	.then(function(response) {
		let results = response.listItems;
		props.allItems = props.allItems.concat(results);

		if (processData.numberToStartAt < props.totalItems) {
			return jsomCreateItemsMetered(props);
		}
		return props.allItems;
	});
	
}

/**
 * update list items on a meter so you dont get throttled
 * url is a site relative url
 * pass listGUID or listTitle not both
 * columnInfo is an array of objects that contain the column data to create the list item
 *
 * column data should be passed as follows
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
 * @param {{url:string, listGUID:string, listTitle:string, columnInfo:object[]}} props
 * @returns {promise} 
 */
export function jsomCreateItemsMetered(props) {
	let processData = null;

	if (!props.configured) {
		let defaults = {
			totalPerTrip: 50,
			numberToStartAt: 0,
			totalItems: props.columnInfo.length,
			allItems: [],
			configured: true
		};
		processData = Object.assign({}, defaults, pros);
	} else {
		processData = props;
	}

	let itemCreator = new jsomCUD(),
		index = processData.numberToStartAt;

	for (index; index < processData.totalItems; index++) {
		
		itemCreator.createItem(processData.columnInfo[index]);

		if (itemCreator.totalRequests() === processData.totalPerTrip) {
			index++;
			break;
		}
	}

	return itemCreator.sendToSever(props.url, props.listGUID)
	.then(function(response) {
		let results = response.listItems;
		props.allItems = props.allItems.concat(results);

		if (processData.numberToStartAt < props.totalItems) {
			return jsomCreateItemsMetered(props);
		}
		return props.allItems;
	});
	
}

