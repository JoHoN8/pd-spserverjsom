(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory(require("pd-sputil"));
	else if(typeof define === 'function' && define.amd)
		define(["pd-sputil"], factory);
	else if(typeof exports === 'object')
		exports["pdspserverjsom"] = factory(require("pd-sputil"));
	else
		root["pdspserverjsom"] = factory(root["pdsputil"]);
})(this, function(__WEBPACK_EXTERNAL_MODULE_1__) {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (immutable) */ __webpack_exports__["jsomGetDataFromSearch"] = jsomGetDataFromSearch;
/* harmony export (immutable) */ __webpack_exports__["jsomListItemRequest"] = jsomListItemRequest;
/* harmony export (immutable) */ __webpack_exports__["jsomEnsureUser"] = jsomEnsureUser;
/* harmony export (immutable) */ __webpack_exports__["jsomGetItemsById"] = jsomGetItemsById;
/* harmony export (immutable) */ __webpack_exports__["jsomGetFilesByRelativeUrl"] = jsomGetFilesByRelativeUrl;
/* harmony export (immutable) */ __webpack_exports__["jsomTaxonomyRequest"] = jsomTaxonomyRequest;
/* harmony export (immutable) */ __webpack_exports__["jsomSendDataToServer"] = jsomSendDataToServer;
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "jsomCUD", function() { return jsomCUD; });
/* harmony export (immutable) */ __webpack_exports__["jsomCreateItemsMetered"] = jsomCreateItemsMetered;
/* harmony export (immutable) */ __webpack_exports__["jsomUpdateItemsMetered"] = jsomUpdateItemsMetered;
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_pd_sputil__ = __webpack_require__(1);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_pd_sputil___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__);
var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

/**
	app name pd-spserverjsom
	requires a polyfill for Object.assign
 */
/*global SP Microsoft Sys*/


var clearRequestDigest = function clearRequestDigest() {
	//this function was to clear the web manager when a taxonomy field was on the dom and you couldnt use jsom across site collections
	//the issue seems to be fixed 7/26/16 and i am commenting out the places where it is call in this file
	var manager = Sys.Net.WebRequestManager;
	if (manager._events !== null && manager._events._list !== null) {
		var invokingRequests = manager._events._list.invokingRequest;

		while (invokingRequests !== null && invokingRequests.length > 0) {
			manager.remove_invokingRequest(invokingRequests[0]);
		}
	}
};
var jsomListItemDataExtractor = function jsomListItemDataExtractor(spItemCollection) {

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
			cleanArray.push(enumerableResponse.get_current().get_fieldValues());
		}

		return cleanArray;
	}

	itemsToTranform.forEach(function (item) {
		cleanArray.push(item.get_fieldValues());
	});
	return cleanArray;
};
var fromSearchWorker = function fromSearchWorker(props) {
	var scriptCheck = null;
	var currentResults = props.allResults || [];
	var glob = Microsoft.SharePoint.Client;
	if (glob && glob.Search) {
		scriptCheck = Promise.resolve();
	} else {
		scriptCheck = Object(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__["loadSPScript"])("SP.Search.js");
	}

	return scriptCheck.then(function () {

		var clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext('/search'),
		    keywordQuery = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(clientContext);

		if (!props.startRow) {
			props.startRow = 0;
		}
		if (!props.rowLimit) {
			props.rowLimit = 250;
		}

		keywordQuery.set_queryText(props.query);
		keywordQuery.set_sourceId(props.sourceId);
		keywordQuery.set_trimDuplicates(props.trimDuplicates || false);
		keywordQuery.set_startRow(props.startRow);
		keywordQuery.set_rowLimit(props.rowLimit);

		if (props.properties) {
			var propertiesObj = keywordQuery.get_selectProperties();
			props.properties.forEach(function (item) {
				propertiesObj.add(item);
			});
		}
		var searchExecute = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(clientContext);
		var results = searchExecute.executeQuery(keywordQuery);

		return jsomSendDataToServer({
			context: clientContext,
			results: results
		});
	}).then(function (response) {
		var tableData = response.results.get_value(),
		    requestProps = tableData.ResultTables[0],
		    results = requestProps.ResultRows;

		if (results.length === 0) {
			props.allResults = [];
		} else {
			props.allResults = currentResults.concat(results);
		}

		if (requestProps.TotalRows > props.startRow + requestProps.RowCount) {
			props.startRow = props.startRow + requestProps.RowCount;
			return fromSearchWorker(props);
		} else {
			return props.allResults;
		}
	});
};
var depCheck = function depCheck() {
	try {
		var dep1 = Object.assign,
		    dep2 = Promise;
	} catch (error) {
		throw new Error("The pd-spserverjsom library requires a polyfill for Object.assign and Promise. Please add both to continue.");
	}
};
depCheck();

/**
 * Retrieves data from the SP search index
 * url is a relative url
 * trimDuplicates, rowLimit and startRow is optional
 * @param {{url:string, query:string, sourceId:string, trimDuplicates:boolean, rowLimit:number, startRow:number}} props
 * @returns {promise}
 */
function jsomGetDataFromSearch(props) {
	return Object(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__["waitForScriptsReady"])('sp.js').then(function () {
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
function jsomListItemRequest(props) {
	//todo make this function recursive
	return Object(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__["waitForScriptsReady"])('SP.js').then(function () {

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
			clientContext.load(listItemCollection, 'Include(' + props.columnsToInclude.join(',') + ')');
		} else {
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
function jsomEnsureUser(url, user) {

	var datatype = Object(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__["getDataType"])(user),
	    startStringCheck = /^i:0#\.f\|membership\|/,
	    verifiedUsers = [],
	    usersToVerify,
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

	usersToVerify.forEach(function (userData, index) {
		//i:0#.f|membership|
		userLogin = userData.AccountName || userData.WorkEmail;

		if (!startStringCheck.test(userLogin)) {
			userData.AccountName = 'i:0#.f|membership|' + userLogin.toLowerCase();
			userLogin = userData.AccountName;
		}

		temp = web.ensureUser(userLogin);
		verifiedUsers[index] = temp;
		context.load(verifiedUsers[index]);
	});

	return jsomSendDataToServer({
		context: context
	}).then(function () {
		var giveBackValue, userTemp;

		usersToVerify.forEach(function (user, index) {
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
		return giveBackValue;
	});
}
/**
 * Gets list items based on id
 * url is a site relative url
 * pass listGUID or listTitle not both
 * @param {{url:string, listGUID:string, listTitle:string, arrayOfIds:number[], columnsToInclude:string[]}} props
 * @returns {promise} 
 */
function jsomGetItemsById(props) {

	return Object(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__["waitForScriptsReady"])('SP.js').then(function () {

		var clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext.get_current(),
		    currentResults = props.allResults || [],
		    totalItemsPerTrip = props.maxPerTrip || 100,
		    totalItemsToGet = props.arrayOfIds.length,
		    ii = props.numberToStartAt || 0,
		    listItemCollection = [],
		    list;

		if (props.listGUID) {
			list = clientContext.get_web().get_lists().getById(props.listGUID);
		} else {
			list = clientContext.get_web().get_lists().getByTitle(props.listTitle);
		}

		while (ii < totalItemsToGet) {
			var item = list.getItemById(props.arrayOfIds[ii]);
			if (props.columnsToInclude) {
				//Include('properties') does not work here;
				clientContext.load(item, props.columnsToInclude);
			} else {
				clientContext.load(item);
			}
			listItemCollection.push(item);

			if (listItemCollection.length === totalItemsPerTrip) {
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
		}).then(function (data) {
			var cleanedResults = jsomListItemDataExtractor(data.listItems);

			props.allResults = currentResults.concat(cleanedResults);

			if (ii < totalItemsToGet) {
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
function jsomGetFilesByRelativeUrl(props) {

	return Object(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__["waitForScriptsReady"])('SP.js').then(function () {

		var clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext.get_current(),
		    web = clientContext.get_web(),
		    totalItemsToGet = props.fileRefs.length,
		    ii = 0,
		    fileObjCollection = [];

		while (ii < totalItemsToGet) {
			var file = web.getFileByServerRelativeUrl(props.fileRefs[ii]);
			if (props.columnsToInclude) {
				//Include('properties') does not work here;
				clientContext.load(file, props.columnsToInclude);
			} else {
				clientContext.load(file);
			}
			fileObjCollection.push(file);
			ii++;
		}

		return jsomSendDataToServer({
			context: clientContext,
			files: fileObjCollection
		}).then(function (data) {
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
function jsomTaxonomyRequest(termStoreId, termSetId) {
	//item.IsAvailableForTagging
	return Object(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__["waitForScriptsReady"])('sp.js').then(function () {
		var tax = void 0;

		if (SP.Taxonomy) {
			//already loaded
			tax = Promise.resolve();
		} else {
			tax = Object(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__["loadSPScript"])('sp.taxonomy.js');
		}
		return tax;
	}).then(function () {
		// termStoreId ex - "5b7c889a745c4087bccb796372e50d36"

		var clientContext = new SP.ClientContext.get_current(),
		    taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(clientContext),
		    termStore = taxonomySession.get_termStores().getById(termStoreId),
		    termSet = termStore.getTermSet(termSetId),
		    terms = termSet.getAllTerms();

		clientContext.load(terms, 'Include(CustomProperties, Id,' + 'IsAvailableForTagging, LocalCustomProperties, Name, PathOfTerm)');

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
function jsomSendDataToServer(serverData) {
	return new Promise(function (resolve, reject) {

		serverData.context.executeQueryAsync(function () {
			//success
			resolve(serverData);
		}, function () {
			reject(arguments);
		}); //end QueryAsync
	});
}
/**
 * Class that allows for batch create, update, recycle, delete
 */
var jsomCUD = function () {
	function jsomCUD() {
		_classCallCheck(this, jsomCUD);

		this.itemsToSend = [];
		this.userRequests = [];
	}

	_createClass(jsomCUD, [{
		key: '_getContext',
		value: function _getContext(site) {

			if (site) {
				this.context = new this.sp.ClientContext(site);
			} else {
				this.context = new this.sp.ClientContext.get_current();
			}
			return this;
		}
	}, {
		key: '_getList',
		value: function _getList(listId) {
			var isGuid = Object(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__["validGuid"])(listId);

			if (isGuid) {
				this.list = this.context.get_web().get_lists().getById(listId);
			} else {
				this.list = this.context.get_web().get_lists().getByTitle(listId);
			}
			return this;
		}
	}, {
		key: '_dataTranmitter',
		value: function _dataTranmitter() {

			return jsomSendDataToServer({
				context: this.context,
				listItems: this.itemsToSend
			});
		}
	}, {
		key: '_addColumnData',
		value: function _addColumnData(listItem, colData) {
			var _this = this;

			colData.forEach(function (colObj) {

				if (colObj.dataType === 'taxBlank') {

					var taxColObj = _this.list.get_fields().getByInternalNameOrTitle(colObj.columnName),
					    taxField = _this.context.castTo(taxColObj, _this.sp.Taxonomy.TaxonomyField);

					taxField.validateSetValue(listItem, null);
				} else if (colObj.dataType === 'taxSingle') {
					var _taxColObj = _this.list.get_fields().getByInternalNameOrTitle(colObj.columnName),
					    _taxField = _this.context.castTo(_taxColObj, _this.sp.Taxonomy.TaxonomyField);

					var taxonomySingle = new _this.sp.Taxonomy.TaxonomyFieldValue();
					taxonomySingle.set_label(colObj.termLabel);
					taxonomySingle.set_termGuid(colObj.termGuid);
					taxonomySingle.set_wssId(-1);
					_taxField.setFieldValueByValue(listItem, taxonomySingle);
				} else if (colObj.dataType === 'taxMulti') {
					var _taxColObj2 = _this.list.get_fields().getByInternalNameOrTitle(colObj.columnName),
					    _taxField2 = _this.context.castTo(_taxColObj2, _this.sp.Taxonomy.TaxonomyField),
					    termPrep = colObj.multiTerms.map(function (termInfo) {
						//-1;#Mamo|10d05b55-6ae5-413b-9fe6-ff11b9b5767c
						return '-1;# ' + termInfo.termLabel + '|' + termInfo.termGuid;
					});

					var terms = _this.sp.Taxonomy.TaxonomyFieldValueCollection(termPrep.join(';#'), _taxField2);

					_taxField2.setFieldValueByValueCollection(listItem, terms);
				} else if (colObj.dataType === 'choiceMulti') {
					listItem.set_item(colObj.columnName, colObj.choices);
				} else if (colObj.dataType === 'lookupSingle') {
					var lookupVal = new _this.sp.FieldLookupValue();
					lookupVal.set_lookupId(colObj.itemId);
					listItem.set_item(colObj.columnName, lookupVal);
				} else if (colObj.dataType === 'lookupMulti') {
					var multiLookupVal = colObj.idArray.map(function (ppId) {
						return _this.sp.FieldUserValue.fromUser(ppId);
					});
					listItem.set_item(colObj.columnName, multiLookupVal);
				} else if (colObj.dataType === 'ppSingle') {
					var ppVal = _this.sp.FieldUserValue.fromUser(colObj.account);
					listItem.set_item(colObj.columnName, ppVal);
				} else if (colObj.dataType === 'ppMulti') {
					var multiPPVal = colObj.accountArray.map(function (ppId) {
						var lookupValue = new _this.sp.FieldLookupValue();
						return lookupValue.set_lookupId(ppId);
					});
					listItem.set_item(colObj.columnName, multiPPVal);
				} else if (colObj.dataType === 'hyperlink') {
					var hyperLink = new _this.sp.FieldUrlValue();
					hyperLink.set_url(colObj.url);
					hyperLink.set_description(colObj.description);
					listItem.set_item(colObj.columnName, hyperLink);
				} else if (colObj.dataType === 'simple') {
					listItem.set_item(colObj.columnName, colObj.columnValue);
				}
			}, this);
			return this;
		}
	}, {
		key: '_loadItem',
		value: function _loadItem(listItem) {
			var nextIndex = this.itemsToSend.length;

			listItem.update();
			this.itemsToSend[nextIndex] = listItem;
			this.context.load(this.itemsToSend[nextIndex]);

			return this;
		}
	}, {
		key: '_createListItems',
		value: function _createListItems() {
			var _this2 = this;

			this.userRequests.forEach(function (obj) {
				var listItem = null;
				if (obj.action === 'create') {
					listItem = _this2.list.addItem(new _this2.sp.ListItemCreationInformation());
					_this2._addColumnData(listItem, obj.columnData)._loadItem(listItem);
				} else if (obj.action === 'update') {
					listItem = _this2.list.getItemById(obj.itemId);
					_this2._addColumnData(listItem, obj.columnData)._loadItem(listItem);
				} else if (obj.action === 'recycle') {
					listItem = _this2.list.getItemById(obj.itemId);
					listItem.recycle();
				} else if (obj.action === 'delete') {
					listItem = _this2.list.getItemById(obj.itemId);
					listItem.deleteObject();
				}
			}, this);

			return this;
		}
	}, {
		key: '_determineDataType',
		value: function _determineDataType(value) {

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
	}, {
		key: '_addDataType',
		value: function _addDataType(colInfo) {
			var _this3 = this;

			var fixedData = [];

			colInfo.forEach(function (item) {
				var copy = Object.assign({}, item);
				_this3._determineDataType(copy);
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

	}, {
		key: '_addItem',
		value: function _addItem(action, columnInfo, itemId) {

			var prepedObj = {};

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
   * @param {object[]} columnInfo
   */

	}, {
		key: 'createItem',
		value: function createItem(columnInfo) {
			this._addItem('create', columnInfo);
		}
		/**
   * Adds a update operation to the queue<br>
   * columnInfo is an array of objects that contain the data to updates the list item<br>
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

	}, {
		key: 'updateItem',
		value: function updateItem(itemId, columnInfo) {
			this._addItem('update', columnInfo, itemId);
		}
		/**
   * Adds a recycle operation to the queue
   * itemId can be a number or an array of number to recycle
   * @param {number|number[]} itemId 
   */

	}, {
		key: 'recycleItem',
		value: function recycleItem(itemId) {
			var _this4 = this;

			var type = Object(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__["getDataType"])(itemId);

			if (type === 'number') {
				this._addItem('recycle', null, itemId);
			} else if (type === 'array') {
				itemId.forEach(function (id) {
					_this4._addItem('recycle', null, id);
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

	}, {
		key: 'deleteItem',
		value: function deleteItem(itemId) {
			var _this5 = this;

			var type = Object(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__["getDataType"])(itemId);

			if (type === 'number') {
				this._addItem('delete', null, itemId);
			} else if (type === 'array') {
				itemId.forEach(function (id) {
					_this5._addItem('delete', null, id);
				}, this);
			} else {
				throw new Error('invalid datatype passed to delete item function');
			}
		}
	}, {
		key: 'totalRequests',
		value: function totalRequests() {
			return this.userRequests.length;
		}
		/**
   * Sends the data added with addItem method to the server
   * @param {string} site relative site url 
   * @param {string} listId guid or title of the list
   */

	}, {
		key: 'sendToSever',
		value: function sendToSever(site, listId) {
			//make sure everything is in place before doing process
			var self = this;
			return Object(__WEBPACK_IMPORTED_MODULE_0_pd_sputil__["waitForScriptsReady"])('sp.js').then(function () {
				self.sp = SP;

				self._getContext(site)._getList(listId)._createListItems();

				return self._dataTranmitter();
			});
		}
	}]);

	return jsomCUD;
}();
/**
 * Create list items on a meter so you dont get throttled
 * url is a site relative url
 * pass listGUID or listTitle not both
 * columnInfo is an array of arrays. the inner array contains 
 * objects that contain the column data to create the list item
 * ex. [
 * 		[{columnName: "some", columnValue: 3}], <--1 item created
 * 		[{columnName: "something", columnValue: 8}] <-- 2 item created
 * 		]
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
function jsomCreateItemsMetered(props) {
	var processData = null;

	if (!props.configured) {
		var defaults = {
			totalPerTrip: 50,
			numberToStartAt: 0,
			totalItems: props.columnInfo.length,
			allItems: [],
			configured: true
		};
		processData = Object.assign({}, defaults, props);
	} else {
		processData = props;
	}

	var itemCreator = new jsomCUD(),
	    index = processData.numberToStartAt;

	for (index; index < processData.totalItems; index++) {

		itemCreator.createItem(processData.columnInfo[index]);

		var setupToCreate = itemCreator.totalRequests();
		if (setupToCreate === processData.totalPerTrip || index === processData.totalItems) {
			index++;
			break;
		}
	}
	processData.numberToStartAt = index;

	return itemCreator.sendToSever(processData.url, processData.listGUID).then(function (response) {
		var results = response.listItems;
		processData.allItems = processData.allItems.concat(results);

		if (processData.numberToStartAt < processData.totalItems) {
			return jsomCreateItemsMetered(processData);
		}
		return processData.allItems;
	});
}

/**
 * update list items on a meter so you dont get throttled
 * url is a site relative url
 * pass listGUID or listTitle not both
 * updateInfo is an array of objects that contain the column data and item id to update
 * ex [
 * 		{itemId: 3, columnInfo: [{columnName: "col1", columnValue: "uuumm"}]}
 * 		{itemId: 5, columnInfo: [{columnName: "col3", columnValue: "woohoo"}]}
 * 		]
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
 * @param {{url:string, listGUID:string, listTitle:string, updateInfo:object[]}} props
 * @returns {promise} 
 */
function jsomUpdateItemsMetered(props) {
	var processData = null;

	if (!props.configured) {
		var defaults = {
			totalPerTrip: 50,
			numberToStartAt: 0,
			totalItems: props.updateInfo.length,
			allItems: [],
			configured: true
		};
		processData = Object.assign({}, defaults, props);
	} else {
		processData = props;
	}

	var itemCreator = new jsomCUD(),
	    index = processData.numberToStartAt;

	for (index; index < processData.totalItems; index++) {
		var current = processData.updateInfo[index];
		itemCreator.updateItem(current.itemId, current.columnInfo);

		var setupToCreate = itemCreator.totalRequests();
		if (setupToCreate === processData.totalPerTrip || index === processData.totalItems) {
			index++;
			break;
		}
	}
	processData.numberToStartAt = index;

	return itemCreator.sendToSever(processData.url, processData.listGUID).then(function (response) {
		var results = response.listItems;
		processData.allItems = processData.allItems.concat(results);

		if (processData.numberToStartAt < processData.totalItems) {
			return jsomUpdateItemsMetered(processData);
		}
		return processData.allItems;
	});
}

/***/ }),
/* 1 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_1__;

/***/ })
/******/ ]);
});