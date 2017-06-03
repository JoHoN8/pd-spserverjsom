(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory(require("jquery"));
	else if(typeof define === 'function' && define.amd)
		define(["jquery"], factory);
	else if(typeof exports === 'object')
		exports["pdspserverjsom"] = factory(require("jquery"));
	else
		root["pdspserverjsom"] = factory(root["$"]);
})(this, function(__WEBPACK_EXTERNAL_MODULE_0__) {
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
/******/ 	// identity function for calling harmony imports with the correct context
/******/ 	__webpack_require__.i = function(value) { return value; };
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
/******/ 	return __webpack_require__(__webpack_require__.s = 2);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_0__;

/***/ }),
/* 1 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (immutable) */ __webpack_exports__["spSaveForm"] = spSaveForm;
/* harmony export (immutable) */ __webpack_exports__["getDataType"] = getDataType;
/* harmony export (immutable) */ __webpack_exports__["elementTagName"] = elementTagName;
/* harmony export (immutable) */ __webpack_exports__["argsConverter"] = argsConverter;
/* harmony export (immutable) */ __webpack_exports__["arrayInsertAtIndex"] = arrayInsertAtIndex;
/* harmony export (immutable) */ __webpack_exports__["arrayRemoveAtIndex"] = arrayRemoveAtIndex;
/* harmony export (immutable) */ __webpack_exports__["encodeAccountName"] = encodeAccountName;
/* harmony export (immutable) */ __webpack_exports__["promiseDelay"] = promiseDelay;
/* harmony export (immutable) */ __webpack_exports__["exportToCSV"] = exportToCSV;
/* harmony export (immutable) */ __webpack_exports__["getPageInfo"] = getPageInfo;
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_jquery__ = __webpack_require__(0);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_jquery___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0_jquery__);
/**
    app name sputil
 */


const processRow = function (row) {
    var finalVal = '';
    for (var j = 0; j < row.length; j++) {
        var innerValue = row[j] === null ? '' : row[j].toString();
        if (row[j] instanceof Date) {
            innerValue = row[j].toLocaleString();
        }
        var result = innerValue.replace(/"/g, '""');
        if (result.search(/("|,|\n)/g) >= 0) {
            result = '"' + result + '"';
        }
        if (j > 0) {
            finalVal += ',';
        }
        finalVal += result;
    }
    return finalVal + '\r\n';
};
const profileProps = ['PreferredName','SPS-JobTitle','WorkPhone','OfficeNumber',
    'WorkEmail','doeaSpecialAccount','SPS-Department','AccountName','SPS-Location',
    'PositionID','Manager','Office', "LastName", "FirstName"];
/* harmony export (immutable) */ __webpack_exports__["profileProps"] = profileProps;


function spSaveForm(formId, saveButtonValue) {
    if (!PreSaveItem()) {return false;}
    if (formId && SPClientForms.ClientFormManager.SubmitClientForm(formId)) {return false;}
    WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions(saveButtonValue, "", true, "", "", false, true));
}
function getDataType(item) {

	return Object.prototype.toString.call(item);
}
function elementTagName(element) {
	var ele;
	if (element instanceof __WEBPACK_IMPORTED_MODULE_0_jquery__) {
		ele = element.prop('tagName');
	}else {
		ele = element.tagName;
	}

	return ele.toLowerCase();
}
function argsConverter(args, startAt) {
	var giveBack = [],
		numberToStartAt,
		total = args.length;
	for (numberToStartAt = startAt || 0; numberToStartAt < total; numberToStartAt++){
		giveBack.push(args[numberToStartAt]);
	  }
	  return giveBack;
}
function arrayInsertAtIndex(array, index) {
	//all items past index will be inserted starting at index number
	var arrayToInsert = Array.prototype.splice.apply(arguments, [2]);
	Array.prototype.splice.apply(array, [index, 0].concat(arrayToInsert));
	return array;
}
function arrayRemoveAtIndex(array, index) {
	Array.prototype.splice.apply(array, [index, 1]);
	return array;
}
function encodeAccountName(acctName) {
	var check = /^i:0\#\.f\|membership\|/,
		formattedName;

	if (check.test(acctName)) {
		formattedName = acctName;
	} else {
		formattedName = 'i:0#.f|membership|' + acctName;
	}

	return encodeURIComponent(formattedName);
}
function promiseDelay(time) {
	var def = __WEBPACK_IMPORTED_MODULE_0_jquery__["Deferred"](),
		amount = time || 5000;

	setTimeout(function() {
		def.resolve();
	}, amount);
	return def.promise();
}
class sesStorage {
	//frontEnd to session Storage
    constructor() {
        this.storageAdaptor = sessionStorage;
    }
	toType(obj) {
		return ({}).toString.call(obj).match(/\s([a-z|A-Z]+)/)[1].toLowerCase();
	}
	getItem(key) {
		var item = this.storageAdaptor.getItem(key);

		try {
			item = JSON.parse(item);
		} catch (e) {}

		return item;
	}
	setItem(key, value) {
		var type = this.toType(value);

		if (/object|array/.test(type)) {
			value = JSON.stringify(value);
		}

		this.storageAdaptor.setItem(key, value);
	}
	removeItem(key) {
		this.storageAdaptor.removeItem(key);
	}
}
/* harmony export (immutable) */ __webpack_exports__["sesStorage"] = sesStorage;

class sublish {
    constructor() {
        this.cache = {};
    }
    publish(id) {
        var args = argsConverter(arguments, 1),
            ii,
            total;
        if (!this.cache[id]) {
            this.cache[id] = [];
        }
        total = this.cache[id].length;
        for (ii=0; ii < total; ii++) {
            this.cache[id][ii].apply(this, args);
        }

    }
    subscribe(id, fn) {
        if (!this.cache[id]) {
            this.cache[id] = [fn];
        } else {
            this.cache[id].push(fn);
        }
    }
    unsubscribe(id, fn) {
        var ii,
            total;
        if (!this.cache[id]) {
            return;
        }
        total = this.cache[id].length;
        for(ii = 0; ii < total; ii++){
            if (this.cache[id][ii] === fn) {
                this.cache[id].splice(ii, 1);
            }
        }
    }
    clear(id) {
        if (!this.cache[id]) {
            return;
        }
        this.cache[id] = [];
    }
}
/* harmony export (immutable) */ __webpack_exports__["sublish"] = sublish;

function exportToCSV(filename, rows) {
    /*
        rows should be
        exportToCsv('export.csv', [
            ['name','description'],	
            ['david','123'],
            ['jona','""'],
            ['a','b'],

        ])
    
    */
    var csvFile = '';
    for (var i = 0; i < rows.length; i++) {
        csvFile += processRow(rows[i]);
    }

    var blob = new Blob([csvFile], { type: 'text/csv;charset=utf-8;' });
    if (navigator.msSaveBlob) { // IE 10+
        navigator.msSaveBlob(blob, filename);
    } else {
        var link = document.createElement("a");
        if (link.download !== undefined) { // feature detection
            // Browsers that support HTML5 download attribute
            var url = URL.createObjectURL(blob);
            link.setAttribute("href", url);
            link.setAttribute("download", filename);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    }
}
function getPageInfo() {
    
    return _spPageContextInfo;
};

/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


Object.defineProperty(exports, "__esModule", {
    value: true
});
exports.jsomGetDataFromSearch = jsomGetDataFromSearch;
exports.jsomListItemRequest = jsomListItemRequest;
exports.jsomEnsureUser = jsomEnsureUser;
exports.jsomGetItemsById = jsomGetItemsById;
exports.jsomGetFilesByRelativeUrl = jsomGetFilesByRelativeUrl;
exports.jsomTaxonomyRequest = jsomTaxonomyRequest;
exports.jsomSendDataToServer = jsomSendDataToServer;
exports.jsomListItemDataExtractor = jsomListItemDataExtractor;

var _jquery = __webpack_require__(0);

var $ = _interopRequireWildcard(_jquery);

var _pdSputil = __webpack_require__(1);

function _interopRequireWildcard(obj) { if (obj && obj.__esModule) { return obj; } else { var newObj = {}; if (obj != null) { for (var key in obj) { if (Object.prototype.hasOwnProperty.call(obj, key)) newObj[key] = obj[key]; } } newObj.default = obj; return newObj; } }

/**
    app name pd-spserverjsom
    need to import ajax and use with getUserData
 */
var waitForScriptsReady = function waitForScriptsReady(scriptName) {
    var def = $.Deferred();

    ExecuteOrDelayUntilScriptLoaded(function () {
        return def.resolve('Ready');
    }, scriptName);

    return def.promise();
};
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
var jsomToObj = function jsomToObj(spItemCollection) {
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

function jsomGetDataFromSearch(props, currentResults) {
    // props {
    //     url: ,
    //     properties: []
    //     query: "EmpPositionNumber=\""+ posNumber + "\"",
    //     sourceId: ,
    //     trimDuplicates: optional,
    //     rowLimit: optional,
    //     startRow: optional,
    // }

    var scriptCheck = null;
    var allResults = currentResults || [];
    var glob = Microsoft.SharePoint.Client;
    if (glob && glob.Search) {
        scriptCheck = $.Deferred().resolve();
    } else {
        scriptCheck = (0, _pdSputil.loadSPScript)("SP.Search.js");
    }

    return scriptCheck.then(function () {

        var clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext.get_current(),
            keywordQuery = new Microsoft.SharePoint.Client.Search.Query.keywordQuery(clientContext);

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

        allResults = allResults.concat(results.ResultRows);

        if (requestProps.TotalRows > props.startRow + requestProps.RowCount) {
            props.startRow = props.startRow + requestProps.RowCount;
            return jsomGetDataFromSearch(props, allResults);
        } else {
            return allResults;
        }
    });
}
function jsomListItemRequest(props) {
    //props is obj {url, listId, query, columnsToInclude}
    return waitForScriptsReady('SP.js').then(function () {
        //clearRequestDigest();

        var clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext.get_current(),
            list = clientContext.get_web().get_lists().getById(props.listId),
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
function jsomEnsureUser(user, url) {
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

    jsomSendDataToServer({
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
        def.resolve(giveBackValue);
    }).fail(function () {
        def.reject();
    });

    return def.promise();
}
function jsomGetItemsById(props) {
    //props is obj {url, listId || listName, arrayOfIDs, numberToStartAt columnsToInclude}

    return waitForScriptsReady('SP.js').then(function () {
        //clearRequestDigest();

        var clientContext = props.url ? new SP.ClientContext(props.url) : new SP.ClientContext.get_current(),
            arrayOfResults = props.previousResults || [],
            totalItemsPerTrip = props.maxPerTrip || 200,
            totalItemsToGet = props.arrayOfIDs.length,
            ii = props.numberToStartAt || 0,
            listItemCollection = [],
            list = clientContext.get_web().get_lists();

        if (props.listId) {
            list = list.getById(props.listId);
        } else {
            list = list.getByTitle(props.listName);
        }

        while (ii < totalItemsToGet) {
            var item = list.getItemById(props.arrayOfIDs[ii]);
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
            var cleanedResults = jsomToObj(data.listItems),
                combinedArray = arrayOfResults.concat(cleanedResults);

            if (ii < totalItemsToGet) {
                props.numberToStartAt = ii;
                props.previousResults = combinedArray;
                return jsomGetItemsById(props);
            }

            return combinedArray;
        });
    });
}
function jsomGetFilesByRelativeUrl(props) {
    //props is obj {url, fileRefs, numberToStartAt columnsToInclude}

    return waitForScriptsReady('SP.js').then(function () {
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
function jsomTaxonomyRequest(termSetID) {
    //item.IsAvailableForTagging
    return (0, _pdSputil.loadSPScript)('sp.taxonomy.js').then(function () {
        //clearRequestDigest();

        var clientContext = new SP.ClientContext.get_current(),
            taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(clientContext),
            termStore = taxonomySession.get_termStores().getById("5b7c889a745c4087bccb796372e50d36"),
            termSet = termStore.getTermSet(termSetID),
            terms = termSet.getAllTerms();

        clientContext.load(terms, 'Include(CustomProperties, Id,' + 'IsAvailableForTagging, LocalCustomProperties, Name, PathOfTerm)');

        return jsomSendDataToServer({
            context: clientContext,
            terms: terms
        });
    });
}
function jsomSendDataToServer(serverData) {
    var def = $.Deferred();

    serverData.context.executeQueryAsync(function () {
        //success
        def.resolve(serverData);
    }, function () {
        def.reject(arguments);
    }); //end QueryAsync
    return def.promise();
}
function jsomListItemDataExtractor(listItemCollection) {

    return jsomToObj(listItemCollection);
}

/***/ })
/******/ ]);
});
//# sourceMappingURL=pd-spserverjsom.js.map