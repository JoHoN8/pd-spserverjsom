// Type definitions for [~THE LIBRARY NAME~] [~OPTIONAL VERSION NUMBER~]
// Project: [~THE PROJECT NAME~]
// Definitions by: [~YOUR NAME~] <[~A URL FOR YOU~]>

/*~ This is the module template file for function modules.
 *~ You should rename it to index.d.ts and place it in a folder with the same name as the module.
 *~ For example, if you were writing a file for "super-greeter", this
 *~ file should be 'super-greeter/index.d.ts'
 */

/*~ Note that ES6 modules cannot directly export callable functions.
 *~ This file should be imported using the CommonJS-style:
 *~   import x = require('someLibrary');
 *~
 *~ Refer to the documentation to understand common
 *~ workarounds for this limitation of ES6 modules.
 */

/*~ If this module is a UMD module that exposes a global variable 'myFuncLib' when
 *~ loaded outside a module loader environment, declare that global here.
 *~ Otherwise, delete this declaration.
 */
export as namespace pdspserverjsom;

//interfaces
interface anyOject {
	[key:string]: any
}
declare interface toServerData {
	context: any
}
declare interface props {
	url: string
}
declare interface listPropsBase extends props {
	query: string,
	columnsToInclude?: string[],
	folderRelativeUrl?: string
}
declare interface listPropGuid extends listPropsBase {
	listGUID: string
}
declare interface listPropTitle extends listPropsBase {
	listTitle: string
}
declare interface searchProps extends props {
	query: string,
	sourceId: string,
	trimDuplicates?: boolean,
	rowLimit?: number,
	startRow?: number
}
declare interface userObjAccount extends anyOject {
	AcountName: string
}
declare interface userObjEmail extends anyOject {
	WorkEmail: string
}
declare interface listItemById_base extends props {
	arrayOfIds: number[],
	columnsToInclude: string[]
}
declare interface listItemById_Guid extends listItemById_base {
	listGUID: string,
}
declare interface listItemById_Title extends listItemById_base {
	listTitle: string,
}
declare interface getFiles extends props {
	fileRefs: string[],
	columnsToInclude: string[]
}
declare interface taxReturn extends toServerData {
	terms: any[]
}
declare interface metered_obj {
	columnName: string,
	columnValue: any
}
declare interface id_obj {
	itemId: number
}
declare interface updateObj extends id_obj {
	columnInfo: metered_obj[]
}

declare interface metered_guid extends props {
	listGUID: string
}
declare interface metered_title extends props {
	listTitle: string
}
declare interface metered_create_guid extends metered_guid {
	columnInfo: metered_obj[][]
}
declare interface metered_create_title extends metered_title {
	columnInfo: metered_obj[][]
}
declare interface metered_update_guid extends metered_guid {
	updateInfo: updateObj[][]
}
declare interface metered_update_title extends metered_title {
	updateInfo: updateObj[][]
}
declare interface metered_recycle_guid extends metered_guid {
	recycleIds: number[]
}
declare interface metered_recycle_title extends metered_title {
	recycleIds: number[]
}


export function jsomGetDataFromSearch(props:searchProps): Promise<toServerData>;
export function jsomListItemRequest(props:listPropGuid | listPropTitle): Promise<toServerData>;
export function jsomEnsureUser(url: string, user: userObjAccount | userObjEmail | userObjAccount[] | userObjEmail[]): Promise<toServerData>;
export function jsomGetItemsById(props:listItemById_Guid | listItemById_Title): Promise<anyOject[]>;
export function jsomGetFilesByRelativeUrl(props:getFiles): Promise<anyOject>;
export function jsomTaxonomyRequest(termStoreId: string, termSetId: string): Promise<taxReturn>;
export function jsomSendDataToServer(toServerData: toServerData): Promise<anyOject>;

export class jsomCreateItemsMetered {
	constructor (props: metered_create_guid | metered_create_title);
	totalRequests(): number;
	sendData(): Promise<toServerData>
}
export class jsomUpdateItemsMetered {
	constructor (props: metered_update_guid | metered_update_title);
	totalRequests(): number;
	sendData(): Promise<toServerData>
}
export class jsomRecycleItemsMetered {
	constructor (props: metered_recycle_guid | metered_recycle_title);
	totalRequests(): number;
	sendData(): Promise<toServerData>
}
