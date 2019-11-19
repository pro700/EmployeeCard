import "core-js/es6/object";
import "core-js/es6/array"; 
import "core-js/es6/map"; 
import "core-js/es6/set"; 
import "core-js/es6/index"; 
import "core-js/es6/promise"; 
//import "core-js/modules/es6.array.iterator.js";
//import "core-js/modules/es6.array.from.js";
//import "core-js/modules/es6.promise";
import "whatwg-fetch";
//import 'sp-build-tasks/dist/webpack/polyfills';

import * as $ from "jquery";

import * as React from 'react';
import * as ReactDOM from "react-dom";

//import * as sp from "@pnp/sp";
import { sp, List, Item, ListEnsureResult, ItemAddResult, FieldAddResult, SiteUserProps, UserProfileQuery, Web, ContextInfo, SPConfiguration, SPRest, PermissionKind, Fields, Field, FieldCreationProperties, XmlSchemaFieldCreationInformation, ItemUpdateResult } from '@pnp/sp';

import * as pnp from '@pnp/pnpjs';

import { taxonomy, Session, TermSet, TermStore, TermStores, ITermStore, ITermSetData, ITermSet, ITermData, ITerm, ITermGroup, ITermStoreData } from "@pnp/sp-taxonomy";

import { TreeView, ItemRenderProps } from '@progress/kendo-react-treeview';

//import '@progress/kendo-theme-material/scss/treeview.scss';
//import '@progress/kendo-theme-bootstrap/scss/treeview.scss';
import '@progress/kendo-theme-default/scss/treeview.scss';
import './EmployeeCard.less'
import { ISearchQueryBuilder, SearchQuery, SearchResult, SearchResults, SearchQueryBuilder, SortDirection } from "@pnp/sp";
import { SPRestAddIn } from "@pnp/sp-addinhelpers";
//import Date from "core-js";
//import "./SPServices/jquery.SPServices.js";
//var spservices = require("./SPServices/jquery.SPServices.js");
//import { InitSPServices, SPServices } from  "./SPServices/jquery.SPServices.js";

export interface EmployeeCardRow {
    id: string;
    parent: string;
    text: string;
    icon: string;
    type: string;
    title: string;
    email: string;
    birthday: string;
    workphone: string;
    WorkType: string;
    TabNum: string;
    idFirm: string;
    FirmName: string;
    Auto_Card: string;
    CuratorFullName: string;
    CuratorAutoCard: string;
    MobilePhone: string;
    InternalPhone: string;
    Photo: string;
}

export interface TreeViewItem {
    key?: string;
    text: string;
    items?: TreeViewItem[];
    expanded?: boolean;
    checked?: boolean;
    checkIndeterminate?: boolean;
    disabled?: boolean;
    hasChildren?: boolean;
    selected?: boolean;
    type?: string;
    pictureURL?: string;
    email?: string;
    username?: string;
    accountname?: string;
    id?: string;
}

export interface EmployeeCardLayoutProps {
}

export interface EmployeeCardLayoutState {
    message: string;
    error: string;
    treeViewData: TreeViewItem[];
}


export class EmployeeCardLayout extends React.Component<EmployeeCardLayoutProps, EmployeeCardLayoutState> {

    constructor(props: EmployeeCardLayoutProps) {
        super(props);

        this.state = { message: "", error: "", treeViewData: [] };

    }


//    onExpandChange = {(event) => { }}
//itemRender = { (props) => {
//    return <span> {props.item.text} </span>;
//} }

    // 
    // <img className={"k-icon"} style={{ height: "2em", width: "2em", borderRadius: "50%" }} src={props.item.pictureURL} key={props.item.key} />


    render(): JSX.Element {

        var props: ItemRenderProps;

        return (
            <div style={{ background: "lightgray" }}>
                <table>
                    <tr>
                        <td>
                            <div id="treecontainer"> 
                                <TreeView data={this.state.treeViewData}
                                    expandIcons={true}
                                    checkboxes={false}
                                    onExpandChange={this.onExpandChange}
                                    onCheckChange={this.onCheckChange}
                                    onItemClick={this.onItemClick}
                                    itemRender={props => 
                                        props.item.type == "department" &&
                                        <span>
                                            <span className={"k-icon k-i-folder"} key={props.item.key}></span>
                                            {props.item.text}
                                        </span> ||
                                        props.item.type == "user" &&
                                        <span className="k-mid">
                                            <img className={"k-icon"} style={{ height: "2em", width: "2em", borderRadius: "50%" }} src={this.getUrlParamByName("SPHostUrl") + "/_layouts/userphoto.aspx?size=M&accountname=" + props.item.username} key={props.item.key} />
                                            {props.item.text}
                                        </span>
                                    }
                                />;
                            </div>
                        </td>
                        <td> <div id="panecontainer">  </div> </td>
                    </tr>
                </table>

                <div id="message" style={{ color: "blue" }}>
                    {this.state.message}
                </div>

                <div id="error" style={{ color: "red" }}>
                    {this.state.error}
                </div>
            </div>
        );
    }

    onItemClick = (event) => {
        event.item.selected = !event.item.selected;
        this.forceUpdate();
    }

    onExpandChange = (event) => {
        event.item.expanded = !event.item.expanded;
        this.forceUpdate();
    }

    onCheckChange = (event) => {
        event.item.checked = !event.item.checked;
        this.forceUpdate();
    }

    componentDidMount() {

        var hostweb = this.getUrlParamByName("SPHostUrl");
        var addinweb = this.getUrlParamByName("SPAppWebUrl");

        sp.setup({
            sp: { baseUrl: addinweb }
        });

        taxonomy.setup({
            sp: { baseUrl: addinweb }
        });

        this.ensureLists()
            .then(res => {
                /////////////////
                // this.setState({ message: "Ensure lists ok!" });
                ////////////////
                this.populateTree()
                    .then((data: TreeViewItem[]) => {
                        this.setState({ treeViewData: data });

                        window.setTimeout(this.updateUserProfiles.bind(this), 500, 0);
                    })
                    .catch(err => {
                        this.setState({ error: "componentDidMount: populateRootAndFirstLevel() error=" + JSON.stringify(err) });
                    });
            })
            .catch(err => {
                this.setState({ error: "componentDidMount: ensureLists() error=" + JSON.stringify(err) });
            });

    }

    async updateUserProfiles() {

        var hostweb = this.getUrlParamByName("SPHostUrl");
        var addinweb = this.getUrlParamByName("SPAppWebUrl");

        sp.setup({
            sp: { baseUrl: addinweb }
        });

        function isValidDate(date) {
            return date && Object.prototype.toString.call(date) === "[object Date]" && !isNaN(date);
        };

        //const crossDomainWeb: Web = pnp.sp.crossDomainWeb(addinweb, hostweb);

        try {
            var lastItems: any[] = await sp.web.lists.getByTitle("EC_Cards").items.select("Id", "LastModifiedTime").top(1).orderBy("LastModifiedTime", false).get();

            //this.setState({ message: "get lastItems ok!" });

            let lastModifiedTime: Date = new Date(1900, 0, 1);
            if (lastItems.length > 0) {
                var date = new Date(lastItems[0]["LastModifiedTime"]);
                if (isValidDate(date)) {
                    lastModifiedTime = date;
                }
            }

            var searchRes: pnp.SearchResults = await sp.search({
                Querytext: `LastModifiedTime>"${lastModifiedTime.toISOString()}"`,
                SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
                RowLimit: 10,
                SortList: [{ Property: "LastModifiedTime", Direction: SortDirection.Ascending }],
                SelectProperties: ['AccountName', 'LastModifiedTime', 'Department', 'JobTitle', 'WorkEmail', 'Path', 'PictureURL', 'PreferredName', 'UserProfile_GUID', 'OriginalPath']
            });

            //this.setState({ message: this.state.message + ", search ok! count=" + searchRes.PrimarySearchResults.length });

            for (var index: number = 0; index < searchRes.PrimarySearchResults.length; index++) {

                var searchResItem: pnp.SearchResult = searchRes.PrimarySearchResults[index];

                var userProfileProps: any = await sp.profiles.getPropertiesFor(searchResItem["AccountName"]);

              //  this.setState({ message: this.state.message + `, index=${index}, getPropertiesFor(${searchResItem["AccountName"]}) ok!` });

                var userProfileObject = {};
                userProfileProps.UserProfileProperties.forEach((property: any) => {
                    userProfileObject[property.Key] = property.Value;
                });

                let SID = userProfileObject["SID"];

                let itemsBySID: any[] = await sp.web.lists.getByTitle("EC_Cards").items.filter(`SID eq '${SID}'`).top(1).get();

                //this.setState({ message: this.state.message + `, index=${index}, getItemsBySID(${SID}) ok! lenght=${itemsBySID.length}` });

                if (itemsBySID.length > 0) {
                    let Id: number = parseInt(itemsBySID[0]["ID"]);
                    let updateRes: ItemUpdateResult = await sp.web.lists.getByTitle("EC_Cards").items.getById(Id)
                        .update({
                            Title: userProfileObject["PreferredName"],
                            FullName: userProfileObject["PreferredName"],
                            FirstName: userProfileObject["FirstName"],
                            LastName: userProfileObject["LastName"],
                            AccountName: userProfileObject["AccountName"],
                            UserName: userProfileObject["UserName"],
                            EMail: userProfileObject["WorkEmail"],
                            PictureURL: {
                                "__metadata": { type: "SP.FieldUrlValue" },
                                Description: "",
                                Url: userProfileObject["PictureURL"]
                            },
                            LastModifiedTime: searchResItem["LastModifiedTime"]
                        });
                    //this.setState({ message: this.state.message + `, ${index} update ok!` });
                }
                else {
                    let addRes: ItemAddResult = await sp.web.lists.getByTitle("EC_Cards").items
                        .add({
                            SID: SID,
                            Title: userProfileObject["PreferredName"],
                            FullName: userProfileObject["PreferredName"],
                            FirstName: userProfileObject["FirstName"],
                            LastName: userProfileObject["LastName"],
                            AccountName: userProfileObject["AccountName"],
                            UserName: userProfileObject["UserName"],
                            EMail: userProfileObject["WorkEmail"],
                            PictureURL: {
                                "__metadata": { type: "SP.FieldUrlValue" },
                                Description: "",
                                Url: userProfileObject["PictureURL"]
                            },
                            LastModifiedTime: searchResItem["LastModifiedTime"]
                        });

                    //this.setState({ message: this.state.message + `, ${index} add ok!` });
                }
            }

            if (searchRes.PrimarySearchResults.length > 0) {
                window.setTimeout(this.updateUserProfiles.bind(this), 500, 0);
            }
        }
        catch(err)
        {
            this.setState({ error: "updateUserProfiles() try catch error=" + err });
        }
    }

    getUrlParamByName(name) {
        name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
        var regex = new RegExp("[\\?&]" + name + "=([^&#]*)");
        var results = regex.exec(location.search);
        return results === null ? "" : decodeURIComponent(results[1].replace(/\+/g, " "));
    }  


    getAddInHostWeb() {
        var addinweb = this.getUrlParamByName("SPAppWebUrl");
        var hostweb = this.getUrlParamByName("SPHostUrl");
        var web = sp.web;
        if (addinweb.length > 0 && hostweb.length > 0) {
            web = pnp.sp.crossDomainWeb(addinweb, hostweb);
        }
        return web;
    }

    getAddInHostSite() {
        var addinweb = this.getUrlParamByName("SPAppWebUrl");
        var hostweb = this.getUrlParamByName("SPHostUrl");
        var site = sp.site;
        if (addinweb.length > 0 && hostweb.length > 0) {
            site = pnp.sp.crossDomainSite(addinweb, hostweb);
        }
        return site;
    }

    ensureLists(): Promise<any>{

        return new Promise<any>((resolve, reject) => {

            var hostweb = this.getUrlParamByName("SPHostUrl");
            var addinweb = this.getUrlParamByName("SPAppWebUrl");

            sp.setup({
                sp: { baseUrl: addinweb }
            });

            taxonomy.setup({
                sp: { baseUrl: addinweb }
            });

            //const web: Web = pnp.sp.crossDomainWeb(addinweb, hostweb);

            sp.web.lists.ensure("EC_Cards")
                .then(res => {
                    res.list.fields.select("Title", "InternalName", "TypeAsString").get()
                        .then((fields: any[]) => {

                            const session1 = new Session(addinweb);
                            const p1 = session1.getDefaultSiteCollectionTermStore().get();
                            const p2 = session1.getDefaultSiteCollectionTermStore().groups.getByName('People').termSets.getByName('Department').get();
                            const p3 = session1.getDefaultSiteCollectionTermStore().groups.getByName('People').termSets.getByName('Job Title').get();

                            Promise.all([p1, p2, p3]).then(async ([termStoreData, termSetData1, termSetData2]: [ITermStoreData, ITermSetData, ITermSetData]) => {

                                try {
                                    if (!this.isInFieldsByInternalName(fields, "SID")) { await res.list.fields.createFieldAsXml(this.getTextFieldXml("91542991-7F8B-4F5F-8B4F-9519CA9660BB", "SID", "SID", "EmployeeCard", false)); }
                                    if (!this.isInFieldsByInternalName(fields, "AccountName")) { await res.list.fields.createFieldAsXml(this.getTextFieldXml("11542991-7FB8-4F5F-8B4F-9519CA9BC0BB", "AccountName", "AccountName", "EmployeeCard", true)); }
                                    if (!this.isInFieldsByInternalName(fields, "UserName")) { await res.list.fields.createFieldAsXml(this.getTextFieldXml("11514299-9FB8-4F7F-9B4F-9519CA9BC0AB", "UserName", "UserName", "EmployeeCard", true)); }
                                    if (!this.isInFieldsByInternalName(fields, "FullName")) { await res.list.fields.createFieldAsXml(this.getTextFieldXml("9BD418AE-6026-48CA-9D68-F03749331C09", "FullName", "FullName", "EmployeeCard", false)); }
                                    if (!this.isInFieldsByInternalName(fields, "FirstName")) { await res.list.fields.createFieldAsXml(this.getTextFieldXml("DB207EE2-9FD4-439C-917B-2FA19AD14C24", "FirstName", "FirstName", "EmployeeCard", false)); }
                                    if (!this.isInFieldsByInternalName(fields, "LastName")) { await res.list.fields.createFieldAsXml(this.getTextFieldXml("77C24297-C373-47A4-A92A-B504F5DBD748", "LastName", "LastName", "EmployeeCard", false)); }
                                    if (!this.isInFieldsByInternalName(fields, "EMail")) { await res.list.fields.createFieldAsXml(this.getTextFieldXml("F6B5C72F-2030-443B-A3A5-A65F68B390CE", "EMail", "EMail", "EmployeeCard", false)); }
                                    if (!this.isInFieldsByInternalName(fields, "Department")) { await res.list.fields.createFieldAsXml(this.getTaxonomyFieldXml("1F2EF8CD-A280-4DAD-B915-21CC72E13974", "Department", "Department", "EmployeeCard", this.cleanGuid(termStoreData.Id), this.cleanGuid(termSetData1.Id), false)); }
                                    if (!this.isInFieldsByInternalName(fields, "JobTitle")) { await res.list.fields.createFieldAsXml(this.getTaxonomyFieldXml("76DA43C2-6AA7-4F66-8CEB-34D4C318DAD5", "JobTitle", "JobTitle", "EmployeeCard", this.cleanGuid(termStoreData.Id), this.cleanGuid(termSetData2.Id), false)); }
                                    if (!this.isInFieldsByInternalName(fields, "Gender")) { await res.list.fields.createFieldAsXml(this.getChoiceFieldXml("E5BBF051-5122-4A9C-94B9-4D08FBBD48EC", "Gender", "Gender", "EmployeeCard", false, ["Male", "Female"])); }
                                    if (!this.isInFieldsByInternalName(fields, "LastModifiedTime")) { await res.list.fields.createFieldAsXml(this.getDateTimeFieldXml("F6B5C72F-2030-443B-A3A5-A65F68C390DF", "LastModifiedTime", "LastModifiedTime", "EmployeeCard", false, "DateTime")); }
                                    if (!this.isInFieldsByInternalName(fields, "PictureURL")) { await res.list.fields.createFieldAsXml(`<Field ID="{F6A5C72F-2730-443B-B9A5-A65F68C390DF}" Type="URL" Format="Image" Name="PictureURL" DisplayName="PictureURL" Required="false" Group="EmployeeCard"></Field>`); }
                                    resolve();
                                }
                                catch (err) {
                                    reject(err);
                                }
                            });
                        })
                        .catch(err => reject(err));
                })
                .catch(err => reject({ catch: "sp.web.lists.ensure('EC_Cards')", err: err }));
        });
    }

    private isInFieldsByInternalName(fields: any[], name: string) {
        return fields.filter((field: any) => { return field["InternalName"] == name; }).length > 0;
    }

    private getItemsTreeAndMap(map: Map<string, TreeViewItem>, parentPath: string, terms: ITermData[], expanded: boolean = false): TreeViewItem[] {
        return terms
            .filter((term: ITermData) => { return term.PathOfTerm == parentPath + (parentPath == "" ? "" : ";") + term.Name })
            .map(term => {
                var items: TreeViewItem[] = this.getItemsTreeAndMap(map, term.PathOfTerm, terms);
                var item = {
                    key: term.PathOfTerm,
                    text: term.Name,
                    expanded: expanded,
                    hasChildren: (items.length > 0),
                    type: "department",
                    id: term.Id,
                    items: items
                };
                map.set(this.cleanGuid(term.Id), item);
                return item;
            });
    }

    private populateTree(): Promise<TreeViewItem[]> {
        return new Promise<any[]>((resolve, reject) => {

            var hostweb = this.getUrlParamByName("SPHostUrl");
            var addinweb = this.getUrlParamByName("SPAppWebUrl");

            sp.setup({
                sp: { baseUrl: addinweb }
            });

            taxonomy.setup({
                sp: { baseUrl: addinweb }
            });

            const session1 = new Session(addinweb);

            let p2 = session1.getDefaultSiteCollectionTermStore().groups.getByName('People').termSets.getByName('Department').terms.get();

            p2.then((terms: ITermData[]) => {

                /////////////////
                //this.setState({ message: "Terms ok!" });
                ////////////////

                var map: Map<string, TreeViewItem> = new Map();
                var items: TreeViewItem[] = this.getItemsTreeAndMap(map, "", terms);
                var item = {
                    key: "-Ы",
                    text: "Компанія",
                    expanded: true,
                    hasChildren: (items.length > 0),
                    type: "department",
                    id: "-1",
                    items: items
                }
                let data: TreeViewItem[] = [item];
                map.set("-1", item);
                resolve(data);

                //var str: string = "-------------map----------------";
                //map.forEach((value, key, map) => {
                //    str += ", key=" + key + " value.text=" + JSON.stringify(value.text);
                //});
                //this.setState({ message: this.state.message + ", MAP=" + str });

                sp.web.lists.getByTitle("EC_Cards").items.getPaged()
                    .then(async paged_items => {
                        ///////////////////////////
                        //this.setState({ message: "-----------------get cards paged------------ paged_items.hasNext=" + paged_items.hasNext + ", paged_items.results.length=" + paged_items.results.length });

                        do {
                            paged_items.results.forEach((card, card_index) => {
                                //this.setState({ message: this.state.message + `,------------------card(${card_index})=${JSON.stringify(card)}` });

                                var itemByDep = map.get("-1");
                                if (card["Department"] != null) {
                                    var dep = card["Department"];
                                    //this.setState({ message: this.state.message + "|||,dep:" + JSON.stringify(dep) });
                                    var guid = dep["TermGuid"];
                                    //this.setState({ message: this.state.message + ",TermGuid:" + guid});
                                    itemByDep = map.get(guid);
                                    //this.setState({ message: this.state.message + ",itemByDep:" + JSON.stringify(itemByDep) });
                                }

                                itemByDep.items.push({
                                    key: card["SID"],
                                    email: card["EMail"],
                                    username: card["UserName"],
                                    accountname: card["AccountName"],
                                    text: card["FullName"],
                                    expanded: false,
                                    hasChildren: false,
                                    type: "user",
                                    id: card["ID"],
                                    items: [],
                                    pictureURL: card["PictureURL"] == null ? "" : decodeURI(card["PictureURL"]["Url"])
                                });
                                //this.setState({ message: this.state.message + ",---------------Title:" + card["Title"] + ", Department.TermGuid=" + this.cleanGuid(card["Department"]["TermGuid"]) + ", PictureURL" + card["PictureURL"]});
                            });

                            this.setState({ treeViewData: data });

                            if (paged_items.hasNext) {
                                paged_items = await paged_items.getNext();
                            }
                            else {
                                break;
                            }
                        } while (true);
                    })
                    .catch(err => {
                        this.setState({ error: "error=" + JSON.stringify(err) });
                    });
            }).catch(error => {
                this.setState({ error: "error=" + JSON.stringify(error) });
                resolve([]);
            });
        });
    }

    private getItems(parentPath: string, terms: ITermData[], expanded: boolean = false): TreeViewItem[] {
        return terms
            .filter((term: ITermData) => { return term.PathOfTerm == parentPath + (parentPath == "" ? "" : ";") + term.Name })
            .map(term => {
                var items: TreeViewItem[] = this.getItems(term.PathOfTerm, terms);
                return {
                    key: term.PathOfTerm,
                    text: term.Name,
                    expanded: expanded,
                    hasChildren: (items.length > 0),
                    type: "department",
                    id: term.Id,
                    items: items
                };
            });
    }

    private getItemsByDepartment = (itemsTree: TreeViewItem[], departmentText: string): TreeViewItem[] => {
        let resItems: TreeViewItem[] = [];
        itemsTree.forEach(item => {
            if (item.type == "department") {
                if (item.text == departmentText) {
                    resItems.push(item);
                }
                resItems.push(...this.getItemsByDepartment(item.items, departmentText));
            }
        });
        return resItems;
    }

    private populateRootAndFirstLevel(): Promise<TreeViewItem[]> {
        return new Promise<any[]>((resolve, reject) => {

            var hostweb = this.getUrlParamByName("SPHostUrl");
            var addinweb = this.getUrlParamByName("SPAppWebUrl");

            sp.setup({
                sp: { baseUrl: addinweb }
            });

            taxonomy.setup({
                sp: { baseUrl: addinweb }
            });

            const session1 = new Session(addinweb);

            let p2 = session1.getDefaultSiteCollectionTermStore().groups.getByName('People').termSets.getByName('Department').terms.get();
            p2.then((terms: ITermData[]) => {

                var items: TreeViewItem[] = this.getItems("", terms);

                let data: TreeViewItem[] = [{
                    key: "-Ы",
                    text: "Компанія",
                    expanded: true,
                    hasChildren: (items.length > 0),
                    type: "department",
                    id: "",
                    items: items
                }];

                sp.search({
                    Querytext: '*',
                    SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
                    RowLimit: 1000,
                    RowsPerPage: 1000,
                    SelectProperties: ['AccountName', 'Department', 'JobTitle', 'WorkEmail', 'Path', 'PictureURL', 'PreferredName', 'UserProfile_GUID', 'OriginalPath']
                })
                    .then(res => {

                        let promises: Promise<any>[] = res.PrimarySearchResults.map((user: any) => {
                            return sp.profiles.getPropertiesFor(user.AccountName);
                        });

                        Promise.all(promises)
                            .then((allUsersProps: any[]) => {

                                //this.setState({ message: "allUsersProps=" + JSON.stringify(allUsersProps) });

                                //var searchresusers = [];





                                allUsersProps.forEach((userProps: any) => {

                                    userProps.UserProfileProperties.forEach((property: any) => {
                                        userProps[property.Key] = property.Value;
                                    });

                                    //searchresusers.push({
                                    //    "SPSDep": userProps["SPS-Department"],
                                    //    "Acc": userProps["AccountName"],
                                    //    "Name": userProps["PreferredName"],
                                    //    "Url": decodeURI(userProps["PictureUrl"])
                                    //});

                                    this.getItemsByDepartment(data, userProps['SPS-Department']).forEach(item => {
                                        item.items.push({
                                            key: userProps["AccountName"],
                                            email: userProps["Email"],
                                            username: userProps["UserName"],
                                            text: userProps["PreferredName"],
                                            expanded: false,
                                            hasChildren: false,
                                            type: "user",
                                            id: "",
                                            items: [],
                                            pictureURL: decodeURI(userProps["PictureUrl"])
                                        });
                                    });
                                });

                                //allUsersProps.forEach((userProps: any) => {

                                //    //let parent: any = 'root';
                                //    //if (Rows.filter((row: EmployeeCardRow) => { return (row.id == userProps['Department']); }).length > 0) {
                                //    //    parent = userProps['Department'];
                                //    //}

                                //    //Rows.push({
                                //    //    id: userProps["AccountName"],
                                //    //    parent: parent,
                                //    //    text: userProps["DisplayName"],
                                //    //    icon: "jstree-icon jstree-file",
                                //    //    type: "person",
                                //    //    title: userProps["Title"],
                                //    //    email: userProps["Email"],
                                //    //    birthday: userProps["SPS-Birthday"],
                                //    //    workphone: userProps["WorkPhone"],
                                //    //    WorkType: "",
                                //    //    TabNum: "",
                                //    //    idFirm: "1",
                                //    //    FirmName: "Talan Systems",
                                //    //    Auto_Card: userProps["AccountName"],
                                //    //    CuratorFullName: "",
                                //    //    CuratorAutoCard: userProps["Manager"],
                                //    //    MobilePhone: userProps["CellPhone"],
                                //    //    InternalPhone: "",
                                //    //    Photo: userProps["PictureURL"]
                                //    //});
                                //});

                                resolve(data);

                            })
                            .catch((err: any) => {
                                reject(err);
                            });
                    })
                    .catch(error => {
                        this.setState({ error: "error=" + JSON.stringify(error) });
                        resolve(data);
                    });

            })
            .catch(error => {
                this.setState({ error: "error=" + JSON.stringify(error) });
                resolve([]);
            });

            
            //let p2 = taxonomy1.getDefaultSiteCollectionTermStore().groups.getByName('People').termSets.getByName('Department').terms.get();
            //p2.then((terms: ITermData[]) => {
            //    var items: TreeViewItem[] = this.getItems("", terms);
            //    let data: TreeViewItem[] = [{
            //        key: "-Ы",
            //        text: "Компанія",
            //        expanded: true,
            //        hasChildren: (items.length > 0),
            //        type: "department",
            //        items: items
            //    }];

            //    sp.search({
            //        Querytext: '*',
            //        SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
            //        RowLimit: 1000,
            //        RowsPerPage: 1000,
            //        SelectProperties: ['AccountName', 'Department', 'JobTitle', 'WorkEmail', 'Path', 'PictureURL', 'PreferredName', 'UserProfile_GUID', 'OriginalPath']
            //    })
            //        .then(res => {

            //            let promises: Promise<any>[] = res.PrimarySearchResults.map((user: any) => {
            //                return sp.profiles.getPropertiesFor(user.AccountName);
            //            });

            //            Promise.all(promises)
            //                .then((allUsersProps: any[]) => {

            //                    var searchresusers = [];

            //                    allUsersProps.forEach((userProps: any) => {


            //                        let getItemsByDepartment = (itemsTree: TreeViewItem[], departmentText: string) => {
            //                            let resItems: TreeViewItem[] = [];
            //                            itemsTree.forEach(item => {
            //                                if (item.type == "department") {
            //                                    if (item.text == departmentText) {
            //                                        resItems.push(item);
            //                                    }
            //                                    resItems.push(...getItemsByDepartment(item.items, departmentText));
            //                                }
            //                            });
            //                            return resItems;
            //                        };

            //                        userProps.UserProfileProperties.results.forEach((property: any) => {
            //                            userProps[property.Key] = property.Value;
            //                        });

            //                        searchresusers.push({
            //                            "SPSDep": userProps["SPS-Department"],
            //                            "Acc": userProps["AccountName"],
            //                            "Name": userProps["PreferredName"],
            //                            "Url": decodeURI(userProps["PictureUrl"])
            //                        });

            //                        getItemsByDepartment(data, userProps['SPS-Department']).forEach(item => {
            //                            item.items.push({
            //                                key: userProps["AccountName"],
            //                                text: userProps["PreferredName"],
            //                                expanded: false,
            //                                hasChildren: false,
            //                                type: "user",
            //                                items: [],
            //                                pictureURL: decodeURI(userProps["PictureUrl"])
            //                            });
            //                        });

            //                        //let parent: any = 'root';
            //                        //if (Rows.filter((row: EmployeeCardRow) => { return (row.id == userProps['Department']); }).length > 0) {
            //                        //    parent = userProps['Department'];
            //                        //}

            //                        //Rows.push({
            //                        //    id: userProps["AccountName"],
            //                        //    parent: parent,
            //                        //    text: userProps["DisplayName"],
            //                        //    icon: "jstree-icon jstree-file",
            //                        //    type: "person",
            //                        //    title: userProps["Title"],
            //                        //    email: userProps["Email"],
            //                        //    birthday: userProps["SPS-Birthday"],
            //                        //    workphone: userProps["WorkPhone"],
            //                        //    WorkType: "",
            //                        //    TabNum: "",
            //                        //    idFirm: "1",
            //                        //    FirmName: "Talan Systems",
            //                        //    Auto_Card: userProps["AccountName"],
            //                        //    CuratorFullName: "",
            //                        //    CuratorAutoCard: userProps["Manager"],
            //                        //    MobilePhone: userProps["CellPhone"],
            //                        //    InternalPhone: "",
            //                        //    Photo: userProps["PictureURL"]
            //                        //});
            //                    });

            //                    resolve(data);

            //                })
            //                .catch((err: any) => {
            //                    reject(err);
            //                });
            //        })
            //        .catch((err: any) => {
            //            reject(err);
            //        });

            //});
            

        });
    }

    public cleanGuid(guid: string): string {
        if (guid !== undefined) {
            return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
        } else {
            return '';
        }
    }


    private getChoiceFieldXml(ID: string, DispalyName: string, Name: string, Group: string, required: boolean, choices: string[]) {
        return `<Field ID="{${ID}}" Type="Choice" Name="${Name}" DisplayName="${DispalyName}" Required="${required}" Group="${Group}">
                    <choices>
                        ${choices.map((choice: string) => { return `<choice>${choice}</choice>`;}).join()}
                    </choices>
                </Field>`;
    }

    private getTextFieldXml(ID: string, DispalyName: string, Name: string, Group: string, required: boolean) {
        return `<Field ID="{${ID}}" Type="Text" Name="${Name}" DisplayName="${DispalyName}" Required="${required}" Group="${Group}"></Field>`;
    }

    private getDateTimeFieldXml(ID: string, DispalyName: string, Name: string, Group: string, required: boolean, Format: string) {
        return `<Field ID="{${ID}}" Type="DateTime" Name="${Name}" Format="${Format}" DisplayName="${DispalyName}" Required="${required}" Group="${Group}"></Field>`;
    }

    private getTaxonomyFieldXml(ID: string, DispalyName: string, Name: string, Group: string, SspId: string, TermSetId: string, required: boolean): string {
        return `<Field ID="{${ID}}" Type="TaxonomyFieldType"  Name="${Name}" DisplayName="${DispalyName}" ShowField="Term1033" Required="${required}" Group="${Group}" >
                    <Customization>
                        <ArrayOfProperty>
                            <Property>
                                <Name>SspId</Name>
                                <Value xmlns:q1="http://www.w3.org/2001/XMLSchema" p4:type="q1:string" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">${SspId}</Value>
                            </Property>
                            <Property>
                                <Name>TermSetId</Name>
                                <Value xmlns:q2="http://www.w3.org/2001/XMLSchema" p4:type="q2:string" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">${TermSetId}</Value>
                            </Property>
                        </ArrayOfProperty>
                    </Customization>
                </Field>`;
    }

    private populate(): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            this.getData()
                .then((Rows: EmployeeCardRow[]) => {
                    this.setState({ message: this.state.message + "Rows final=" + JSON.stringify(Rows) });
                })
                .catch((err) => {
                    this.setState({ error: err });
                });
        });
    }

    private getData(): Promise<EmployeeCardRow[]> {

        return new Promise<EmployeeCardRow[]>((resolve, reject) => {
            let Rows: EmployeeCardRow[] = [{ id: "root", parent: "#", text: "Talan Systems", icon: "jstree-icon jstree-folder", type: "department", title: "", email: "", birthday: "", workphone: "", WorkType: "", TabNum: "", idFirm: "1", FirmName: "Talan Systems", Auto_Card: "", CuratorFullName: "", CuratorAutoCard: "", MobilePhone: "", InternalPhone: "", Photo: "" }];
            //let Deps: any[] = [];
            //let Pers: any[] = [];

            //let web = new Web(this.context.pageContext.web.absoluteUrl);

            this.setState({ message: this.state.message + ", before getData taxonomy.getDefaultSiteCollectionTermStore()" });
            try {
                this.setState({ message: this.state.message + ", this.context=" + JSON.stringify(this.context) });
            }
            catch (e) {
                this.setState({ error: this.state.error + ", this.context error=" + JSON.stringify(e) });
            }

            //var config: SPConfiguration = {
            //    sp: {
            //        baseUrl: "https://sokolsofop.sharepoint.com",
            //        headers: {
            //            Accept: "application/json;odata=verbose"
            //        }
            //    }
            //};
            //taxonomy.setup(config);
            let p1 = sp.site.getContextInfo();

            p1.then((info: ContextInfo) => {

                this.setState({ message: this.state.message + "\n ContextInfo SiteFullUrl ok=" + JSON.stringify(info.SiteFullUrl) });

                taxonomy.setup({ sp: { headers: { Accept: "application/json;odata=verbose" }, baseUrl: info.SiteFullUrl}});
                sp.setup({ sp: { headers: { Accept: "application/json;odata=verbose" }, baseUrl: info.SiteFullUrl } });

                let p2 = taxonomy.getDefaultSiteCollectionTermStore().groups.get(); //.getByName('People').termSets.getByName('Department').terms.get();
                //let p3 = sp.site.getWebUrlFromPageUrl(window.location.href);

                p2.then((terms: ITermGroup[]) => {

                    this.setState({ message: this.state.message + "\n terms ok=" + JSON.stringify(terms) });

                    //var contextInfo = "";
                    //for (var key in info) {
                    //    contextInfo += key + " : " + info[key] + "; ";
                    //}

                    //terms.forEach((term: ITermData & ITerm) => {
                    //    Rows.push({
                    //        id: term.Name,
                    //        parent: (term['Parent'] ? term["Parent"].Name : 'root'),
                    //        text: term.Name,
                    //        icon: "jstree-icon jstree-folder",
                    //        type: "department",
                    //        title: "", email: "", birthday: "", workphone: "", WorkType: "", TabNum: "", idFirm: "1", FirmName: "Talan Systems", Auto_Card: "", CuratorFullName: "", CuratorAutoCard: "", MobilePhone: "", InternalPhone: "", Photo: ""
                    //    });
                    //});

                    this.setState({ message: this.state.message + "\n getData Rows ok=" + JSON.stringify(Rows) });
                    this.setState({ message: this.state.message + "\n window.location.href=" + window.location.href });

                    /*
                    let query: SearchQuery = {
                        Querytext: '*',
                        SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
                        RowLimit: 1000,
                        RowsPerPage: 1000,
                        SelectProperties: ['AccountName', 'Department', 'JobTitle', 'WorkEmail', 'Path', 'PictureURL', 'PreferredName', 'UserProfile_GUID', 'OriginalPath']
                    };

                    this.setState({ message: this.state.message + "<br/> query.Querytext=" + query.Querytext });
                    */

                    sp.search({
                        Querytext: '*',
                        SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
                        RowLimit: 1000,
                        RowsPerPage: 1000,
                        SelectProperties: ['AccountName', 'Department', 'JobTitle', 'WorkEmail', 'Path', 'PictureURL', 'PreferredName', 'UserProfile_GUID', 'OriginalPath']
                    })
                        .then(res => {

                            var users = res.PrimarySearchResults.map((user: any) => { return user.AccountName; }).join(", ");


                            this.setState({ message: this.state.message + "sp.search ok, users=" + users });



                            let promises: Promise<any>[] = res.PrimarySearchResults.map((user: any) => { return sp.profiles.getPropertiesFor(user.AccountName); });
                            Promise.all(promises)
                                .then((allUsersProps: any[]) => {
                                    allUsersProps.forEach((userProps: any) => {

                                        let parent: any = 'root';

                                        userProps.UserProfileProperties.results.forEach((property: any) => {
                                            userProps[property.Key] = property.Value;
                                        });

                                        if (Rows.filter((row: EmployeeCardRow) => { return (row.id == userProps['Department']); }).length > 0) {
                                            parent = userProps['Department'];
                                        }

                                        Rows.push({
                                            id: userProps["AccountName"],
                                            parent: parent,
                                            text: userProps["DisplayName"],
                                            icon: "jstree-icon jstree-file",
                                            type: "person",
                                            title: userProps["Title"],
                                            email: userProps["Email"],
                                            birthday: userProps["SPS-Birthday"],
                                            workphone: userProps["WorkPhone"],
                                            WorkType: "",
                                            TabNum: "",
                                            idFirm: "1",
                                            FirmName: "Talan Systems",
                                            Auto_Card: userProps["AccountName"],
                                            CuratorFullName: "",
                                            CuratorAutoCard: userProps["Manager"],
                                            MobilePhone: userProps["CellPhone"],
                                            InternalPhone: "",
                                            Photo: userProps["PictureURL"]
                                        });
                                    });

                                    resolve(Rows);

                                })
                                .catch((err: any) => {
                                    reject(err);
                                });
                        })
                        .catch((err: any) => {
                            reject(err);
                        });
                })
                .catch((err) => {
                    reject(err);
                });

            })
            .catch((err) => {
                reject(err);
            });

        });
    }

    /*
    getUserPrifileCount(): Promise<any> {

        return new Promise<any>((resolve, reject) => {

            var hostweb = this.getUrlParamByName("SPHostUrl");
            var addinweb = this.getUrlParamByName("SPAppWebUrl");
            var webServiceURL = hostweb + "/_vti_bin/UserProfileService.asmx";
            //var webServiceURL = hostweb + "/_vti_bin/UserProfileService.asmx";

            var SOAPEnvelopeHeader = "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'><soap:Body>";
            var SOAPEnvelopeFooter = "</soap:Body></soap:Envelope>";

            $.ajax({
                url: hostweb + "/_vti_bin/UserProfileService.asmx",
                type: "POST",
                dataType: "xml",
                beforeSend: function (xhr) {
                    xhr.setRequestHeader("SOAPAction", "http://microsoft.com/webservices/SharePointPortalServer/UserProfileService/GetUserProfileCount")
                },
                data: SOAPEnvelopeHeader + `<GetUserProfileCount xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService" />` + SOAPEnvelopeFooter,
                contentType: "text/xml;charset='utf-8'",
                success: function (data, status) {
                    resolve(data.d);
                },
                error: function (request, status, error) {
                    reject("request=" + JSON.stringify(request) + ", status=" + status + ", error=" + error);
                }
            });
        });
    }

    getUserPrifileByIndex = (userProfileIndex: number) => {
        try {

            this.setState({ message: "getUserPrifileByIndex start $()=" + JSON.stringify($()) });

            var hostweb = this.getUrlParamByName("SPHostUrl");
            var addinweb = this.getUrlParamByName("SPAppWebUrl");

            var webServiceURL = addinweb + "/_vti_bin/UserProfileService.asmx?op=GetUserProfileCount";
            var soapMessage =
                `<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
                    <soap12:Body>
                        <GetUserProfileCount xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService" />
                    </soap12:Body>
                </soap12:Envelope>`;

            $.ajax({
                url: webServiceURL,
                type: "POST",
                dataType: "xml",
                data: soapMessage,
                contentType: "text/xml; charset=\"utf-8\"",
                success: function (data, status) {
                    alert(data.d);
                },
                error: function (request, status, error) {
                    alert('error');
                }
            });

        }
        catch (e) {
            this.setState({ error: "error=" + JSON.stringify(e) });
        }
    }

     */

    //private adjustSize() {
    //    //var divh = document.getElementById('EmployeeCardMain').offsetHeight;
    //    //var divw = document.getElementById('EmployeeCardMain').offsetWidth;
    //    this.resize($("#EmployeeCardMain").width(), $("#EmployeeCardMain").height())
    //}

    //private resize(width, height) {
    //    //var target = parent["postMessage"] ? parent : (parent.document["postMessage"] ? parent.document : undefined);
    //    var regex = new RegExp(/[Ss]ender[Ii]d=([\daAbBcCdDeEfF]+)/);
    //    var results = regex.exec(window.location.search);
    //
    //    if (null != results && null != results[1]) {
    //        window.parent.postMessage('<message senderId=' + results[1] + '>resize(' + width + ',' + height + ')</message>', '*');
    //    }
    //} 
}


/*
var spAppIFrameSenderInfo = new Array(1);

var SPAppIFramePostMsgHandler = function (e) {
    if (e.data.length > 100)
        return;

    var regex = RegExp(/(<\s*[Mm]essage\s+[Ss]ender[Ii]d\s*=\s*([\dAaBbCcDdEdFf]{8})(\d{1,3})\s*>[Rr]esize\s*\(\s*(\s*(\d*)\s*([^,\)\s\d]*)\s*,\s*(\d*)\s*([^,\)\s\d]*))?\s*\)\s*<\/\s*[Mm]essage\s*>)/);
    var results = regex.exec(e.data);
    if (results == null)
        return;

    var senderIndex = results[3];
    if (senderIndex >= spAppIFrameSenderInfo.length)
        return;

    var senderId = results[2] + senderIndex;

    var iframeId = unescape(spAppIFrameSenderInfo[senderIndex][1]);

    var senderOrigin = unescape(spAppIFrameSenderInfo[senderIndex][2]);

    if (senderId != spAppIFrameSenderInfo[senderIndex][0] || senderOrigin != e.origin)
        return;

    var width = results[5];
    var height = results[7];
    if (width == "") {
        width = '300px';
    }
    else {
        var widthUnit = results[6];
        if (widthUnit == "")
            widthUnit = 'px';

        width = width + widthUnit;
    }

    if (height == "") {
        height = '150px';
    }
    else {
        var heightUnit = results[8];
        if (heightUnit == "")
            heightUnit = 'px';

        height = height + heightUnit;
    }

    var widthCssText = "";
    var resizeWidth = ('False' == spAppIFrameSenderInfo[senderIndex][3]);
    if (resizeWidth) {
        widthCssText = 'width:' + width + ' !important;';
    }

    var cssText = widthCssText;
    var resizeHeight = ('False' == spAppIFrameSenderInfo[senderIndex][4]);
    if (resizeHeight) {
        cssText += 'height:' + height + ' !important';
    }

    if (cssText != "") {
        var webPartInnermostDivId = spAppIFrameSenderInfo[senderIndex][5];
        if (webPartInnermostDivId != "") {
            var webPartDivId = 'WebPart' + webPartInnermostDivId;

            var webPartDiv = document.getElementById(webPartDivId);
            if (null != webPartDiv) {
                webPartDiv.style.cssText = cssText;
            }

            cssText = "";
            if (resizeWidth) {
                var webPartChromeTitle = document.getElementById(webPartDivId + '_ChromeTitle');
                if (null != webPartChromeTitle) {
                    webPartChromeTitle.style.cssText = widthCssText;
                }

                cssText = 'width:100% !important;'
            }

            if (resizeHeight) {
                cssText += 'height:100% !important';
            }

            var webPartInnermostDiv = document.getElementById(webPartInnermostDivId);
            if (null != webPartInnermostDiv) {
                webPartInnermostDiv.style.cssText = cssText;
            }
        }

        var iframe = document.getElementById(iframeId);
        if (null != iframe) {
            iframe.style.cssText = cssText;
        }
    }
}

if (typeof window.addEventListener != 'undefined') {
    window.addEventListener('message', SPAppIFramePostMsgHandler, false);
}
else if (typeof window.attachEvent != 'undefined') {
    window.attachEvent('onmessage', SPAppIFramePostMsgHandler);
}

spAppIFrameSenderInfo[0] = new Array("CC4CB9140", "g_9f0007c9_cd40_442e_a09f_0f6a27860e3b", "https:\u002f\u002fsokolsofop-008ee5ffa8fe3b.sharepoint.com", "False", "False", "ctl00_ctl33_g_527ab16c_1941_4eeb_98be_508dc434975f");
*/
