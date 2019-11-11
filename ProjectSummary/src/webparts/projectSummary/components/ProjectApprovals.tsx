import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Label, TextField, PrimaryButton, DefaultButton, Checkbox, Panel, PanelType, Spinner } from 'office-ui-fabric-react/lib';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';

import styles from './ProjectSummary.module.scss';
//import { IProjectSummaryProps, IProjectSummaryState, IListItem, IProjectSpace } from './IProjectSummaryProps';
import { IProjectApprovalsProps, IProjectApprovalsState, IListItem, } from './IProjectApprovalProps';
import { escape } from '@microsoft/sp-lodash-subset';
// import * as CustomJS from 'CustomJS';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPHttpClient } from '@microsoft/sp-http';
import * as moment from 'moment';
import { sp, Web } from "@pnp/sp";


import ProjectSpace from './ProjectSpace';


import {
    Dropdown,
    IDropdown,
    DropdownMenuItemType,
    IDropdownOption
} from "office-ui-fabric-react/lib/Dropdown";
import { ListFormService } from '../../../Commonfiles/Services/CommonMethods';
import { IListFormService } from '../../../Commonfiles/Services/ICommonMethods';

import { element } from 'prop-types';
import { App } from 'sp-pnp-js';

export default class ProjectApprovals extends React.Component<IProjectApprovalsProps, IProjectApprovalsState, {}> {

    private listFormService: IListFormService;
    private fields = [];
    private items = [];
    private selectedApprovals = [];
    private delteItems = [];
    private Category = [];
    public ItemId: number;
    private ActionTakenKey: number;
    public liaisonofficer: number = null;
    public investor: number = null;
    public PjtState: string;
    public pjtSpace: string;
    public Approvals: any[];

    constructor(props: IProjectApprovalsProps) {
        super(props);
        // Initiate the component state
        this.state = {
            items: [],
            hideDialog: true,
            Category: [],
            Agency: null,
            pjtItem: {},
            showPanel: true,
            spinner:false

        };


        this.listFormService = new ListFormService(props.context.spHttpClient);
        this.ItemId = Number(this.props.context.pageContext.web.absoluteUrl.split("/ProjectSpace")[1]);
        const restApi = `${this.props.context.pageContext.site.absoluteUrl}/_api/web/lists/GetByTitle('Projects')/items(${this.ItemId})`;
        this.listFormService._getListItem(this.props.context, restApi)
            .then((response) => {
                //    this.props.context.statusRenderer.displayLoadingIndicator(document.getElementById("ApprovalsList"), "Approvals..."); 

                this.setState({
                    pjtItem: response
                });
            }).then(() => {

                const url: string = `${this.props.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('Approvals Master')/items?$select=*,Agency/Title,Agency/ShortName&$expand=Agency`;
                this.listFormService._getListitems(this.props.context, url)
                    .then((resp) => {
                        let data = resp.value;
                        if (window.navigator.userAgent.indexOf("Trident/") > 0) {
                            let cat: any[] = [...new Set(data.map(x => x.Category))];
                            this.Category = cat[0]._values;
                        }
                        else {
                            this.Category = [...new Set(data.map(x => x.Category))];
                        }


                        this.setState({
                            items: data
                        });

                        if (this.state.pjtItem.ApprovalsId.length > 0) {
                            this.items = this.state.pjtItem.ApprovalsId;
                            this.selectedApprovals = this.state.pjtItem.ApprovalsId;
                        }

                        //   this.props.context.statusRenderer.clearLoadingIndicator(document.getElementById("ApprovalsList"));

                    });
            });

        this._onCheckboxChange = this._onCheckboxChange.bind(this);

    }

    /**
       * Gets the schema for all relevant fields for a specified SharePoint list form.     
       * @param event to capture the type of event.     
       * @  Method to capture updated in the form.
       */
    private handleChange(event) {
        if (event.target.value !== "") {
            this.fields.push(event.target.id);
            this.items.push(event.target.id);
        }
    }

    private _onCheckboxChange(event, isChecked: boolean): void {

        if (isChecked) {
            let value = event.currentTarget.id;
            if (this.selectedApprovals.indexOf(Number(value)) == -1) {
                this.fields.push(event.currentTarget.name);
            }
            this.items.push(event.currentTarget.id);
        }
        else {

            let value = event.currentTarget.name;
            if (this.fields.indexOf(value) > 0) {
                this.fields = this.fields.filter(function (ele) {
                    return ele != value;
                });
            }

            let itemID = event.currentTarget.id;
            if ((this.items.indexOf(Number(itemID)) > -1) || (this.items.indexOf(itemID) > -1)) {
                this.items = this.items.filter(function (ele) {
                    return ele != itemID;
                });

                this.delteItems.push(event.currentTarget.id);
            }
            console.log(`The option has been changed to ${isChecked}.`);
        }
    }

    public _submitData() {
        let web = new Web(`${this.props.context.pageContext.web.absoluteUrl}`);
        sp.site.getContextInfo().then(d => {
            console.log(d.FormDigestValue);
        });
        let list = web.lists.getByTitle("Approvals");
        list.getListItemEntityTypeFullName().then(entityTypeFullName => {
            let batch = web.createBatch();
            for (const index of this.fields) {
                let item = this.state.items[index];
                list.items.inBatch(batch).add(
                    {
                        Title: item.Title,
                        ApprovalCategory: `${item.Category}`,
                        AgencyResponsible: `${item.Agency.Title}`,
                        ProjectName: this.state.pjtItem.Title,
                        InvestorId: this.state.pjtItem.InvestorId,
                        LiasonOfficerId: this.state.pjtItem.LiaisonOfficerId,
                        ApprovalOrder: `${item.ApprovalOrder}`,
                        ApprovalShortName: `${item.Agency.ShortName}`,
                        ApprovalID: item.Id.toString()
                    },
                    entityTypeFullName).then(b => {
                        console.log(b);
                    });
            }
            batch.execute()
                .then(
                    d => {
                        console.log("Done");
                        alert("Approval list has successfuly saved.");
                        this.props.onDissmissPanel();
                        // this.setState({ showPanel: false });                        
                    });
        });
    }


    private _closeDialog = (): void => {
        this.setState({ hideDialog: true });
    }

    //function to generate dynamic data body to create an item.
    private _getContentBody(listItemEntityTypeName: string) {
        let _fields = [...new Set(this.fields)];
        var bodyContent = {
            '__metadata': {
                'type': listItemEntityTypeName
            },
        };

        for (let id of _fields) {
            if (id == "ProposedStartDate-label") {
                let value = (document.getElementById(id) as HTMLInputElement).value;
                let vDate = moment(value, "DD/MM/YYYY").format("MM/DD/YYYY");
                bodyContent["ProposedStartDate"] = new Date(vDate);
            }
            else {
                let value = (document.getElementById(id) as HTMLInputElement).value;
                bodyContent[id] = value;
            }
        }
        bodyContent["PromotionType"] = "Direct";
        let body: string = JSON.stringify(bodyContent);
        return body;
    }
    //function to submit the Project summary and for updates
    private async SaveData() {

        if (this.delteItems.length > 0) {
            for (let ditem of this.delteItems) {
                let itemID = ditem;
                if ((this.items.indexOf(Number(itemID)) > -1) || (this.items.indexOf(itemID) > -1)) {
                    this.delteItems = this.delteItems.filter(function (ele) {
                        return ele != itemID;
                    });
                }
            }

        }


        let web = new Web(`${this.props.context.pageContext.site.absoluteUrl}`);
        let list = web.lists.getByTitle("Projects");
        list.items.getById(this.ItemId).update({
            ApprovalsId: {
                results: this.items
            }
        }).then(async i => {
            console.log(i);
           await this._submitData();
            if (this.delteItems.length > 0) {
                await this._deleteItems();
            }
        });
    }

    private async _deleteItems() {
        let web = new Web(`${this.props.context.pageContext.web.absoluteUrl}`);
        for (let dItem of this.delteItems) {
            await web.lists.getByTitle("Approvals").items.top(1).filter(`ApprovalID eq '${dItem}'`).get().then(async (items: any[]) => {
                if (items.length > 0) {
                   await web.lists.getByTitle("Approvals").items.getById(items[0].Id).delete().then(_ => { });
                }
            });
        }

    }

    private _submitform() {
        this.SaveData()
            .then((resp) => {
            });
    }


    public render(): React.ReactElement<IProjectApprovalsProps> {
        var itemIndex: number = -1;
        return (
            <div id="ApprovalsList" className={styles.projectSummary} >
                <div className="widget-header">
                    <h4 className="widget-title lighter smaller"> PROJECT SUMMARY APPROVALS</h4>
                </div>
                <div id='reactForm' className={styles.body}>
                    {this.Category.map((value, index) => {
                        return (
                            <div>
                                <div style={{ background: "#fbaf33", color: "#fff", margin: "0 auto", padding: "0px 8px", lineHeight: "25px" }}>
                                    <h4>{value}</h4>
                                </div>
                                {this.state.items.filter(item => item.Category == `${value}`).map((AppValue, index) => {
                                    ++itemIndex;
                                    var isChecked: boolean = null;
                                    if (this.state.pjtItem.ApprovalsId.indexOf(AppValue.Id) > -1)
                                        isChecked = true;

                                    return (
                                        <div className={styles.row}>
                                            <div className="ApprovalCheckList">
                                                {
                                                    this.state.pjtItem.ApprovalsId.indexOf(AppValue.Id) > -1 ?
                                                        <Checkbox name={itemIndex.toString()} defaultChecked id={AppValue.Id.toString()} label={AppValue.Title} onChange={this._onCheckboxChange} /> :
                                                        <Checkbox name={itemIndex.toString()} id={AppValue.Id.toString()} label={AppValue.Title} onChange={this._onCheckboxChange} />

                                                }

                                            </div>
                                            <div className="AgencyName">
                                                <label>{AppValue.AgencyId ? AppValue.Agency.Title : ""}</label>
                                            </div>
                                        </div>
                                    );
                                })}

                            </div>
                        );
                    })}

                </div>

                <PrimaryButton
                    text="Save"
                    onClick={() => this.SaveData()}
                ></PrimaryButton>

                {/* <DefaultButton
                    text="Cancel"></DefaultButton>                */}

                <div>
                    <Panel
                        isOpen={this.state.spinner}
                        type={PanelType.custom}
                        headerText=""
                        closeButtonAriaLabel="Close"
                    >
                        <div>
                            <Spinner label="We are working, please wait..." ariaLive="assertive" labelPosition="right" />
                        </div>
                    </Panel>
                </div>

            </div>
        );
    }
}

