import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Label, TextField, PrimaryButton, DefaultButton, Checkbox } from 'office-ui-fabric-react/lib';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';

import styles from './ProjectSummary.module.scss';
//import { IProjectSummaryProps, IProjectSummaryState, IListItem, IProjectSpace } from './IProjectSummaryProps';
import { IProjectApprovalsProps, IProjectApprovalsState, IListItem } from './IProjectApprovalProps';
import { escape } from '@microsoft/sp-lodash-subset';
// import * as CustomJS from 'CustomJS';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPHttpClient } from '@microsoft/sp-http';
import * as moment from 'moment';
import { sp, Web } from "@pnp/sp";


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
    private Category = [];
    public ItemId: number;
    private ActionTakenKey: number;
    public liaisonofficer: number = null;
    public investor: number = null;
    public PjtState: string;
    public pjtSpace: string;

    constructor(props: IProjectApprovalsProps) {
        super(props);
        // Initiate the component state
        this.state = {
            items: [],
            hideDialog: true,
            Category: [],
            Agency: null

        };

        this.listFormService = new ListFormService(props.context.spHttpClient);
        this.ItemId = Number(window.location.search.split("PID=")[1]);
        const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Projects')/items(${this.ItemId})?$select=*,LiaisonOfficer/Id,LiaisonOfficer/EMail,Approvals/Title&$expand=LiaisonOfficer,Approvals`;
        this.listFormService._getListItem(this.props.context, restApi)
            .then((response) => {
                this.pjtSpace = response.ProjectURL;
            });

        const url: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Approvals Master')/items?$select=*,Agency/Title&$expand=Agency`;
        this.listFormService._getListitems(this.props.context, url)
            .then((resp) => {
                let data = resp.value;
                this.Category = [...new Set(data.map(x => x.Category))];
                this.setState({
                    items: data
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
            this.fields.push(event.currentTarget.name);
            this.items.push(event.currentTarget.id)
        }
        else {
            let value = event.currentTarget.name
            this.fields = this.fields.filter(function (ele) {
                return ele != value;
            });

            let itemID = event.currentTarget.id
            this.items = this.items.filter(function (ele) {
                return ele != itemID;
            });
            console.log(`The option has been changed to ${isChecked}.`);
        }
    }

    public _submitData() {
        let web = new Web(`${this.pjtSpace}`);
        let list = web.lists.getByTitle("Approvals");
        list.getListItemEntityTypeFullName().then(entityTypeFullName => {
            let batch = web.createBatch();

            for (const index of this.fields) {
                let item = this.state.items[index];
                list.items.inBatch(batch).add(
                    {
                        Title: `${item.Title}`,
                        ApprovalCategory:`${item.Category}`,
                        ApprovalOrder:`${item.ApprovalOrder}`,
                        AgencyResponsible:`${item.Agency.Title}`,
                        AgencyGroupId:`${item.AgencyGroupId}`
                        
                        //ProjectName
                        //Investor
                        //ApprovalID

                    },
                    entityTypeFullName).then(b => {
                        console.log(b);
                    });
            }
            batch.execute().then(d =>
                console.log("Done")
            );
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
        return this.listFormService._getListItemEntityTypeName(this.props.context, "Projects")
            .then(listItemEntityTypeName => {
                let vbody: string = this._getContentBody(listItemEntityTypeName);
                return this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Projects')/items`, SPHttpClient.configurations.v1, {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-type': 'application/json;odata=verbose',
                        'odata-version': ''
                    },
                    body: vbody
                });
            }).then(response => {
                return response.json();
            });
    }

    private _submitform() {
        this.SaveData()
            .then((resp) => {
            })
    }


    public render(): React.ReactElement<IProjectApprovalsProps> {
        return (
            <div className={styles.projectSummary}>
                <div className={styles.header}>
                    <h4 className={styles.title}>PROJECT SUMMARY APPROVALS</h4>
                </div>
                <div id='reactForm' className={styles.body}>
                    {this.Category.map((value, index) => {
                        return (
                            <div>
                                <div style={{ background: "#fbaf33", color: "#fff", margin: "0 auto", padding: "0px 8px", lineHeight:"25px" }}>
                                    <h4>{value}</h4>
                                </div>
                                {this.state.items.filter(item => item.Category == `${value}`).map((AppValue, index) => {
                                    return (
                                        <div className={styles.row}>
                                            <div className="ApprovalCheckList">
                                                <Checkbox name={index.toString()} id={AppValue.Id.toString()} label={AppValue.Title} onChange={this._onCheckboxChange} />
                                            </div>
                                            <div className="AgencyName">
                                                <label>{AppValue.AgencyId ? AppValue.Agency.Title : ""}</label>
                                            </div>
                                        </div>
                                    )
                                })}

                            </div>
                        )
                    })}

                </div>

                <PrimaryButton
                    text="Submit"
                    onClick={() => this._submitData()}
                ></PrimaryButton>

                <DefaultButton
                    text="Cancel"></DefaultButton>               
                  
            </div>
        )
    }
}
