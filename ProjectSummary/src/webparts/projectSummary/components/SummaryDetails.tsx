import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Label, TextField, PrimaryButton, DefaultButton, DatePicker, Spinner } from 'office-ui-fabric-react/lib';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';

import styles from './ProjectSummary.module.scss';
import { IProjectSummaryProps, IProjectSummaryState, IListItem, IProjectSpace } from './IProjectSummaryProps';
import { escape } from '@microsoft/sp-lodash-subset';
//import * as CustomJS from 'CustomJS';
import * as $ from 'jQuery';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as moment from 'moment';
import {
    Dropdown,
    IDropdown,
    DropdownMenuItemType,
    IDropdownOption
} from "office-ui-fabric-react/lib/Dropdown";
import { ListFormService } from '../../../Commonfiles/Services/CommonMethods';
import { IListFormService } from '../../../Commonfiles/Services/ICommonMethods';
//import '../../../Commonfiles/Services/customStyles.css';
import '../../../Commonfiles/Services/Custom.css';
import { ListItemAttachments } from '@pnp/spfx-controls-react/lib/ListItemAttachments';
import { sp, Web, ItemAddResult } from "@pnp/sp";

export interface IProjectSummarySubmissionState {
    startDate: Date;
    items: IListItem;
    disabled: boolean;
    isAdmin: boolean;
    pjtAccepted: boolean;
    Actions: any[];
    ActionTaken: number;
    hideDialog: boolean;
    pjtSpace: string;
    listID: string;
    ItemId: number;
    spinner: boolean;
}


export default class SummaryDetails extends React.Component<IProjectSummaryProps, IProjectSummarySubmissionState, {}> {

    private listFormService: IListFormService;
    private fields = [];
    public ItemId: number;
    private ActionTakenKey: number;
    public etag: string = undefined;
    public liaisonofficer: number = null;
    public PjtState: string;
    public isActivityChanged: boolean = false;
    public PjtStatus: string;



    constructor(props: IProjectSummaryProps) {
        super(props);
        // Initiate the component state
        this.state = {
            startDate: null,
            items: {},
            disabled: false,
            isAdmin: false,
            pjtAccepted: false,
            Actions: [],
            ActionTaken: null,
            hideDialog: true,
            pjtSpace: null,
            listID: null,
            ItemId: null,
            spinner: false
        };
      
        this.listFormService = new ListFormService(props.context.spHttpClient);
        this.ItemId = Number(window.location.search.split("PID=")[1]);
        this._getProjectActions();

        if (this.ItemId) {
            const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Projects')/items(${this.ItemId})?$select=*,LiaisonOfficer/Id,LiaisonOfficer/EMail&$expand=LiaisonOfficer`;
            this.listFormService._getListItem(this.props.context, restApi)
                .then((response) => {
                    this.setState({
                        items: response,
                        disabled: true,
                        startDate: response.ProposedStartDate ? new Date(response.ProposedStartDate) : null,
                        pjtAccepted: response.ActionTakenId == 1 ? true : false,
                        ActionTaken: response.ActionTakenId ? Number(response.ActionTakenId) : null,
                        pjtSpace: response.ProjectURL,
                        ItemId: this.ItemId
                    });
                });


            this.listFormService._getListItem_etag(this.props.context, "Projects", this.ItemId)
                .then((resp) => {
                    this.etag = resp;
                });
        }

        this.listFormService._getloginusergroups(this.props.context)
            .then((response) => {
                response.Groups.map(((item: any, inc) => {
                    if (item.Title === "IF Admin") {
                        this.setState({
                            isAdmin: true
                        });
                        return false;
                    }
                }));
            });

        const listrestApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Projects')`;
        this.listFormService._getListItem(this.props.context, listrestApi)
            .then((response) => {
                this.setState({
                    listID: response.Id
                });
            });

       
    }

    private _onFormatDate = (date: Date): string => {
        return date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear();
    }
    
    //function to get the project Actions passing Webcontext and restAPI url
    public _getProjectActions() {
        const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Actions Master')/items`;
        this.listFormService._getListitems(this.props.context, restApi)
            .then((response) => {
                let items = response.value;
                this.setState({
                    Actions: items
                });
            });
    }


   
 

    public render(): React.ReactElement<IProjectSummaryProps> {
        return (
            <div className={styles.projectSummary}>
                <div className="widget-box widget-color-blue2">
                    <div className="widget-header">
                        <h4 className="widget-title lighter smaller">Project Summary Submission</h4>
                    </div>
                    <div className="widget-Summary">
                        <div className="widget-body">
                            <div className="widget-main padding-8">
                                <div className="row">
                                    <div className="profile-user-info profile-user-info-striped">
                                        <div className="profile-info-row">
                                            <div className="profile-info-name">Project Title</div>
                                            <div className="profile-info-value">
                                                <TextField id="Title" underlined placeholder="Project Title" value={this.state.items.Title} readOnly />
                                            </div>
                                        </div>
                                        <div className="profile-info-row">
                                            <div className="profile-info-name">Short Description </div>
                                            <div className="profile-info-value">
                                                <TextField id="ProjectDescription" underlined multiline rows={3} value={this.state.items.ProjectDescription} readOnly />
                                            </div>
                                        </div>

                                        <div className="profile-info-row">
                                            <div className="profile-info-name">List of investors / Partners</div>
                                            <div className="profile-info-value">
                                                <TextField id="Listofinvestors" underlined placeholder="List of Investors/Partners" value={this.state.items.Listofinvestors} readOnly />
                                            </div>
                                        </div>
                                        <div className="profile-info-row">
                                            <div className="profile-info-name">Proposed Start Date </div>
                                            <div className="profile-info-value">
                                                <DatePicker placeholder=""
                                                    id="ProposedStartDate"
                                                    value={this.state.startDate}
                                                    formatDate={this._onFormatDate}
                                                    minDate={new Date(2000, 12, 30)}
                                                    isMonthPickerVisible={false}
                                                    disabled={true} />
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div className="widget-subheader" style={{ background: "#fbaf33", color: "#fff", width: "95%", margin: "0 auto", padding: "1px" }}>
                            <h4 className="widget-title lighter smaller" style={{ margin: "5px" }}>Project Specifications</h4>
                        </div>
                        <div className="widget-body" style={{ width: "95%", margin: "0 auto" }}>
                            <div className="widget-main " style={{ padding: "0 0px 8px 0px" }}>
                                <div className="row">
                                    <div className="profile-user-info profile-user-info-striped">
                                        <div className="profile-info-row">
                                            <div className="profile-info-name">Products  &amp; Associated Quantity</div>
                                            <div className="profile-info-value">
                                                <TextField id="Productsandassociatedquantities"
                                                    className="wd100"
                                                    name="Productsandassociatedquantities"
                                                    multiline
                                                    rows={3}
                                                    placeholder="Products & Associated Quantity"
                                                    underlined
                                                    value={this.state.items.Productsandassociatedquantities}
                                                    readOnly={true} />
                                            </div>
                                            <div className="profile-info-name">Capital Expenditure </div>
                                            <div className="profile-info-value">
                                                <TextField id="CapitalExpenditure" className="wd100" underlined label="" placeholder="Capital Expenditure" value={this.state.items.CapitalExpenditure} readOnly suffix="US$MM" />
                                            </div>
                                        </div>
                                        <div className="profile-info-row">
                                        <div className="profile-info-name">Port Requirements </div>
                                            <div className="profile-info-value">
                                                <TextField id="Port" className="wd100"  multiline rows={3} label="" underlined value={this.state.items.Port} readOnly />
                                            </div>
                                            <div className="profile-info-name">Natural Gas Usage</div>
                                            <div className="profile-info-value">
                                                <TextField id="Naturalgas" className="wd100" suffix="mmscf/d" underlined value={this.state.items.Naturalgas} readOnly />
                                            </div>
                                            
                                        </div>
                                        <div className="profile-info-row">
                                        <div className="profile-info-name">Warehousing Requirements </div>
                                            <div className="profile-info-value">
                                                <TextField id="WarehousingRequirements" className="wd100" multiline rows={3} label="" underlined value={this.state.items.WarehousingRequirements} readOnly />
                                            </div>
                                            <div className="profile-info-name">Electricity </div>
                                            <div className="profile-info-value">
                                                <TextField type="text" id="ElectricityMW" className="Electricity ms-TextField-field wd100" suffix="MW" underlined value={this.state.items.ElectricityMW} readOnly />
                                                <TextField type="text" id="ElectricityKW" className="Electricity ms-TextField-field wd100" suffix="kVA" underlined value={this.state.items.ElectricityKW} readOnly />
                                            </div>
                                            
                                        </div>
                                        <div className="profile-info-row">
                                        <div className="profile-info-name">If Energy Efficient Project, Potential Savings </div>
                                            <div className="profile-info-value">
                                                <TextField id="PotentialSaving" className="wd100" label="" underlined value={this.state.items.PotentialSaving} readOnly />
                                            </div>
                                            <div className="profile-info-name">Water Consumption</div>
                                            <div className="profile-info-value">
                                                <TextField id="Water" className="wd100" label="" suffix="m³/month" underlined value={this.state.items.Water} readOnly />
                                            </div>                                          
                                          
                                        </div>
                                        <div className="profile-info-row" >                                           
                                            <div className="profile-info-name">Other </div>
                                            <div className="profile-info-value">
                                                <TextField id="Other" className="wd100" label="" multiline rows={3} underlined value={this.state.items.Other} readOnly />
                                            </div>
                                            
                                            <div className="profile-info-name">Land Requirements </div>
                                            <div className="profile-info-value">
                                                <TextField id="Land" className="wd100" label="" suffix="hectares" underlined value={this.state.items.Land} readOnly />
                                            </div>
                                        </div>
                                    </div >
                                </div >
                            </div >
                        </div >


                    </div>

                    <div className="widget-Actions">
                        <div className="widget-body">
                            <div className="widget-main padding-8">
                                <div className="row">
                                    <div className="profile-user-info profile-user-info-striped">

                                        <div className="profile-info-row" style={(this.state.isAdmin || (this.state.items.ActionTakenId || this.state.items.ActionTakenId == '1')) ? {} : { display: 'none' }}>
                                            <div className="profile-info-name">Project Actions:</div>
                                            <div className="profile-info-value">
                                                <Dropdown id='ActionTaken'
                                                    defaultSelectedKey={this.state.ActionTaken}
                                                    placeholder="Select an Action"
                                                    label=''
                                                    disabled={true}
                                                    options={this.state.Actions.map((item: any) => { return { key: item.ID, text: item.Title }; })}
                                                   
                                                />
                                            </div>
                                        </div>
                                       
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                    {(this.state.listID && this.state.ItemId) ? (
                        <div className={styles.row}>
                            <ListItemAttachments listId={this.state.listID}
                                itemId={this.state.ItemId}
                                context={this.props.context}
                                disabled={false} />
                        </div>) : (
                            <div></div>
                        )
                    }

                    <div className={styles.pullright}>
                        &nbsp;&nbsp;<PrimaryButton title="Back" text="Back" allowDisabledFocus onClick={()=>{window.history.back()}}></PrimaryButton>
                    </div>

                </div>
            </div >
        );
    }
}