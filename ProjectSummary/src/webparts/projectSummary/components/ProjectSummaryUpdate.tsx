import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Label, TextField, PrimaryButton, DefaultButton, DatePicker, Spinner } from 'office-ui-fabric-react/lib';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';

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
import '../../../Commonfiles/Services/customStyles.css';
import { ListItemAttachments } from '@pnp/spfx-controls-react/lib/ListItemAttachments';
import { sp, Web, ItemAddResult } from "@pnp/sp";
import "@pnp/polyfill-ie11";

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
    spinner:boolean;
}


export default class ProjectSummaryUpdate extends React.Component<IProjectSummaryProps, IProjectSummarySubmissionState, {}> {

    private listFormService: IListFormService;
    private fields = [];
    public ItemId: number;
    private ActionTakenKey: number;
    public etag: string = undefined;
    public liaisonofficer: number = null;
    public PjtState: string;
    public isActivityChanged: boolean = false;
    public PjtStatus:string;



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
            spinner:false          
        };
        // SPComponentLoader.loadScript('https://ttengage.sharepoint.com/sites/ttEngage_Dev/SiteAssets/jquery.js', {
        //     globalExportsName: 'jQuery'
        // }).catch((error) => {

        // }).then((): Promise<{}> => {
        //     return SPComponentLoader.loadScript('https://ttengage.sharepoint.com/sites/ttEngage_Dev/SiteAssets/jquery.MultiFile.js', {
        //         globalExportsName: 'jQuery'
        //     });
        // }).catch((error) => {

        // });

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

        /**
          this.setState({
            loginUser:this.props.context.pageContext.user.email
          })
      
          let data = moment("2/10/2019", "DD/MM/YYYY").format("MM/DD/YYYY");
          console.log(data);
      
          */

    }

    private _onFormatDate = (date: Date): string => {
        return date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear();
    }

    /**
       * Gets the schema for all relevant fields for a specified SharePoint list form.     
       * @param event to capture the type of event.     
       * @  Method to capture updated in the form.
       */
    private handleChange(event) {
        if (event.target.value !== "") {
            this.fields.push(event.target.id);
        }
    }

    private _getChanges = (internalName: string, event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        //console.log(`Selection change: ${item.text}  ${item.key} ${item.selected ? 'selected' : 'unselected'}`);
        this.fields.push(internalName);
        if (internalName == "ActionTaken") {
            this.ActionTakenKey = Number(item.key);
            this.PjtStatus = item.text;

            if (item.key == 1) {
                this.setState({
                    pjtAccepted: true,
                    ActionTaken: Number(item.key),
                    hideDialog: false
                });
            }
            else {
                this.setState({
                    pjtAccepted: false,
                    ActionTaken: Number(item.key)
                });
            }
        }
    }


    private _closeDialog = (): void => {
        this.setState({ hideDialog: true });
    }


    private _getupdateBodyContent(listItemEntityTypeName: string) {
        let _fields = [...new Set(this.fields)];
        var bodyContent = {
            '__metadata': {
                'type': listItemEntityTypeName
            },
        };

        if (window.navigator.userAgent.indexOf("Trident/") > 0) {
            _fields = _fields[0]._values;
        }

        for (let id of _fields) {

            if (id == "ActionTaken") {
                bodyContent["ActionTakenId"] = this.ActionTakenKey;
                bodyContent["ProjectStatus"] = this.PjtStatus;
                if (this.ActionTakenKey == 1) {
                    bodyContent["LiaisonOfficerId"] = this.liaisonofficer;
                }
            }
            else if (id == "ProjectURL")
                bodyContent["ProjectURL"] = `https://ttengage.sharepoint.com${this.state.pjtSpace}`;           
            else {
                let value = (document.getElementById(id) as HTMLInputElement).value;
                bodyContent[id] = value;
            }
        }        

        let body: string = JSON.stringify(bodyContent);
        return body;
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


    //function to capture People picker.
    private _getPeoplePickerItems(items: any[]) {
        for (let item of items) {
            this.liaisonofficer = item.id;
        }
    }

    //function to submit the Project summary and for updates   
    public async updateData() {
        return this.listFormService._getListItemEntityTypeName(this.props.context, "Projects")
            .then(listItemEntityTypeName => {
                let vbody: string = this._getupdateBodyContent(listItemEntityTypeName);
                return this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Projects')/items(${this.ItemId})`, SPHttpClient.configurations.v1, {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-type': 'application/json;odata=verbose',
                        'odata-version': '',
                        'IF-MATCH': this.etag,
                        'X-HTTP-Method': 'MERGE'
                    },
                    body: vbody
                });
            }).then((response: SPHttpClientResponse): void => {
                alert("Updated Successfully...");
                window.location.href = this.props.context.pageContext.web.absoluteUrl;

            }, (error: any): void => {

            });
    }

    private _submitform() {
        this.updateData()

    }

    public ProjectSpace() {
        this.setState({
            spinner:true
        });
        let vsiteurl = `ProjectSpace${this.ItemId}`;
        let vsiteTitle = this.state.items.Title;
        let vsiteDesp = this.state.items.ProjectDescription;

        this.listFormService._creatProjectSpace(this.props.context, vsiteTitle, vsiteurl)
            .then((responseJSON) => {
                this.fields.push("ProjectURL");
                this.setState({
                    hideDialog: true,
                    pjtSpace: responseJSON.ServerRelativeUrl,
                    spinner:false

                });
            });
    }


    public render(): React.ReactElement<IProjectSummaryProps> {
        return (
            <div className={styles.projectSummary}>
                <div className="widget-box widget-color-blue2">
                    <div className="widget-header">
                        <h4 className="widget-title lighter smaller">PROJECT SUMMARY SUBMISSION</h4>
                    </div>
                    <div className="widget-body">
                        <div className="widget-main padding-8">
                            <div className="row">
                                <div className="profile-user-info profile-user-info-striped">
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Project Title</div>
                                        <div className="profile-info-value">
                                            <TextField id="Title" placeholder="Project Title" value={this.state.items.Title} disabled />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Short Description </div>
                                        <div className="profile-info-value">
                                            <TextField id="ProjectDescription" multiline rows={3} value={this.state.items.ProjectDescription} disabled />
                                        </div>
                                    </div>

                                    <div className="profile-info-row">
                                        <div className="profile-info-name">List of investors / Partners</div>
                                        <div className="profile-info-value">
                                            <TextField id="Listofinvestors" placeholder="List of Investors/Partners" value={this.state.items.Listofinvestors} disabled />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Products  &amp; Associated Quantity</div>
                                        <div className="profile-info-value">
                                            <TextField id="Productsandassociatedquantities"
                                                name="Productsandassociatedquantities"
                                                multiline
                                                rows={3}
                                                placeholder="Products & Associated Quantity"
                                                value={this.state.items.Productsandassociatedquantities}
                                                disabled={true} />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Capital Expenditure </div>
                                        <div className="profile-info-value">
                                            <TextField id="CapitalExpenditure" label="" placeholder="Capital Expenditure" value={this.state.items.CapitalExpenditure} disabled suffix="US$MM" />
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
                                    <div className="profile-info-row" style={this.state.items.Naturalgas ? {} : { display: 'none' }}>
                                        <div className="profile-info-name">Natural Gas</div>
                                        <div className="profile-info-value">
                                            <TextField id="Naturalgas" suffix="mmscf/d" value={this.state.items.Naturalgas} disabled />
                                        </div>
                                    </div>
                                    <div className="profile-info-row" style={(this.state.items.ElectricityMW || this.state.items.ElectricityKW) ? {} : { display: 'none', padding: '0px 5px' }}>
                                        <div className="profile-info-name">Electricity </div>
                                        <div className="profile-info-value">
                                            <TextField type="text" id="ElectricityMW" className="Electricity ms-TextField-field" suffix="MW" value={this.state.items.ElectricityMW} disabled />
                                            <TextField type="text" id="ElectricityKW" className="Electricity ms-TextField-field" suffix="kVA" value={this.state.items.ElectricityKW} disabled />

                                        </div>
                                    </div>
                                    <div className="profile-info-row" style={(this.state.items.Water) ? {} : { display: 'none' }}>
                                        <div className="profile-info-name">Water consumption</div>
                                        <div className="profile-info-value">
                                            <TextField id="Water" label="" suffix="m³/month" value={this.state.items.Water} disabled />
                                        </div>
                                    </div>
                                    <div className="profile-info-row" style={(this.state.items.Land) ? {} : { display: 'none' }}>
                                        <div className="profile-info-name">Land requirement </div>
                                        <div className="profile-info-value">
                                            <TextField id="Land" label="" suffix="hectares" value={this.state.items.Land} disabled />
                                        </div>
                                    </div>
                                    <div className="profile-info-row" style={(this.state.items.Port) ? {} : { display: 'none' }}>
                                        <div className="profile-info-name">Port requirements </div>
                                        <div className="profile-info-value">
                                            <TextField id="Port" label="" value={this.state.items.Port} disabled />
                                        </div>
                                    </div>
                                    <div className="profile-info-row" style={(this.state.items.WarehousingRequirements) ? {} : { display: 'none' }}>
                                        <div className="profile-info-name">Warehousing requirement </div>
                                        <div className="profile-info-value">
                                            <TextField id="WarehousingRequirements" label="" value={this.state.items.WarehousingRequirements} disabled />
                                        </div>
                                    </div>
                                    <div className="profile-info-row" style={(this.state.items.PotentialSaving) ? {} : { display: 'none' }
                                    }>
                                        <div className="profile-info-name">If Energy Efficient Project, Potential Saving </div>
                                        <div className="profile-info-value">
                                            <TextField id="PotentialSaving" label="" value={this.state.items.PotentialSaving} disabled />
                                        </div>
                                    </div >

                                    <div className="profile-info-row" style={(this.state.items.Other) ? {} : { display: 'none' }}>
                                        <div className="profile-info-name">Other </div>
                                        <div className="profile-info-value">
                                            <TextField id="Other" label="" value={this.state.items.Other} disabled />
                                        </div>
                                    </div >
                                </div >
                            </div >
                        </div >
                    </div >

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
                                                disabled={(this.state.items.ActionTakenId !== '1' && this.state.isAdmin) ? false : true}
                                                options={this.state.Actions.map((item: any) => { return { key: item.ID, text: item.Title }; })}
                                                onChange={this._getChanges.bind(this, "ActionTaken")}

                                            />
                                        </div>
                                    </div>
                                    <div className="profile-info-row" style={this.state.pjtAccepted ? {} : { display: 'none' }}>
                                        <div className="profile-info-name">Liaison Officer</div>
                                        <div className="profile-info-value">
                                            <PeoplePicker context={this.props.context}
                                                personSelectionLimit={1}
                                                groupName={""}
                                                showtooltip={true}
                                                isRequired={true}
                                                ensureUser={true}
                                                selectedItems={this._getPeoplePickerItems.bind(this)}
                                                showHiddenInUI={false}
                                                principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                                                resolveDelay={1500}
                                            />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Comments</div>
                                        <div className="profile-info-value">
                                            <TextField id="Comments" label="" multiline rows={3} onBlur={this.handleChange.bind(this)} disabled={this.state.isAdmin ? false : true} />
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


                    <div className="pull-right mtp">

                        <PrimaryButton title="Submit" text="Submit" onClick={() => this._submitform()}></PrimaryButton>
                        &nbsp;&nbsp;<PrimaryButton title="Cancel" text="Cancel" allowDisabledFocus href={this.props.context.pageContext.web.absoluteUrl}></PrimaryButton>
                    </div>

                    {/* <div className={styles.row}>
            <PrimaryButton text="Submit"
                           onClick={this._submitform.bind(this)}></PrimaryButton>
        </div> */}


                    <div>
                        <Dialog hidden={this.state.hideDialog}
                            onDismiss={this._closeDialog}
                            dialogContentProps={{
                                type: DialogType.normal,
                                title: 'Project Confirmation',
                                subText: 'Do you want to make this project Accepted for Facilitation?'
                            }}
                            modalProps={{
                                isBlocking: true,
                                styles: { main: { maxWidth: 450 } }
                            }}>
                            <DialogFooter>
                                <PrimaryButton onClick={this.ProjectSpace.bind(this)} text="Yes" />
                                <DefaultButton onClick={this._closeDialog} text="Cancel" />
                            </DialogFooter>
                        </Dialog>

                        <Spinner label="We are Working Please wait..."  style={(this.state.spinner) ? {} : { display: 'none' }}/>
                    </div>


                </div>
            </div >
        );
    }
}