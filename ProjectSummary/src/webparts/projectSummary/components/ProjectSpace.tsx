import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Label, TextField, PrimaryButton, DefaultButton, DatePicker } from 'office-ui-fabric-react/lib';
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


export default class ProjectSpace extends React.Component<IProjectSummaryProps, IProjectSummaryState, {}> {

    private listFormService: IListFormService;
    private fields = [];
    public ItemId: number;
    private ActionTakenKey: number;
    private StageKey: number;
    private ActivityKey: number;
    public etag: string = undefined;
    public liaisonofficer: number = null;
    public PjtState: string;
    public isActivityChanged: boolean = false;



    constructor(props: IProjectSummaryProps) {
        super(props);
        // Initiate the component state
        this.state = {
            multiline: false,
            startDate: null,
            addUsers: [],
            items: {},
            status: null,
            disabled: false,
            isAdmin: false,
            pjtAccepted: false,
            Actions: [],
            Stages: [],
            Activities: [],
            ActionTaken: null,
            Stage: null,
            Activity: null,
            showState: false,
            hideDialog: true,
            formType: "New",
            pjtSpace: null,
            listID: null,
            ItemId: null,
            liaisonEmail: null,
            stageStartDate: null,
            isLiaison: false
        };
        SPComponentLoader.loadScript('https://ttengage.sharepoint.com/sites/ttEngage_Dev/SiteAssets/jquery.js', {
            globalExportsName: 'jQuery'
        }).catch((error) => {

        }).then((): Promise<{}> => {
            return SPComponentLoader.loadScript('https://ttengage.sharepoint.com/sites/ttEngage_Dev/SiteAssets/jquery.MultiFile.js', {
                globalExportsName: 'jQuery'
            });
        }).catch((error) => {

        });


        this.listFormService = new ListFormService(props.context.spHttpClient);
        this.ItemId = Number(window.location.search.split("PID=")[1]);

        if (this.ItemId) {
            const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Projects')/items(${this.ItemId})?$select=*,LiaisonOfficer/Id,LiaisonOfficer/EMail&$expand=LiaisonOfficer`;
            this.listFormService._getListItem(this.props.context, restApi)
                .then((response) => {

                    let vShowState = false;


                    if (response.LiaisonOfficerId) {
                        this.setState({
                            liaisonEmail: response.LiaisonOfficer.EMail,
                            isLiaison: (response.LiaisonOfficer.EMail == this.props.context.pageContext.user.email) ? true : false
                        });

                        let state = response.StageId;
                        if (((this.state.isAdmin || this.state.isLiaison) || !state) && response.ActionTakenId == 1) {
                            vShowState = true;
                        }
                    }
                    this._getProjectState();
                    if (response.StageId) {
                        this._getActivities(response.StageId);
                    }

                    this.setState({
                        items: response,
                        startDate: response.ProposedStartDate ? new Date(response.ProposedStartDate) : null,
                        Stage: response.StageId ? Number(response.StageId) : null,
                        Activity: response.ActivityId ? Number(response.ActivityId) : null,
                        showState: vShowState,
                        pjtSpace: response.ProjectURL
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
        if (internalName == "Step") {
            this.StageKey = Number(item.key);
            this.setState({
                Stage: Number(item.key),
                stageStartDate: null
            });
            this._getActivities(item.key.toString());
        }
        else if (internalName == "Activity")
            this.ActivityKey = Number(item.key);
        this.setState({
            Activity: Number(item.key)
        });
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

            if (id == "Step")
                bodyContent["StageId"] = this.state.Stage;
            else if (id == "Activity")
                bodyContent["ActivityId"] = this.state.Activity;
            else {
                let value = (document.getElementById(id) as HTMLInputElement).value;
                bodyContent[id] = value;
            }
        }

        bodyContent["ProjectStatus"] = $('#StageId span')[0].textContent + "-" + $('#ActivityId span')[0].textContent

        let body: string = JSON.stringify(bodyContent);
        return body;
    }

    //function to get the project Actions passing Webcontext and restAPI url
    public _getProjectState() {
        const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Stages Master')/items`;
        this.listFormService._getListitems(this.props.context, restApi)
            .then((response) => {
                let items = response.value;
                this.setState({
                    Stages: items
                });
            });
    }

    //function to get the project Actions passing Webcontext and restAPI url
    public _getActivities(StageId: string) {
        const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Activities Master')/items?$filter= StageId eq ${StageId}`;
        this.listFormService._getListitems(this.props.context, restApi)
            .then((response) => {
                let items = response.value;
                this.setState({
                    Activities: items
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
                let vStage = $('#StageId span')[0].textContent;
                let Activity = $("#ActivityId span")[0].textContent;
                // let stageDate = item.StageStartDate;
                let vComments = (document.getElementById('Comments') as HTMLInputElement).value;

                let web = new Web(`${this.state.items.ProjectURL}`);
                let list = web.lists.getByTitle("Activities").items.add({
                    Title: Activity,
                    Stage: vStage,
                    Comments: vComments,
                    StartDate: new Date()
                }).then((iar: ItemAddResult) => {
                    console.log(iar);
                    alert("Updated Successfully...");
                    window.location.href = this.props.context.pageContext.web.absoluteUrl;

                });

            }, (error: any): void => {

            });
    }

    private _submitform() {
        this.updateData()

    }

    public render(): React.ReactElement<IProjectSummaryProps> {
        return (
            <div className={styles.projectSummary}>
                <div className="widget-box widget-color-blue2" style={{display:"flow-root"}}>
                    <div className="widget-header">
                        <h4 className="widget-title lighter smaller">PROJECT SPACE</h4>
                    </div>

                    <div className="widget-body left">
                        <div className="widget-main padding-8">
                            <div className="row">
                                <div className="profile-user-info profile-user-info-striped">

                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Current Stage:</div>
                                        <div className="profile-info-value">
                                            <Dropdown id='StageId'
                                                defaultSelectedKey={this.state.Stage}
                                                placeholder="Select a state"
                                                label=''
                                                disabled={(this.state.isAdmin || this.state.isLiaison) ? false : true}
                                                options={this.state.Stages.map((item: any) => { return { key: item.ID, text: item.Title }; })}
                                                onChange={this._getChanges.bind(this, "Step")}
                                            />

                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Activity:</div>
                                        <div className="profile-info-value">
                                            <Dropdown id='ActivityId'
                                                defaultSelectedKey={this.state.Activity}
                                                placeholder="Select a Activity"
                                                label=''
                                                disabled={(this.state.isAdmin || this.state.isLiaison) ? false : true}
                                                options={this.state.Activities.map((item: any) => { return { key: item.ID, text: item.Title }; })}
                                                onChange={this._getChanges.bind(this, "Activity")}
                                            />
                                        </div>
                                    </div>

                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Next Stage:</div>
                                        <div className="profile-info-value">
                                            <Dropdown id='StageId'
                                                defaultSelectedKey={this.state.Stage + 1}
                                                label=''
                                                disabled={true}
                                                options={this.state.Stages.map((item: any) => { return { key: item.ID, text: item.Title }; })}
                                                onChange={this._getChanges.bind(this, "Step")}
                                            />

                                        </div>
                                    </div>

                                    <div className="profile-info-row" >
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
                                                defaultSelectedUsers={[`${this.state.liaisonEmail}`]}
                                                disabled={this.state.isAdmin ? false : true} />
                                        </div>
                                    </div>


                                    <div className="profile-info-row">
                                        <div className="profile-info-name">lates Comments</div>
                                        <div className="profile-info-value">
                                            <TextField label="" multiline rows={3} onBlur={this.handleChange.bind(this)} disabled value={this.state.items.Comments} />
                                        </div>
                                    </div>

                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Comments</div>
                                        <div className="profile-info-value">
                                            <TextField id="Comments" label="" multiline rows={3} onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>

                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Required Action</div>
                                        <div className="profile-info-value">
                                            <TextField id="ReqAction" label="" onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>

                                </div>
                            </div>
                        </div>


                    </div>

                    <div className="widget-body right">
                        <div className="widget-main padding-8">
                            <div className="">
                                <div className="profile-info-row">
                                    <h5>{this.state.items.Title}</h5>
                                </div>
                                <div className="profile-info-row">
                                    {this.state.items.ProjectDescription}
                                </div>
                                <div className="profile-info-row">
                                    <label className="blod">Investors : </label>
                                    {this.state.items.Listofinvestors}
                                </div>
                                <div className="profile-info-row">
                                    <label className="blod">Products : </label>
                                    {this.state.items.Productsandassociatedquantities}
                                </div>
                                <div className="profile-info-row">
                                    <label className="blod">CapEx : </label>
                                    {this.state.items.CapitalExpenditure}
                                </div>
                                <div className="profile-info-row">
                                    <label className="blod">Start Date : </label>
                                    {moment(this.state.startDate).format("DD/MM/YYYY")}
                                </div>
                                <div className="profile-info-row">
                                    <label className="blod">Project Action : </label>
                                    {'Accepted for Facilitation'}
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

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
                            title: 'Porject Conformation',
                            subText: 'Do you want to make this project Accepted for Facilitation?'
                        }}
                        modalProps={{
                            isBlocking: true,
                            styles: { main: { maxWidth: 450 } }
                        }}>
                        <DialogFooter>
                            <PrimaryButton onClick={this._closeDialog.bind(this)} text="Yes" />
                            <DefaultButton onClick={this._closeDialog} text="Cancel" />
                        </DialogFooter>
                    </Dialog>
                </div>


            </div>

        );
    }
}