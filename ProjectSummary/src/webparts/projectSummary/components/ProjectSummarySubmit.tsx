import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Label, TextField, PrimaryButton, DefaultButton, DatePicker } from 'office-ui-fabric-react/lib';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';

import styles from './ProjectSummary.module.scss';
//import { IProjectSummaryProps, IProjectSummaryState, IListItem, IProjectSpace } from './IProjectSummaryProps';
import { IProjectSummarySubmitProps, IProjectSummarySubmitState, IListItem } from './IProjectSummarySubmitProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as CustomJS from 'CustomJS';
//import * as $ from 'jQuery';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPHttpClient } from '@microsoft/sp-http';
import * as moment from 'moment';
import {
    Dropdown,
    IDropdown,
    DropdownMenuItemType,
    IDropdownOption
} from "office-ui-fabric-react/lib/Dropdown";
import { ListFormService } from '../../../Commonfiles/Services/CommonMethods';
import { IListFormService } from '../../../Commonfiles/Services/ICommonMethods';
import { sp, Web } from "@pnp/sp";
import "@pnp/polyfill-ie11";
import '../../../Commonfiles/Services/customStyles.css';


export default class ProjectSummarySubmit extends React.Component<IProjectSummarySubmitProps, IProjectSummarySubmitState, {}> {

    private listFormService: IListFormService;
    private fields = [];
    public ItemId: number;
    private ActionTakenKey: number = null;
    public liaisonofficer: number = null;
    public investor: number = null;
    public PjtState: string;

    constructor(props: IProjectSummarySubmitProps) {
        super(props);
        // Initiate the component state
        this.state = {
            multiline: false,
            startDate: null,
            items: {},
            status: null,
            isAdmin: false,
            pjtAccepted: false,
            Actions: [],
            ActionTaken: null,
            hideDialog: true,
            pjtSpace: null
           
        };
        SPComponentLoader.loadScript('https://ttengage.sharepoint.com/sites/ttEngage_Dev/SiteAssets/jquery.js', {
            globalExportsName: 'jQuery'
        }).catch((error) => {

        }).then((): Promise<{}> => {
            return SPComponentLoader.loadScript('https://ttengage.sharepoint.com/sites/ttEngage_Dev/SiteAssets/jquery.MultiFile.js', {
                globalExportsName: 'jQuery'
            });
        }).catch((error) => {

        })

        this.listFormService = new ListFormService(props.context.spHttpClient);
        this._getProjectActions();

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

        setTimeout(function () {
            CustomJS.load();
        }, 3000);

        /**
          this.setState({
            loginUser:this.props.context.pageContext.user.email
          })
      
          let data = moment("2/10/2019", "DD/MM/YYYY").format("MM/DD/YYYY");
          console.log(data);
      
          */

    }

    //Method to convert single line text to multy line field in html.
    private _onChange = (ev: any, newText: string): void => {
        const newMultiline = newText.length > 50;
        if (newMultiline !== this.state.multiline) {
            this.setState({ multiline: newMultiline });
        }
    }

    private _onSelectDate = (date: Date | null | undefined): void => {
        this.setState({ startDate: date });
        this.fields.push("ProposedStartDate-label");
    }

    private _onFormatDate = (date: Date): string => {
        return date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear();
    }

    //   public componentDidMount(): void {
    //     setTimeout(function () {
    //       CustomJS.load();
    //     }, 3000);
    //   }

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

    //function to generate dynamic data body to create an item.
    private _getContentBody(listItemEntityTypeName: string) {
        let _fields = [...new Set(this.fields)];

        if(window.navigator.userAgent.indexOf("Trident/") > 0)
        {
            _fields = _fields[0]._values;
        }

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
            else if (id == "ActionTaken") {
                bodyContent["ActionTakenId"] = this.ActionTakenKey;
                if (this.ActionTakenKey == 1) {
                    bodyContent["LiaisonOfficerId"] = this.liaisonofficer;
                }

            }
            else {
                let value = (document.getElementById(id) as HTMLInputElement).value;
                bodyContent[id] = value;
            }
        }
        bodyContent["PromotionType"] = "Direct";
        bodyContent["InvestorId"] = Number($('.ms-Persona')[0].id);
        bodyContent["ProjectStatus"] = this.ActionTakenKey == 1 ? "Accepted for Facilitation" : "Submitted";

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

    //function to capture People picker.
    private _getInvestorItems(items: any[]) {
        for (let item of items) {
            this.investor = item.id;
        }
    }

    private _buttonClear() {
        $("input").val("");
        $("textarea").val("");
    }

    public ItemAttachments() {
        let attachemnts = $("#Attachments input:file");
        if (attachemnts.length > 1) {
            var itemAttachments = []
            $.each(attachemnts, function (index, file) {
                let afile = file as HTMLInputElement;
                if (afile.files.length > 0) {
                    itemAttachments.push({
                        name: afile.files[0].name,
                        content: afile.files[0]
                    });
                }
            });

            let ListItem = sp.web.lists.getByTitle(`Projects`).items.getById(this.ItemId);
            ListItem.attachmentFiles.addMultiple(itemAttachments)
                .then(r => {
                    console.log(r);
                    alert("Successfully submitted....");
                    window.location.href = this.props.context.pageContext.web.absoluteUrl;
                });

        }
        else {
            alert("Successfully submitted....");
            window.location.href = this.props.context.pageContext.web.absoluteUrl;
        }

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
        var pjtTitle = $("#Title").val().toString().trim();
        let pjtDesc = $("#ProjectDescription").val().toString().trim();
        let listInvst = $("#Listofinvestors").val().toString().trim();
        let product = $("#Productsandassociatedquantities").val().toString().trim();
        let capital = $("#CapitalExpenditure").val().toString().trim();

        if (pjtTitle == "")  {
            return alert("Please Enter Project Title");            
        }
        if (pjtDesc == "") {
            return alert("Please Enter Shot Description");            
        }
        if (listInvst == "") {
            return alert("Please Enter list of Investors");            
        }
        if (product == "") {
            return alert("Please Enter Products & Associated Quantity");            
        }
        if (capital == "") {
            return alert("Please Enter Capital Expenditure");            
        }



        this.SaveData()
            .then((resp) => {
                this.ItemId = resp.Id;
                if (this.state.pjtAccepted == true) {
                    this.ProjectSpace();
                }
                else {
                    this.ItemAttachments();
                }
            })
    }

    public ProjectSpace() {
        let vsiteurl = `ProjectSpace${this.ItemId}`;
        let vsiteTitle = this.state.items.Title;
        let vsiteDesp = this.state.items.ProjectDescription;

        this.listFormService._creatProjectSpace(this.props.context, vsiteTitle, vsiteurl)
            .then((responseJSON) => {
                this.fields.push("ProjectURL");
                this.setState({
                    pjtSpace: responseJSON.ServerRelativeUrl
                })

                this.listFormService._getListItemEntityTypeName(this.props.context, "Projects")
                    .then(listItemEntityTypeName => {
                        const body: string = JSON.stringify({
                            '__metadata': {
                                'type': listItemEntityTypeName
                            },
                            'ProjectURL': `https://ttengage.sharepoint.com${this.state.pjtSpace}`
                        });
                        return this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Projects')/items(${this.ItemId})`, SPHttpClient.configurations.v1, {
                            headers: {
                                'Accept': 'application/json;odata=nometadata',
                                'Content-type': 'application/json;odata=verbose',
                                'odata-version': '',
                                'IF-MATCH': '*',
                                'X-HTTP-Method': 'MERGE'
                            },
                            body: body
                        });
                    }).then((resp) => {
                        console.log(resp);
                        this.ItemAttachments();
                    })
            });
    }


    public render(): React.ReactElement<IProjectSummarySubmitProps> {
        return (
            <div className={styles.projectSummary}>

                <div className="widget-box widget-color-blue2">
                    <div className="widget-header">
                        <h4 className="widget-title lighter smaller">PROJECT SUMMARY </h4>
                    </div>
                    <div className="widget-body">
                        <div className="widget-main padding-8">
                            <div className="row">
                                <div className="profile-user-info profile-user-info-striped">
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Name of Project <span style={{color:"red"}}>*</span></div>
                                        <div className="profile-info-value">
                                            <TextField id="Title" placeholder="Project Title" onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Short Description <span style={{color:"red"}}>*</span></div>
                                        <div className="profile-info-value">
                                            <TextField id="ProjectDescription" multiline rows={3} onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>

                                    <div className="profile-info-row">
                                        <div className="profile-info-name">List of investors / Partners <span style={{color:"red"}}>*</span></div>
                                        <div className="profile-info-value">
                                            <TextField id="Listofinvestors" placeholder="List of Investors/Partners" onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Products  &amp; Associated Quantity <span style={{color:"red"}}>*</span></div>
                                        <div className="profile-info-value">
                                            <TextField id="Productsandassociatedquantities"
                                                name="Productsandassociatedquantities"
                                                multiline
                                                rows={3}
                                                placeholder="Products & Associated Quantity"
                                                onBlur={this.handleChange.bind(this)}
                                            />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Capital Expenditure <span style={{color:"red"}}>*</span></div>
                                        <div className="profile-info-value">
                                            <TextField id="CapitalExpenditure" placeholder="Capital Expenditure" onBlur={this.handleChange.bind(this)} suffix="US$MM" />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Proposed Start Date </div>
                                        <div className="profile-info-value">
                                            <DatePicker placeholder="Select a start date..."
                                                id="ProposedStartDate"
                                                onSelectDate={this._onSelectDate}
                                                value={this.state.startDate}
                                                formatDate={this._onFormatDate}
                                                minDate={new Date()}
                                                isMonthPickerVisible={false} />
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
                                        <div className="profile-info-name">Natural Gas</div>
                                        <div className="profile-info-value">
                                            <TextField id="Naturalgas" onBlur={this.handleChange.bind(this)} suffix="mmscf/d" />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                    </div>
                                    <div className="profile-info-row">
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Electricity </div>
                                        <div className="profile-info-value">
                                            <TextField type="text" id="ElectricityMW" className="Electricity ms-TextField-field" onBlur={this.handleChange.bind(this)} suffix="MW" />
                                            <TextField type="text" id="ElectricityKW" className="Electricity ms-TextField-field" onBlur={this.handleChange.bind(this)} suffix="kVA" />

                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Water consumption</div>
                                        <div className="profile-info-value">
                                            <TextField id="Water" label="" onBlur={this.handleChange.bind(this)} suffix="m³/month" />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Land requirement </div>
                                        <div className="profile-info-value">
                                            <TextField id="Land" label="" onBlur={this.handleChange.bind(this)} suffix="hectares" />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Port requirements </div>
                                        <div className="profile-info-value">
                                            <TextField id="Port" label="" onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Warehousing requirement </div>
                                        <div className="profile-info-value">
                                            <TextField id="WarehousingRequirements" label="" onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">If Energy Efficient Project, Potential Saving </div>
                                        <div className="profile-info-value">
                                            <TextField id="PotentialSaving" label="" onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>

                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Other </div>

                                        <div className="profile-info-value">

                                            <TextField id="Other" label="" onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>


                    <div className="widget-body">
                        <div className="widget-main padding-8">
                            <div className="row">
                                <div className="profile-user-info profile-user-info-striped">

                                    <div className="profile-info-row" style={(this.state.isAdmin) ? {} : { display: 'none' }}>
                                        <div className="profile-info-name">Project Actions:</div>
                                        <div className="profile-info-value">
                                            <Dropdown
                                                id='ActionTaken'
                                                defaultSelectedKey={this.state.ActionTaken}
                                                placeholder="Select an Action"
                                                disabled={this.state.items.ActionTakenId == '1' ? true : false}
                                                options={this.state.Actions.map((item: any) => { return { key: item.ID, text: item.Title }; })}
                                                onChange={this._getChanges.bind(this, "ActionTaken")}
                                            />
                                        </div>
                                    </div>
                                    <div className="profile-info-row" style={(this.state.isAdmin) ? {} : { display: 'none' }}>
                                        <div className="profile-info-name">Investor</div>
                                        <div className="profile-info-value">
                                            <PeoplePicker context={this.props.context}
                                                titleText=""
                                                personSelectionLimit={1}
                                                groupName={""}
                                                showtooltip={true}
                                                isRequired={true}
                                                ensureUser={true}
                                                selectedItems={this._getInvestorItems.bind(this)}
                                                showHiddenInUI={false}
                                                principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                                                resolveDelay={1500}
                                                defaultSelectedUsers={[`${this.props.context.pageContext.user.email}`]} />
                                        </div>
                                    </div>
                                    <div className="profile-info-row" style={this.state.pjtAccepted ? {} : { display: 'none' }}>
                                        <div className="profile-info-name">Liaison Officer</div>
                                        <div className="profile-info-value">
                                            <PeoplePicker context={this.props.context}
                                                titleText=""
                                                personSelectionLimit={1}
                                                groupName={""}
                                                showtooltip={true}
                                                isRequired={true}
                                                disabled={this.state.isAdmin ? false : true}
                                                ensureUser={true}
                                                selectedItems={this._getPeoplePickerItems.bind(this)}
                                                showHiddenInUI={false}
                                                principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                                                resolveDelay={1500} />
                                        </div>
                                    </div>

                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Comments</div>
                                        <div className="profile-info-value">
                                            <TextField id="Comments" label="" multiline rows={3} onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>

                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <div>
                    <div className="profile-info-row">
                        <div id='txtAttachemtns'>
                            <input id='Attachments' type='file' className='multy'></input>
                        </div>
                    </div>

                    <div className="pull-right mtp">
                        <PrimaryButton title="Clear" text="Clear" allowDisabledFocus onClick={() => this._buttonClear()}></PrimaryButton>
                        &nbsp;&nbsp;<PrimaryButton title="Submit" text="Submit" onClick={() => this._submitform()}></PrimaryButton>
                        &nbsp;&nbsp;<PrimaryButton title="Cancel" text="Cancel" allowDisabledFocus href={this.props.context.pageContext.web.absoluteUrl}></PrimaryButton>
                    </div>

                </div>

                <div>
                    <Dialog
                        hidden={this.state.hideDialog}
                        onDismiss={this._closeDialog}
                        dialogContentProps={{
                            type: DialogType.normal,
                            title: 'Porject Conformation',
                            subText: 'Do you want to make this project Accepted for Facilitation?'
                        }}
                        modalProps={{
                            isBlocking: true,
                            styles: { main: { maxWidth: 450 } }
                        }}
                    >
                        <DialogFooter>
                            <PrimaryButton onClick={this._closeDialog} text="Yes" />
                            <DefaultButton onClick={this._closeDialog} text="Cancel" />
                        </DialogFooter>
                    </Dialog>
                </div>
            </div>
        )
    }
}
