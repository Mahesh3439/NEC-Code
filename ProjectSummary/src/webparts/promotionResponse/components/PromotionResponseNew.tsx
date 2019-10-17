import * as React from 'react';
import styles from './PromotionResponse.module.scss';
import { IPromotionResponseProps, IPromotionResponseState, IListItem } from './IPromotionResponseProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Label, TextField, PrimaryButton, DefaultButton, DatePicker } from 'office-ui-fabric-react/lib';
import { SPComponentLoader } from '@microsoft/sp-loader';
//import * as $ from 'jquery';
import * as CustomJS from 'CustomJS';

import { ListFormService } from '../../../Commonfiles/Services/CommonMethods';
import { IListFormService } from '../../../Commonfiles/Services/ICommonMethods';
import * as moment from 'moment';
import { SPHttpClient } from '@microsoft/sp-http';
import { string } from 'prop-types';
import { sp, Web } from "@pnp/sp";
import '../../../Commonfiles/Services/customStyles.css';

export default class PromotionResponseNew extends React.Component<IPromotionResponseProps, IPromotionResponseState, {}> {

    private listFormService: IListFormService;
    private fields = [];
    public PItemId: number;

    constructor(props: IPromotionResponseProps) {
        super(props);
        // Initiate the component state
        this.state = {
            multiline: false,
            startDate: null,
            addUsers: [],
            items: {},
            status: null,
            crtPjtSpace: false,
            isAdmin: false,
            pjtAccepted: false,
            showState: false,
            hideDialog: true,
            formType: "New",
            pjtSpace: null,
            PromotionType: null,
            listID: null,
            ItemId: null
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
        this.PItemId = Number(window.location.search.split("PRID=")[1]);

        if (this.PItemId) {
            const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Investment Promotions')/items(${this.PItemId})?$select=Id,Title,PromotionType`;
            this.listFormService._getListItem(this.props.context, restApi)
                .then((response) => {
                    this.setState({
                        items: response
                    });

                    CustomJS.load();
                });
        }

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

    private handleChange(event) {
        if (event.target.value !== "") {
            this.fields.push(event.target.id);
        }
    }

    private _getContentBody(listItemEntityTypeName: string) {
        let _fields = [...new Set(this.fields)];
        var bodyContent = {
            '__metadata': {
                'type': listItemEntityTypeName
            },
        };

        if(window.navigator.userAgent.indexOf("Trident/") > 0)
        {
            _fields = _fields[0]._values;
        }

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
        bodyContent["Title"] = this.state.items.Title;
        bodyContent["PromotionID"] = this.PItemId.toString();

        // bodyContent["PromotionType"] = this.state.items.PromotionType;
        let body: string = JSON.stringify(bodyContent);
        return body;
    }

    private _buttonClear() {
        $("input").val("");
        $("textarea").val("");
    }

    private _submitform() {
        var pjtTitle = $("#PjtTitle").val().toString().trim();
        let pjtDesc = $("#ProjectDescription").val().toString().trim();
        let listInvst = $("#Listofinvestors").val().toString().trim();
        let product = $("#Productsandassociatedquantities").val().toString().trim();
        let capital = $("#CapitalExpenditure").val().toString().trim();

        if (pjtTitle == "") {
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
     this._SaveData();
    }

    private _SaveData() {

        var listTitle = null;

        if (this.state.items.PromotionType == "EOI")
            listTitle = "EOI Responses";
        else if (this.state.items.PromotionType == "RFPP")
            listTitle = "RFPP Responses";

        this.listFormService._getListItemEntityTypeName(this.props.context, listTitle)
            .then(listItemEntityTypeName => {
                let vbody: string = this._getContentBody(listItemEntityTypeName);
                return this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`, SPHttpClient.configurations.v1, {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-type': 'application/json;odata=verbose',
                        'odata-version': ''
                    },
                    body: vbody
                });
            }).then(response => {
                return response.json();
            })
            .then((item) => {
                console.log(item.Id);
                let itemID = item.Id;
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

                    let ListItem = sp.web.lists.getByTitle(`${listTitle}`).items.getById(itemID);
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
            });
    }


    public render(): React.ReactElement<IPromotionResponseProps> {
        return (
            <div className={styles.promotionResponse}>

                <div className="widget-box widget-color-blue2">
                    <div className="widget-header">
                        <h4 className="widget-title lighter smaller">SUBMIT PROMOTION INTEREST </h4>
                    </div>
                    <div className="widget-body">
                        <div className="widget-main padding-8">
                            <div className="row">
                                <div className="profile-user-info profile-user-info-striped">
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Promotion Title</div>
                                        <div className="profile-info-value">
                                            <TextField id="Title" placeholder={this.state.items.Title} onBlur={this.handleChange.bind(this)} defaultValue={this.state.items.Title} disabled />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Project Title *</div>
                                        <div className="profile-info-value">
                                            <TextField id="PjtTitle" placeholder='Project Title' onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Short Description *</div>
                                        <div className="profile-info-value">
                                            <TextField id="ProjectDescription" multiline rows={3} onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>

                                    <div className="profile-info-row">
                                        <div className="profile-info-name">List of investors / Partners *</div>
                                        <div className="profile-info-value">
                                            <TextField id="Listofinvestors" placeholder="List of Investors/Partners" onBlur={this.handleChange.bind(this)} />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Products  &amp; Associated Quantity *</div>
                                        <div className="profile-info-value">
                                            <TextField id="Productsandassociatedquantities"
                                                name="Productsandassociatedquantities"
                                                multiline
                                                rows={3}
                                                placeholder="Products & Associated Quantity"
                                                onBlur={this.handleChange.bind(this)}
                                                value="" />
                                        </div>
                                    </div>
                                    <div className="profile-info-row">
                                        <div className="profile-info-name">Capital Expenditure *</div>
                                        <div className="profile-info-value">
                                            <TextField id="CapitalExpenditure" placeholder="Capital Expenditure" onBlur={this.handleChange.bind(this)} />
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
                                                isMonthPickerVisible={false}
                                            />
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
                </div>

                <div>
                    <div>
                        <div id='txtAttachemtns'>
                            <input id='Attachments' type='file' className='multy' multiple></input>
                        </div>
                    </div>

                    <div className="pull-right mtp">
                        <PrimaryButton title="Clear" text="Clear" allowDisabledFocus onClick={() => this._buttonClear()}></PrimaryButton>
                        &nbsp;&nbsp;<PrimaryButton title="Submit" text="Submit" onClick={() => this._submitform()}></PrimaryButton>
                        &nbsp;&nbsp;<PrimaryButton title="Cancel" text="Cancel" allowDisabledFocus href={this.props.context.pageContext.web.absoluteUrl}></PrimaryButton>
                    </div>

                    {/* <div className={styles.row}>
                        <PrimaryButton
                            text="Submit"
                            onClick={() => this.uploadAttachemnts()}
                        ></PrimaryButton>
                    </div> */}
                </div>

            </div>
        );
    }
}
