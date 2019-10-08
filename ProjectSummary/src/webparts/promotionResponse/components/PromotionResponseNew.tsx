import * as React from 'react';
import styles from './PromotionResponse.module.scss';
import { IPromotionResponseProps, IPromotionResponseState, IListItem } from './IPromotionResponseProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Label, TextField, PrimaryButton, DefaultButton, DatePicker } from 'office-ui-fabric-react/lib';
import { SPComponentLoader } from '@microsoft/sp-loader';

import { ListFormService } from '../../../Commonfiles/Services/CommonMethods';
import { IListFormService } from '../../../Commonfiles/Services/ICommonMethods';
import * as moment from 'moment';
import { SPHttpClient } from '@microsoft/sp-http';
import { string } from 'prop-types';

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
            PromotionType: null
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
        this.PItemId = Number(window.location.search.split("PId=")[1]);


        if (this.PItemId) {
            const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Investment Promotions')/items(${this.PItemId})?$select=Id,Title,PromotionType`;
            this.listFormService._getListItem(this.props.context, restApi)
                .then((response) => {
                    this.setState({
                        items: response
                    });
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
        }

        for (let id of _fields) {
            if (id == "ProposedStartDate-label") {
                let value = (document.getElementById(id) as HTMLInputElement).value;
                let vDate = moment(value, "DD/MM/YYYY").format("MM/DD/YYYY")
                bodyContent["ProposedStartDate"] = new Date(vDate);
            }
            else {
                let value = (document.getElementById(id) as HTMLInputElement).value;
                bodyContent[id] = value;
            }
        }
        bodyContent["Title"] = this.state.items.Title;
        bodyContent["PromotionID"] = this.PItemId;
        
       // bodyContent["PromotionType"] = this.state.items.PromotionType;
        let body: string = JSON.stringify(bodyContent);
        return body;
    }


    private _submitform() {
        var listTitle = null;

        if (this.state.items.PromotionType == "EOI")
            listTitle = "EOI Responses"
        else if (this.state.items.PromotionType == "RFPP")
            listTitle = "RFPP Responses"


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
            .then((item: IListItem): void => {
                this.setState({
                    status: `Item '${item.Title}' successfully created`
                });
            }, (error: any): void => {
                this.setState({
                    status: 'Error while creating the item: ' + error
                });
            });
    }


    public render(): React.ReactElement<IPromotionResponseProps> {
        return (
            <div className={styles.promotionResponse}>
                <div id='reactForm'>
                    <div className={styles.row}>
                        <TextField id="Title" label="Name of Project" placeholder="Project Title" required onBlur={this.handleChange.bind(this)} value={this.state.items.Title} disabled />
                    </div>
                    <div className={styles.row}>
                        <TextField id="ProjectDescription" label="Shot Description" multiline rows={3} onBlur={this.handleChange.bind(this)} />
                    </div>
                    <div className={styles.row}>
                        <TextField id="Listofinvestors" label="Investors/Partners" placeholder="List of Investors/Partners" required onBlur={this.handleChange.bind(this)} />
                    </div>
                    <div className={styles.row}>
                        <TextField
                            id="Productsandassociatedquantities"
                            label="Products & Quantity"
                            name="Productsandassociatedquantities"
                            multiline={this.state.multiline}
                            onChange={this._onChange}
                            placeholder="Products & Associated Quantity"
                            onBlur={this.handleChange.bind(this)}
                            value=""

                        />
                    </div>
                    <div className={styles.row}>
                        <TextField id="CapitalExpenditure" label="Capital Expenditure" placeholder="Capital Expenditure" onBlur={this.handleChange.bind(this)} />
                    </div>
                    <div className={styles.row}>
                        <Label >Proposed Start Date</Label>
                        <DatePicker placeholder="Select a start date..."
                            id="ProposedStartDate"
                            onSelectDate={this._onSelectDate}
                            value={this.state.startDate}
                            formatDate={this._onFormatDate}
                            minDate={new Date(2000, 12, 30)}
                            isMonthPickerVisible={false}
                        />
                    </div>
                    <div className={styles.Requirement}>
                        <div className={styles.subHeader}><span>Project Specifications</span></div>
                        <div className={styles.row} >
                            <TextField id="Naturalgas" label="Natural Gas usage" onBlur={this.handleChange.bind(this)} suffix="mmscf/d" />
                        </div>
                        <div className="{styles.row} ms-Grid-row" id="Electricity">
                            {/* <Label className="ms-Label">Electricity consumption</Label>
              <input type="text" id="ElectricityMW" className="Electricity ms-TextField-field" onBlur={this.handleChange.bind(this)} value={this.state.items.ElectricityMW} disabled={this.state.disabled}  suffix="MW" />
              <input type="text" id="ElectricityKW" className="Electricity ms-TextField-field" onBlur={this.handleChange.bind(this)} value={this.state.items.ElectricityKW} disabled={this.state.disabled} placeholder="KVA" /> */}
                            <Label className="ms-Label">Electricity consumption</Label>
                            <TextField type="text" id="ElectricityMW" className="Electricity ms-TextField-field" onBlur={this.handleChange.bind(this)} suffix="MW" />
                            <TextField type="text" id="ElectricityKW" className="Electricity ms-TextField-field" onBlur={this.handleChange.bind(this)} suffix="KVA" />

                        </div>
                        <div className={styles.row} >
                            <TextField id="Water" label="Water consumption" onBlur={this.handleChange.bind(this)} suffix="Cubic meters/Month" />
                        </div>
                        <div className={styles.row}>
                            <TextField id="Land" label="Land requirements" onBlur={this.handleChange.bind(this)} suffix="Hectores" />
                        </div>
                        <div className={styles.row}>
                            <TextField id="Port" label="Port requirements" onBlur={this.handleChange.bind(this)} />
                        </div>
                        <div className={styles.row}>
                            <TextField id="WarehousingRequirements" label="Warehousing requirements" onBlur={this.handleChange.bind(this)} />
                        </div>
                        <div className={styles.row}>
                            <TextField id="PotentialSaving" label="If Energy Efficient Project, Potential Saving" onBlur={this.handleChange.bind(this)} />
                        </div>
                        <div className={styles.row}>
                            <TextField id="Other" label="Other" onBlur={this.handleChange.bind(this)} />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <PrimaryButton
                            text="Submit"
                            onClick={() => this._submitform()}
                        ></PrimaryButton>
                    </div>
                </div>
            </div>
        );
    }
}
