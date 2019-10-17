import * as React from 'react';
import styles from './PromotionResponse.module.scss';
import { IPromotionResponseProps, IPromotionResponseState, IListItem } from './IPromotionResponseProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Label, TextField, PrimaryButton, DefaultButton, DatePicker, Checkbox } from 'office-ui-fabric-react/lib';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  Dropdown,
  IDropdown,
  DropdownMenuItemType,
  IDropdownOption
} from "office-ui-fabric-react/lib/Dropdown";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';

import { ListFormService } from '../../../Commonfiles/Services/CommonMethods';
import { IListFormService } from '../../../Commonfiles/Services/ICommonMethods';
import * as moment from 'moment';
import { SPHttpClient } from '@microsoft/sp-http';
import { string } from 'prop-types';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ListItemAttachments } from '@pnp/spfx-controls-react/lib/ListItemAttachments';
import '../../../Commonfiles/Services/customStyles.css';

export default class PromotionResponseEdit extends React.Component<IPromotionResponseProps, IPromotionResponseState, {}> {

  private listFormService: IListFormService;
  private fields = [];
  public PItemId: number;
  public PType: string;
  public liaisonofficer: number = null;
  public responseTitle: string;


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
    this.PItemId = Number(window.location.search.split("ItemId=")[1].split("&PType")[0]);
    this.PType = window.location.search.split("PType=")[1];


    if (this.PItemId) {

      if (this.PType == "EOI")
        this.responseTitle = "EOI Responses";
      else if (this.PType == "RFPP")
        this.responseTitle = "RFPP Responses";

      const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.responseTitle}')/items(${this.PItemId})`;
      this.listFormService._getListItem(this.props.context, restApi)
        .then((response) => {
          this.setState({
            items: response,
            ItemId: this.PItemId
          });
        });


      const listrestApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.responseTitle}')`;
      this.listFormService._getListItem(this.props.context, listrestApi)
        .then((response) => {
          this.setState({
            listID: response.Id
          });
        });
    }

    this._onCheckboxChange = this._onCheckboxChange.bind(this);

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

  private _getChanges = (internalName: string, event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    //console.log(`Selection change: ${item.text}  ${item.key} ${item.selected ? 'selected' : 'unselected'}`);
    this.fields.push(internalName);
    if (item.text == "Proceed with Project Development") {
      this.setState({
        hideDialog: false,
        pjtAccepted: true,
        status: item.text
      });
    }
    else {
      this.setState({
        pjtAccepted: false,
        status: item.text
      });
    }
  }



  //function to capture People picker.
  private _getPeoplePickerItems(items: any[]) {
    for (let item of items) {
      this.liaisonofficer = item.id;
    }
  }

  private _onCheckboxChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean): void {
    this.setState(
      {
        crtPjtSpace: true
      });

    this.fields.push("withPjtSpace");
    //console.log(`The option has been changed to ${isChecked}.`);
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
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
      if (id == "withPjtSpace") {
        bodyContent["PjtSpace"] = 1;
      }
      else if (id == "EOIStatus" || id == "RFPPStatus") {
        bodyContent[id] = this.state.status;
      }
      else {
        let value = (document.getElementById(id) as HTMLInputElement).value;
        bodyContent[id] = value;
      }
    }
    //bodyContent["Title"] = this.state.items.Title;
    // bodyContent["PromotionType"] = this.state.items.PromotionType;
    let body: string = JSON.stringify(bodyContent);
    return body;
  }

  private _getBodyforPDA(listItemEntityTypeName: string) {

    var bodyContent = {
      '__metadata': {
        'type': listItemEntityTypeName
      },
    };

    bodyContent["Title"] = this.state.items.Title;
    bodyContent["ProjectDescription"] = this.state.items.ProjectDescription;
    bodyContent["Listofinvestors"] = this.state.items.Listofinvestors;
    bodyContent["Productsandassociatedquantities"] = this.state.items.Productsandassociatedquantities;
    bodyContent["CapitalExpenditure"] = this.state.items.CapitalExpenditure;
    bodyContent["ProposedStartDate"] = new Date(this.state.items.ProposedStartDate);
    bodyContent["Naturalgas"] = this.state.items.Naturalgas;
    bodyContent["ElectricityMW"] = this.state.items.ElectricityMW;
    bodyContent["ElectricityKW"] = this.state.items.ElectricityKW;
    bodyContent["Water"] = this.state.items.Water;
    bodyContent["Land"] = this.state.items.Land;
    bodyContent["Port"] = this.state.items.Port;
    bodyContent["WarehousingRequirements"] = this.state.items.WarehousingRequirements;
    bodyContent["PotentialSaving"] = this.state.items.PotentialSaving;
    bodyContent["Other"] = this.state.items.Other;
    bodyContent["ActionTakenId"] = 1;
    bodyContent["LiaisonOfficerId"] = this.liaisonofficer;
    bodyContent["PromotionType"] = this.PType;
    bodyContent["ProjectStatus"] = "Accepted for Facilitation";
    bodyContent["InvestorId"] = this.state.items.AuthorId;

    // bodyContent["Title"] = this.state.items.Title;
    // bodyContent["PromotionType"] = this.state.items.PromotionType;
    let body: string = JSON.stringify(bodyContent);
    return body;
  }

  private _submitform() {
    if (this.state.crtPjtSpace == true) {
      this._createProject()
        .then((resp) => {
          let itemID = resp.Id;
          let vsiteurl = `ProjectSpace${itemID}`;
          let vsiteTitle = resp.Title;
          this.listFormService._creatProjectSpace(this.props.context, vsiteTitle, vsiteurl)
            .then((responseJSON) => {
              this.setState({
                pjtSpace: responseJSON.ServerRelativeUrl
              });

              this.listFormService._getListItemEntityTypeName(this.props.context, "Projects")
                .then(listItemEntityTypeName => {
                  const body: string = JSON.stringify({
                    '__metadata': {
                      'type': listItemEntityTypeName
                    },
                    'ProjectURL': `https://ttengage.sharepoint.com${this.state.pjtSpace}`
                  });
                  return this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Projects')/items(${itemID})`, SPHttpClient.configurations.v1, {
                    headers: {
                      'Accept': 'application/json;odata=nometadata',
                      'Content-type': 'application/json;odata=verbose',
                      'odata-version': '',
                      'IF-MATCH': '*',
                      'X-HTTP-Method': 'MERGE'
                    },
                    body: body
                  });
                })
                // .then(response => {
                //   return response.json();
                // })
                .then((resp) => {
                  console.log(resp);
                  this.updateResponse();
                });

            });
        });
    }
    else {
      this.updateResponse();
    }

  }

  public _createProject() {
    var listTitle = "Projects";

    return this.listFormService._getListItemEntityTypeName(this.props.context, listTitle)
      .then(listItemEntityTypeName => {
        let vbody: string = this._getBodyforPDA(listItemEntityTypeName);
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
      });
  }

  public updateResponse() {
    this.listFormService._getListItemEntityTypeName(this.props.context, this.responseTitle)
      .then(listItemEntityTypeName => {
        let vbody: string = this._getContentBody(listItemEntityTypeName);
        return this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.responseTitle}')/items(${this.PItemId})`, SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: vbody
        });
      }).then((resp) => {
        console.log(resp);
        alert("Updated Successfully...");
      });


  }

  public render(): React.ReactElement<IPromotionResponseProps> {
    return (
      <div className={styles.promotionResponse} >

        <div className="widget-box widget-color-blue2">
          <div className="widget-header">
            <h4 className="widget-title lighter smaller">UPDATE PROMOTION INTEREST </h4>
          </div>
          <div className="widget-body">
            <div className="widget-main padding-8">
              <div className="row">
                <div className="profile-user-info profile-user-info-striped">
                  <div className="profile-info-row">
                    <div className="profile-info-name">Promotion Title</div>
                    <div className="profile-info-value">
                      <TextField id="Title" placeholder="Project Title" value={this.state.items.Title} disabled />
                    </div>
                  </div>
                  <div className="profile-info-row">
                    <div className="profile-info-name">Project Title</div>
                    <div className="profile-info-value">
                      <TextField id="PjtTitle" value={this.state.items.PjtTitle} disabled />
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
                      <TextField id="CapitalExpenditure" label="" placeholder="Capital Expenditure" value={this.state.items.CapitalExpenditure} disabled />
                    </div>
                  </div>
                  <div className="profile-info-row">
                    <div className="profile-info-name">Proposed Start Date </div>
                    <div className="profile-info-value">
                      <DatePicker placeholder=""
                        id="ProposedStartDate"
                        onSelectDate={this._onSelectDate}
                        value={this.state.items.ProposedStartDate ? new Date(this.state.items.ProposedStartDate) : null}
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
                  <div className="profile-info-row" style={(this.state.items.PotentialSaving) ? {} : { display: 'none' }}>
                    <div className="profile-info-name">If Energy Efficient Project, Potential Saving </div>
                    <div className="profile-info-value">
                      <TextField id="PotentialSaving" label="" value={this.state.items.PotentialSaving} disabled />
                    </div>
                  </div>

                  <div className="profile-info-row" style={(this.state.items.Other) ? {} : { display: 'none' }}>
                    <div className="profile-info-name">Other </div>
                    <div className="profile-info-value">
                      <TextField id="Other" label="" value={this.state.items.Other} disabled />
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

                  <div className="profile-info-row" style={this.PType == "EOI" ? {} : { display: 'none' }}>
                    <div className="profile-info-name">Status</div>
                    <div className="profile-info-value">
                      <Dropdown label=""
                        id="EOIStatus"
                        onChange={this._getChanges.bind(this, "EOIStatus")}
                        placeholder="Select an option"
                        options={[
                          { key: '1', text: 'Proceed to RFPP' },
                          { key: '2', text: 'Proceed with Project Development' },
                          { key: '3', text: 'Rejected' },
                        ]} />
                    </div>
                  </div>
                  <div className="profile-info-row" style={this.PType == "RFPP" ? {} : { display: 'none' }}>
                    <div className="profile-info-name">Status</div>
                    <div className="profile-info-value">
                      <Dropdown label=""
                        id="EOIStatus"
                        onChange={this._getChanges.bind(this, "EOIStatus")}
                        placeholder="Select an option"
                        options={[
                          { key: '1', text: 'Proceed with Project Development' },
                          { key: '2', text: 'Rejected' }
                        ]} />
                    </div>
                  </div>
                  <div className="profile-info-row" style={this.state.pjtAccepted ? {} : { display: 'none' }}>
                    <div className="profile-info-name">Status</div>
                    <div className="profile-info-value">
                      <Checkbox label="Create a Project and its project space" onChange={this._onCheckboxChange} />
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
                    <div className="profile-info-name">lates Comments</div>
                    <div className="profile-info-value">
                      <TextField label="" multiline rows={3} onBlur={this.handleChange.bind(this)} disabled value={this.state.items.Comments} />
                    </div>
                  </div>

                  <div className="profile-info-row" >
                    <div className="profile-info-name">Comments</div>
                    <div className="profile-info-value">
                      <TextField id="Comments" label="" multiline rows={3} onBlur={this.handleChange.bind(this)} />
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
            <PrimaryButton
              text="Submit"
              onClick={this._submitform.bind(this)}
            ></PrimaryButton>
          </div> */}


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


        </div >


      </div>
    );
  }
}
