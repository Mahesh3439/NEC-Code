import * as React from 'react';
import styles from './PromotionResponse.module.scss';

import { IPromotionResponseProps, IPromotionResponseState, IListItem } from './IPromotionResponseProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Label, TextField, PrimaryButton, DefaultButton, DatePicker, Checkbox } from 'office-ui-fabric-react/lib';

import {
  Dropdown,
  IDropdown,
  DropdownMenuItemType,
  IDropdownOption
} from "office-ui-fabric-react/lib/Dropdown";

import { ListFormService } from '../../../Commonfiles/Services/CommonMethods';
import { IListFormService } from '../../../Commonfiles/Services/ICommonMethods';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ListItemAttachments } from '@pnp/spfx-controls-react/lib/ListItemAttachments';
import '../../../Commonfiles/Services/customStyles.css';

export default class PromotionResponse extends React.Component<IPromotionResponseProps, IPromotionResponseState, {}> {
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
      ItemId: null,
      spinner:false,
      disable:false
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
                      <TextField id="Title" value={this.state.items.Title} disabled />
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
                        value={this.state.items.ProposedStartDate ? new Date(this.state.items.ProposedStartDate) : null}
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
                      <TextField id="Water" label="" suffix="cubic meters/month" value={this.state.items.Water} disabled />
                    </div>
                  </div>
                  <div className="profile-info-row" style={(this.state.items.Land) ? {} : { display: 'none' }}>
                    <div className="profile-info-name">Land requirement </div>
                    <div className="profile-info-value">
                      <TextField id="Land" label="" suffix="hectores" value={this.state.items.Land} disabled />
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

                  <div className="profile-info-row">
                    <div className="profile-info-name">Status</div>
                    <div className="profile-info-value">
                      <TextField id="Comments" label="" multiline rows={3} value={this.PType == "EOI" ? this.state.items.EOIStatus : this.state.items.RFPPStatus} disabled />
                    </div>
                  </div>
                  <div className="profile-info-row" >
                    <div className="profile-info-name">Comments</div>
                    <div className="profile-info-value">
                      <TextField id="Comments" label="" multiline rows={3} value={this.state.items.Comments} disabled />
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
            <PrimaryButton title="Cancel" text="Cancel" allowDisabledFocus onClick={() => window.history.back()}></PrimaryButton>
          </div>
        </div >
      </div>
    );
  }
}
