import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Label, TextField, PrimaryButton, DatePicker, Spinner, SpinnerSize, PeoplePickerItemSuggestion, thProperties } from 'office-ui-fabric-react/lib';
import styles from './ProjectSummary.module.scss';
import { IProjectSummaryProps, IProjectSummaryState, IListItem, IFieldSchema } from './IProjectSummaryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as CustomJS from 'CustomJS';
import * as $ from 'jQuery';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { SPHttpClient } from '@microsoft/sp-http';
import { Web, List, ItemAddResult } from "sp-pnp-js/lib/pnp";
import * as moment from 'moment';
import {
  Dropdown,
  IDropdown,
  DropdownMenuItemType,
  IDropdownOption
} from "office-ui-fabric-react/lib/Dropdown";
import { ListFormService } from '../Services/CommonMethods';
import { IListFormService } from '../Services/ICommonMethods';



export default class ProjectSummary extends React.Component<IProjectSummaryProps, IProjectSummaryState, {}> {

  private ProjectActions: IDropdownOption[] = [];
  private listItemEntityTypeName: string = undefined;
  private listFormService: IListFormService;
  private fields = [];
  public ItemId: number;
  public vWeb: Web = new Web(this.props.context.pageContext.web.absoluteUrl);
  private ActionTakenKey: number;
  private StageKey: number;
  private ActivityKey: number;
  public styleOptions: any;
  public etag: string = undefined;

  constructor(props: IProjectSummaryProps) {
    super(props);
    // Initiate the component state
    this.state = {
      multiline: false,
      startDate: null,
      addUsers:[],
      items: {},
      status: null,
      fieldData: [],
      disabled: false,
      isAdmin: false,
      pjtAccepted: false,
      Actions: []

    };


    // SPComponentLoader.loadScript('//www.microsofttranslator.com/ajax/v3/WidgetV3.ashx?siteData=ueOIGRSKkd965FeEGM5JtQ**', { globalExportsName: 'Translator' }).then((): void => {
    // });

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
    this.ItemId = Number(window.location.search.split("ItemId=")[1]);
    this._getProjectActions();

    if (this.ItemId) {
      this.listFormService._getListItem(this.props.context, "Projects", this.ItemId)
        .then((response) => {
          this.setState({
            items: response,
            disabled: true,
            startDate: response.ProposedStartDate ? new Date(response.ProposedStartDate) : null,
            pjtAccepted: response.ActionTakenId == 1 ? true : false
          });
          this.ActionTakenKey = Number(this.state.items.ActionTakenId);
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
            })
            return false;
          }
        }));
      });

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
  };

  private _onSelectDate = (date: Date | null | undefined): void => {
    this.setState({ startDate: date });
    this.fields.push("ProposedStartDate-label")
  };

  private _onFormatDate = (date: Date): string => {
    return date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear();
  };

  public componentDidMount(): void {
    setTimeout(function () {
      CustomJS.load();
    }, 3000);
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
      if (item.key == 1) {
        this.setState({
          pjtAccepted: true
        });
      }
      else {
        this.setState({
          pjtAccepted: false
        });
      }
    }
    else if (internalName == "Step")
      this.StageKey = Number(item.key);
    else if (internalName == "Activity")
      this.ActivityKey = Number(item.key);


  };


  private onTextChange = (newText: string) => {
    return newText;
  }


  //function to generate dynamic data body to create an item.
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
    bodyContent["PromotionType"] = "Direct"
    let body: string = JSON.stringify(bodyContent);
    return body;
  }

  private _getupdateBodyContent(listItemEntityTypeName: string) {
    let _fields = [...new Set(this.fields)];
    var bodyContent = {
      '__metadata': {
        'type': listItemEntityTypeName
      },
    }

    for (let id of _fields) {

      if (id == "ActionTaken")
      {
        bodyContent["ActionTakenId"] = this.ActionTakenKey;
        if(this.ActionTakenKey == 1)
        {
          bodyContent["LiaisonOfficerId"]= this.state.addUsers;
        }
      }
      else if (id == "Step")
        bodyContent["StageId"] = this.StageKey;
      else if (id == "Activity")
        bodyContent["ActivityId"] = this.ActivityKey;
      else {
        let value = (document.getElementById(id) as HTMLInputElement).value;
        bodyContent[id] = value;
      }

    }

    let body: string = JSON.stringify(bodyContent);
    return body;
  }

  //function to get the project Actions
  public _getProjectActions() {
    this.listFormService._getListitems(this.props.context, "Actions Master")
      .then((response) => {
        let items = response.value;
        this.setState({
          Actions: items
        });
      });
  }

  //function to capture People picker.
  private _getPeoplePickerItems(items: any[]) {
    this.state.addUsers.length = 0;
    let tempuserMngArr = [];
    for (let item in items) {
      tempuserMngArr.push(items[item].id);
    }
    this.setState({ addUsers: items });
    console.log('Items:', items);
  }

  //function to submit the Project summary and for updates
  private async SaveData(): Promise<void> {

    this.listFormService._getListItemEntityTypeName(this.props.context, "Projects")
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
  private async updateData(): Promise<void> {
    this.listFormService._getListItemEntityTypeName(this.props.context, "Projects")
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

  private _submitform() {
    if (this.ItemId)
      this.updateData();
    else
      this.SaveData();
  }


  public render(): React.ReactElement<IProjectSummaryProps> {

    // return (
    //   <div>
    //    // {this.state.items.map((item: any, inc) => {
    return (
      <div className={styles.projectSummary}>
        <div id='reactForm'>
          <div className={styles.row}>
            <TextField id="Title" label="Name of Project" placeholder="Project Title" required onBlur={this.handleChange.bind(this)} value={this.state.items.Title} disabled={this.state.disabled} />
          </div>
          <div className={styles.row}>
            <TextField id="ProjectDescription" label="Shot Description" multiline rows={3} onBlur={this.handleChange.bind(this)} value={this.state.items.ProjectDescription} disabled={this.state.disabled} />
          </div>
          <div className={styles.row}>
            <TextField id="Listofinvestors" label="Investors/Partners" placeholder="List of Investors/Partners" required onBlur={this.handleChange.bind(this)} value={this.state.items.Listofinvestors} disabled={this.state.disabled} />
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
              value={this.state.items.Productsandassociatedquantities}
              disabled={this.state.disabled}
            />
          </div>
          <div className={styles.row}>
            <TextField id="CapitalExpenditure" label="Capital Expenditure" placeholder="Capital Expenditure" onBlur={this.handleChange.bind(this)} value={this.state.items.CapitalExpenditure} disabled={this.state.disabled} />
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
              disabled={this.state.disabled}
            />
          </div>
          <div className={styles.Requirement}>
            <div className={styles.subHeader}><span>Project Specifications</span></div>
            <div className={styles.row}>
              <TextField id="Naturalgas" label="Natural Gas usage" onBlur={this.handleChange.bind(this)} value={this.state.items.Naturalgas} disabled={this.state.disabled} />
            </div>
            <div className={styles.row} id="Electricity">
              <Label className="ms-Label">Electricity consumption</Label>
              <input type="text" id="ElectricityMW" className="Electricity ms-TextField-field" onBlur={this.handleChange.bind(this)} value={this.state.items.ElectricityMW} disabled={this.state.disabled} />
              <input type="text" id="ElectricityKW" className="Electricity ms-TextField-field" onBlur={this.handleChange.bind(this)} value={this.state.items.ElectricityKW} disabled={this.state.disabled} />
            </div>
            <div className={styles.row}>
              <TextField id="Water" label="Water consumption" onBlur={this.handleChange.bind(this)} value={this.state.items.Water} disabled={this.state.disabled} />
            </div>
            <div className={styles.row}>
              <TextField id="Land" label="Land requirement" onBlur={this.handleChange.bind(this)} value={this.state.items.Land} disabled={this.state.disabled} />
            </div>
            <div className={styles.row}>
              <TextField id="Port" label="Port requirements" onBlur={this.handleChange.bind(this)} value={this.state.items.Port} disabled={this.state.disabled} />
            </div>
            <div className={styles.row}>
              <TextField id="WarehousingRequirements" label="Warehousing requirements" onBlur={this.handleChange.bind(this)} value={this.state.items.WarehousingRequirements} disabled={this.state.disabled} />
            </div>
            <div className={styles.row}>
              <TextField id="PotentialSaving" label="If Energy Efficient Project, Potential Saving" onBlur={this.handleChange.bind(this)} value={this.state.items.PotentialSaving} disabled={this.state.disabled} />
            </div>
            <div className={styles.row}>
              <TextField id="Other" label="Other" onBlur={this.handleChange.bind(this)} value={this.state.items.Other} disabled={this.state.disabled} />
            </div>
          </div>

          <div className={styles.row} style={this.state.isAdmin ? {} : { display: 'none' }}>
            <Dropdown
              id='ActionTaken'
              defaultSelectedKey={this.ActionTakenKey}
              placeholder="Select an Action"
              label='Porject Actions:'
              disabled={this.state.items.ActionTakenId == '1' ? true : false}
              options={this.state.Actions.map((item: any) => { return { key: item.ID, text: item.Title }; })}
              onChange={this._getChanges.bind(this, "ActionTaken")}
            />
          </div>

          <div className={styles.row}>
            <TextField id="Comments" label="Comments" multiline rows={3} onBlur={this.handleChange.bind(this)} value={this.state.items.Comments} />
          </div>

          {/* <div className={styles.row}>
            <RichText value={this.state.items.Comments}            
              styleOptions= {this.styleOptions}
              onChange={(text) => this.onTextChange(text)}
            />
          </div> */}


          <div className={styles.row} style={this.state.pjtAccepted ? {} : { display: 'none' }}>
            <PeoplePicker
              context={this.props.context}
              titleText="Liaison Officer"
              personSelectionLimit={1}
              groupName={""} // Leave this blank in case you want to filter from all users
              showtooltip={true}
              isRequired={true}
              disabled={this.state.isAdmin ? false : true}
              ensureUser={true}
              selectedItems={this._getPeoplePickerItems}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
              resolveDelay={1500}
            />
          </div>


        </div>

        <div className={styles.row}>
          <div id='txtAttachemtns'>
            <input id='Attachments' type='file' className='multy' multiple></input>
          </div>
        </div>

        <PrimaryButton
          text="Submit"
          onClick={() => this._submitform()}
        ></PrimaryButton>

      </div>
    )
    //   })}

    // </div>
    //);
  }
}
