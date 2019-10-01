import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Label, TextField, DatePicker, DefaultButton, PrimaryButton, Spinner, SpinnerSize, PeoplePickerItemSuggestion } from 'office-ui-fabric-react/lib';
import styles from './ProjectSummary.module.scss';
import { IProjectSummaryProps, IProjectSummaryState, IListItem, IFieldSchema } from './IProjectSummaryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as CustomJS from 'CustomJS';
import * as $ from 'jQuery';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPHttpClient } from '@microsoft/sp-http';
import {
  Dropdown,
  IDropdown,
  DropdownMenuItemType,
  IDropdownOption
} from "office-ui-fabric-react/lib/Dropdown";
import { ListFormService } from '../Services/CommonMethods';
import { IListFormService } from '../Services/ICommonMethods';



export default class ProjectSummary extends React.Component<IProjectSummaryProps, IProjectSummaryState, {}> {

  private ProjectActions: any[] = [];
  private listItemEntityTypeName: string = undefined;
  private listFormService: IListFormService;
  private fields = [];
  public ItemId: number;

  constructor(props: IProjectSummaryProps) {
    super(props);
    // Initiate the component state
    this.state = {
      multiline: false,
      startDate: null,
      addUsers: [],
      items: {},
      status: null,
      fieldData: [],
      disabled: false,
      isAdmin:false

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
    let ItemId = Number(window.location.search.split("ItemId=")[1]);
    this._getProjectActions();

    if (ItemId) {
      this.listFormService._getListItem(this.props.context, "Projects", ItemId)
        .then((response) => {
          this.setState({
            items: response,
            disabled: true,
            startDate: response.ProposedStartDate ? new Date(response.ProposedStartDate) : null
          });

        })
    }

    // this.setState({
    //   loginUser:this.props.context.pageContext.user.email
    // })

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

  //function to generate dynamic data body to create an item.
  private _getContentBody(listItemEntityTypeName: string) {
    let _fields = [...new Set(this.fields)];
    var bodyContent = {
      '__metadata': {
        'type': listItemEntityTypeName
      },
    }

    for (let id of _fields) {
      let value = (document.getElementById(id) as HTMLInputElement).value;
      bodyContent[id] = value;
    }
    let body: string = JSON.stringify(bodyContent);
    return body;
  }

  //function to get the project Actions
  public _getProjectActions() {
    let ProjectActions = this.listFormService._getListitems(this.props.context, "Actions Master")
      .then((response) => {
        let items = response.value;
        items.forEach(item => {
          this.ProjectActions.push({
            key: item.Id,
            text: item.Title
          })
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
    //this.setState({ addUsers: tempuserMngArr });
    console.log('Items:', items);
  }

  //function to submit the Project summary and for updates

  private async _submitform(): Promise<void> {
    this.setState({
      status: 'Project Submitting...',
    });

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
            <div className={styles.subHeader}><span>Requirement</span></div>
            <div className={styles.row}>
              <TextField id="Naturalgas" label="Natural Gas" onBlur={this.handleChange.bind(this)} value={this.state.items.Naturalgas} disabled={this.state.disabled} />
            </div>
            <div className={styles.row}>
              <TextField id="Electricity" label="Electricity" onBlur={this.handleChange.bind(this)} value={this.state.items.Electricity} disabled={this.state.disabled} />
            </div>
            <div className={styles.row}>
              <TextField id="Water" label="Water" onBlur={this.handleChange.bind(this)} value={this.state.items.Water} disabled={this.state.disabled} />
            </div>
            <div className={styles.row}>
              <TextField id="Land" label="Land" onBlur={this.handleChange.bind(this)} value={this.state.items.Land} disabled={this.state.disabled} />
            </div>
            <div className={styles.row}>
              <TextField id="Port" label="Port" onBlur={this.handleChange.bind(this)} value={this.state.items.Port} disabled={this.state.disabled} />
            </div>
            <div className={styles.row}>
              <TextField id="Other" label="Other" onBlur={this.handleChange.bind(this)} value={this.state.items.Other} disabled={this.state.disabled} />
            </div>
          </div>

          <div className={styles.row} style={this.state.isAdmin ? {} : { display: 'none' }}>
            <PeoplePicker
              context={this.props.context}
              titleText="Investor"
              personSelectionLimit={3}
              groupName={""} // Leave this blank in case you want to filter from all users
              showtooltip={true}
              isRequired={true}
              disabled={this.state.disabled}
              ensureUser={true}
              selectedItems={this._getPeoplePickerItems}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
              resolveDelay={1500}
            />
          </div>

          <div className={styles.row}>
            <Dropdown
              label='Porject Actions:'
              options={this.ProjectActions}
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
