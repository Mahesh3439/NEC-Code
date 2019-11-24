import * as React from 'react';
import styles from './SiteContent.module.scss';
import { ISiteContentProps } from './ISiteContentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

import { Icon, IconButton } from 'office-ui-fabric-react';



export interface ISiteContentState {
  listData: any[];
  listType: any[]
}



export default class SiteContent extends React.Component<ISiteContentProps, ISiteContentState, {}> {

  public listType = [101, 100];

  constructor(props: ISiteContentProps) {
    super(props);
    // Initiate the component state
    this.state = {
      listData: [],
      listType: []
    }

    if (!sessionStorage.getItem("UserGroups")) {
      const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/currentuser/?$expand=groups`;
      this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
        .then(resp => { return resp.json(); })
        .then((data) => {
          sessionStorage.setItem("UserGroups", JSON.stringify(data.Groups));

          for (const group of data.Groups) {
            if (group.Title == "IF Admin") {
              sessionStorage.setItem("loginuser", "IF Admin");
              this.getLists();
            }
          }
        });
    }
    else {    
      if (sessionStorage.getItem('loginuser')) {
        this.getLists();
      }
    }
  }

  public async getLists() {
    const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=((BaseTemplate eq 100) or (BaseTemplate eq 101))&$select=BaseTemplate,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    await this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); })
      .then((data) => {
        let vListData: any[] = data.value;
        let type: any[];
        if (window.navigator.userAgent.indexOf("Trident/") > 0) {
          let cat: any[] = [...new Set(vListData.map(x => x.BaseTemplate))];
          type = cat[0]._values;
        }
        else {
          type = [...new Set(vListData.map(x => x.BaseTemplate))];
        }
        this.setState({
          listData: vListData,
          listType: type

        });
      });
  }



  public render(): React.ReactElement<ISiteContentProps> {
    return (
      <div className={styles.siteContent}>
        {
          this.listType.map((value, index) => {
            return (
              <div>
                {this.state.listData.filter(item => item.BaseTemplate == value).map((list, index) => {
                  return (
                    <div>
                      <div className="" style={{display:"table-cell"}}>
                        <div>
                          {
                            list.BaseTemplate == 101 ?
                              <div className="iconsDiv FabricFolder">
                                <IconButton iconProps={{ iconName: 'FabricFolder' }} title="View Details" ariaLabel="Info" />
                              </div> :
                              <div className="iconsDiv viewlist">
                                <IconButton iconProps={{ iconName: 'ViewList' }} title="View Details" ariaLabel="Info" />
                              </div>
                          }
                        </div>
                      </div>
                      <div style={{display:"table-cell",verticalAlign:"middle"}}>
                        <a href={list.RootFolder.ServerRelativeUrl} target='_blank'>{list.Title}</a>
                      </div>
                    </div>
                  );
                })}

              </div>
            );
          })
        }
      </div>
    );
  }
}
