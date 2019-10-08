import * as React from 'react';
import { Label, TextField, Checkbox} from 'office-ui-fabric-react/lib';
import styles from './ProjectSummary.module.scss';
import { ListFormService } from '../../../Commonfiles/Services/CommonMethods';
import { IListFormService } from '../../../Commonfiles/Services/ICommonMethods';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IApprovalsProps {   
    context: WebPartContext;
  }

  export interface IApprovalsState {
      items:any[]
  }


export default class HTMLContent extends React.Component<IApprovalsProps,IApprovalsState, {}> {

    constructor(props:IApprovalsProps) {
        super(props);
        this.state = {
            items:[]
        }

        this._getApprovalsList.bind(this);
    }

    public _getApprovalsList()
    {   
        const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Approvals Master')/items`
        return this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
          .then(resp => { return resp.json(); });  

    }

    public change()
    {
        alert("Checkbox");
    }
    
    render() {
        return (
            <div>
                <div className={styles.row}>
                    <TextField label="Investors/Partners" placeholder="List of Investors/Partners" onChange={this.change.bind(this)} required />
                </div>
            </div>
        )

    }
}

