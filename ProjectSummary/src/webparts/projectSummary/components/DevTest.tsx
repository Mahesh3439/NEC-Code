import * as React from 'react';
import { Label, TextField, Checkbox } from 'office-ui-fabric-react/lib';
import styles from './ProjectSummary.module.scss';
import { ListFormService } from '../../../Commonfiles/Services/CommonMethods';
import { IListFormService } from '../../../Commonfiles/Services/ICommonMethods';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import * as pnp from 'sp-pnp-js';
import { array } from 'prop-types';
import { Item } from 'sp-pnp-js';

export interface IApprovalsProps {
    context: WebPartContext;
}

export interface IApprovalsState {
    listitems: any[]
}


export default class HTMLContent extends React.Component<IApprovalsProps, IApprovalsState, {}> {

    constructor(props: IApprovalsProps) {
        super(props);
        this.state = {
            listitems: []
        }

    }

    public componentDidMount() {
        this._getListItems();
    }

    public _getListItems() {

        const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Investment Promotions')/items?$select=Title,DeadlineDate,PromotionType`
        this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
            .then(resp => { return resp.json(); })
            .then((response) => {
                let data = response.value;
                for (let item of data) {
                    var api = "";
                    if (item.PromotionType == "EOI") {
                        api = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EOI Responses')/items?$filter=Title eq '${item.Title}'`;
                    }
                    else {
                        api = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('RFPP Responsed')/items?$filter=Title eq '${item.Title}'`;
                    }

                    this.props.context.spHttpClient.get(api, SPHttpClient.configurations.v1)
                        .then(resp => { return resp.json(); })
                        .then((response) => {
                            let data = response.value;
                            let count = data.length;
                            if (count > 0) {
                                let title = data[0].Title;
                                this.setState({
                                    listitems: [...this.state.listitems, {
                                        'Title': title,
                                        'Count': count
                                    }],
                                })
                            }
                        })
                }
            })

        /**
                pnp.sp.web.lists.getByTitle("Investment Promotions").items.select("Title", "DeadlineDate").orderBy("Id").get()
                    .then((response: any[]) => {
                        response.forEach(function (item) {
                            pnp.sp.web.lists.getByTitle("EOI Responses").items.select("Title,PromotionType").orderBy("Id").filter("Title eq '" + item.Title + "'").getAll()
                                .then((allItems) => {
                                    if (allItems.length > 0) {
                                        let title = allItems[0]["Title"];
                                        let promotiontype = allItems[0]["PromotionType"];
                                        let count: any = String(allItems.length);
                                        this.setState({
                                            listitems: [...this.state.listitems, {
                                                'Title': title,
                                                'NumberofResponses': count,
                                                'PromotionType': promotiontype
                                            }],
                                        });
                                    }
                                });
                            // console.log(this.state.listitems);
                            //  this.setState({listitems:listofitems});
                        })
        
                    });
        
        
        */


    }

    public change() {
        alert("Checkbox");
    }

    public render(): React.ReactElement<IApprovalsProps> {
        const viewFields: IViewField[] = [
            {
                name: 'Title',
                displayName: 'Title',
                sorting: true,
                maxWidth: 350,
                isResizable: true,
            },
            {
                name: 'Count',
                displayName: "Count",
                sorting: true,
                maxWidth: 80
            }
        ];


        return (
            <div>
                <ListView
                    items={this.state.listitems}
                    viewFields={viewFields}
                    iconFieldName="ServerRelativeUrl"
                    compact={true}
                    selectionMode={SelectionMode.multiple}
                />
            </div>
        );
    }
}

