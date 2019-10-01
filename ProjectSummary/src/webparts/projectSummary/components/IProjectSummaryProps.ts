import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IProjectSummaryProps {
  description: string;
  context: WebPartContext;
}

export interface IProjectSummaryState {
  multiline: boolean;
  startDate: Date;
  addUsers: number[];
  items: IListItem;
  status:string;
  fieldData:IFieldSchema[];
  disabled : boolean;
  isAdmin:boolean;
  
}

export interface IListItem {
  Id?:number;
  Title?: string;
  InvestorName?: string;
  CompanyName?: string;
  PromotionType?: string;
  ProjectName?: string;
  PromotionTypeSubject?: string;
  ProjectDescription?: string;
  Listofinvestors?: string;
  Productsandassociatedquantities?: string;
  Naturalgas?: string;
  Electricity?: string;
  Water?: string;
  Land?: string;
  Port?: string;
  Other?: string;
  CapitalExpenditure?: string;
  ProposedStartDate?: string;
  ActionTaken?: string;
  LiaisonOfficer?: string;
  Comments?: string;
  Status?: string;
  Stage?: string;
  Activity?: string;
}


export interface IFieldSchema {
internalName:string;
value:string|number|Date;
}
