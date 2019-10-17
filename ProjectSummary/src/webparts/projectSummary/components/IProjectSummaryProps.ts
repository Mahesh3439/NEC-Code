import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IProjectSummaryProps {
  context: WebPartContext;
}

export interface IProjectSummaryState {
  multiline: boolean;
  startDate: Date;
  addUsers: number[];
  items: IListItem;
  status: string;  
  disabled: boolean;
  isAdmin: boolean;
  pjtAccepted: boolean;
  Actions: any[];
  Stages: any[];
  Activities: any[];
  ActionTaken: number;
  Stage: number;
  Activity: number;
  showState: boolean;
  hideDialog: boolean;
  formType:string;
  pjtSpace:string;
  listID:string;
  ItemId:number;
  liaisonEmail:string;
  stageStartDate:Date;
  isLiaison:boolean;
}

export interface IListItem {
  Id?: number;
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
  ActionTakenId?: string;
  LiaisonOfficerId?: string;
  Comments?: string;
  Status?: string;
  StageId?: string;
  ActivityId?: string;
  PotentialSaving?: string;
  WarehousingRequirements?: string;
  ElectricityMW?: string;
  ElectricityKW?: string;
  ProjectURL?:string;
  StageStartDate?:string;
}


export interface IProjectSpace {
  siteTitle?: string;
  siteURL?: string;
  siteOwner?: number;
  siteDesp?: string;
  investor?: number;
}
