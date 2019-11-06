import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IProjectSummarySubmitProps {
  context: WebPartContext;
}

export interface IProjectSummarySubmitState {
  multiline: boolean;
  startDate: Date; 
  items: IListItem;
  status: string;
  isAdmin: boolean;
  pjtAccepted: boolean;
  Actions: any[];
  ActionTaken: number;
  hideDialog: boolean;
  pjtSpace:string;
  spinner:boolean;
  listID:string;
  ItemId:number;
  defVale:string;
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
}


export interface IErrorLog{
  component?:string;
  Module?:string;
  page?:string;
  exception?:string;
}
