import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPromotionResponseProps { 
  context:WebPartContext;
}

export interface IPromotionResponseState {
  multiline: boolean;
  startDate: Date;
  addUsers: number[];
  items: IListItem;
  status: string;  
  crtPjtSpace: boolean;
  isAdmin: boolean;
  pjtAccepted: boolean;   
  showState: boolean;
  hideDialog: boolean;
  formType:string;
  pjtSpace:string;
  PromotionType:string;
  listID:string;
  ItemId:number;
  spinner:boolean;
  disable:boolean;

   
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
  LiaisonOfficerId?: string;
  Comments?: string;
  Status?: string;
  PromotionID?:string;
  RFPPStatus?:string;
  EOIStatus?:string; 
  PotentialSaving?: string;
  WarehousingRequirements?: string;
  ElectricityMW?: string;
  ElectricityKW?: string;
  ProjectURL?:string;
  AuthorId?:number;
  PjtTitle?:string;
  DeadlineDate?:Date;
}

export interface IErrorLog{
  component?:string;
  Module?:string;
  page?:string;
  exception?:string;
}
