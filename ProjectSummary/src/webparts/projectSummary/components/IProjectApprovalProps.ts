import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IProjectSpace } from './IProjectSummaryProps';

export interface IProjectApprovalsProps {
  context: WebPartContext;
  onDissmissPanel:()=>void;
}

export interface IProjectApprovalsState {
  items:IListItem[];
  hideDialog:boolean;
  Category:any[];
  Agency:string;
  pjtItem:IProjectListItem;
  showPanel:boolean;
  spinner: boolean;
}

export interface IListItem {
  Id?: number;
  Title?: string;
  AgencyName?:string;
  Category?:string;
  DescriptionofApprovalProcess?:string;
  AgencyContactPersonId?:string;
  AgencyGroupId?:string;
  ApprovalOrder?:string;
  ApprovalName?:string;
  OfficeNo?:string;
  MobileNo?:string;
  Email?:string;
  ApplythruttBizLink?:boolean;
  FormLink?:string; 
  Agency?:ILookupItem;
  AgencyId?:string;
  ApprovalShortName?:string;
}

export interface IProjectListItem {
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
  InvestorId?:string;
  ApprovalsId?:any[];
}


export interface ILookupItem {
  Id?:number;
  Title?:string;
  ShortName?:string;

}





