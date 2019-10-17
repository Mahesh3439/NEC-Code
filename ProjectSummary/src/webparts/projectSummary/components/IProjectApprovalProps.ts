import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IProjectApprovalsProps {
  context: WebPartContext;
}

export interface IProjectApprovalsState {
  items:IListItem[],
  hideDialog:boolean,
  Category:any[],
  Agency:string
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
  AgencyId?:string
}


export interface ILookupItem {
  Id?:number;
  Title?:string;

}





