export interface IODataList
{
  Id:string;
  Title:string;
  BaseTemplate:number;
}

export interface IODataView
{
  Id:string;
  Title:string;
}

export interface IOField
{
  Id:string;
  Title:string;
  Hidden: boolean;
  ReadOnlyField: boolean;
  InternalName:string;
  TypeAsString: string;
}