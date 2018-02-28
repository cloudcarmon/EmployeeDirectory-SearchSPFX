export interface IEmployeeDirectoryProps {
  description: string;
  searchQuerys:string;
  UserName:string;
  UserContact:string;
  siteurl:string;
  UserArray: Array<string>[];
  options:string;
  loading: boolean;
  Nodata:boolean;
  selectedValue:string;
}
