export interface IGetListDataWebPartProps {
  description: string;
  SharepointList: string;
}


export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;

}