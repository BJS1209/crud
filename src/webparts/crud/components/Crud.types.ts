import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IColor, IColumn } from "office-ui-fabric-react";

export interface ICrudProps {
  description: string;
  spcontext: WebPartContext;
}

export interface ICrudState {
  associateName: string;
  age: string;
  date: Date;
  allAssociates: any[];
  open: boolean;
  selectedAssociate: any;
  color: IColor;
  columns: IColumn[];
}
