import { IListItems } from "./IListems";


export interface IFileViewerState {
 ListItems:IListItems[];
 DistinctCategories: String[];
 docUrl:string;
 show:boolean;
}