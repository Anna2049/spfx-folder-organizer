import { GroupingStrategy } from "./GroupingType";

export interface IListInfo {
  id: string;
  title: string;
  itemCount: number;
  rootItemCount: number;
  folderCreationEnabled: boolean;
  rootFolderUrl: string;
  // UI state
  selected: boolean;
  groupingStrategy: GroupingStrategy;
  levels: number;
  sourceFieldInternalName: string;
  sourceFieldTitle: string;
  availableFields: IFieldInfo[];
  status: ListProcessingStatus;
  statusMessage: string;
  progress: number;
}

export interface IFieldInfo {
  id: string;
  internalName: string;
  title: string;
  typeAsString: string;
}

export enum ListProcessingStatus {
  Idle = "Idle",
  Processing = "Processing",
  Done = "Done",
  Error = "Error",
}
