import { BaseListViewCommandSet, IListViewCommandSetListViewUpdatedParameters, IListViewCommandSetExecuteEventParameters } from "@microsoft/sp-listview-extensibility";
export interface IExportLisrItemsCommandSetProperties {
}
export default class ExportListItemsCommandSet extends BaseListViewCommandSet<IExportListItemsCommandSetProperties> {
    private _wb;
    private _viewColumns;
    private _listTitle;
    onInit(): Promise<void>;
    onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void;
    onExecute(event: IListViewCommandSetExecuteEventParameters): void;
    private _getFieldValueAsText;
    private writeToExcel;
    private getViewColumns;
    private Initiate;
}
