import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { sp } from "@pnp/pnpjs";
import CustomDialog from '../components/CustomDialog';

import * as strings from 'ClosePeriodCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IClosePeriodCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ClosePeriodCommandSet';

export interface IPeriodYearItem {
  Id: number;
  Quarter: number;
  Period: number;
  Status: string;
}


let selectedItem = {} as IPeriodYearItem;

export default class ClosePeriodCommandSet extends BaseListViewCommandSet<IClosePeriodCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ClosePeriodCommandSet');

    sp.setup({
      spfxContext: this.context
    });

    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const cmd: Command = this.tryGetCommand('OPEN_PERIOD');
    const listTitle = this.context.pageContext.list.title;
    if (cmd) {
      cmd.visible = event.selectedRows.length == 1 && listTitle == "Period Year";
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'OPEN_PERIOD':
        this.getSelectedItem(event);
        let dialog: CustomDialog =  new CustomDialog();
        dialog.message = `Are you sure you want to open Quarter ${selectedItem.Quarter} for Period ${selectedItem.Period}?`;
        dialog.selectedItem = selectedItem;
        dialog.show();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private getSelectedItem(event: IListViewCommandSetExecuteEventParameters): void {
    //console.log("Start");
    selectedItem.Id = event.selectedRows[0].getValueByName("ID") as number;
    selectedItem.Quarter = event.selectedRows[0].getValueByName("Quarter_x0020_Num") as number;
    selectedItem.Period = event.selectedRows[0].getValueByName("Period_x0020_Year") as number;
    selectedItem.Status = event.selectedRows[0].getValueByName("Status");
    //console.log(selectedItem);
  }
}
