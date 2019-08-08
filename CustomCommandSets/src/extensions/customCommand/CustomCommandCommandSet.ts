import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import {sp} from '@pnp/sp';

import * as strings from 'CustomCommandCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICustomCommandCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'CustomCommandCommandSet';

export default class CustomCommandCommandSet extends BaseListViewCommandSet<ICustomCommandCommandSetProperties> {

  private enabled:boolean;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized CustomCommandCommandSet');
    sp.setup({
      spfxContext: this.context
    });
    return this.isCommandSetEnabled();
  }

  private isCommandSetEnabled(): Promise<void> {
    return sp.web.select("AllProperties").expand("AllProperties").get().then(props => {
      return sp.web.getList(this.context.pageContext.list.serverRelativeUrl).select("Id").get().then(list => {
         let lists:any[] = (props["AllProperties"]["spfxcmdsetlists"] as string).split(',');
          this.enabled = lists.indexOf(list.Id) > -1;
          return;
      });
    });
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {

    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    const compareTwoCommand: Command = this.tryGetCommand('COMMAND_2');
    if(this.enabled)
    {
      
      if (compareOneCommand) {
        // This command should be hidden unless exactly one row is selected.
        compareOneCommand.visible = event.selectedRows.length === 1;
      }
    }
    else{
      compareOneCommand.visible = compareTwoCommand.visible = false;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        Dialog.alert(`${this.properties.sampleTextOne}`);
        break;
      case 'COMMAND_2':
        Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
