import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";    // Required for web operations
import "@pnp/sp/lists";   // Required for list operations
import "@pnp/sp/fields";  // Required if working with fields, though not directly used here
import "@pnp/sp/items";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPreserveButtonCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'PreserveButtonCommandSet';

export default class PreserveButtonCommandSet extends BaseListViewCommandSet<IPreserveButtonCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized PreserveButtonCommandSet');

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('PRESERVE');
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'PRESERVE':
        // Using Dialog.alert just to show an alert; remove if not needed
        Dialog.alert('Preserve command executed!').catch(() => {
          // handle any error from the Dialog.alert
          console.error('Failed to show alert.');
        });

        // Safely logging selected items data
        if ((this.context.listView.selectedRows?.length ?? 0) > 0) {
          this.context.listView.selectedRows?.forEach(row => {
            const isadData = row.fields.filter(i=>i.displayName.includes("ISAD-"))
            // Log each selected item's ID and all other column values
            console.log(`Selected item ID: ${row.getValueByName('ID')}`);
            console.log(`Selected item data: `, row);
            console.log(`Selected item ISAD Data: `, isadData);
            const apiKeyService = new ApiKeyService(this.context);
            apiKeyService.getApiKey().then(key => {
                console.log("API Key:", key);
                // Further actions such as API calls can be handled here
            });
          });
        } else {
          console.log('No items selected.');
        }
        break;
      default:
        throw new Error('Unknown command');
    }
  }
  

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareOneCommand: Command = this.tryGetCommand('PRESERVE');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = (this.context.listView.selectedRows?.length ?? 0) > 0;

    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
class ApiKeyService {
  private sp: SPFI;

  constructor(context: any) {
    this.sp = spfi().using(SPFx(context));
  }

  public async getApiKey(): Promise<string> {
    try {
      // Fetching items with 'Key' field and 'Active' property set to true
      const items = await this.sp.web.lists.getByTitle("curate-api-key").items
        .select("Key", "Active", "Created")
        .filter("Active eq 'Active'")
        .orderBy("Created", false)
        .top(5)();
      console.log(items);  // Check the output in the console to see the order of items

      if (items.length > 0) {
        return items[0].Key; // Return the key from the first item where 'Active' is true
      }
      return ''; // Return empty if no keys found or none are active
    } catch (error) {
      console.error("Error fetching API key:", error);
      return ''; // Return empty or throw error as needed
    }
  }
}
