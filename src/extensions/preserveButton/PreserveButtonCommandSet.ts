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
        Dialog.alert('Preserve command executed!').catch(() => {
          console.error('Failed to show alert.');
        });
  
        if ((this.context.listView.selectedRows?.length ?? 0) > 0) {
          this.context.listView.selectedRows?.forEach(row => {
            const itemId = row.getValueByName('ID');
            //const itemRelativeUrl = row.getValueByName('FileRef');
            
            // Ensure that site ID and list ID are defined

            const listId = this.context.pageContext.list?.id;
            const siteId = `${window.location.hostname},${this.context.pageContext.site?.id},${this.context.pageContext.web?.id}`
            console.log("site: ", this.context)
            
            const listServe = new ListServe(this.context);
            listServe.getList().then(key => {
                console.log("List id:", key);
            });
            console.log("file: ", row)
            if (!siteId || !listId) {
              console.error('Site ID or List ID is undefined');
              return; // Exit if we don't have essential IDs
            }
  
            // Assuming Drive ID is the same as List ID for document libraries
            const driveId = listId; // You might want to verify or retrieve this programmatically
  
            // Constructing the Graph API URL
            const graphApiUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/content`;
            console.log(`Graph API URL for item content: ${graphApiUrl}`);
  
            const apiKeyService = new ApiKeyService(this.context);
            apiKeyService.getApiKey().then(key => {
                console.log("API Key:", key);
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

class ListServe {
  private sp: SPFI;

  constructor(context: any) {
    this.sp = spfi().using(SPFx(context));
  }
  public async getList(): Promise<string> {
    try {
      // Fetching items with 'Key' field and 'Active' property set to true
      const items = await this.sp.web.lists
        .top(5)();
      console.log(items);
      if (items.length > 0) {
        return items[0].Id// Return the key from the first item where 'Active' is true
      }
      return ''; // Return empty if no keys found or none are active
    } catch (error) {
      console.error("Error fetching lists:", error);
      return ''; // Return empty or throw error as needed
    }  // Check the output in the console to see the order of items
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
