import { Log } from '@microsoft/sp-core-library';
import { ITheme, getTheme } from '@fluentui/react';
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

const ThemeState = (<any>window).__themeState__;
​
// Get theme from global UI fabric state object if exists, if not fall back to using uifabric    
export function getThemeColor(slot: string) {
    if (ThemeState && ThemeState.theme && ThemeState.theme[slot]) {
        return ThemeState.theme[slot];
    }
    const theme = getTheme();
    
    return theme[slot as keyof ITheme];
}

export interface IPreserveButtonCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'PreserveButtonCommandSet';

export default class PreserveButtonCommandSet extends BaseListViewCommandSet<IPreserveButtonCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized PreserveButtonCommandSet');
    console.log('PreserveButtonCommandSet.onInit memememememememememememe');
    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('PRESERVE');
    compareOneCommand.visible = false;

    // get theme colour
    const themeBackground = getTheme().semanticColors.bodyBackground
    const fillColor = getThemeColor("themeDarkAlt").replace('#', '%23');

    // function to check if the colour is light or dark
    const isLight = (hex: any) => (0.299 * parseInt(hex.substr(1, 2), 16) + 0.587 * parseInt(hex.substr(3, 2), 16) + 0.114 * parseInt(hex.substr(5, 2), 16)) > 150;
    let exportSvg
    // if the theme is light, use a darker colour to generate the icon svg
    if (isLight(themeBackground)){
      exportSvg = `data:image/svg+xml,%3Csvg width='252.99999999999997' height='204' xmlns='http://www.w3.org/2000/svg' version='1.1'%3E %3Cg%3E %3Ctitle%3ELayer 1%3C/title%3E %3Cpath stroke-width='0' id='svg_1' fill-rule='evenodd' fill='${fillColor}' d='m86.919,0.975c-15.326,2.713 -30.444,8.936 -42.309,17.416c-6.923,4.948 -21.637,19.855 -26.264,26.609c-11.209,16.362 -17.346,37.262 -17.346,59.072l0,8.928l30.127,0l30.127,0l-0.51,-5.25c-1.78,-18.306 7.716,-36.42 23.05,-43.969c6.458,-3.18 7.172,-3.302 23.184,-3.957c9.087,-0.373 25.462,-0.412 36.388,-0.089l19.866,0.589l6.859,3.287c8.975,4.301 15.155,10.594 19.604,19.962c3.368,7.093 3.555,8.034 3.555,17.927c0,9.893 -0.187,10.834 -3.555,17.927c-4.537,9.553 -10.65,15.706 -19.989,20.118c-2.402,1.13467 -4.804,2.26933 -7.206,3.404l-36.25,0.685l-36.25,0.684l0,29.412l0,29.413l39.75,-0.391c39.138,-0.384 39.883,-0.43 48.38,-2.943c16.645,-4.925 30.966,-13.156 43.343,-24.913c14.341,-13.623 23.134,-28.106 28.256,-46.538c2.492,-8.967 2.74,-11.395 2.74,-26.858c0,-15.463 -0.248,-17.891 -2.74,-26.858c-9.521,-34.263 -36.333,-61.019 -71.599,-71.45c-8.531,-2.523 -9.078,-2.555 -47.63,-2.79c-21.45,-0.131 -41.061,0.127 -43.581,0.573m-86.919,172.025l0,31l31,0l31,0l0,-31l0,-31l-31,0l-31,0l0,31'/%3E %3C/g%3E %3C/svg%3E`
    }else{
      exportSvg = `data:image/svg+xml,%3Csvg width='252.99999999999997' height='204' xmlns='http://www.w3.org/2000/svg' version='1.1'%3E %3Cg%3E %3Ctitle%3ELayer 1%3C/title%3E %3Cpath stroke-width='0' id='svg_1' fill-rule='evenodd' fill='%23ecfff8' d='m86.919,0.975c-15.326,2.713 -30.444,8.936 -42.309,17.416c-6.923,4.948 -21.637,19.855 -26.264,26.609c-11.209,16.362 -17.346,37.262 -17.346,59.072l0,8.928l30.127,0l30.127,0l-0.51,-5.25c-1.78,-18.306 7.716,-36.42 23.05,-43.969c6.458,-3.18 7.172,-3.302 23.184,-3.957c9.087,-0.373 25.462,-0.412 36.388,-0.089l19.866,0.589l6.859,3.287c8.975,4.301 15.155,10.594 19.604,19.962c3.368,7.093 3.555,8.034 3.555,17.927c0,9.893 -0.187,10.834 -3.555,17.927c-4.537,9.553 -10.65,15.706 -19.989,20.118c-2.402,1.13467 -4.804,2.26933 -7.206,3.404l-36.25,0.685l-36.25,0.684l0,29.412l0,29.413l39.75,-0.391c39.138,-0.384 39.883,-0.43 48.38,-2.943c16.645,-4.925 30.966,-13.156 43.343,-24.913c14.341,-13.623 23.134,-28.106 28.256,-46.538c2.492,-8.967 2.74,-11.395 2.74,-26.858c0,-15.463 -0.248,-17.891 -2.74,-26.858c-9.521,-34.263 -36.333,-61.019 -71.599,-71.45c-8.531,-2.523 -9.078,-2.555 -47.63,-2.79c-21.45,-0.131 -41.061,0.127 -43.581,0.573m-86.919,172.025l0,31l31,0l31,0l0,-31l0,-31l-31,0l-31,0l0,31'/%3E %3C/g%3E %3C/svg%3E`
    }
   
    // Set the icon
    compareOneCommand.iconImageUrl = exportSvg;
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
        case 'PRESERVE':
            console.log("event info: ", event);
            if (!((this.context.listView.selectedRows?.length ?? 0) > 0)) {
                Dialog.alert('No items selected. Please select at least one item to preserve.').catch((e) => {
                    console.error('Failed to show alert, details: ', e);
                });
                return; // Exit if the selection is empty
            }

            const listId = this.context.pageContext.list?.id; // Get the list ID
            const siteId = `${window.location.hostname},${this.context.pageContext.site?.id},${this.context.pageContext.web?.id}`; // Get the site ID

            if (!siteId || !listId) {
                Dialog.alert('Site ID or List ID is undefined, please contact the extension developer').catch((e) => {
                    console.error('Failed to show alert, details: ', e);
                });
                return; // Exit if we don't have essential IDs
            }

            const reqUrls: Array<{ id: string, name: string }> = this.context.listView.selectedRows?.map(row => {
                const itemId: string = row.getValueByName('UniqueId').replaceAll("{", "").replaceAll("}", "");
                const itemName: string = row.getValueByName('FileLeafRef');
                return { id: itemId, name: itemName };
            }) ?? [];

            const apiKeyService = new ApiKeyService(this.context);
            apiKeyService.getApiKey() // Get the linked Curate API key and host url from the organisations list
                .then(curateDetails => {
                    const sharepointRequest = {
                        curateDetails: curateDetails,
                        sharepointDetails: {
                            drivePath: `${window.location.origin}/${window.location.pathname.split("/")[1]}`,
                            siteId: siteId
                        },
                        uploadItems: reqUrls
                    };

                    // Send the request to the linked Curate API
                    return fetch('http://127.0.0.1:3030/uploadSharePointPackage', { // replace with the actual URL of the server curateDetails.siteUrl in production
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify(sharepointRequest)
                    });
                })
                .then(response => {
                    if (!response.ok) {
                        throw new Error('Network response was not ok');
                    }
                    return response.json();
                })
                .then(data => {
                    console.log(data);
                    if (data.message === 'All uploads initiated successfully') {
                        Dialog.alert('All uploads initiated successfully.').catch((e) => {
                            console.error('Failed to show alert, details: ', e);
                        });
                    } else if (data.errors && data.errors.length > 0) {
                        Dialog.alert(`Some uploads failed to initiate: ${data.errors.map((error: { item: string, error: string }) => `${error.item}: ${error.error}`).join(', ')}`).catch((e) => {
                            console.error('Failed to show alert, details: ', e);
                        });
                    }
                })
                .catch(error => {
                    console.error('Error:', error);
                    // Show error
                    Dialog.alert("An error occurred while executing the preserve command: " + error.message).catch(alertError => {
                        console.error('Failed to show alert, details: ', alertError);
                    });
                });
            break;
        default:
            throw new Error('Unknown command');
    }
}


  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');
    const compareOneCommand: Command = this.tryGetCommand('PRESERVE');
    if (compareOneCommand) {
      // The command should be hidden unless one or more rows are selected.
      compareOneCommand.visible = (this.context.listView.selectedRows?.length ?? 0) > 0;
    }
    this.raiseOnChange();
  }
}

class ApiKeyService {
  private sp: SPFI;

  constructor(context: any) {
    this.sp = spfi().using(SPFx(context));
  }

  public async getApiKey(): Promise<{apiKey: string, siteUrl: string}> {
    try {
      // Fetching items with 'Key' field and 'Active' property set to true
      const items = await this.sp.web.lists.getByTitle("curate-api-key").items
        .select("Key", "Active", "Created", "CurateURL")
        .filter("Active eq 'Active'")
        .orderBy("Created", false) // Order by created date
        .top(5)(); // Get the first 5 items
  
      if (items.length > 0) {
        return { apiKey: items[0].Key, siteUrl: items[0].CurateURL }; // Return the key from the first item where 'Active' is true
      }
  
      const noItemsError = new Error("No active API key found");
      console.error(noItemsError.message);
      throw noItemsError; // Throw error if no items found or none are active
  
    } catch (error) {
      if (error instanceof Error) {
        const errorMessage = `Error fetching API key: ${error.message}`;
        Dialog.alert(errorMessage).catch((alertError) => {
          console.error('Failed to show alert, details: ', alertError);
        });
        console.error(errorMessage, error);
        throw error; // Rethrow the original error
      } else {
        const unexpectedError = new Error("Unexpected error occurred");
        Dialog.alert(unexpectedError.message).catch((alertError) => {
          console.error('Failed to show alert, details: ', alertError);
        });
        console.error(unexpectedError.message, error);
        throw unexpectedError;
      }
    }
  }
  
}
