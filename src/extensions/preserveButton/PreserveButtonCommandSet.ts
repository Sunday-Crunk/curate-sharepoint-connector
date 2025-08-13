import { Log } from "@microsoft/sp-core-library";
import { ITheme, getTheme } from "@fluentui/react";
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";

import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs"; // Required for web operations
import "@pnp/sp/lists"; // Required for list operations
import "@pnp/sp/fields"; // Required if working with fields, though not directly used here
import "@pnp/sp/items";

import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */

const ThemeState = (window as unknown as { __themeState__: { theme: ITheme } })
  .__themeState__;
// Get theme from global UI fabric state object if exists, if not fall back to using uifabric
export function getThemeColor(slot: string): string {
  // Added return type
  if (
    ThemeState &&
    ThemeState.theme &&
    ThemeState.theme[slot as keyof ITheme]
  ) {
    return ThemeState.theme[slot as keyof ITheme] as string;
  }
  const theme = getTheme();
  return theme[slot as keyof ITheme] as string; // Ensure return type matches
}

export interface IPreserveButtonCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = "PreserveButtonCommandSet";

// ApiKeyService class definition moved above PreserveButtonCommandSet
// to resolve "no-use-before-define" linting error for ApiKeyService.
class ApiKeyService {
  private sp: SPFI;

  constructor(context: ListViewCommandSetContext) {
    // Typed context
    this.sp = spfi().using(SPFx(context));
  }

  public async getApiKey(): Promise<{ apiKey: string; siteUrl: string }> {
    try {
      const listName = "soteria-details";

      const list = await this.sp.web.lists.getByTitle(listName);
      if (!list) {
        throw new Error(`The '${listName}' list was not found.`);
      }

      const items = await list.items
        .select("Key", "Active", "Created", "SoteriaURL")
        .filter("Active eq 'Active'")
        .orderBy("Created", false)
        .top(1)();

      console.log(`Found ${items.length} active items`);
      if (items.length > 0) {
        console.log("Item data:", items[0]);
      }

      if (items.length > 0) {
        const item = items[0];
        if (!item.SoteriaURL) {
          throw new Error(
            "Active configuration found but SoteriaURL is missing."
          );
        }

        const result = {
          apiKey: item.Key,
          siteUrl: item.SoteriaURL,
        };

        console.log("Returning result:", result);
        return result;
      }

      const noItemsError = new Error(
        `No active API key configuration found in '${listName}' list.`
      );
      console.error(noItemsError.message);
      throw noItemsError;
    } catch (error) {
      let errorMessage = "Error fetching API key configuration";
      if (error instanceof Error) {
        errorMessage = `${errorMessage}: ${error.message}`;
      } else if (typeof error === "string") {
        errorMessage = `${errorMessage}: ${error}`;
      }

      throw new Error(errorMessage);
    }
  }
}

export default class PreserveButtonCommandSet extends BaseListViewCommandSet<IPreserveButtonCommandSetProperties> {
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized PreserveButtonCommandSet");
    const preserveCommand: Command | undefined = this.tryGetCommand("PRESERVE"); // Renamed for clarity
    if (preserveCommand) {
      preserveCommand.visible = false;

      const themeBackground = getTheme().semanticColors.bodyBackground;
      // Ensure fillColor is derived correctly and safely
      const themeDarkAltColor = getThemeColor("themeDarkAlt") || "#005a9e"; // Fallback color
      const fillColor = themeDarkAltColor.replace("#", "%23");

      const isLight = (hex: string | undefined): boolean => {
        // Added undefined check for hex
        if (!hex || hex.length < 6) return true; // Default to light if hex is invalid
        return (
          0.299 * parseInt(hex.substr(1, 2), 16) +
            0.587 * parseInt(hex.substr(3, 2), 16) +
            0.114 * parseInt(hex.substr(5, 2), 16) >
          150
        );
      };

      let exportSvg: string;
      if (isLight(themeBackground)) {
        exportSvg = `data:image/svg+xml,%3Csvg width='252.99999999999997' height='204' xmlns='http://www.w3.org/2000/svg' version='1.1'%3E %3Cg%3E %3Ctitle%3ELayer 1%3C/title%3E %3Cpath stroke-width='0' id='svg_1' fill-rule='evenodd' fill='${fillColor}' d='m86.919,0.975c-15.326,2.713 -30.444,8.936 -42.309,17.416c-6.923,4.948 -21.637,19.855 -26.264,26.609c-11.209,16.362 -17.346,37.262 -17.346,59.072l0,8.928l30.127,0l30.127,0l-0.51,-5.25c-1.78,-18.306 7.716,-36.42 23.05,-43.969c6.458,-3.18 7.172,-3.302 23.184,-3.957c9.087,-0.373 25.462,-0.412 36.388,-0.089l19.866,0.589l6.859,3.287c8.975,4.301 15.155,10.594 19.604,19.962c3.368,7.093 3.555,8.034 3.555,17.927c0,9.893 -0.187,10.834 -3.555,17.927c-4.537,9.553 -10.65,15.706 -19.989,20.118c-2.402,1.13467 -4.804,2.26933 -7.206,3.404l-36.25,0.685l-36.25,0.684l0,29.412l0,29.413l39.75,-0.391c39.138,-0.384 39.883,-0.43 48.38,-2.943c16.645,-4.925 30.966,-13.156 43.343,-24.913c14.341,-13.623 23.134,-28.106 28.256,-46.538c2.492,-8.967 2.74,-11.395 2.74,-26.858c0,-15.463 -0.248,-17.891 -2.74,-26.858c-9.521,-34.263 -36.333,-61.019 -71.599,-71.45c-8.531,-2.523 -9.078,-2.555 -47.63,-2.79c-21.45,-0.131 -41.061,0.127 -43.581,0.573m-86.919,172.025l0,31l31,0l31,0l0,-31l0,-31l-31,0l-31,0l0,31'/%3E %3C/g%3E %3C/svg%3E`;
      } else {
        exportSvg = `data:image/svg+xml,%3Csvg width='252.99999999999997' height='204' xmlns='http://www.w3.org/2000/svg' version='1.1'%3E %3Cg%3E %3Ctitle%3ELayer 1%3C/title%3E %3Cpath stroke-width='0' id='svg_1' fill-rule='evenodd' fill='%23ecfff8' d='m86.919,0.975c-15.326,2.713 -30.444,8.936 -42.309,17.416c-6.923,4.948 -21.637,19.855 -26.264,26.609c-11.209,16.362 -17.346,37.262 -17.346,59.072l0,8.928l30.127,0l30.127,0l-0.51,-5.25c-1.78,-18.306 7.716,-36.42 23.05,-43.969c6.458,-3.18 7.172,-3.302 23.184,-3.957c9.087,-0.373 25.462,-0.412 36.388,-0.089l19.866,0.589l6.859,3.287c8.975,4.301 15.155,10.594 19.604,19.962c3.368,7.093 3.555,8.034 3.555,17.927c0,9.893 -0.187,10.834 -3.555,17.927c-4.537,9.553 -10.65,15.706 -19.989,20.118c-2.402,1.13467 -4.804,2.26933 -7.206,3.404l-36.25,0.685l-36.25,0.684l0,29.412l0,29.413l39.75,-0.391c39.138,-0.384 39.883,-0.43 48.38,-2.943c16.645,-4.925 30.966,-13.156 43.343,-24.913c14.341,-13.623 23.134,-28.106 28.256,-46.538c2.492,-8.967 2.74,-11.395 2.74,-26.858c0,-15.463 -0.248,-17.891 -2.74,-26.858c-9.521,-34.263 -36.333,-61.019 -71.599,-71.45c-8.531,-2.523 -9.078,-2.555 -47.63,-2.79c-21.45,-0.131 -41.061,0.127 -43.581,0.573m-86.919,172.025l0,31l31,0l31,0l0,-31l0,-31l-31,0l-31,0l0,31'/%3E %3C/g%3E %3C/svg%3E`;
      }
      preserveCommand.iconImageUrl = exportSvg;
    }

    this.context.listView.listViewStateChangedEvent.add(
      this,
      this._onListViewStateChanged
    );

    return Promise.resolve();
  }

  private extractDriveAndItemId(url: string): {
    driveId: string;
    itemId: string;
  } {
    const regex = /drives\/([^/]+)\/items\/([^?]+)/;
    const match = url.match(regex);

    if (match && match.length === 3) {
      const driveId = match[1];
      const itemId = match[2];
      return { driveId, itemId };
    } else {
      // Log the problematic URL for debugging without throwing an error that might break item processing.
      // Consider returning a default or error object if this is critical.
      console.warn(`Could not extract driveId and itemId from URL: ${url}`);
      return { driveId: "unknown", itemId: "unknown" }; // Fallback
    }
  }

  public getTenantName(context: ListViewCommandSetContext): string {
    const siteUrl = context.pageContext.web.absoluteUrl;
    const tenantName = new URL(siteUrl).hostname.split(".")[0];
    return tenantName;
  }

  public async onExecute(
    event: IListViewCommandSetExecuteEventParameters
  ): Promise<void> {
    switch (event.itemId) {
      case "PRESERVE": {
        if (!((this.context.listView.selectedRows?.length ?? 0) > 0)) {
          Dialog.alert(
            "No items selected. Please select at least one item to preserve."
          ).catch((e: Error) => {
            console.error(
              "Failed to show alert for no items selected, details: ",
              e
            );
          });
          return;
        }

        const listId: string | undefined =
          this.context.pageContext.list?.id.toString(); // Ensure it's a string
        const siteId: string = `${
          window.location.hostname
        },${this.context.pageContext.site?.id.toString()},${this.context.pageContext.web?.id.toString()}`;

        interface UserInfo {
          name: string;
          email: string;
        }

        const userInfo: UserInfo = {
          name: this.context.pageContext.user.displayName,
          email: this.context.pageContext.user.email,
        };

        if (!siteId || !listId) {
          Dialog.alert(
            "Site ID or List ID is undefined, please contact the extension developer."
          ).catch((e: Error) => {
            console.error("Failed to show alert for missing IDs, details: ", e);
          });
          return;
        }

        const reqUrls: Array<Record<string, unknown>> =
          this.context.listView.selectedRows?.map((row) => {
            const itemId: string = row
              .getValueByName("UniqueId")
              .toString()
              .replaceAll("{", "")
              .replaceAll("}", "");
            const itemName: string = row
              .getValueByName("FileLeafRef")
              .toString();
            const spUrl: string = row.getValueByName(".spItemUrl").toString();

            // Safely call extractDriveAndItemId
            let ids = { driveId: "unknown", itemId: "unknown" };
            try {
              ids = this.extractDriveAndItemId(spUrl);
            } catch (e: unknown) {
              console.error(
                `Error extracting drive/item ID for ${itemName}:`,
                e
              );
            }

            const type: string =
              row.getValueByName("FSObjType").toString() === "1"
                ? "Folder"
                : "File";
            const fileSize: string =
              row.getValueByName("File_x0020_Size")?.toString() ?? "0";

            const metadata: Record<string, string> = {};
            const allFields = row.fields;

            allFields.forEach((field) => {
              const fieldName = field.displayName;
              if (fieldName.toLowerCase().startsWith("soteria-")) {
                const namespaceKey = fieldName.substring(8);
                metadata[namespaceKey] =
                  row.getValueByName(field.internalName)?.toString() ?? "";
              }
            });

            return {
              id: itemId,
              spId: ids?.itemId,
              driveId: ids?.driveId,
              name: itemName,
              fileSize,
              type,
              metadata,
            };
          }) ?? [];

        const apiKeyService = new ApiKeyService(this.context);

        try {
          const { apiKey, siteUrl } = await apiKeyService.getApiKey();

          const sharepointRequestBody = {
            curateDetails: {
              siteUrl: siteUrl,
              apiKey: apiKey,
            },
            sharepointDetails: {
              drivePath: `${window.location.origin}/${
                window.location.pathname.split("/")[1]
              }`,
              siteId: siteId,
            },
            uploadItems: reqUrls,
            userInfo,
          };

          const response = await this.sendRequest(
            sharepointRequestBody,
            siteUrl,
            apiKey
          );

          if (!response.ok) {
            const errorText = await response.text();
            throw new Error(
              `Network response was not ok. Status: ${response.status}. Message: ${errorText}`
            );
          }

          const data = await response.json();

          if (data.success === true) {
            await Dialog.alert(
              (data?.message as string) || "Operation successful."
            );
          } else if (
            data.errors &&
            Array.isArray(data.errors) &&
            data.errors.length > 0
          ) {
            await Dialog.alert(
              `Some uploads failed to initiate: ${(
                data.errors as Array<{ item: string; error: string }>
              )
                .map(
                  (error: { item: string; error: string }) =>
                    `${error.item}: ${error.error}`
                )
                .join(", ")}`
            );
          } else {
            await Dialog.alert(
              (data?.message as string) ||
                "An unknown issue occurred during the operation."
            );
          }
        } catch (error) {
          console.error("Error during PRESERVE command execution:", error);
          await Dialog.alert(
            "An error occurred while executing the preserve command: " +
              (error instanceof Error ? error.message : "Unknown error")
          );
        }
        break;
      }
      default:
        throw new Error("Unknown command");
    }
  }

  private sendRequest(
    requestBody: Record<string, unknown>,
    targetApiUrlBase: string,
    tokenForHeader: string
  ): Promise<Response> {
    const endpointUrl = `${targetApiUrlBase.replace(
      /\/$/,
      ""
    )}/api/sharepoint/uploadSharePointPackage`;
    console.log("endpointUrl", endpointUrl);
    return fetch(
      "http://127.0.0.1:8000/api/sharepoint/uploadSharePointPackage",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${tokenForHeader}`,
        },
        body: JSON.stringify(requestBody),
      }
    );
  }

  private _onListViewStateChanged = (
    args: ListViewStateChangedEventArgs
  ): void => {
    Log.info(LOG_SOURCE, "List view state changed");
    const preserveCommand: Command | undefined = this.tryGetCommand("PRESERVE");
    if (preserveCommand) {
      preserveCommand.visible =
        (this.context.listView.selectedRows?.length ?? 0) > 0;
    }
    this.raiseOnChange();
  };
}
