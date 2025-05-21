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

// Modal creation function
function createApiKeyModal(onSubmit: (value: string) => void) {
  // Modified onSubmit to specify string value
  // CSS styles
  const modalStyles = `
    .modal-backdrop {
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background-color: rgba(0, 0, 0, 0.5);
      display: flex;
      justify-content: center;
      align-items: center;
      z-index: 1000;
    }
    .modal-content {
      background-color: #ffffff;
      padding: 20px;
      border-radius: 4px;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
      width: 300px;
    }
    .modal-title {
      font-size: 18px;
      font-weight: 600;
      margin-bottom: 15px;
    }
    .modal-input {
      width: 100%;
      padding: 8px;
      margin-bottom: 15px;
      border: 1px solid #c8c8c8;
      border-radius: 2px;
      box-sizing: border-box;
    }
    .modal-button {
      padding: 8px 16px;
      margin-right: 10px;
      border: none;
      border-radius: 2px;
      cursor: pointer;
    }
    .modal-button-primary {
      background-color: #0078d4;
      color: white;
    }
    .modal-button-secondary {
      background-color: #f4f4f4;
      color: #333;
    }
    `;

  // Create and inject stylesheet
  const styleSheet = document.createElement("style");
  styleSheet.innerText = modalStyles;
  document.head.appendChild(styleSheet);
  const modalHtml = `
      <div class="modal-backdrop">
        <div class="modal-content">
          <div class="modal-title">Enter API Key</div>
          <input type="text" id="apiKeyInput" class="modal-input" placeholder="Enter your API key here">
          <button id="cancelButton" class="modal-button modal-button-secondary">Cancel</button>
          <button id="submitButton" class="modal-button modal-button-primary">Submit</button>
        </div>
      </div>
    `;

  const modalContainer = document.createElement("div");
  modalContainer.innerHTML = modalHtml;
  document.body.appendChild(modalContainer);

  const apiKeyInput = <HTMLInputElement>document.getElementById("apiKeyInput");
  const cancelButton = <HTMLButtonElement>(
    document.getElementById("cancelButton")
  );
  const submitButton = <HTMLButtonElement>(
    document.getElementById("submitButton")
  );

  function closeModal() {
    document.body.removeChild(modalContainer);
    document.head.removeChild(styleSheet); // Clean up stylesheet
  }

  cancelButton?.addEventListener("click", closeModal);

  submitButton?.addEventListener("click", () => {
    const apiKey = apiKeyInput?.value;
    if (apiKey) {
      // Ensure apiKey is not undefined
      onSubmit(apiKey);
    }
    closeModal();
  });

  return {
    open: () => {
      document.body.appendChild(modalContainer); // Ensure modal is in DOM before focusing
      apiKeyInput?.focus();
    },
    close: closeModal,
  };
}

const ThemeState = (<any>window).__themeState__;
// Get theme from global UI fabric state object if exists, if not fall back to using uifabric
export function getThemeColor(slot: string): string {
  // Added return type
  if (ThemeState && ThemeState.theme && ThemeState.theme[slot]) {
    return ThemeState.theme[slot];
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

interface OIDCConfig {
  clientId: string;
  redirectUri: string;
  authorizationEndpoint: string;
  tokenEndpoint: string;
}

class SimpleOIDCClient {
  private config: OIDCConfig;

  constructor(config: OIDCConfig) {
    this.config = config;
  }

  getAuthorizationUrl(scope: string = "openid profile email"): string {
    const params = new URLSearchParams({
      client_id: this.config.clientId,
      redirect_uri: this.config.redirectUri,
      scope,
      response_type: "code",
      state: this.generateRandomState(),
    });

    return `${this.config.authorizationEndpoint}?${params.toString()}`;
  }

  async exchangeCodeForToken(code: string): Promise<any> {
    const params = new URLSearchParams({
      grant_type: "authorization_code",
      code,
      redirect_uri: this.config.redirectUri,
      client_id: this.config.clientId,
    });

    const response = await fetch(this.config.tokenEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: params.toString(),
    });

    if (!response.ok) {
      const errorBody = await response.text(); // Read error body for more details
      throw new Error(
        `Failed to exchange code for token. Status: ${response.status}. Body: ${errorBody}`
      );
    }

    return response.json();
  }

  private generateRandomState(): string {
    return Math.random().toString(36).substring(2, 15);
  }
}

function startAuthFlow(curateUrl: string): void {
  const config: OIDCConfig = {
    clientId: "cells-client",
    redirectUri: "https://" + curateUrl + "/oauth2/oob",
    authorizationEndpoint: "https://" + curateUrl + "/oidc/oauth2/auth",
    tokenEndpoint: "https://" + curateUrl + "/oidc/oauth2/token",
  };
  const client = new SimpleOIDCClient(config);

  // Start auth flow
  const authUrl = client.getAuthorizationUrl();
  window.open(authUrl, "_blank");
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
    const regex = /drives\/([^\/]+)\/items\/([^?]+)/;
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
      case "PRESERVE":
        if (!((this.context.listView.selectedRows?.length ?? 0) > 0)) {
          Dialog.alert(
            "No items selected. Please select at least one item to preserve."
          ).catch((e) => {
            console.error(
              "Failed to show alert for no items selected, details: ",
              e
            );
          });
          return;
        }

        const listId = this.context.pageContext.list?.id.toString(); // Ensure it's a string
        const siteId = `${
          window.location.hostname
        },${this.context.pageContext.site?.id.toString()},${this.context.pageContext.web?.id.toString()}`;

        interface UserInfo {
          name: string; // Changed String to string
          email: string; // Changed String to string
        }

        const userInfo: UserInfo = {
          name: this.context.pageContext.user.displayName,
          email: this.context.pageContext.user.email,
        };

        if (!siteId || !listId) {
          Dialog.alert(
            "Site ID or List ID is undefined, please contact the extension developer."
          ).catch((e) => {
            console.error("Failed to show alert for missing IDs, details: ", e);
          });
          return;
        }

        const reqUrls: Array<any> = // Consider a more specific type for items in reqUrls
          this.context.listView.selectedRows?.map((row) => {
            const itemId: string = row
              .getValueByName("UniqueId")
              .toString() // Ensure string
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
            } catch (e) {
              console.error(
                `Error extracting drive/item ID for ${itemName}:`,
                e
              );
              // Potentially skip this item or mark it as problematic
            }

            const type: string =
              row.getValueByName("FSObjType").toString() === "1"
                ? "Folder"
                : "File";
            const fileSize: string =
              row.getValueByName("File_x0020_Size")?.toString() ?? "0"; // Handle potential undefined

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
        apiKeyService
          .getApiKey() // This returns: Promise<{ apiKey: string; siteUrl: string }>
          .then(async (curateServiceDetails) => {
            // curateServiceDetails contains { apiKey, siteUrl }
            const { apiKey: initialApiKey, siteUrl: curateApiBaseUrl } =
              curateServiceDetails;

            const curateDetailsForBody = {
              siteUrl: curateApiBaseUrl,
              // Add any other non-sensitive properties from curateServiceDetails if needed for the body
            };

            const sharepointRequestBody = {
              curateDetails: curateDetailsForBody, // API key is NOT in here
              sharepointDetails: {
                drivePath: `${window.location.origin}/${
                  window.location.pathname.split("/")[1]
                }`,
                siteId: siteId,
              },
              uploadItems: reqUrls,
              userInfo,
            };

            let finalApiKeyForHeader: string;

            if (initialApiKey === "client") {
              return new Promise<Response>((resolvePromise, rejectPromise) => {
                // Added rejectPromise
                const openAuthModal = () => {
                  // Renamed to avoid conflict with createApiKeyModal
                  const modal = createApiKeyModal((authCode: string) => {
                    // authCode from modal
                    const config: OIDCConfig = {
                      clientId: "cells-client",
                      redirectUri: curateApiBaseUrl + "/oauth2/oob", // Using https already in startAuthFlow
                      authorizationEndpoint:
                        curateApiBaseUrl + "/oidc/oauth2/auth",
                      tokenEndpoint: curateApiBaseUrl + "/oidc/oauth2/token",
                    };
                    const client = new SimpleOIDCClient(config);
                    client
                      .exchangeCodeForToken(authCode)
                      .then((tokenResponse) => {
                        if (!tokenResponse.access_token) {
                          throw new Error(
                            "Access token not found in OIDC response."
                          );
                        }
                        finalApiKeyForHeader = tokenResponse.access_token;
                        resolvePromise(
                          this.sendRequest(
                            sharepointRequestBody,
                            curateApiBaseUrl,
                            finalApiKeyForHeader
                          )
                        );
                      })
                      .catch((oidcError) => {
                        // Catch errors from OIDC exchange
                        console.error("OIDC token exchange failed:", oidcError);
                        Dialog.alert(
                          `Authentication failed: ${oidcError.message}`
                        ).catch((e) => console.error(e));
                        rejectPromise(oidcError); // Propagate error
                      });
                  });
                  modal.open();
                };
                openAuthModal();
                // Ensure curateApiBaseUrl for startAuthFlow does not include "https://" prefix if startAuthFlow adds it
                const CfgCurateUrl = curateApiBaseUrl.startsWith("https://")
                  ? curateApiBaseUrl.substring(8)
                  : curateApiBaseUrl;
                startAuthFlow(CfgCurateUrl);
              });
            } else {
              finalApiKeyForHeader = initialApiKey;
              return this.sendRequest(
                sharepointRequestBody,
                curateApiBaseUrl, // This is the full https://... URL
                finalApiKeyForHeader
              );
            }
          })
          .then((response) => {
            if (!response) return; // Guard against undefined response if promise was rejected earlier
            if (!response.ok) {
              // Try to get error message from response body for better diagnostics
              return response.text().then((text) => {
                throw new Error(
                  `Network response was not ok. Status: ${response.status}. Message: ${text}`
                );
              });
            }
            return response.json();
          })
          .then((data) => {
            if (!data) return; // Guard against undefined data
            if (data.success === true) {
              Dialog.alert(data?.message || "Operation successful.").catch(
                (e) => {
                  console.error("Failed to show success alert, details: ", e);
                }
              );
            } else if (data.errors && data.errors.length > 0) {
              Dialog.alert(
                `Some uploads failed to initiate: ${data.errors
                  .map(
                    (error: { item: string; error: string }) =>
                      `${error.item}: ${error.error}`
                  )
                  .join(", ")}`
              ).catch((e) => {
                console.error(
                  "Failed to show partial failure alert, details: ",
                  e
                );
              });
            } else {
              // Handle cases where success is not true, but no specific errors are provided
              Dialog.alert(
                data?.message ||
                  "An unknown issue occurred during the operation."
              ).catch((e) => console.error(e));
            }
          })
          .catch((error) => {
            console.error("Error during PRESERVE command execution:", error);
            Dialog.alert(
              "An error occurred while executing the preserve command: " +
                (error.message || "Unknown error")
            ).catch((alertError) => {
              console.error(
                "Failed to show general error alert, details: ",
                alertError
              );
            });
          });
        break;
      default:
        throw new Error("Unknown command");
    }
  }

  private sendRequest(
    requestBody: any,
    targetApiUrlBase: string, // e.g., "https://curate.example.com"
    tokenForHeader: string // The Bearer token
  ): Promise<Response> {
    // Ensure targetApiUrlBase is a valid base URL before appending path
    const endpointUrl = `${targetApiUrlBase.replace(
      /\/$/,
      ""
    )}/api/sharepoint/uploadSharePointPackage`;

    return fetch(endpointUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${tokenForHeader}`, // Using Bearer token
      },
      body: JSON.stringify(requestBody),
    });
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

class ApiKeyService {
  private sp: SPFI;

  constructor(context: ListViewCommandSetContext) {
    // Typed context
    this.sp = spfi().using(SPFx(context));
  }

  public async getApiKey(): Promise<{ apiKey: string; siteUrl: string }> {
    try {
      const list = await this.sp.web.lists.getByTitle("soteria-details");
      if (!list) {
        // This check might not be hit if getByTitle throws, but good for safety
        throw new Error("The 'soteria-details' list was not found.");
      }

      const items = await list.items
        .select(
          "Key",
          "Active",
          "Created",
          "SoteriaURL",
          "EnableAPIKeyAuthentication"
        )
        .filter("Active eq 'Active'") // Assuming 'Active' is a choice or text field with this exact value
        .orderBy("Created", false)
        .top(1)(); // Get only the most recent active item

      if (items.length > 0) {
        const item = items[0];
        if (!item.SoteriaURL) {
          throw new Error(
            "Active configuration found but SoteriaURL is missing."
          );
        }
        return {
          apiKey: item.EnableAPIKeyAuthentication
            ? item.Key || "client"
            : "client", // Fallback if Key is empty but auth enabled
          siteUrl: item.SoteriaURL, // This should be the full base URL e.g. https://curate.example.com
        };
      }

      const noItemsError = new Error(
        "No active API key configuration found in 'soteria-details' list."
      );
      console.error(noItemsError.message);
      // Dialog.alert might be too intrusive here if this is a background check; consider logging primarily.
      // await Dialog.alert(noItemsError.message);
      throw noItemsError;
    } catch (error) {
      let errorMessage = "Error fetching API key configuration";
      if (error instanceof Error) {
        errorMessage = `${errorMessage}: ${error.message}`;
      } else if (typeof error === "string") {
        errorMessage = `${errorMessage}: ${error}`;
      }
      console.error(errorMessage, error);
      // Avoid cascading Dialog.alert if the caller will also show one.
      // await Dialog.alert(errorMessage);
      throw new Error(errorMessage); // Re-throw a generic or specific error
    }
  }
}
