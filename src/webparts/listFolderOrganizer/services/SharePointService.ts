import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { IFieldInfo, IListInfo, ListProcessingStatus } from "../models/IListInfo";
import { GroupingStrategy } from "../models/GroupingType";

export class SharePointService {
  private _spHttpClient: SPHttpClient;
  private _siteUrl: string;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this._spHttpClient = spHttpClient;
    this._siteUrl = siteUrl;
  }

  /* ------------------------------------------------------------------ */
  /*  Lists                                                              */
  /* ------------------------------------------------------------------ */

  /** Fetch all non-hidden, non-catalog custom lists (BaseTemplate 100). */
  public async getLists(): Promise<IListInfo[]> {
    const apiUrl =
      `${this._siteUrl}/_api/web/lists?` +
      `$filter=Hidden eq false and IsCatalog eq false and BaseTemplate eq 100` +
      `&$select=Id,Title,ItemCount,EnableFolderCreation,RootFolder/ServerRelativeUrl` +
      `&$expand=RootFolder` +
      `&$orderby=Title`;

    const response: SPHttpClientResponse = await this._spHttpClient.get(
      apiUrl,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Failed to fetch lists: ${response.statusText} – ${errText}`);
    }

    const data = await response.json();

    return (data.value || []).map(
      (list: any): IListInfo => ({
        id: list.Id,
        title: list.Title,
        itemCount: list.ItemCount,
        rootItemCount: 0,
        folderCreationEnabled: list.EnableFolderCreation,
        rootFolderUrl: list.RootFolder.ServerRelativeUrl,
        selected: false,
        groupingStrategy: GroupingStrategy.Date,
        levels: 1,
        sourceFieldInternalName: "",
        sourceFieldTitle: "",
        availableFields: [],
        status: ListProcessingStatus.Idle,
        statusMessage: "",
        progress: 0,
      })
    );
  }

  /**
   * Count items sitting directly in the list root folder (not in subfolders).
   * Returns -1 when the list exceeds the 5 000-item view threshold and the
   * filtered query is rejected by SharePoint.
   */
  public async getRootItemCount(listId: string, rootFolderUrl: string): Promise<number> {
    const escaped = SharePointService._escapeODataString(rootFolderUrl);
    const apiUrl =
      `${this._siteUrl}/_api/web/lists(guid'${listId}')/items?` +
      `$filter=FSObjType eq 0 and FileDirRef eq '${escaped}'` +
      `&$select=Id` +
      `&$top=5001`;

    const response: SPHttpClientResponse = await this._spHttpClient.get(
      apiUrl,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      // 500 / 503 typically means the list view threshold was exceeded
      return -1;
    }

    const data = await response.json();
    const items = data.value || [];

    // If we got exactly 5001 back, there are more than we can count this way
    return items.length > 5000 ? -1 : items.length;
  }

  /* ------------------------------------------------------------------ */
  /*  Fields                                                             */
  /* ------------------------------------------------------------------ */

  /** Get fields for a list filtered by grouping strategy type. */
  public async getListFields(
    listId: string,
    strategy: GroupingStrategy
  ): Promise<IFieldInfo[]> {
    let typeFilter: string;
    switch (strategy) {
      case GroupingStrategy.Date:
        typeFilter = "TypeAsString eq 'DateTime'";
        break;
      case GroupingStrategy.Text:
        typeFilter =
          "(TypeAsString eq 'Text' or TypeAsString eq 'Note' or TypeAsString eq 'Choice')";
        break;
      case GroupingStrategy.Choice:
        typeFilter =
          "(TypeAsString eq 'Choice' or TypeAsString eq 'MultiChoice')";
        break;
      default:
        typeFilter = "TypeAsString eq 'Text'";
    }

    const apiUrl =
      `${this._siteUrl}/_api/web/lists(guid'${listId}')/fields?` +
      `$filter=Hidden eq false and ${typeFilter}` +
      `&$select=Id,InternalName,Title,TypeAsString` +
      `&$orderby=Title`;

    const response: SPHttpClientResponse = await this._spHttpClient.get(
      apiUrl,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch fields: ${response.statusText}`);
    }

    const data = await response.json();

    return (data.value || []).map(
      (field: any): IFieldInfo => ({
        id: field.Id,
        internalName: field.InternalName,
        title: field.Title,
        typeAsString: field.TypeAsString,
      })
    );
  }

  /* ------------------------------------------------------------------ */
  /*  Enable folders                                                     */
  /* ------------------------------------------------------------------ */

  /** Turn on folder creation for a list. */
  public async enableFolderCreation(listId: string): Promise<void> {
    const apiUrl = `${this._siteUrl}/_api/web/lists(guid'${listId}')`;
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE",
      },
      body: JSON.stringify({ EnableFolderCreation: true }),
    };

    const response: SPHttpClientResponse = await this._spHttpClient.post(
      apiUrl,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`Failed to enable folder creation: ${response.statusText}`);
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Folders                                                            */
  /* ------------------------------------------------------------------ */

  /**
   * Ensure the full folder path exists under the list root.
   * Creates each level one-by-one if it doesn't already exist.
   */
  public async ensureFolder(
    listRootFolderUrl: string,
    relativeFolderPath: string
  ): Promise<void> {
    const parts = relativeFolderPath.split("/").filter((p) => p.length > 0);
    let currentPath = listRootFolderUrl;

    for (const part of parts) {
      currentPath = `${currentPath}/${part}`;
      const escaped = SharePointService._escapeODataString(currentPath);

      // Check whether folder exists
      const checkUrl = `${this._siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escaped}')`;
      const checkResp: SPHttpClientResponse = await this._spHttpClient.get(
        checkUrl,
        SPHttpClient.configurations.v1
      );

      if (checkResp.ok) {
        const folderData = await checkResp.json();
        if (folderData.Exists !== false) {
          continue; // already there
        }
      }

      // Create
      const createUrl = `${this._siteUrl}/_api/web/folders/add('${escaped}')`;
      const createResp: SPHttpClientResponse = await this._spHttpClient.post(
        createUrl,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } }
      );

      if (!createResp.ok) {
        const errText = await createResp.text();
        if (errText.toLowerCase().indexOf("already exists") === -1) {
          throw new Error(
            `Failed to create folder "${currentPath}": ${errText}`
          );
        }
      }
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Items                                                              */
  /* ------------------------------------------------------------------ */

  /** Retrieve every item (not folders) from a list, handling paging. */
  public async getListItems(
    listId: string,
    sourceFieldInternalName: string
  ): Promise<any[]> {
    const allItems: any[] = [];
    let apiUrl: string | null =
      `${this._siteUrl}/_api/web/lists(guid'${listId}')/items?` +
      `$select=Id,${sourceFieldInternalName},FileDirRef,FileRef` +
      `&$filter=FSObjType eq 0` +
      `&$top=5000`;

    while (apiUrl) {
      const response: SPHttpClientResponse = await this._spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch list items: ${response.statusText}`);
      }

      const data = await response.json();
      allItems.push(...(data.value || []));

      // Follow paging link if present
      apiUrl =
        data["odata.nextLink"] || data["@odata.nextLink"] || null;
    }

    return allItems;
  }

  /** Move a list item into a target folder via ValidateUpdateListItem. */
  public async moveItemToFolder(
    listId: string,
    itemId: number,
    targetFolderServerRelativeUrl: string
  ): Promise<void> {
    const apiUrl =
      `${this._siteUrl}/_api/web/lists(guid'${listId}')` +
      `/items(${itemId})/ValidateUpdateListItem()`;

    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
      },
      body: JSON.stringify({
        formValues: [
          {
            FieldName: "FileDirRef",
            FieldValue: targetFolderServerRelativeUrl,
          },
        ],
        bNewDocumentUpdate: false,
      }),
    };

    const response: SPHttpClientResponse = await this._spHttpClient.post(
      apiUrl,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(
        `Failed to move item ${itemId}: ${response.statusText} – ${errText}`
      );
    }

    // Inspect field-level results for errors
    const result = await response.json();
    if (result.value) {
      for (const fieldResult of result.value) {
        if (fieldResult.HasException) {
          throw new Error(
            `Error moving item ${itemId}: ${fieldResult.ErrorMessage}`
          );
        }
      }
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Folder-path generation (static helpers)                            */
  /* ------------------------------------------------------------------ */

  /** Calculate the target folder path for an item based on grouping config. */
  public static generateFolderPath(
    item: any,
    sourceFieldInternalName: string,
    strategy: GroupingStrategy,
    levels: number
  ): string {
    const value = item[sourceFieldInternalName];
    if (value === null || value === undefined || value === "") {
      return "_Uncategorized";
    }

    switch (strategy) {
      case GroupingStrategy.Date:
        return SharePointService._generateDatePath(value, levels);
      case GroupingStrategy.Text:
        return SharePointService._generateTextPath(value, levels);
      case GroupingStrategy.Choice:
        return SharePointService._generateChoicePath(value);
      default:
        return "_Uncategorized";
    }
  }

  private static _generateDatePath(dateValue: string, levels: number): string {
    const date = new Date(dateValue);
    if (isNaN(date.getTime())) return "_Uncategorized";

    const year = date.getFullYear().toString();
    const monthNum = SharePointService._pad2(date.getMonth() + 1);
    const monthName = date.toLocaleString("en-US", { month: "long" });
    const day = SharePointService._pad2(date.getDate());

    const parts: string[] = [year];
    if (levels >= 2) parts.push(`${monthNum} - ${monthName}`);
    if (levels >= 3) parts.push(day);

    return parts.join("/");
  }

  private static _generateTextPath(textValue: string, levels: number): string {
    const str = String(textValue).trim();
    if (!str) return "_Uncategorized";

    const upper = str.toUpperCase();
    const firstChar = upper.charAt(0);
    const isLetter = /^[A-Z]$/.test(firstChar);
    const level1 = isLetter ? firstChar : "#";

    const parts: string[] = [level1];
    if (levels >= 2) {
      const chars2 = upper.substring(0, 2);
      parts.push(chars2.length >= 2 ? chars2 : level1);
    }
    if (levels >= 3) {
      const chars3 = upper.substring(0, 3);
      parts.push(chars3.length >= 3 ? chars3 : parts[parts.length - 1]);
    }

    return parts.join("/");
  }

  private static _generateChoicePath(value: string): string {
    if (!value) return "_Uncategorized";
    const parts = String(value)
      .split(";#")
      .filter((v) => v.trim().length > 0);
    const choiceValue = parts[0] || "_Uncategorized";
    return SharePointService._sanitizeFolderName(choiceValue);
  }

  /* ------------------------------------------------------------------ */
  /*  Utility                                                            */
  /* ------------------------------------------------------------------ */

  /** Remove characters that are invalid in SharePoint folder names. */
  private static _sanitizeFolderName(name: string): string {
    return name.replace(/[~#%&*{}\\:<>?\/+|"]/g, "_").trim() || "_Unnamed";
  }

  /** Escape single quotes for OData string literals. */
  private static _escapeODataString(value: string): string {
    return value.replace(/'/g, "''");
  }

  /** Zero-pad a number to two digits (ES5-safe). */
  private static _pad2(n: number): string {
    return n < 10 ? "0" + n : "" + n;
  }
}
