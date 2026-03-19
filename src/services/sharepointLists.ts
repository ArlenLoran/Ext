declare global {
  interface Window {
    _spPageContextInfo?: any;
  }
}

type SharePointFieldType =
  | "Text"
  | "Note"
  | "Number"
  | "Boolean"
  | "DateTime"
  | "Choice"
  | "Currency"
  | "URL"
  | "User";

interface CreateListField {
  title: string;
  type: SharePointFieldType;
  required?: boolean;
  maxLength?: number;
  choices?: string[];
  defaultValue?: string | number | boolean;
}

interface GetItemsOptions {
  select?: string[];
  expand?: string[];
  filter?: string;
  orderBy?: string;
  top?: number;
}

const spContext = {
  get webAbsoluteUrl(): string {
    const url = window._spPageContextInfo?.webAbsoluteUrl;
    if (!url) {
      throw new Error("Contexto do SharePoint não encontrado (_spPageContextInfo.webAbsoluteUrl).");
    }
    return url.replace(/\/$/, "");
  },

  get webServerRelativeUrl(): string {
    const url = window._spPageContextInfo?.webServerRelativeUrl;
    if (url === undefined) {
      throw new Error("Contexto do SharePoint não encontrado (_spPageContextInfo.webServerRelativeUrl).");
    }
    return url || "/";
  },

  get userEmail(): string {
    return window._spPageContextInfo?.userEmail || "desconhecido@dhl.com";
  },

  get userDisplayName(): string {
    return window._spPageContextInfo?.userDisplayName || "Usuário Desconhecido";
  },

  isAvailable(): boolean {
    return !!window._spPageContextInfo;
  },
};

function buildApiUrl(path: string): string {
  return `${spContext.webAbsoluteUrl}/_api${path}`;
}

async function getRequestDigest(): Promise<string> {
  const response = await fetch(buildApiUrl("/contextinfo"), {
    method: "POST",
    headers: {
      Accept: "application/json;odata=verbose",
    },
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Erro ao obter FormDigestValue: ${response.status} - ${text}`);
  }

  const data = await response.json();
  return data.d.GetContextWebInformation.FormDigestValue;
}

async function spFetch(
  url: string,
  options: RequestInit = {},
  needsDigest = false
): Promise<any> {
  const headers: Record<string, string> = {
    Accept: "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose",
    ...(options.headers as Record<string, string>),
  };

  if (needsDigest) {
    headers["X-RequestDigest"] = await getRequestDigest();
  }

  const response = await fetch(url, {
    ...options,
    headers,
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Erro SharePoint: ${response.status} - ${text}`);
  }

  if (response.status === 204) return null;
  
  const text = await response.text();
  if (!text) return null;

  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

function getListItemEntityTypeFullName(listTitle: string): Promise<string> {
  const url = buildApiUrl(`/web/lists/getbytitle('${encodeURIComponent(listTitle)}')?$select=ListItemEntityTypeFullName`);
  return spFetch(url).then((data) => data.d.ListItemEntityTypeFullName);
}

function mapFieldType(type: SharePointFieldType): number {
  const types: Record<SharePointFieldType, number> = {
    Text: 2,
    Note: 3,
    Number: 9,
    DateTime: 4,
    Boolean: 8,
    Choice: 6,
    Currency: 10,
    URL: 11,
    User: 20,
  };

  return types[type];
}

export const SharePointListsService = {
  isContextAvailable(): boolean {
    return spContext.isAvailable();
  },

  getUserInfo(): { email: string; displayName: string } {
    return {
      email: spContext.userEmail,
      displayName: spContext.userDisplayName
    };
  },

  async listExists(listTitle: string): Promise<boolean> {
    try {
      const url = buildApiUrl(`/web/lists/getbytitle('${encodeURIComponent(listTitle)}')?$select=Id,Title`);
      await spFetch(url);
      return true;
    } catch {
      return false;
    }
  },

  async createList(
    listTitle: string,
    description = "",
    template = 100
  ): Promise<void> {
    const url = buildApiUrl("/web/lists");
    const body = {
      __metadata: { type: "SP.List" },
      AllowContentTypes: true,
      BaseTemplate: template,
      ContentTypesEnabled: true,
      Description: description,
      Title: listTitle,
    };

    await spFetch(
      url,
      {
        method: "POST",
        body: JSON.stringify(body),
      },
      true
    );
  },

  async createField(listTitle: string, field: CreateListField): Promise<void> {
    const url = buildApiUrl(`/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields`);
    const fieldTypeKind = mapFieldType(field.type);

    const basePayload: any = {
      __metadata: { type: "SP.Field" },
      Title: field.title,
      FieldTypeKind: fieldTypeKind,
      Required: !!field.required,
    };

    if (field.type === "Text" && field.maxLength) {
      basePayload.MaxLength = field.maxLength;
    }

    await spFetch(
      url,
      {
        method: "POST",
        body: JSON.stringify(basePayload),
      },
      true
    );

    // Ajustes extras para Choice
    if (field.type === "Choice" && field.choices?.length) {
      const updateUrl = buildApiUrl(
        `/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields/getbyinternalnameortitle('${encodeURIComponent(field.title)}')`
      );

      const updateBody = {
        __metadata: { type: "SP.FieldChoice" },
        Choices: {
          results: field.choices,
        },
        EditFormat: 0,
      };

      await spFetch(
        updateUrl,
        {
          method: "POST",
          headers: {
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: JSON.stringify(updateBody),
        },
        true
      );
    }
  },

  async fieldExists(listTitle: string, fieldTitle: string): Promise<boolean> {
    try {
      const url = buildApiUrl(`/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields/getbyinternalnameortitle('${encodeURIComponent(fieldTitle)}')?$select=Id`);
      await spFetch(url);
      return true;
    } catch {
      return false;
    }
  },

  async ensureList(
    listTitle: string,
    description = "",
    fields: CreateListField[] = []
  ): Promise<void> {
    const exists = await this.listExists(listTitle);
    if (!exists) {
      await this.createList(listTitle, description);
    }
    
    // Check and add missing fields
    for (const field of fields) {
      const fExists = await this.fieldExists(listTitle, field.title);
      if (!fExists) {
        await this.createField(listTitle, field);
      }
    }
  },

  async getItems<T = any>(
    listTitle: string,
    options: GetItemsOptions = {}
  ): Promise<T[]> {
    const query: string[] = [];

    if (options.select?.length) query.push(`$select=${options.select.join(",")}`);
    if (options.expand?.length) query.push(`$expand=${options.expand.join(",")}`);
    if (options.filter) query.push(`$filter=${encodeURIComponent(options.filter)}`);
    if (options.orderBy) query.push(`$orderby=${encodeURIComponent(options.orderBy)}`);
    if (options.top) query.push(`$top=${options.top}`);

    const qs = query.length ? `?${query.join("&")}` : "";
    const url = buildApiUrl(`/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items${qs}`);

    const data = await spFetch(url);
    return data.d.results as T[];
  },

  async getItemById<T = any>(
    listTitle: string,
    id: number,
    options: Omit<GetItemsOptions, "filter" | "top"> = {}
  ): Promise<T | null> {
    const query: string[] = [];

    if (options.select?.length) query.push(`$select=${options.select.join(",")}`);
    if (options.expand?.length) query.push(`$expand=${options.expand.join(",")}`);

    const qs = query.length ? `?${query.join("&")}` : "";
    const url = buildApiUrl(`/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${id})${qs}`);

    try {
      const data = await spFetch(url);
      return data.d as T;
    } catch {
      return null;
    }
  },

  async getItemsByFilter<T = any>(
    listTitle: string,
    filter: string,
    options: Omit<GetItemsOptions, "filter"> = {}
  ): Promise<T[]> {
    return (this as any).getItems(listTitle, {
      ...options,
      filter,
    });
  },

  async createItem<TPayload extends Record<string, any>>(
    listTitle: string,
    payload: TPayload
  ): Promise<number> {
    const entityType = await getListItemEntityTypeFullName(listTitle);
    const url = buildApiUrl(`/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items`);

    const body = {
      __metadata: { type: entityType },
      ...payload,
    };

    const data = await spFetch(
      url,
      {
        method: "POST",
        body: JSON.stringify(body),
      },
      true
    );

    return data.d.Id;
  },

  async updateItem<TPayload extends Record<string, any>>(
    listTitle: string,
    id: number,
    payload: TPayload
  ): Promise<void> {
    const entityType = await getListItemEntityTypeFullName(listTitle);
    const url = buildApiUrl(`/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${id})`);

    const body = {
      __metadata: { type: entityType },
      ...payload,
    };

    await spFetch(
      url,
      {
        method: "POST",
        headers: {
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: JSON.stringify(body),
      },
      true
    );
  },

  async deleteItem(listTitle: string, id: number): Promise<void> {
    const url = buildApiUrl(`/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${id})`);

    await spFetch(
      url,
      {
        method: "POST",
        headers: {
          "IF-MATCH": "*",
          "X-HTTP-Method": "DELETE",
        },
      },
      true
    );
  },

  async upsertItem<TPayload extends Record<string, any>>(
    listTitle: string,
    filter: string,
    payload: TPayload
  ): Promise<{ id: number; created: boolean }> {
    const items = await (this as any).getItemsByFilter(listTitle, filter, {
      select: ["Id"],
      top: 1,
    });

    if (items.length > 0) {
      await this.updateItem(listTitle, items[0].Id, payload);
      return { id: items[0].Id, created: false };
    }

    const id = await this.createItem(listTitle, payload);
    return { id, created: true };
  },

  async downloadFile(serverRelativeUrl: string): Promise<Blob> {
    const url = buildApiUrl(`/web/getfilebyserverrelativeurl('${encodeURIComponent(serverRelativeUrl)}')/$value`);
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/octet-stream",
      },
    });

    if (!response.ok) {
      throw new Error(`Erro ao baixar arquivo: ${response.status}`);
    }

    return await response.blob();
  },

  async moveFile(sourceUrl: string, destUrl: string): Promise<void> {
    const decodedSource = decodeURIComponent(sourceUrl);
    const decodedDest = decodeURIComponent(destUrl);
    const url = buildApiUrl(`/web/GetFileByServerRelativePath(decodedurl='${decodedSource.replace(/'/g, "''")}')/moveto(newurl='${decodedDest.replace(/'/g, "''")}',flags=1)`);
    await spFetch(
      url,
      {
        method: "POST",
      },
      true
    );
  },
};
