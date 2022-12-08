import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IOffer } from "../model/IOffer";

export interface ISPService {
  createOffer(offer: IOffer, siteUrl: string, siteDomain: string): Promise<any>;
}

export class SPService implements ISPService {
  public static readonly serviceKey: ServiceKey<SPService> =
    ServiceKey.create<SPService>('react-office-create-offer', SPService);

  private _spHttpClient: SPHttpClient;
  private teamSiteUrl: string;
  private teamSiteDomain: string;
  private teamSiteRelativeUrl: string;

  constructor(serviceScope: ServiceScope) {
    
    serviceScope.whenFinished(() => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);        
    });
  }

  public async createOffer(offer: IOffer, siteUrl: string, siteDomain: string): Promise<any> {
    this.teamSiteUrl = siteUrl;
    this.teamSiteDomain = siteDomain;
    this.teamSiteRelativeUrl = this.teamSiteUrl.split(this.teamSiteDomain)[1];
    const tmplFile = await this.loadTemplate(offer);
    const newFile = await this.createOfferFile(tmplFile);
    const newFileUrl = this.teamSiteUrl + newFile.ServerRelativeUrl;
    const fileListItemInfo = await this.getFileListItem(tmplFile.name);    
    const fileListItem = await this.updateFileListItem(fileListItemInfo.id, fileListItemInfo.type, offer);
    console.log(fileListItem);
    return Promise.resolve({ item: null, fileUrl: newFileUrl });
  }

  private async loadTemplate (offer: IOffer): Promise<any> {
    const requestUrl: string = `${this.teamSiteUrl}/_api/web/GetFileByServerRelativeUrl('${this.teamSiteRelativeUrl}/_cts/Offering/Offering.dotx')/OpenBinaryStream()`;
    const response = await this._spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);
    const fileBlob = await response.blob();
    const respFile = { data: fileBlob.arrayBuffer, name: `${offer.title}.docx`, size: fileBlob.size };
    return respFile;
  }

  private async createOfferFile(tmplFile: any): Promise<any> {
    const uploadUrl = `${this.teamSiteUrl}/_api/web/GetFolderByServerRelativeUrl('${this.teamSiteRelativeUrl}/Shared Documents')/files/add(overwrite=true,url='${tmplFile.name}')` ;

    let spOpts : ISPHttpClientOptions  = {
      headers: {
        "Accept": "application/json",
        "Content-Length": tmplFile.size,
        "Content-Type": "application/json"
      },
      body: tmplFile.data        
    };
    const response = await this._spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, spOpts);
    return response;
  }

  private async getFileListItem(fileName: string): Promise<any> {
    const requestUrl = `${this.teamSiteUrl}/_api/web/GetFileByServerRelativeUrl('${this.teamSiteRelativeUrl}/Shared Documents/${fileName}')/ListItemAllFields`;
    const response = await this._spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);
    const jsonResp = await response.json();
    const itemID = jsonResp.ID;
    return { id: itemID, type: jsonResp["@odata.type"] }; // ToDo: ServerRedirectedEmbedUri  
  }

  private async updateFileListItem(itemID: string, itemType: string, offer: IOffer): Promise<any> {
    const requestUrl = `${this.teamSiteUrl}/_api/web/lists/GetByTitle('Documents')/items(${itemID})`;
    let spOpts : ISPHttpClientOptions  = {
      headers: {
        "Content-Type": "application/json;odata=verbose",
        "Accept": "application/json;odata=verbose",
        "If-Match": "*",
        "X-HTTP-Method": "MERGE"
      },
      body: JSON.stringify({
        "__metadata": {
            "type": itemType
        },
        "Title": offer.title,
        "OfferingDescription": offer.description,
        "OfferingVAT": offer.vat,
        "OfferingNetPrice": offer.price,
        "OfferingDate": offer.date
      })
    };
    const response = await this._spHttpClient.post(requestUrl, SPHttpClient.configurations.v1, spOpts);
    return response.json();
  }
}


