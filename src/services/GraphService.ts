import { Client } from "@microsoft/microsoft-graph-client";

export class GraphService {
  private _token: any;
  graphClient: Client;
  constructor(token) {
    if (!token || !token.trim()) {
      throw new Error('SimpleGraphClient: Invalid token received.');
    }
    this._token = token;
    // Get an Authenticated Microsoft Graph client using the token issued to the user.
    this.graphClient = Client.init({
      authProvider: (done) => {
        done(null, this._token); // First parameter takes an error if you can't get an access token.
      }
    });
  }

  async getSiteId(hostName, siteUrl) {
    let siteId = await this.graphClient.api(`/sites/${hostName}:/${siteUrl}`).get();
    return siteId.id;
  }
  async getProductSite(siteId) {
    let site = await this.graphClient.api(`/sites/${siteId}/sites?search=Product`).get();
    return site.value;
  }

  async getProducts(site, listName, searchText) {
    let products = await this.graphClient.api(`/sites/${site[0].id}/lists/${listName}/items?expand=fields&$filter=startswith(fields/Title,'${searchText}')`).get();
    return products;
  }

  async getretailCategories(site, listName) {
    let column = await this.graphClient.api(`/sites/${site[0].id}/lists/${listName}/columns/RetailCategory`).get();
    return column.choice.choices;
  }









}