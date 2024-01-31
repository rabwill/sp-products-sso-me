import { Client } from "@microsoft/microsoft-graph-client";
import { ProductItem } from "../types/ProductItems";
import config from "../config";
const listFields = [
  "id",
  "fields/Title",
  "fields/RetailCategory",
  "fields/PhotoSubmission",
  "fields/CustomerRating",
  "fields/ReleaseDate"
];

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
  async getSharePointStieId(): Promise<string> {
    const { sharepointIds } = await this.graphClient.api(`/sites/${config.sharepointHost}:/${config.sharepointSite}`).select("sharepointIds").get();
    return sharepointIds.siteId;
  }

  async getProducts(searchText): Promise<ProductItem[]>{
    const siteId = await this.getSharePointStieId();
    let products = await this.graphClient.api(`/sites/${siteId}/lists/Products/items?expand=fields&select=${listFields.join(",")}&$filter=startswith(fields/Title,'${searchText}')`).get();
    const productItems: ProductItem[]= products.value.map((item) => {
      return {
        Id: item.id,
        Title: item.fields.Title,
        RetailCategory: item.fields.RetailCategory,
        PhotoSubmission: item.fields.PhotoSubmission,
        CustomerRating: item.fields.CustomerRating,
        ReleaseDate: item.fields.ReleaseDate
      };
    }
    );
    return productItems;
  }

  async getretailCategories() {
    const siteId = await this.getSharePointStieId();
    let column = await this.graphClient.api(`/sites/${siteId}/lists/Products/columns/RetailCategory`).get();
    return column.choice.choices;
  }









}