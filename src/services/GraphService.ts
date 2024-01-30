import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { Client } from "@microsoft/microsoft-graph-client";
import axios from "axios";


export class GraphService {
    private permissionList=["User.Read.All", "User.Read", "Sites.ReadWrite.All"]; //todo move to env    
    getGraphClient(tokenCredential) {       
        const authProvider = new TokenCredentialAuthenticationProvider(
            tokenCredential,
            {
                scopes: this.permissionList,
            }
        );
        // Initialize Graph client instance with authProvider
        return Client.initWithMiddleware({
            authProvider: authProvider,
        });
    }
   // testing ME endpoint
    async getMyProfile(token) {       
        
        const endpoint = 'https://graph.microsoft.com/v1.0/me';
        const options = {
          headers: {
            Authorization: `Bearer ${token}`
          }
        };      
        // Make the request and return the response data
        try {
          const response = await axios.get<any>(endpoint, options);
          return response.data;
        } catch (error) {
          throw error;
        }
      
    }
    async getSiteId(token, hostName, siteUrl) {        
       
        const endpoint = `https://graph.microsoft.com/v1.0/sites/${hostName}:/${siteUrl}`;
        const options = {
          headers: {
            Authorization: `Bearer ${token}`
          }
        };      
        // Make the request and return the response data
        try {
          const response = await axios.get<any>(endpoint, options);
          return response.data.id;
        } catch (error) {
          throw error;
        }
      
    } 
    async getProductSite(token, siteId) {   
        
        const endpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/sites?search=Product`;
        const options = {
          headers: {
            Authorization: `Bearer ${token}`
          }
        };      
        // Make the request and return the response data
        try {
          const response = await axios.get<any>(endpoint, options);
          return response.data.value;
        } catch (error) {
          throw error;
        }
      
    }
    async getProducts(token, site, listName, searchText) {     
      const endpoint = `https://graph.microsoft.com/v1.0/sites/${site[0].id}/lists/${listName}/items?expand=fields&$filter=startswith(fields/Title,'${searchText}')`;
      const options = {
        headers: {
          Authorization: `Bearer ${token}`
        }
      };      
      // Make the request and return the response data
      try {
        const response = await axios.get<any>(endpoint, options);
        return response.data;
      } catch (error) {
        throw error;
      }
    }

    async getretailCategories(token, site, listName) {
     
      const endpoint = `https://graph.microsoft.com/v1.0/sites/${site[0].id}/lists/Products/columns/RetailCategory`;
      const options = {
        headers: {
          Authorization: `Bearer ${token}`
        }
      };      
      // Make the request and return the response data
      try {
        const response = await axios.get<any>(endpoint, options);
        return response.data.choice.choices
      } catch (error) {
        throw error;
      }
    }




    // async getSiteId(graphClient, hostName, siteUrl) {
    //     let siteId = await graphClient.api(`/sites/${hostName}:/${siteUrl}`).get();
    //     return siteId.id;
    // }

    // async getProductSite(graphClient, siteId) {
    //     let site = await graphClient.api(`/sites/${siteId}/sites?search=Product`).get();
    //     return site.value;
    // }

    // async getProducts(graphClient, site, listName, searchText) {
    //     let products = await graphClient.api(`/sites/${site[0].id}/lists/${listName}/items?expand=fields&$filter=startswith(fields/Title,'${searchText}')`).get();
    //     return products;
    // }

    // async getretailCategories(graphClient, site, listName) {
    //     let column = await graphClient.api(`/sites/${site[0].id}/lists/Products/columns/RetailCategory`).get();
    //     return column.choice.choices;
    // }

}