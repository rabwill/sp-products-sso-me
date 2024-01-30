import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  InvokeResponse,
} from "botbuilder";
import * as ACData from "adaptivecards-templating";
import helloWorldCard from "./adaptiveCards/helloWorldCard.json";
import { AuthService } from "./services/AuthService";
import { GraphService } from "./services/GraphService";
import config from "./config";


export class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Search.
  public async handleTeamsMessagingExtensionQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<any> {
   
    const searchQuery = query.parameters[0].value;
    const credentials = new AuthService(context);
    const token = await credentials.getUserToken(query);
    if (!token) {
      // There is no token, so the user has not signed in yet.
      return credentials.getSignInComposeExtension();
  }    
  const graphService = new GraphService(); 
//  const graphClient = graphService.getGraphClient(token);
// const me = await graphService.getMyProfile(graphClient);

  const hostName = config.sharepointHost;
  const siteUrl = config.sharepointSite;
  const listName = config.sharepointList;
  const siteId = await graphService.getSiteId(token, hostName, siteUrl)
  let site = await graphService.getProductSite(token, siteId);
  
  const products = await graphService.getProducts(token, site, listName, searchQuery);
  const categories= await graphService.getretailCategories(token,site,listName);
  const attachments = [];
  products.value.forEach((obj) => {
    const template = new ACData.Template(helloWorldCard);
    const card = template.expand({
      $root: {
        title: obj.fields.Title,
        category: obj.fields.RetailCategory,
        categories:categories
      },
    });
    const preview = CardFactory.heroCard(obj.fields.Title);
    const attachment = { ...CardFactory.adaptiveCard(card), preview };
    attachments.push(attachment);
  });

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: attachments,
    },
  };
}
}
