import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery
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
    const hostName = config.sharepointHost;
    const siteUrl = config.sharepointSite;
    const listName = config.sharepointList;
    const searchQuery = query.parameters[0].value;
    const credentials = new AuthService(context);
    const token = await credentials.getUserToken(query);
    if (!token) {
      // There is no token, so the user has not signed in yet.
      return credentials.getSignInComposeExtension();
  }    
  const graphService = new GraphService(token);  
  const products = await graphService.getProducts(searchQuery);
  const categories= await graphService.getretailCategories();
  const attachments = [];
  products.forEach((obj) => {
    const template = new ACData.Template(helloWorldCard);
    const card = template.expand({
      $root: {
        title: obj.Title,
        category: obj.RetailCategory,
        categories:categories
      },
    });
    const preview = CardFactory.heroCard(obj.Title);
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
