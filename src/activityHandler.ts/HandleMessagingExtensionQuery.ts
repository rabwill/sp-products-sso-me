import {TeamsActivityHandler, CardFactory, TurnContext, MessagingExtensionQuery} from "botbuilder";
import * as AdaptiveCards from "adaptivecards-templating";
import viewProduct from "../adaptiveCards/viewProduct.json";
import { AuthService } from "../services/AuthService";
import { GraphService } from "../services/GraphService";

export const HandleMessagingExtensionQuery = async (context: TurnContext, query: MessagingExtensionQuery): Promise<any> => {   
    const searchQuery = query.parameters[0].value;
    const credentials = new AuthService(context);
    const token = await credentials.getUserToken(query);
    if (!token) {
      // There is no token, so the user has not signed in yet.
      return credentials.getSignInComposeExtension();
  }    
  const graphService = new GraphService(token);  
  const products = await graphService.getProducts(searchQuery);
  //const categories= await graphService.getretailCategories();
  const attachments = [];
  products.forEach((obj) => {
    const template = new AdaptiveCards.Template(viewProduct);
    const card = template.expand({
      $root: {
        Product: obj
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
