import {TeamsActivityHandler, CardFactory, TurnContext, MessagingExtensionQuery} from "botbuilder";
import * as AdaptiveCards from "adaptivecards-templating";
import viewProduct from "../adaptiveCards/viewProduct.json";
import { AuthService } from "../services/AuthService";
import { GraphService } from "../services/GraphService";
import { cleanupParam } from "../util";
let queryCount = 0;
export const HandleMessagingExtensionQuery = async (context: TurnContext, query: MessagingExtensionQuery): Promise<any> => {  
 
  let productName=null, retailCategory=null;
 
  if (query.parameters.length === 1 && query.parameters[0]?.name === "productName") {
    [productName, retailCategory] = (query.parameters[0]?.value.split(','));
  } else {
    productName = cleanupParam(query.parameters.find((element) => element.name === "productName")?.value);
    retailCategory = cleanupParam(query.parameters.find((element) => element.name === "targetAudience")?.value);
  }
  console.log(`🔎 Query #${++queryCount}:\nproductName=${productName}, retailCategory=${retailCategory}`);   

    const credentials = new AuthService(context);
    const token = await credentials.getUserToken(query);
    if (!token) {
      // There is no token, so the user has not signed in yet.
      return credentials.getSignInComposeExtension();
  }    
  const graphService = new GraphService(token);  
  const products = await graphService.getProducts(productName,retailCategory);
  const categories= await graphService.getretailCategories();
  const attachments = [];
  for (const obj of products) {
    const template = new AdaptiveCards.Template(viewProduct);
    const imageUrl= await graphService.getPhotoFromSharePoint("Product Imagery",obj.PhotoSubmission);
    const card = template.expand({
      $root: {
        Product: obj,
        RetailCategories: categories,
        UserId:context.activity.from.id,
        Imageurl:imageUrl
      },
    });
 
    const preview = CardFactory.heroCard(`<b>${obj.Title}</b>,
      of catgegory <br/> ${obj.RetailCategory}<br/> `,
      [imageUrl],null,null);
    const attachment = { ...CardFactory.adaptiveCard(card), preview };
    attachments.push(attachment);
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: attachments,
    },
  };
}
