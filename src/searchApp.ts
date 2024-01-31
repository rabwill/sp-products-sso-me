import {TeamsActivityHandler,TurnContext, MessagingExtensionQuery, AdaptiveCardInvokeValue, AdaptiveCardInvokeResponse, CardFactory, MessageFactory, InvokeResponse} from "botbuilder";
import { HandleMessagingExtensionQuery } from "./activityHandler.ts/HandleMessagingExtensionQuery";
import { HandleTeamsTaskModuleFetch } from "./activityHandler.ts/HandleTeamsTaskModuleFetch";
import * as AdaptiveCards from "adaptivecards-templating";
export class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
    this.onMessage(async (context, next) => {
      await next();
    });
  }
  public override async handleTeamsMessagingExtensionQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<any> {
    return await HandleMessagingExtensionQuery(context, query);
  }
  public override async handleTeamsTaskModuleFetch(context, invokeValue): Promise<any> {
    return await HandleTeamsTaskModuleFetch(context, invokeValue);
  }
  public override async handleTeamsTaskModuleSubmit(context, taskModuleRequest):Promise<any> {
    const obj = taskModuleRequest.data;      
    const userName = context.activity.from.name;  
    const mention = {
      type: "mention",
      mentioned: context.activity.from,
      text: `<at>${userName}</at>`,
    };
    const topLevelMessage = MessageFactory.text(`Thank you for updating the work item  ${mention.text}`);
    topLevelMessage.entities = [mention];    


    const templateJson = require('./adaptiveCards/viewProduct.json')
    const template = new AdaptiveCards.Template(templateJson);
    const card = template.expand({
      $root: {
        Product: obj.Product
      }
    });
    const resultCard = CardFactory.adaptiveCard(card);
    await context.sendActivity(topLevelMessage);
    await context.sendActivity({
      type: 'message',
      attachments: [resultCard],
    });
  }
}
