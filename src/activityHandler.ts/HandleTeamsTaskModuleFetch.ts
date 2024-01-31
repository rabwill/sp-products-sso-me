import { AdaptiveCardInvokeValue, TurnContext } from "botbuilder";
import * as AdaptiveCards from "adaptivecards-templating";
import { CardFactory } from "botbuilder";
import { setTaskInfo } from "../util";
import { AuthService } from "../services/AuthService";
import { GraphService } from "../services/GraphService";


export const HandleTeamsTaskModuleFetch = async (context: TurnContext, invokeValue: any): Promise<any> => { 
    const obj = invokeValue.data.data;
    const credentials = new AuthService(context);
    const token = await credentials.getUserToken();
    if (!token) {
        return credentials.getSignInAdaptiveCardInvokeResponse();
    }
    const graphService = new GraphService(token);     
    const categories= await graphService.getretailCategories();
    let taskInfo: any = {};
    const templateJson = require('../adaptiveCards/editProduct.json')
    const template = new AdaptiveCards.Template(templateJson);
    const card = template.expand({
        $root: {
            Product: obj.Product,
            RetailCategories: categories
        }
    });
    const resultCard = CardFactory.adaptiveCard(card);
    taskInfo.card = resultCard;
    setTaskInfo(taskInfo);
    return {
      task: {
        type: 'continue',
        value: taskInfo
      }
    }
}