import { AdaptiveCardInvokeResponse, AdaptiveCardInvokeValue, InvokeResponse, MessageFactory, TurnContext } from "botbuilder";
import { CreateActionErrorResponse, CreateAdaptiveCardInvokeResponse } from "../util";
import { AuthService } from "../services/AuthService";
import { GraphService } from "../services/GraphService";
import * as AdaptiveCards from "adaptivecards-templating";
import viewProduct from "../adaptiveCards/viewProduct.json";
import success from "../adaptiveCards/success.json";
export const HandleInvokeActivity = async (context: TurnContext): Promise<any> => {
    const invokeValue = context.activity.value;
    if (invokeValue.action.type !== 'Action.Execute') {
        return CreateActionErrorResponse(
            400,
            0,
            `ActionTypeNotSupported: ${invokeValue.action.type} is not a supported action.`
        );
    }

    const credentials = new AuthService(context);
    const token = await credentials.getUserToken();
    if (!token) {
        // There is no token, so the user has not signed in yet.
        return credentials.getSignInAdaptiveCardInvokeResponse();
    }
    const graphService = new GraphService(token); 
    const categories= await graphService.getretailCategories();
    const verb = invokeValue.action.verb;
    const data:any=invokeValue.action.data;     
    try {
        switch (verb) {
            case 'save':          
                const updatedProduct = {Id:data.productId,Title:data.Title,RetailCategory:data.RetailCategory,ReleaseDate:data.ReleaseDate};
                const product = await graphService.updateProduct(updatedProduct);
                const successTemplate = new AdaptiveCards.Template(success);
                var cardS = successTemplate.expand({
                    $root: {
                        Product: product,
                        message:"Product updated successfully",                       
                    }
                });         
                                    
                return CreateAdaptiveCardInvokeResponse(200,cardS);                    
            case 'cancel':
                return await refreshCard(data.productId);   
            case 'refresh':
                    return await refreshCard(data.productId);           ;
            default:
                return CreateActionErrorResponse(400, 0, `ActionVerbNotSupported: ${verb} is not a supported action verb.`);
        }
    } catch (error) {       
        return CreateActionErrorResponse(500, 0, error.message);
    }

    async function refreshCard(id) {
        const response = await graphService.getProduct(id);
        const viewTemplate = new AdaptiveCards.Template(viewProduct);
        var cardP = viewTemplate.expand({
            $root: {
                Product: response,
                RetailCategories: categories
            }
        });
        return CreateAdaptiveCardInvokeResponse(200, cardP);
    }
};
