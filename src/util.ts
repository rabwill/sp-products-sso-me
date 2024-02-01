import { AdaptiveCardInvokeResponse, InvokeResponse } from "botbuilder";

export const CreateInvokeResponse = (status: number, body?: unknown): InvokeResponse => {
    return { status, body };
};
export const CreateAdaptiveCardInvokeResponse = (statusCode: number, body?: Record<string, unknown>): AdaptiveCardInvokeResponse => {
   return {
            statusCode: statusCode,
            type: 'application/vnd.microsoft.card.adaptive',
            value: body
        };
};
export const CreateActionErrorResponse = ( statusCode: number, errorCode: number = -1,    errorMessage: string = 'Unknown error'): AdaptiveCardInvokeResponse => {
    return {
        statusCode: statusCode,
        type: 'application/vnd.microsoft.error',
        value: {
            error: {
                code: errorCode,
                message: errorMessage,
            },
        },
    };
};

export const CreateInvokeErrorResponse = (statusCode: number, errorCode: number = -1, errorMessage: string = 'Unknown error'): InvokeResponse => {
    return CreateInvokeResponse(statusCode, {
        error: {
            code: errorCode,
            message: errorMessage,
        },
    });
};
export const setTaskInfo = (taskInfo) => {
    taskInfo.height = 350;
    taskInfo.width = 800;
    taskInfo.title = "";
  }
