const { AdaptiveCardInvokeResponse } = require ('botbuilder');

async function CreateAdaptiveCardInvokeResponse (statusCode, body) {
    return {
             statusCode: statusCode,
             type: 'application/vnd.microsoft.card.adaptive',
             value: body
         };
};

async function CreateActionErrorResponse (statusCode, errorCode, errorMessage) {
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
async function CreateInvokeResponse (body) {
    return { status: 200, body }
};
module.exports = { CreateInvokeResponse, CreateAdaptiveCardInvokeResponse, CreateActionErrorResponse };
