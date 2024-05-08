const {TurnContext, CardFactory} = require("botbuilder");
// const SNIncidents = require("./SNIncidents.json");
// const successCard = require ("./successCard.json");
// const errorCard = require ("./errorCard.json");
const ACData = require ("adaptivecards-templating");
const ACLookupTeam = require("./ACLookupTeam.json");
// const searchIncident = require("./SearchIncident");
// const { bearerToken, cookie, baseURL, access_token } = require("../config");
// const { CreateInvokeResponse, getInventoryStatus } = require("./utils");
// const CommonFunctions = require("../CommonFunctions");
// const axios = require("axios");
const { runQuery } = require('../CommonFunctions');
function getEditCard(result) {

    var template = new ACData.Template(SNIncidents);
    var card = template.expand({
        $root: {
            number: result.number,
            short_description: result.short_description,
            severity: result.severity,
            link: 'https://ven01957.service-now.com/incident.do?sysparm_query=number='+result.number,
          },
    });
    return CardFactory.adaptiveCard(card);
}
// let readQuery='';
// let config = {
//     method: 'get',
//     maxBodyLength: Infinity,
//     url: readQuery,
//     headers: { 
//       'Authorization': '', 
//       'Cookie': cookie    
//     },
//   };

async function GetTeamInfo(context,accessToken) {
    // try {
        console.log('\r\nInside handleTeamsCardActionGetTeamInfo');
        let workerId = context.activity.value.action.data.id;
        // await context.sendActivity({ attachments: [errorBody] });
        // await context.sendActivity('Looking up team information...');
        let readQuery = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers/'+workerId.toString()+'/directReports';
        const response = await runQuery(readQuery, 'get', '', accessToken);
        if (response.data.total === 0) {
            var responseBody = { statusCode: 200, type: "application/vnd.microsoft.activity.message", value: "Sorry, could not find any Team Informtion" }
            return { status: 200, responseBody }
        } else {
            console.log(JSON.stringify(response.data));
            const template = new ACData.Template(ACLookupTeam);
            const attachments = [];
            let json = response.data;
            const resultCard = template.expand({$root: json});
            const preview = CardFactory.heroCard('Team');
            const attachment = { ...CardFactory.adaptiveCard(resultCard), preview };
            attachments.push(attachment);
            console.log(resultCard);

            var responseBody = { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: resultCard }
            // return CreateInvokeResponse(responseBody);
            // await context.sendActivity({ attachments: [attachment] });
            return { status: 200, responseBody }
            
        }
    // } catch(error) {
    //     console.log(error);
    //     var errorBody = { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: errorCard }
    //     // return CreateInvokeResponse(errorBody);
    //     await context.sendActivity({ attachments: [errorBody] });
    //     return { status: 200, errorBody }
    // }
}

module.exports = { getEditCard, GetTeamInfo };