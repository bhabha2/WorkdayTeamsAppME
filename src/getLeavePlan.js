const axios = require("axios");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const {getEditCard} = require("./adaptiveCards/cardHandler");
const COMMAND_ID = "getLeavePlan";
const CommonFunctions = require("./CommonFunctions");
const { runQuery } = require('./CommonFunctions');
  // Message extension Code
  // define function to search incident
async function handleTeamsMessagingExtensionQuery(context, query,accessToken) {
  // Add your code here
  //await authorizeUser.authorizeUser();
  const searchQuery = query.parameters[0].value;
  // query1: get userid from name    
  let workerId = await CommonFunctions.getWorkerid(query.parameters[0].value, accessToken);
  console.log("id: " + workerId);
  if (workerId.attachments) {
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: workerId.attachments,
      },
    };
  } else if (workerId.id) {
    let readQuery = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers/' + workerId.id + '/timeOffSummary';
    //query2: https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?limit=10&view=TimeOffSummary&search={username}

    try {
      // const response = await axios.request(config);
      const response = await runQuery(readQuery, 'get', '', accessToken);
      console.log(JSON.stringify(response.data));
      const attachments = [];
      let json = response.data;
      for (let i = 0; i < json.data.length; i++) {
        let item = json.data[i];
        console.log(item);
        const template = new ACData.Template(ACLookupCoworker);
        const resultCard = template.expand({
          $root: {
            timeOffType: item.timeOffType.descriptor,
            reason: item.reason.descriptor || '',
            quantity: item.quantity,
            status: item.status.descriptor || '',
           },
        });

        const preview = CardFactory.heroCard(item.descriptor, item.businessTitle);
        const attachment = { ...CardFactory.adaptiveCard(resultCard), preview };
        attachments.push(attachment);
      }
      return {
        composeExtension: {
          type: "result",
          attachmentLayout: "list",
          attachments: attachments,
        },
      }
    } catch (error) {
      console.log(error);
    };
  }
}
module.exports ={ COMMAND_ID, handleTeamsMessagingExtensionQuery };
