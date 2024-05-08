const axios = require("axios");

const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const ACLookupCoworker = require("./adaptiveCards/ACLookupCoworker.json");
let {bearerToken, cookie, baseURL } = require("./config");
const {getEditCard} = require("./adaptiveCards/cardHandler");
const COMMAND_ID = "getLeaveBalance";
const authorizeUser = require("./AuthorizeUser");
const { authenticate, axiosInstance } = require('./axiosConfig');
const { runQuery } = require('./CommonFunctions');
const { access } = require("fs");
  // Message extension Code
  // define function to search incident

  async function handleTeamsMessagingExtensionQuery(context, query, accessToken) {
    // Add your code here
    //await authorizeUser.authorizeUser();
console.log('\r\nInside getLeaveBalance');
    const searchQuery = query.parameters[0].value;
    let searchValue='';
    // look for 'incident_no', 'short_description' and 'assigned_to' in query and assign the value to SearchParameter and SearchValue
    let readQuery = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?limit=10&view=TimeOffSummary&search=';
    try {
      searchName = query.parameters.find((element) => element.name === "user")?.value || '';
      readQuery += searchName;
    } catch (error) {
      console.log('No Search value found');
    }
    try {
      const response = await runQuery(readQuery,'get','',accessToken);
      const attachments = [];
      let json = response.data;
      for (let i = 0; i < json.data.length; i++) {
        let item = json.data[i];
        console.log(item);
        const template = new ACData.Template(ACLookupCoworker);
        const resultCard = template.expand({
          $root: {
            link: item.href,
            id: item.id,
            idVisibility: false,
            user: item.descriptor,
            totalHourlyBalance: item.totalHourlyBalance || 0,
            businessDetailsVisible: false,
            idVisible: false,
            leaveInfoVisible: true
          },
        });

        const preview = CardFactory.heroCard(item.descriptor, item.totalHourlyBalance + 'hours');
        const attachment = { ...CardFactory.adaptiveCard(resultCard), preview };
        attachments.push(attachment);
      }
      return {
        composeExtension: {
          type: "result",
          attachmentLayout: "list",
          attachments: attachments,
        },
      };
    } catch (error) {
      console.error('Error processing request. Error: ', error.message);
      // response.send(500, 'Internal Server Error');
      return (response.send(500, 'Internal Server Error'));
    }

}

module.exports ={ COMMAND_ID, handleTeamsMessagingExtensionQuery };
