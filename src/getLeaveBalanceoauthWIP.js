const axios = require("axios");

const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const ACLookupCoworker = require("./adaptiveCards/ACLookupCoworker.json");
const { bearerToken, cookie, baseURL } = require("./config");
const {getEditCard} = require("./adaptiveCards/cardHandler");
const COMMAND_ID = "getLeaveBalance";
const authorizeUser = require("./AuthorizeUser");
const { authenticate, axiosInstance } = require('./axiosConfig');

  // Message extension Code
  // define function to search incident

  async function handleTeamsMessagingExtensionQuery(context, query) {
    // Add your code here
    //await authorizeUser.authorizeUser();
    // Call the authenticate function
authenticate()
  .then(token => {
    if (token) {
      console.log('Authenticated, token:', token);
      const searchQuery = query.parameters[0].value;
      // look for 'incident_no', 'short_description' and 'assigned_to' in query and assign the value to SearchParameter and SearchValue
      let readQuery = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?limit=10&view=TimeOffSummary&search=';
      try {
        searchName = query.parameters.find((element) => element.name === "user")?.value || '';
        readQuery += searchName;
      } catch (error) {
        console.log('No Search value found');
      }
      // console.log(config.url);
      axiosInstance.get(readQuery)
        .then(response => {
          console.log('Response:', response.data);
        })
        .catch(error => {
          console.error('Error:', error);
        });
      //
      // console.log(JSON.stringify(response.data));
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
    } else {
      console.log('Token is empty');
    }
  })
  .catch(error => {
    console.log(error);
  });
}

// module.exports ={ COMMAND_ID, handleTeamsMessagingExtensionQuery };
