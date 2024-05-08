const axios = require("axios");

const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const ACLookupCoworker = require("./adaptiveCards/ACLookupCoworker.json");
let {bearerToken, cookie, baseURL } = require("./config");
const {getEditCard} = require("./adaptiveCards/cardHandler");
const COMMAND_ID = "getLeaveBalance1";
const authorizeUser = require("./AuthorizeUser");
const { authenticate, axiosInstance } = require('./axiosConfig');
const { getAccessTokenWithRefreshToken } = require('./CommonFunctions');
  // Message extension Code
  // define function to search incident

  async function handleTeamsMessagingExtensionQuery(context, query) {
    // Add your code here
    //await authorizeUser.authorizeUser();

    const searchQuery = query.parameters[0].value;
    let searchValue='';
    // look for 'incident_no', 'short_description' and 'assigned_to' in query and assign the value to SearchParameter and SearchValue
    let readQuery = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?limit=10&view=TimeOffSummary&search=';

    
//query2: https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?limit=10&view=TimeOffSummary&search={username}

    try {
      searchName = query.parameters.find((element) => element.name === "user")?.value||'';
        readQuery+=searchName;
    } catch (error) {
      console.log('No Search value found');
    }

    let config = {
              method: 'get',
              maxBodyLength: Infinity,
              url: readQuery,
              headers: { 
                'Authorization': bearerToken, 
                'Cookie': cookie    
              },
            };
            // console.log(readQuery);
            console.log(config.url);
        // axios.request(config)
        // .then((response) => {
        try {
          const response = await axios.request(config);
        } catch (error) {
          console.log(error);
          if (error.response.status === 401) {
            console.error('Unauthorized request. Error: ', error.message);
            // call function to get new token using refresh token
            access_token = await getAccessTokenWithRefreshToken();
            if (access_token) {
              console.log('New access token gathered, please retry: ', access_token);
              // update the authorization header with the new access token
              config.headers.Authorization = `Bearer ${access_token}`;
                bearerToken = `Bearer ${config.access_token}`;
                // update bearerToken in config file under Config with the new bearer token
                require('fs').writeFileSync('./config.js', `module.exports = ${JSON.stringify(config)}`);
              // retry the request
              try {
                const response = await axios.request(config);
                console.log(JSON.stringify(response.data));
                // rest of the code...
              
        // console.log(JSON.stringify(response.data));
        const attachments = [];
        let json = response.data;
        for (let i = 0; i < json.data.length; i++) {
          let item = json.data[i];
          console.log(item);
          const template = new ACData.Template(ACLookupCoworker);
          const resultCard = template.expand({
            $root: {
              link:item.href,
              id:item.id,
              idVisibility: false,
              user: item.descriptor,
              totalHourlyBalance: item.totalHourlyBalance || 0,
              businessDetailsVisible: false,
              idVisible: false,
              leaveInfoVisible:true
            },
            });

          const preview = CardFactory.heroCard(item.descriptor, item.totalHourlyBalance+'hours');
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

    // })
  } catch (error) {
    console.error('Error processing request after retry. Error: ', error.message);
    // response.send(500, 'Internal Server Error');
    return (response.send(500, 'Internal Server Error'));
  }
} else {
  console.error('Error refreshing access token. Error: ', error.message);
  // response.send(401, 'Unauthorized');
  return (response.send(401, 'Unauthorized'));
}
} else {
console.error('Error processing request. Error: ', error.message);
// response.send(500, 'Internal Server Error');
return (response.send(500, 'Internal Server Error'));
}
}
}

module.exports ={ COMMAND_ID, handleTeamsMessagingExtensionQuery };
