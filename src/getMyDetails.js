const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const ACLookupCoworker = require("./adaptiveCards/ACLookupCoworker.json");
const { bearerToken, cookie, baseURL } = require("./config");
const {getEditCard} = require("./adaptiveCards/cardHandler");
const COMMAND_ID = "getMyDetails";
const authorizeUser = require("./AuthorizeUser");
  // Message extension Code
  // define function to search incident

  async function handleTeamsMessagingExtensionQuery(context, query) {
    // Add your code here
    //await authorizeUser.authorizeUser();

    const searchQuery = query.parameters[0].value;
    let searchValue='';
    // look for 'incident_no', 'short_description' and 'assigned_to' in query and assign the value to SearchParameter and SearchValue
    
  let readQuery = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?query=';


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
            console.log(readQuery);
            console.log(config.url);
        // axios.request(config)
        // .then((response) => {
        try
        {
        const response = await axios.request(config);
        console.log(JSON.stringify(response.data));

      const attachments = [];
      response.data.forEach(() => {
        
        const template = new ACData.Template(ACLookupCoworker);
        const resultCard = template.expand({
          $root: {
            user: descriptor,
            businessTitle: businessTitle,
            primaryWorkEmail:primaryWorkEmail,
            primaryWorkPhone:primaryWorkPhone,
            primarySupervisoryOrganization:primarySupervisoryOrganization,
            id:id,
            primaryWorkAddressText:primaryWorkAddressText
          },
          });
        
          const preview = CardFactory.heroCard(descriptor, businessTitle);
          //
          const attachment = { ...CardFactory.adaptiveCard(resultCard), preview };
          // const attachment = { preview };
          //
          // const resultCard = getEditCard(result);
          // const attachment = { ...resultCard, preview };
          // console.log(preview)
          attachments.push(attachment);
      });

      return {
        composeExtension: {
          type: "result",
          attachmentLayout: "list",
          attachments: attachments,
        },
      };

    // })
    }
    catch(error) {
      console.log(error);
    };
}
module.exports ={ COMMAND_ID, handleTeamsMessagingExtensionQuery };
