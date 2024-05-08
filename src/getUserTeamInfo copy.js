const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const ACLookupCoworker = require("./adaptiveCards/ACLookupCoworker.json");
const { bearerToken, cookie, baseURL } = require("./config");
const {getEditCard} = require("./adaptiveCards/cardHandler");
const COMMAND_ID = "getUserTeamInfo";
const authorizeUser = require("./AuthorizeUser");
  // Message extension Code
  // define function to search incident

  async function handleTeamsMessagingExtensionQuery(context, query) {
  // Add your code here
  //await authorizeUser.authorizeUser();

  const id = query.parameters[0].value;
  let searchValue='';
  // fetch username from query

  // query1: get userid from name    
// let readQuery1 = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?limit=10&view=TimeOffSummary&search=';

//query2:https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers/3aa5550b7fe348b98d7b5741afc65534/directReports
let readQuery2 = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers/'+id+'/directReports';


//query2: https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?limit=10&view=TimeOffSummary&search={username}

  // try {
  //   searchName = query.parameters.find((element) => element.name === "user")?.value||'';
  //     readQuery1+=searchName;
  // } catch (error) {
  //   console.log('No Search value found');
  // }

  let config = {
            method: 'get',
            maxBodyLength: Infinity,
            url: readQuery2,
            headers: { 
              'Authorization': bearerToken, 
              'Cookie': cookie    
            },
          };
          // console.log(readQuery1);
          console.log(config.url);
      // axios.request(config)
      // .then((response) =>

      try
      {
      // const response = await axios.request(config);
      // console.log(JSON.stringify(response.data));

      // searchValue = response.data[0].id;
      // console.log(searchValue);
      // config.url = readQuery2;
      const response2 = await axios.request(config);
      console.log(JSON.stringify(response2.data));

      const attachments = [];
      let json = response2.data;
      for (let i = 0; i < json.data.length; i++) {
        let item = json.data[i];
        console.log(item);
        console.log(item);
        const template = new ACData.Template(ACLookupCoworker);
        const resultCard = template.expand({
          $root: {
            link:item.href,
            idVisibility: false,
            id:item.id,
            user: item.descriptor,
            businessDetailsVisible: true,
            businessTitle: item.businessTitle || 'Not Available',
            primaryWorkEmail:item.primaryWorkEmail || 'Not Available',
            primaryWorkPhone:item.primaryWorkPhone || 'Not Available',
            primarySupervisoryOrganization:item.primarySupervisoryOrganization.descriptor || 'Not Available',
            idVisible: false,
            leaveInfoVisible:false,
            totalHourlyBalance: item.totalHourlyBalance || 0
          },
          });
        // const resultCard = template.expand({
        //   $root: {
        //     user: item.descriptor,
        //     totalHourlyBalance: item.totalHourlyBalance,
        //     id:item.id
        //   },
        //   });
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
  }
  catch(error) {
    console.log(error);
  };
}

module.exports ={ COMMAND_ID1, handleTeamsMessagingExtensionQuery1 };
