const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const ACTeamInfo = require("./adaptiveCards/ACTeamInfo.json");
const ACLookupCoworker = require("./adaptiveCards/ACLookupCoworker.json");
const COMMAND_ID = "getUserTeamInfo";
const CommonFunctions = require("./CommonFunctions");
var resultCard = '';
const {baseURL } = require("./config");
const { CreateInvokeResponse, CreateAdaptiveCardInvokeResponse, CreateActionErrorResponse } = require("./adaptiveCards/utils");

const axios = require("axios");
let config = {
  method: 'get',
  maxBodyLength: Infinity,
  url: '',
  headers: { 
    'Authorization': '', 
    'Cookie': ''    
  },
};
  // Message extension Code
  // define function to search incident
  async function handleTeamsMessagingExtensionQuery(context, query,accessToken) {
//assuming that request will always be based on worker name here as no one would know worker id
//call function to search for worker string using readQuery1

let workerId = await CommonFunctions.getWorkerid(context, query.parameters[0].value, accessToken);
if (workerId.attachments){
  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: workerId.attachments,
    },
  };
}else if (workerId.id) {

let readQuery = baseURL +'/'+ workerId.id.toString()+'/directReports';
      try
      {

      config.method = 'get';
      config.url = readQuery;
      config.headers.Authorization = `Bearer ${accessToken}`;
      const response2 = await axios.request(config);
      const attachments = [];
      let json = response2.data;
      for (let i = 0; i < json.data.length; i++) {
        let item = json.data[i];
        const template = new ACData.Template(ACLookupCoworker);
        resultCard = template.expand({
          $root: {
            link:item.href,
            idVisibility: false,
            id:item.id,
            user: item.descriptor,
            businessDetailsVisible: true,
            businessTitle: item.businessTitle || '',
            primaryWorkEmail:item.primaryWorkEmail || '',
            primaryWorkPhone:item.primaryWorkPhone || '',
            primarySupervisoryOrganization:item.primarySupervisoryOrganization.descriptor || '',
            // idVisible: false,
            leaveInfoVisible:false,
            totalHourlyBalance: item.totalHourlyBalance || 0,
            teamMembers:''
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
    };

  // })
  }
  catch(error) {
    console.log(error);
  };
}
  }

  async function onInvokeActivity(context,accessToken) {
    console.log('\r\nInside getUserTeamInfo onInvoke');
    try {
      let response2;
      // Backup the context object
      let backupContext = Object.assign({}, context);
      //capture primary user's information passed in the adaptive card to be sent back along with team member information
      let workerId = context.activity.value.action.data.id;
      let userName = context.activity.value.action.data.user;
      let businessTitle = context.activity.value.action.data.businessTitle;
      let primaryWorkEmail = context.activity.value.action.data.primaryWorkEmail;
      let primaryWorkPhone = context.activity.value.action.data.primaryWorkPhone;
      let primarySupervisoryOrganization = context.activity.value.action.data.primarySupervisoryOrganization;
      
      //run query to gather reportees
      let readQuery = baseURL+'/'+workerId.toString()+'/directReports';
      config.url = readQuery;
      config.headers.Authorization = `Bearer ${accessToken}`;
      response2 = await axios.request(config);
      // context = backupContext;
      // const attachments = [];
      const teamMembers = [];
      let json = response2.data;
      for (let i = 0; i < json.data.length; i++) {
        let item = json.data[i];
        //populate the fact set with the item.descriptor and item.businessTitle values
        teamMembers.push({ title: item.descriptor, value: item.businessTitle});
      }
      const template = new ACData.Template(ACTeamInfo);
      const resultCard = template.expand({
        $root: {
          id:workerId,
          user: userName,
          businessTitle: businessTitle,
          primaryWorkEmail:primaryWorkEmail,
          primaryWorkPhone:primaryWorkPhone,
          primarySupervisoryOrganization:primarySupervisoryOrganization,
          teamMembers: teamMembers
        },
        });
    var responseBody = { statusCode: 200, type: 'application/vnd.microsoft.card.adaptive', value: resultCard }
    // console.log('\r\rresponseBody: ', responseBody)
    return CreateInvokeResponse(responseBody);
    // const response = utils.CreateAdaptiveCardInvokeResponse(200,resultCard);
    // await context.sendActivity({
    //   type: 'invokeResponse',
    //   value: response
  // });
    // await context.sendActivity('Processing your request...'); 
  }
  catch(error) {
    console.log(error);
    // return CreateActionErrorResponse(400,0, "Invalid request");
  }
}

module.exports ={ COMMAND_ID, handleTeamsMessagingExtensionQuery, onInvokeActivity};
