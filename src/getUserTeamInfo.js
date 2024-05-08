const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const ACLookupCoworker = require("./adaptiveCards/ACLookupCoworker.json");
const helloWorldCard = require("./adaptiveCards/helloWorldCard.json");
const COMMAND_ID = "getUserTeamInfo";
const CommonFunctions = require("./CommonFunctions");
var resultCard = '';
var response2 = '';
const cardHandler = require ("./adaptiveCards/cardHandler");
const utils = require ("./adaptiveCards/utils");
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
// const { runQuery } = require('./CommonFunctions');
  // Message extension Code
  // define function to search incident
  async function handleTeamsMessagingExtensionQuery(context, query,accessToken) {
//assuming that request will always be based on worker name here as no one would know worker id
//call function to search for worker string using readQuery1

let workerId = await CommonFunctions.getWorkerid(context, query.parameters[0].value, accessToken);
console.log("\r\nid: "+ workerId);
if (workerId.attachments){
  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: workerId.attachments,
    },
  };
}else if (workerId.id) {
  console.log("\r\nid: "+ workerId.id);

//query2:https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers/3aa5550b7fe348b98d7b5741afc65534/directReports
let readQuery = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers/'+workerId.id.toString()+'/directReports';
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
        console.log(item);
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
            idVisible: false,
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
    //assuming that request will always be based on worker name here as no one would know worker id
    //call function to search for worker string using readQuery1
    try {
      let response2;
      // Backup the context object
      let backupContext = Object.assign({}, context);
      console.log('\ncontext: ',context.activity.value.action.data);
      let workerId = context.activity.value.action.data.id;
      let userName = context.activity.value.action.data.user;
      let businessTitle = context.activity.value.action.data.businessTitle;
  //query2:https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers/3aa5550b7fe348b98d7b5741afc65534/directReports
      let readQuery = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers/'+workerId.toString()+'/directReports';

      config.method = 'get';
      config.url = readQuery;
      config.headers.Authorization = `Bearer ${accessToken}`;
      response2 = await axios.request(config);
      context = backupContext;
      // console.log('\ncontext2.1: ',context);
      // const attachments = [];
      const teamMembers = [];
      let json = response2.data;
      for (let i = 0; i < json.data.length; i++) {
        let item = json.data[i];
        //populate the fact set with the item.descriptor and item.businessTitle values
        teamMembers.push({ title: item.descriptor, value: item.businessTitle});
      }
      const template = new ACData.Template(ACLookupCoworker);
      resultCard = template.expand({
        $root: {
          user: userName,
          businessDetailsVisible: true,
          businessTitle: businessTitle || '',
          teamMembers: teamMembers
        }
        });
    console.log('\nresultCard: ',resultCard);
    // const preview = CardFactory.heroCard(item.descriptor, item.businessTitle);
    // const attachment = { ...CardFactory.adaptiveCard(resultCard) };
    // attachments.push(attachment);
    const response = utils.CreateAdaptiveCardInvokeResponse(200,resultCard);
    await context.sendActivity({
      type: 'invokeResponse',
      value: response
  });
    // await context.sendActivity('Processing your request...'); 
  }
  catch(error) {
    console.log(error);
  };
}

module.exports ={ COMMAND_ID, handleTeamsMessagingExtensionQuery, onInvokeActivity};
