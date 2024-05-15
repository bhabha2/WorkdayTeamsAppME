const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const ACCoworker = require("./adaptiveCards/ACCoworker.json");
const COMMAND_ID = "getUserInfo";
const { runQuery } = require('./CommonFunctions');
const { CreateInvokeResponse, CreateAdaptiveCardInvokeResponse, CreateActionErrorResponse } = require("./adaptiveCards/utils");
// Message extension Code
// define function to search incident

  async function handleTeamsMessagingExtensionQuery(context, query,accessToken) {
    // Add your code here

    const searchQuery = query.parameters[0].value;
    let searchValue='';
// look for 'incident_no', 'short_description' and 'assigned_to' in query and assign the value to SearchParameter and SearchValue
  let readQuery = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?limit=10&view=workerSummary&search=';
//query2: https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?limit=10&view=TimeOffSummary&search={username}

    try {
      searchName = query.parameters.find((element) => element.name === "user")?.value||'';
        readQuery+=searchName;
    } catch (error) {
      console.log('No Search value found');
    }

       console.log(readQuery);
        // .then((response) => {
          try
          {
          // const response = await axios.request(config);
          const response = await runQuery(readQuery, 'get', '', accessToken);
          // console.log('\r\nJSON data: ',JSON.stringify(response.data));
          let attachments = [];
          let json = response.data;
          for (let i = 0; i < json.data.length; i++) {
            let item = json.data[i];
            // console.log('\r\nitem: ',item);
            const template = new ACData.Template(ACCoworker);
            const resultCard = template.expand({
              $root: {
                link:item.href,
                user: item.descriptor,
                // totalHourlyBalance: item.totalHourlyBalance || 0,
                businessDetailsVisible: true,
                businessTitle: item.businessTitle || 'Not Available',
                primaryWorkEmail:item.primaryWorkEmail || 'Not Available',
                primaryWorkPhone:item.primaryWorkPhone || 'Not Available',
                primarySupervisoryOrganization:item.primarySupervisoryOrganization.descriptor || 'Not Available',
                id:item.id,
                // idVisible: false,
                leaveInfoVisible:false
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
    }
  }
    async function onInvokeActivity(context,accessToken) {
      console.log('\r\nInside getUserTeamInfo onInvoke');
      let resultCard ='';
    // look for 'incident_no', 'short_description' and 'assigned_to' in query and assign the value to SearchParameter and SearchValue
      let readQuery = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?limit=10&view=workerSummary&search=';
    //query2: https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?limit=10&view=TimeOffSummary&search={username}
        try {
          searchName = context.activity.value.action.data.user||'';
            readQuery+=searchName;
        } catch (error) {
          console.log('No Search value found');
        }
        console.log('\r\n',readQuery);
        try
        {
        // const response = await axios.request(config);
        const response = await runQuery(readQuery, 'get', '', accessToken);
        // console.log('\r\nJSON data: ',JSON.stringify(response.data));
        let attachments = [];
        let json = response.data;
        for (let i = 0; i < json.data.length; i++) {
          let item = json.data[i];
          // console.log('\r\nitem: ',item);
          const template = new ACData.Template(ACCoworker);
          resultCard = template.expand({
            $root: {
              link:item.href,
              user: item.descriptor,
              // totalHourlyBalance: item.totalHourlyBalance || 0,
              businessDetailsVisible: true,
              businessTitle: item.businessTitle || 'Not Available',
              primaryWorkEmail:item.primaryWorkEmail || 'Not Available',
              primaryWorkPhone:item.primaryWorkPhone || 'Not Available',
              primarySupervisoryOrganization:item.primarySupervisoryOrganization.descriptor || 'Not Available',
              id:item.id,
              // idVisible: false,
              leaveInfoVisible:false
            },
            });
          const attachment = { ...CardFactory.adaptiveCard(resultCard) };
          attachments.push(attachment);
        }
        var responseBody = { statusCode: 200, type: 'application/vnd.microsoft.card.adaptive', value: resultCard }
        return CreateInvokeResponse(responseBody);
      }
      catch(error) {
        console.log(error);
      };
}
module.exports ={ COMMAND_ID, handleTeamsMessagingExtensionQuery, onInvokeActivity };