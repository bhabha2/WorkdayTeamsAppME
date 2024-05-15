const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const ACCoworker = require("./adaptiveCards/ACCoworker.json");
const COMMAND_ID = "getEmployeeDetailsById";
// Message extension Code
// define function to search incident

  async function handleTeamsMessagingExtensionQuery(context, query,accessToken) {
    // Add your code here
          try
          {
      
          let attachments = [];

            const template = new ACData.Template(ACCoworker);
            const resultCard = template.expand({
              $root: {
                link:'https://microsoft.com',
                user:'test user',
                // totalHourlyBalance: item.totalHourlyBalance || 0,
                businessDetailsVisible: true,
                businessTitle: 'test Director',
                primaryWorkEmail:'test email.com',
                primaryWorkPhone:'test90909090',
                primarySupervisoryOrganization:'testOrgn',
                id:'test id',
                // idVisible: false,
                leaveInfoVisible:false
              },
              });

            const preview = CardFactory.heroCard('test user','test director');
            const attachment = { ...CardFactory.adaptiveCard(resultCard), preview };
            attachments.push(attachment);

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
module.exports ={ COMMAND_ID, handleTeamsMessagingExtensionQuery };
