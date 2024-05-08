const axios = require("axios");
// const querystring = require("querystring");
let { bearerToken, cookie, client_id, client_secret, refresh_token,token_endpoint, access_token} = require("./config");
const ACData = require("adaptivecards-templating");
const { CardFactory } = require("botbuilder");
const ACLookupCoworker = require("./adaptiveCards/ACLookupCoworker.json");
// const config1 = require('./config.js');
let readQuery='';

let config = {
    method: 'get',
    maxBodyLength: Infinity,
    url: readQuery,
    headers: { 
      'Authorization': '', 
      'Cookie': cookie    
    },
  };
  
  async function getWorkerid(context, name, accessToken){
  
    const workerName = '';
    // query1: get userid from name    
  readQuery = 'https://wd2-impl-services1.workday.com/ccx/api/v1/microsoft_dpt6/workers?limit=10&search='+name;
  config.url = readQuery;
  try
        {
        // const response = await axios.request(config);
        const response = await runQuery(readQuery, 'get', '', accessToken);
        // console.log(JSON.stringify(response.data));
  
        const attachments = [];
        let json = response.data;
        //get count of workers
        //if count is 1, get worker id and use it to get direct reports using readQuery2
        //if count is more than 1, return all workers and ask user to select one
        //if count is 0, return no worker found
        switch (json.data.length) {
          case 0:
            return 'No worker found';
          case 1:
            console.log(json.data[0].id);
            return {
                    id: json.data[0].id,
                // }
            }
          default:
            for (let i = 0; i < json.data.length; i++) {
              let item = json.data[i];
              console.log(item);
            //   console.log(item);
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
                  primarySupervisoryOrganization: 'Not Available',
                  idVisible: false,
                  leaveInfoVisible:false,
                  totalHourlyBalance: item.totalHourlyBalance || 0
                },
                });
              const preview = CardFactory.heroCard(item.descriptor);
              const attachment = { ...CardFactory.adaptiveCard(resultCard), preview };
              attachments.push(attachment);
            }
      return {
        // ACDetails:{
            attachments: attachments,
        // }
        // composeExtension: {
        //   type: "result",
        //   attachmentLayout: "list",
        //   attachments: attachments,
        // },
      }
    }
    } catch (error) {
      console.log(error);
    }
  }

  // async function getOAuthAccessToken(context) {
  //     // The req.query object has the query params that
  //   // were sent to this route. We want the `code` param
  //   const requestToken = req.query.code;
  //   axios({
  //       // make a POST request
  //       method: "post",
  //       // to the SN/WD/SFDC authentication API, with the client ID, client secret
  //       // and request token
  //       url: `${config.Token_Endpoint}?client_id=${config.Client_ID}&client_secret=${config.Client_Secret}&code=${requestToken}`,
  //       // Set the content type header, so that we get the response in JSOn
  //       headers: {
  //         accept: "application/json",
  //       },
  //     }).then((response) => {
  //       // Once we get the response, extract the access token from
  //       // the response body
  //       const accessToken = response.data.access_token;
  //       // redirect the user to the welcome page, along with the access token
  //       return{
  //           accessToken:accessToken
  //       }
  //     });
  //   }


// const qs = require('querystring');

// async function getAccessTokenWithRefreshToken() {
//   // const requestBody = {
//   //   grant_type: 'refresh_token',
//   //   refresh_token: refresh_token,
//   //   client_id: client_id,
//   //   client_secret: client_secret
//   // };


//   let config = {
//     method: 'post',
//     maxBodyLength: Infinity,
//     url: token_endpoint,
//     headers: { 
//       'Content-Type': 'application/x-www-form-urlencoded',
//       // 'Authorization': bearerToken, 
//       // 'Cookie': cookie    
//     },
//     data: {
//       grant_type: 'refresh_token',
//       refresh_token: refresh_token,
//       client_id: client_id,
//       client_secret: client_secret
//     }
//   };
//   try {
//     // const response = await axios.post(token_endpoint, qs.stringify(requestBody), config);
//     const response = await axios.request(config);
//     return response.data.access_token;
//   } catch (error) {
//     console.error('Error fetching access token with refresh token:', error.message);
//     return null;
//   }
// }

async function runQuery( query, method, data,accessToken) {
  // let config = {
  //   method: "get",
  //   maxBodyLength: Infinity,
  //   url: query,
  //   headers: {
  //     Authorization: bearerToken,
  //     Cookie: cookie,
  //   },
  // };

  // console.log(readQuery);
  // console.log(config.url);
  try {
    // config.url = query;
    // const response = await axios.request(config);
    // return response;
    console.log("\nInside runQuery",accessToken);
    config.method = method;
    config.url = query;
    config.headers.Authorization = `Bearer ${accessToken}`;
    if (data) { config.data = data; }
    // console.log(readQuery);
    console.log(config.url);
    const response = await axios.request(config);
    return response;
  } catch (error) {
    console.log(error);
    return response.send(500, "Internal Server Error");
  }
  // } catch (error) {
  //   console.log(error);
    // if (error.response.status === 401) {
    //   console.error("Unauthorized request. Error: ", error.message);
    //   // call function to get new token using refresh token
    //   let access_token = await getAccessTokenWithRefreshToken();
    //   if (access_token) {
    //     console.log("New access token gathered, retrying: ");
    //     // update the authorization header with the new access token
    //     config.headers.Authorization = `Bearer ${access_token}`;
    //     config1.bearerToken = `Bearer ${access_token}`;
    //     // update bearerToken in config file under Config with the new bearer token
    //     require("fs").writeFileSync(
    //       "./config.js",
    //       `module.exports = ${JSON.stringify(config1)}`
    //     );

    //     // retry the request
    //     try {
    //       const response = await axios.request(config);
    //       console.log(JSON.stringify(response.data));
    //       return response;
    //     } catch (error) {
    //       console.error(
    //         "Error processing request after retry. Error: ",
    //         error.message
    //       );
    //       // response.send(500, 'Internal Server Error');
    //       return response.send(500, "Internal Server Error");
    //     }
    //   } else {
    //     console.error("Error refreshing access token. Error: ", error.message);
    //     // response.send(401, 'Unauthorized');
    //     return response.send(401, "Unauthorized");
    //   }
    // } else {
    //   console.error("Error processing request. Error: ", error.message);
    //   // response.send(500, 'Internal Server Error');
      // return response.send(500, "Internal Server Error");
    // }
  }
module.exports ={ getWorkerid, runQuery};