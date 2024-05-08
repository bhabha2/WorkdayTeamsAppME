const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const helloWorldCard = require("./adaptiveCards/helloWorldCard.json");
// const SNIncidents = require("./adaptiveCards/SNIncidents.json");
const { bearerToken, cookie, baseURL } = require("./config");
const COMMAND_ID = "updateIncidentDetails";
// class UpdateIncident extends TeamsActivityHandler {
//   constructor() {
//     super();
//   }

  // Message extension Code
  // Search.
  async function handleTeamsMessagingExtensionQuery(context, query) {
    const searchQuery = query.parameters[0].value;
    let incidentParam, short_descriptionParam,assigned_toParam, searchParameter, searchValue='';

    try {
      incidentParam = query.parameters.find(p => p.name === 'incident_no').value;
      if (incidentParam) {
        searchParameter='number'
        searchValue=incidentParam 
      }
      short_descriptionParam = query.parameters.find(p => p.name === 'short_description');
      if (short_descriptionParam) {
        searchParameter='short_description'
        searchValue=short_descriptionParam
      }
      assigned_toParam = query.parameters.find(p => p.name === 'assigned_to').value;
      if (assigned_toParam) {
        searchParameter='assigned_to'
        searchValue=assigned_toParam
      }
      console.log(query.commandId);
    } catch (error) {
      console.log('value not found for few variables');
    }
    if (incidentParam && incidentParam.length === 10){

    //read incident number from the query & update the incident
    //first perform get operation to fetch sys_id of incident
    //then perform put operation to update the incident with user input
    let readConfig = {
      method: 'get',
      maxBodyLength: Infinity,
      url: baseURL+'?sysparm_limit=4&sysparm_query='+searchParameter+'LIKE' + searchValue,
      // url: 'https://ven01957.service-now.com/api/now/table/incident?sysparm_limit=4&sysparm_query=assigned_to.nameLIKEAlex',
      headers: { 
        'Authorization': bearerToken, 
        'Cookie': cookie
        },
    };


    try
    {
    const readResponse = await axios.request(readConfig);
    const sys_id = readResponse.data[0].sys_id;
    const short_description = readResponse.data.result[0].short_description;
    // console.log(JSON.stringify(response.data));
  
    let data = JSON.stringify({
      "short_description": "test"+short_description
    });
    // console.log(JSON.stringify(response.data));
    let updateConfig = {
      method: 'put',
      maxBodyLength: Infinity,
      url: baseURL+"/"+sys_id+'?sysparm_exclude_ref_link=true',
      // https://instance.service-now.com/api/now/v1/table/incident/{sys_id}?sysparm_exclude_ref_link=true
      headers: { 
        'Content-Type': 'application/json',
        'Authorization': bearerToken, 
        'Cookie': cookie
        },
      data : data
    };  
    const response = await axios.request(updateConfig);
    console.log(response)
    }
    catch(error) {
      console.log(error);
    };
  }
  }

 async function updateIncident(Incident){
  console.log(Incident);
  let jsonObject={};
  try
  {
  // const readResponse = await axios.request(readConfig);
  // const sys_id = readResponse.data.result[0].sys_id;
  if (Incident.description) {
    jsonObject["short_description"] = Incident.description; 
    }
  if (Incident.priority) {
    jsonObject["priority"] = Incident.priority; 
    }
  // console.log(JSON.stringify(response.data));
  let data = JSON.stringify(jsonObject);
  console.log(JSON.stringify(jsonObject));
  let updateConfig = {
    method: 'put',
    maxBodyLength: Infinity,
    url: baseURL+"/"+Incident.sys_id+'?sysparm_exclude_ref_link=true',
    // https://instance.service-now.com/api/now/v1/table/incident/{sys_id}?sysparm_exclude_ref_link=true
    headers: { 
      'Content-Type': 'application/json',
      'Authorization': bearerToken, 
      'Cookie': cookie
      },
    data : data
  };  
  const response = await axios.request(updateConfig);
  console.log(response)
  }
  catch(error) {
    console.log(error);
  };
 }

module.exports = { COMMAND_ID, handleTeamsMessagingExtensionQuery, updateIncident };
