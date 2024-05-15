// const axios = require("axios");
// const querystring = require("querystring");
const getLeavePlan = require("./getLeavePlan");
const getLeaveBalance = require("./getLeaveBalance");
const getUserInfo = require("./getUserInfo");
const getMyDetails = require("./getMyDetails");
const getUserTeamInfo = require("./getUserTeamInfo");
const getEmployeeDetailsById = require("./getEmployeeDetailsById");
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
  TeamsActivityHandler, CardFactory, ActionTypes, TurnContext,
  ActivityHandler, UserState, ConversationState} = require('botbuilder');
const { LRUCache } = require ('lru-cache');

const {
  ExtendedUserTokenProvider
} = require('botbuilder-core')

const querystring = require('querystring');

// User Configuration property name
const USER_CONFIGURATION = 'userConfigurationProperty';
const cacheOptions = {
  max: 500,
  // for use with tracking overall storage size
  maxSize: 5000,
  sizeCalculation: (value, key) => {
    return 1
  },
  // for use when you need to clean up something when objects are evicted from the cache
  dispose: (value, key) => {
  },

  // how long to live in ms
  ttl: 1000 * 60 * 5,

  // return stale items before removing from cache?
  allowStale: false,

  updateAgeOnGet: false,
  updateAgeOnHas: false,

  // async method to use for cache.fetch(), for
  // stale-while-revalidate type of behavior
  fetchMethod: async (
    key,
    staleValue,
    { options, signal, context }
  ) => { },
}
var cache = new LRUCache(cacheOptions);
//  let cacheStorage = new CacheStorage(cache);
const cacheInitFlag = "Init";
const cacheRevokeFlag = "Revoke";
const { access } = require('fs');


class SearchApp extends TeamsActivityHandler {
  /**
   *
   * @param {UserState} User state to persist configuration settings
   */
  cacheOptions = {
      max: 500,
      // for use with tracking overall storage size
      maxSize: 5000,
      sizeCalculation: (value, key) => {
        return 1
      },
      // for use when you need to clean up something when objects are evicted from the cache
      dispose: (value, key) => {
      },
  
      // how long to live in ms
      ttl: 1000 * 60 * 5,
  
      // return stale items before removing from cache?
      allowStale: false,
  
      updateAgeOnGet: false,
      updateAgeOnHas: false,
  
      // async method to use for cache.fetch(), for
      // stale-while-revalidate type of behavior
      fetchMethod: async (
        key,
        staleValue,
        { options, signal, context }
      ) => { },
  }

  constructor(userState) {
      super();
      // Creates a new user property accessor.
      // See https://aka.ms/about-bot-state-accessors to learn more about the bot state and state accessors.
      this.userConfigurationProperty = userState.createProperty(
          USER_CONFIGURATION
      );
      this.connectionName = process.env.connectionName;
      this.userState = userState;
      this.userProfielAccessor = userState.createProperty(this.UserProfileProperty);
      // this.conversationState = ConversationState;
      this.conversationState = userState;
      this.conversationDataAccessor = this.conversationState.createProperty(this.ConversationDataProperty);

  }

  /**
   * Override the ActivityHandler.run() method to save state changes after the bot logic completes.
   */
  async run(context) {
      await super.run(context);

      // Save state changes
      await this.userState.saveChanges(context);
  }

  // Overloaded function. Receives invoke activities with Activity name of 'composeExtension/queryLink'
  async handleTeamsAppBasedLinkQuery(context, query) {
      const userTokenClient = context.turnState.get(context.adapter.UserTokenClientKey);
          const magicCode =
              context.state && Number.isInteger(Number(context.state))
                  ? context.state
                  : '';

          const tokenResponse = await userTokenClient.getUserToken(
              context.activity.from.id,
              this.connectionName,
              context.activity.channelId,
              magicCode
          );

      if (!tokenResponse || !tokenResponse.token) {
          // There is no token, so the user has not signed in yet.
          // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
          const { signInLink } = await userTokenClient.getSignInResource(
              this.connectionName,
              context.activity
          );

          return {
              composeExtension: {
                  type: 'auth',
                  suggestedActions: {
                      actions: [
                          {
                              type: 'openUrl',
                              value: signInLink,
                              title: 'Bot Service OAuth'
                          },
                      ],
                  },
              },
          };
      }

  }
  async handleTeamsMessagingExtensionConfigurationQuerySettingUrl(
      context,
      query
  ) {
      // The user has requested the Messaging Extension Configuration page settings url.
      const userSettings = await this.userConfigurationProperty.get(
          context,
          ''
      );
      const escapedSettings = userSettings
          ? querystring.escape(userSettings)
          : '';

      return {
          composeExtension: {
              type: 'config',
              suggestedActions: {
                  actions: [
                      {
                          type: ActionTypes.OpenUrl,
                          value: `${process.env.SiteUrl}/public/searchSettings.html?settings=${escapedSettings}`
                      },
                  ],
              },
          },
      };
  }

  // Overloaded function. Receives invoke activities with the name 'composeExtension/setting
  async handleTeamsMessagingExtensionConfigurationSetting(context, settings) {
      // When the user submits the settings page, this event is fired.
      if (settings.state != null) {
          await this.userConfigurationProperty.set(context, settings.state);
      }
  }

  // Overloaded function. Receives invoke activities with the name 'composeExtension/query'.
  async handleTeamsMessagingExtensionQuery(context, query) {
    console.log("\r\nInside handleTeamsMessagingExtensionQuery");
  const userTokeninCache = cache.get(context.activity.from.id);

  const cloudAdapter = context.adapter;

  const userTokenClient = context.turnState.get(cloudAdapter.UserTokenClientKey);

  const magicCode =
    query.state && Number.isInteger(Number(query.state))
      ? query.state
      : '';

  const tokenResponse = await userTokenClient.getUserToken(
    context.activity.from.id,
    this.connectionName,
    context.activity.channelId,
    magicCode
  );

  const { signInLink } = await userTokenClient.getSignInResource(
    this.connectionName,
    context.activity
  );

  //token is not in cache means user has not signed in yet
  if (!userTokeninCache) {

    cache.set(context.activity.from.id, cacheInitFlag);

    return {
      composeExtension: {
        type: 'auth',
        suggestedActions: {
          actions: [
            {
              type: 'openUrl',
              value: signInLink,
              title: 'Bot Service OAuth'
            },
          ],
        },
      },
    };
  }
  //if token in cache, always update the token based on system stored user token
  else if (tokenResponse && tokenResponse.token) {

    if (userTokeninCache.toString().startsWith(cacheRevokeFlag) && userTokeninCache.toString().endsWith(tokenResponse.token)) {
      console.log("\r\nToken is revoked, need to sign in again");
      return {
        composeExtension: {
          type: 'auth',
          suggestedActions: {
            actions: [
              {
                type: 'openUrl',
                value: signInLink,
                title: 'Bot Service OAuth'
              },
            ],
          },
        },
      };
    }
    else {
      cache.set(context.activity.from.id, tokenResponse.token);
      console.log("\r\nCache Status updated in Query: " );
    }
  }
  else if (!tokenResponse || !tokenResponse.token) {
    // There is no system sotred user token, so the user has not signed in yet.
    // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions

    cache.set(context.activity.from.id, cacheInitFlag);

    return {
      composeExtension: {
        type: 'auth',
        suggestedActions: {
          actions: [
            {
              type: 'openUrl',
              value: signInLink,
              title: 'Bot Service OAuth'
            },
          ],
        },
      },
    };
  }
  console.log("\r\nInside handleTeamsMessagingExtensionQuery: ", query.commandId );
  //authorize user
  // return searchIncident.handleTeamsMessagingExtensionQuery(context, query);
  switch (query.commandId) {
    //call the relevant function to handle the query
    case getLeaveBalance.COMMAND_ID:{
      return getLeaveBalance.handleTeamsMessagingExtensionQuery(context, query,tokenResponse.token);
    }
    case getUserInfo.COMMAND_ID:{
      return getUserInfo.handleTeamsMessagingExtensionQuery(context, query,tokenResponse.token);
    }
    case getMyDetails.COMMAND_ID:{
     return getUserInfo.handleTeamsMessagingExtensionQuery(context, query,tokenResponse.token);
    }
    case getUserTeamInfo.COMMAND_ID:{
      return getUserTeamInfo.handleTeamsMessagingExtensionQuery(context, query,tokenResponse.token);
    }
    case getLeavePlan.COMMAND_ID:{
      return getLeavePlan.handleTeamsMessagingExtensionQuery(context, query,tokenResponse.token);
    }
    case 'getEmployeeDetailsById':{
      return getEmployeeDetailsById.handleTeamsMessagingExtensionQuery(context, query,tokenResponse.token);
    }
    default:
      throw new Error("NotImplemented");
  }
}

  // Overloaded function. Receives invoke activities with the name 'composeExtension/selectItem'.
  async handleTeamsMessagingExtensionSelectItem(context, obj) {
      return {
          composeExtension: {
              type: 'result',
              attachmentLayout: 'list',
              attachments: [CardFactory.thumbnailCard(obj.description)]
          },
      };
  }

  // Overloaded function. Receives invoke activities with the name 'composeExtension/fetchTask'
  async handleTeamsMessagingExtensionFetchTask(context, action) {
      if (action.commandId === 'SHOWPROFILE') {
          const userTokenClient = context.turnState.get(context.adapter.UserTokenClientKey);
          const magicCode =
              context.state && Number.isInteger(Number(context.state))
                  ? context.state
                  : '';

          const tokenResponse = await userTokenClient.getUserToken(
              context.activity.from.id,
              this.connectionName,
              context.activity.channelId,
              magicCode
          );

          if (!tokenResponse || !tokenResponse.token) {
              // There is no token, so the user has not signed in yet.
              // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
              const { signInLink } = await userTokenClient.getSignInResource(
                  this.connectionName,
                  context.activity
              );

              return {
                  composeExtension: {
                      type: 'auth',
                      suggestedActions: {
                          actions: [
                              {
                                  type: 'openUrl',
                                  value: signInLink,
                                  title: 'Bot Service OAuth'
                              },
                          ],
                      },
                  },
              };
          }
      }

      return null;
  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
      // This method is to handle the 'Close' button on the confirmation Task Module after the user signs out.
      return {};
  }

async onInvokeActivity(context) {
  // async onInvokeActivity(context) {
      console.log('\r\nonInvoke, ' + context.activity.name);
      let runEvents = true;
      try {
      const valueObj = context.activity.value;
      if (valueObj.authentication) {
          const authObj = valueObj.authentication;
          if (authObj.token) {
              // If the token is NOT exchangeable, then do NOT deduplicate requests.
               if (await this.tokenIsExchangeable(context)) 
               {
                   return await super.onInvokeActivity(context);
               }
               else {
                      const response = 
                      {
                      status: 412
                      };
                  return response;
               }
          }
      }
      let runEvents = true;
      console.log(context.activity.name);
      const userTokeninCache = cache.get(context.activity.from.id);
      // console.log("\r\nuserTokeninCache: " + userTokeninCache);
        if(context.activity.name==='adaptiveCard/action'){
          switch (context.activity.value.action.verb) {
            case 'getTeamInfo': case 'TeamInfoRefresh': {
              console.log('\r\ngetTeamInfo');
             return getUserTeamInfo.onInvokeActivity(context, userTokeninCache);
           
            }
            // }
            case 'refresh': case 'individualRefresh': {
              console.log('\r\nrefresh');
              // console.log('\r\nvalue.action.data: ' + JSON.stringify(context.activity.value.action.data));
             return getUserInfo.onInvokeActivity(context, userTokeninCache);
            }
            default:
              runEvents = false;
              return super.onInvokeActivity(context);
          }
        } else {
            runEvents = false;
            return super.onInvokeActivity(context);
        }
        
      } catch (err) {
        console.error(err);
        if (err.message === 'NotImplemented') {
          return { status: 501 };
        } else if (err.message === 'BadRequest') {
          return { status: 400 };
        }
        throw err;
      }finally {
        if (runEvents) {
          this.defaultNextEvent(context)();
          // return { status: 200 };
        }
      }
}

  async tokenIsExchangeable(context) {
      let tokenExchangeResponse = null;
      try {
          const userId = context.activity.from.id;
          const valueObj = context.activity.value;
          const tokenExchangeRequest = valueObj.authentication;
          // console.log("tokenExchangeRequest.token: " + tokenExchangeRequest.token);

          const userTokenClient = context.turnState.get(context.adapter.UserTokenClientKey);

          tokenExchangeResponse = await userTokenClient.exchangeToken(
              userId,
              this.connectionName,
              context.activity.channelId,
              { token: tokenExchangeRequest.token });

          // console.log('tokenExchangeResponse: ' + JSON.stringify(tokenExchangeResponse));
      } 
      catch (err) 
      {
          console.log('tokenExchange error: ' + err);
          // Ignore Exceptions
          // If token exchange failed for any reason, tokenExchangeResponse above stays null , and hence we send back a failure invoke response to the caller.
      }
      if (!tokenExchangeResponse || !tokenExchangeResponse.token) 
      {
          return false;
      }

      console.log('Exchanged token: ');
      return true;
  }
}

module.exports.SearchApp = SearchApp;
