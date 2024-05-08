let config = {
  botId: process.env.BOT_ID,
  botPassword: process.env.BOT_PASSWORD,
  baseURL: '', 
  bearerToken: '',
  cookie: '',

  REST_API_Endpoint: '',
  token_endpoint: '',
  authURL: '',
  auth_code: '',
  Authorization_response_type: 'code',
  client_id: '',
  client_secret: '',
  redirectURL: 'https://token.botframework.com/.auth/web/redirect',
  access_token: '',
  token_type: "Bearer",
  refresh_token: 
};

module.exports = config;
