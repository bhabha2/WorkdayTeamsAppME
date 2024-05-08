const WDconfig = require("./config");
const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const { bearerToken, cookie, baseURL } = require("./config");

async function authorizeUser() {
    //Authorization Code Grant to workday environment
    // 1. App calls the authorize endpoint to request an auth code.
    //call Authorization_Endpoint with response_type=code & client_id
    //ex. GET [hostUrl/tenant]}/authorize?response_type=code&client_id=[API Client ID]
    //ex. https://impl.workday.com/microsoft_dpt6/authorize?response_type=code&client_id=44095b28b064ca502e676bbbe1cbea04
    // const authorizeurl = config.Authorization_Endpoint + "?response_type=" + config.Authorization_response_type + "&client_id=" + config.client_id + "&redirect_url=" +config.redirectURL+",{redirect:'manual'}";
    // axios.get(authorizeurl)
    //     .then(response => {
    //         // Handle the response here
    //         //show workday login screen

    //         console.log(response.data);
    //     })
    //     .catch(error => {
    //         // Handle the error here
    //         console.error(error);
    //     });

    // 2. Workday prompts the user to login.

    //3. Workday authentication server redirects the browser to the redirect URL registered with the API client, and sends the authorization code in the URL.
    //ex. https://example.com/#code=88xr1xk2sn


    // 4. App calls the token endpoint to exchange auth code for access token.
    //ex. POST [hostUrl/tenant]/token
    // Specify these fields in the request body:
    // client_id - The registered API client ID.
    // client_secret - The registered API Client secret.
    // grant_type - authorization_code
    // code - The returned code value from the authorize endpoint.

    const tokenUrl = 'https://impl.workday.com/microsoft_dpt6/token';
    const requestBody = {
        client_id: WDconfig.Client_ID,
        client_secret: WDconfig.Client_Secret,
        grant_type: 'authorization_code',
        code: WDconfig.auth_code
    };


    axios.post(tokenUrl, querystring.stringify(requestBody))
        .then(response => {
            // Step 5: Auth server returns the access token in response.
            const accessToken = response.data.access_token;
            const refreshToken = response.data.refresh_token;
            const tokenType = response.data.token_type;

            console.log('Access Token:', accessToken);
            console.log('Refresh Token:', refreshToken);
            console.log('Token Type:', tokenType);
        })
        .catch(error => {
            console.error('Error:', error);
        });

    // ex. { "access_token": "7c3obrknwd6nnkxv0r64jdpbx",  
    //   "refresh_token": "yxsiqvdkakj0tp9a4i2xe1fbg4blgrq1noqz2fg",  
    //   "token_type": "Bearer" }
}

module.exports = {authorizeUser};
