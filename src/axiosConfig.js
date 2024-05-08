const axios = require("axios");
const {authURL,client_id, client_secret,user_name,password, token_endpoint, auth_code } = require("./config");

// Create an instance of axios
const axiosInstance = axios.create();

// Add a request interceptor
axiosInstance.interceptors.request.use(async config => {
  // Before each request, check if the access token is expired
  if (isTokenExpired()) {
    try {
      // If it's expired, refresh it
      const response = await axios.post('/refresh_token', {
        refresh_token: getRefreshToken() // getRefreshToken is a function you define to get the current refresh token
      });

      // Save the new access token
      saveAccessToken(response.data.access_token); // saveAccessToken is a function you define to save the new access token

      // Update the config with the new token
      config.headers.Authorization = `Bearer ${response.data.access_token}`;
    } catch (error) {
      // Handle error, e.g. redirect to login page if refresh token is also expired
      console.error(error);
    }
  }

  return config;
}, error => {
  // Do something with request error
  return Promise.reject(error);
});

// function using axios to perform the authentication & return the access token & refresh token
// function to access access token
async function authenticate() {
  // try {
  //   const response = await axios.post(authURL, {
  //     client_id: client_id,
  //     client_secret: client_secret,
  //     grant_type: 'password',
  //     username: user_name,
  //     password: password,
  //   });
  //   return response.data.access_token;
  // } catch (error) {
  //   console.error(error);
  // }
  try {
    const response = await axios.get(token_endpoint, {
      content_type: 'application/x-www-form-urlencoded',
      data:{
        client_id: client_id,
        client_secret: client_secret,
        grant_type: 'authorization_code',
        code: auth_code,
      }
    });
    return response.data.access_token;
  } catch (error) {
    console.error(error);
  }
}
//function to check if token is expired
function isTokenExpired() {
  // Check if the token is expired
  const expirationTime = localStorage.getItem('expirationTime');
  return Date.now() >= expirationTime;
}
// function to access refresh token
async function getRefreshToken(refreshToken) {
  try {
    const response = await axios.post(authURL, {
      client_id: client_id,
      client_secret: client_secret,
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
    });
    return response.data.access_token;
  } catch (error) {
    console.error(error);
  }
}

// function to save access token
function saveAccessToken(accessToken) {
  // Save the access token
  localStorage.setItem('accessToken', accessToken);
  // Set the expiration time
  const expirationTime = Date.now() + 3600 * 1000; // 1 hour
  localStorage.setItem('expirationTime', expirationTime);
}

// Export the authenticate function
module.exports = {authenticate,axiosInstance};