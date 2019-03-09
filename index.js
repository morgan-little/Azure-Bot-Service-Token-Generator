// Common Node JS needed for the application
const express = require('express');
const request = require('request');
const cors = require('cors');
const helmet = require('helmet');

// Azure specific libraries to support Key Vault, Microsoft Rest and Application Insights
const KeyVault = require('azure-keyvault');
const msRestAzure = require('ms-rest-azure');
const appInsights = require("applicationinsights");

// The Vault name to use so we can change it early. Change the name to the correct Vault Name that will be used in production
const vaultName = 'BotKeyVault-Test'

// Create an Express app to handle everything
const app = express();

// Create the list of URLs that should be allowed into CORS. Change to use the actual servers to be used. Use * to allow any
const whitelist = ['http://www.example1.ca', 'http://example1.ca']

// Initialize Application Insights using environment variables
appInsights.setup(process.env.APPINSIGHTS_INSTRUMENTATIONKEY);
appInsights.start();
// Create client to start tracking more information
let client = appInsights.defaultClient;

// We are using a Key Vault for the secret token so we declare the variable here  
var token;
// Pull the app secret from the process environment settings instead of the Key Vault. Comment the above line
// and uncomment the below line.
// const token = process.env.AppToken;

// Overall CORS options and responses
const corsOptions = {
  origin: function(origin, callback) {
    // Track the origin sources. We want to know what origins are coming in so we know if there are other origins
    // being used then the desired one
    client.trackEvent({name: "Request Origin", properties: {Origin: origin}});
    if (whitelist.indexOf(origin) !== -1) {
      callback(null, true)
    } else {
      callback(new Error('Not allowed by CORS'))
    }
  }
}

// Use helmet to provide some security settings
app.use(helmet());

app.get('/', cors(corsOptions), (req, res) => {
  const options = {
    method: 'POST',
    url: 'https://directline.botframework.com/v3/directline/tokens/generate',
    headers:{
      Authorization: "Bearer " + token
    }
  };
  request(options, (error, response, body) => {
    if(response.statusCode == 200){
      res.send(body);
    }else{
      res.status(500).send({error : "Couldn't get Token"});
    }
  });
});

const port = process.env.PORT || 1337;


// Comment this code if you aren't using Key Vault. Look at https://github.com/Azure-Samples/key-vault-node-quickstart 
// to see how to configure the Key Vault and the App Service
msRestAzure.loginWithAppServiceMSI({resource: 'https://vault.azure.net'}).then( (credentials) => {
    const keyVaultClient = new KeyVault.KeyVaultClient(credentials);

    var vaultUri = "https://" + vaultName + ".vault.azure.net/";
    
    keyVaultClient.getSecret(vaultUri, "AppSecret", "").then(function(response){
        // Save the token
        token = response.value;
        
        // Don't start the app until we know we have the token. Don't want to try request a conversation token
        // without the application token
        app.listen(port, () => console.log('Example app listening on port ' + port + '!'));
    });
});


