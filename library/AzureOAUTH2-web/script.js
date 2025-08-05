var $http = require('request');
var assert = require('assert');

getToken(); 

// Authenticate to the tenant using a client ID and secret to capture an authorization token.
// Other scopes, such as https://vault.azure.net/.default, https://manage.office.com/.default for access to O365 APIs or https://graph.microsoft.com/.default for Microsoft Graph could be used

function getToken() {
$http.post('https://login.microsoftonline.com/'+ $secure.AZURE_TENANT +'/oauth2/v2.0/token',
    {
    headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
    },
    form: {
        //'client_id': $secure.NEWRELICSYNTH01UKSNONPRODSERVICEPRINCIPAL,
        //'client_secret': $secure.NEWRELICSYNTH_01_UKS_NONPROD_SERVICEPRINCIPAL_SECRET,

        'client_id': $secure.DEV_SERVICE_PRINCIPLE,
        'client_secret': $secure.DEV_SERVICE_PRINCIPLE_SECRET,
        'grant_type': 'client_credentials',

        // the scope needs to be set as the api:// link in the App Registration
        // The service account needs to be granted Application.Read.All to the App Registration, so the scope can be set as the App Registration API below
        // 'scope': ‘777xxxx……….../.default'
        'scope': $secure.DEV_SERVICE_PRINCIPLE + '/.default'
    }
    }, 
    function(err, resp, body) {
    assert.ok(!err, "Error: Connection error during oauth " + err)
    assert.ok(resp.statusCode == 200, "Error: Error http response code received during oauth " + resp.statusCode);
    const respJson = JSON.parse(resp.body);
    var token = respJson['access_token'];
    getPage(token);
    console.log("INFO: token found : "+ token);
    })
}

function getPage(token) {
$http.get('https://' + "dev" + ‘.myurl.com/healthcheck',
    {
    headers: {
        'Authorization': 'Bearer ' + token
    }
    },
    function(err, resp, body) {
    assert.ok(!err, "Error: Connection error getting site " +err);
    assert.ok(resp.statusCode == 200, "Error: Error http response code received getting site " +resp.statusCode);
    assert.ok(resp.body.toLowerCase().includes('true'), "Error: 'OK', not found in response");
    console.log(resp.body);
    // const respJson = JSON.parse(resp.body);
    })
}
