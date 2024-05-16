require('isomorphic-fetch');
const azure = require('@azure/identity');
const graph = require('@microsoft/microsoft-graph-client');
const authProviders =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

let _settings = undefined;
let _deviceCodeCredential = undefined;
let _userClient = undefined;

function initializeGraphForUserAuth(settings, deviceCodePrompt) {
  // Ensure settings isn't null
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }

  _settings = settings;

  _deviceCodeCredential = new azure.DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.tenantId,
    userPromptCallback: deviceCodePrompt
  });

  const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
    _deviceCodeCredential, {
      scopes: settings.graphUserScopes
    });

  _userClient = graph.Client.initWithMiddleware({
    authProvider: authProvider
  });
}
module.exports.initializeGraphForUserAuth = initializeGraphForUserAuth;

async function getUserTokenAsync() {
    // Ensure credential isn't undefined
    if (!_deviceCodeCredential) {
      throw new Error('Graph has not been initialized for user auth');
    }

    // Ensure scopes isn't undefined
    if (!_settings?.graphUserScopes) {
      throw new Error('Setting "scopes" cannot be undefined');
    }
  
    // Request token with given scopes
    const response = await _deviceCodeCredential.getToken(_settings?.graphUserScopes);
    return response.token;
  }
  module.exports.getUserTokenAsync = getUserTokenAsync;

  async function getUserAsync() {
    // Ensure client isn't undefined
    if (!_userClient) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    return _userClient.api('/me')
      // Only request specific properties
      .select(['displayName', 'mail', 'userPrincipalName'])
      .get();
  }
  module.exports.getUserAsync = getUserAsync;

  async function getInboxAsync() {
    // Ensure client isn't undefined
    if (!_userClient) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    return _userClient.api('/me/mailFolders/inbox/messages')
      .select(['from', 'isRead', 'receivedDateTime', 'subject', 'body'])
      .top(25)
      .orderby('receivedDateTime DESC')
      .get();
  }
  module.exports.getInboxAsync = getInboxAsync;

  async function sendMailAsync(subject, body, recipient) {
    console.log(`Subject: ${subject}`)
    console.log(`Body: ${body}`)
    console.log(`Recipient: ${recipient}`)
    
    // Ensure client isn't undefined
    if (!_userClient) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    // Create a new message
    const message = {
      subject: subject,
      body: {
        content: body,
        contentType: 'text'
      },
      toRecipients: [
        {
          emailAddress: {
            address: recipient
          }
        }
      ]
    };
  
  
    // Send the message
    return _userClient.api('me/sendMail')
      .post({
        message: message
      });
  }
  module.exports.sendMailAsync = sendMailAsync;

  async function deleteMailAsync(id) {
    console.log(`Delete Email: id = ${id}`)
    return _userClient.api(`/me/messages/${id}`)
	.delete(); 
  }
  module.exports.deleteMailAsync = deleteMailAsync;

  async function createEventAsync() {
    const event = {
      subject: 'Let\'s go for lunch',
      body: {
        contentType: 'HTML',
        content: 'Does noon work for you?'
      },
      start: {
          dateTime: '2017-04-15T12:00:00',
          timeZone: 'Pacific Standard Time'
      },
      end: {
          dateTime: '2017-04-15T14:00:00',
          timeZone: 'Pacific Standard Time'
      },
      location: {
          displayName: 'Harry\'s Bar'
      },
      attendees: [
        {
          emailAddress: {
            address: 'samanthab@contoso.com',
            name: 'Samantha Booth'
          },
          type: 'required'
        }
      ],
      allowNewTimeProposals: true,
      transactionId: '7E163156-7762-4BEB-A1C6-729EA81755A7'
    };
    console.log(`Create Event: ${event.subject}`)
    return _userClient.api('me/events').post(event);
  }
  module.exports.createEventAsync = createEventAsync;

  async function listEventsAsync() {
    console.log('Getting events...');
    return _userClient.api('/me/events')
    .header('Prefer','outlook.timezone="Pacific Standard Time"')
    .select('subject,body,bodyPreview,organizer,attendees,start,end,location')
    .get();
  }
  module.exports.listEventsAsync = listEventsAsync;

  
  // This function serves as a playground for testing Graph snippets
// or other code
async function makeGraphCallAsync() {
    // INSERT YOUR CODE HERE
  }
  module.exports.makeGraphCallAsync = makeGraphCallAsync;