const readline = require('readline-sync');

const settings = require('./appSettings');
const graphHelper = require('./graphHelper');

async function main() {
  console.log('JavaScript Graph Tutorial');

  let choice = 0;

  // Initialize Graph
  initializeGraph(settings);

  // Greet the user by name
  await greetUserAsync();

  const choices = [
    'Display access token',
    'List my inbox',
    'Send mail',
    'Delete mail',
    'List events',
    'Create event'
  ];

  while (choice != -1) {
    choice = readline.keyInSelect(choices, 'Select an option', { cancel: 'Exit' });

    switch (choice) {
      case -1:
        // Exit
        console.log('Goodbye...');
        break;
      case 0:
        // Display access token
        await displayAccessTokenAsync();
        break;
      case 1:
        // List emails from user's inbox
        await listInboxAsync();
        break;
      case 2:
        // Send an email message
        await sendMailAsync();
        break;
      case 3:
        // Delete an email message
        await deleteMailAsync();
        break;
      case 4:
        // Run any Graph code
        await listEventsAsync();
        break;
      case 5:
          // Run any Graph code
          await createEventAsync();
          break;
      default:
        console.log('Invalid choice! Please try again.');
    }
  }
}

main();

function initializeGraph(settings) {
    graphHelper.initializeGraphForUserAuth(settings, (info) => {
      // Display the device code message to
      // the user. This tells them
      // where to go to sign in and provides the
      // code to use.
      console.log(info.message);
    });
  }
  
  async function greetUserAsync() {
    try {
      const user = await graphHelper.getUserAsync();
      console.log(`Hello, ${user?.displayName}!`);
      // For Work/school accounts, email is in mail property
      // Personal accounts, email is in userPrincipalName
      console.log(`Email: ${user?.mail ?? user?.userPrincipalName ?? ''}`);
      console.log(`User Principal: ${user.userPrincipalName}`);
    } catch (err) {
      console.log(`Error getting user: ${err}`);
    }
  }
  
  async function displayAccessTokenAsync() {
    try {
      const userToken = await graphHelper.getUserTokenAsync();
      console.log(`User token: ${userToken}`);
    } catch (err) {
      console.log(`Error getting user access token: ${err}`);
    }
  }
  
  async function listInboxAsync() {
    try {
      const messagePage = await graphHelper.getInboxAsync();
      const messages = messagePage.value;
  
      // Output each message's details
      for (const message of messages) {
        console.log(`Message: ${message.subject ?? 'NO SUBJECT'}`);
        console.log(`  From: ${message.from?.emailAddress?.name ?? 'UNKNOWN'}`);
        console.log(`  Status: ${message.isRead ? 'Read' : 'Unread'}`);
        console.log(`  Received: ${message.receivedDateTime}`);
      }
  
      // If @odata.nextLink is not undefined, there are more messages
      // available on the server
      const moreAvailable = messagePage['@odata.nextLink'] != undefined;
      console.log(`\nMore messages available? ${moreAvailable}`);
    } catch (err) {
      console.log(`Error getting user's inbox: ${err}`);
    }
  }

  async function deleteMailAsync(){
    try {
        const messagePage = await graphHelper.getInboxAsync();
        const messages = messagePage.value;
       
        // Output each message's details
        for (i in messages) {
         const message = messages[i];
    
        console.log(`Message: ${message.subject ?? 'NO SUBJECT'}`);
        console.log(`  From: ${message.from?.emailAddress?.name ?? 'UNKNOWN'}`);
        console.log(`  Status: ${message.isRead ? 'Read' : 'Unread'}`);
        console.log(`  Received: ${message.receivedDateTime}`);
        console.log(`  ID: ${message.id ?? 'ID not found'}`);
        console.log(`  i: ${i ?? 'no count'}`);

        }
        
        count = readline.question('Enter count to delete mail:', count =>{
            const mailToDelete = messages[count]
            console.log(`Delete mail: ${mailToDelete.subject}`);
            console.log(`  From: ${mailToDelete.from?.emailAddress?.name ?? 'UNKNOWN'}`);
            console.log(`  Status: ${mailToDelete.isRead ? 'Read' : 'Unread'}`);
            console.log(`  Received: ${mailToDelete.receivedDateTime}`);
            readline.close();

            
        })
        console.log(`mailID: ${count -  messages[count].id}`);
          await graphHelper.deleteMailAsync(messages[count].id);
        

        
        // If @odata.nextLink is not undefined, there are more messages
        // available on the server
        const moreAvailable = messagePage['@odata.nextLink'] != undefined;
        console.log(`\nMore messages available? ${moreAvailable}`);
      } catch (err) {
        console.log(`Error deleting mail: ${err}`);
      }
  }
  
  async function sendMailAsync() {
    try {
      // Send mail to the signed-in user
      // Get the user for their email address
      const user = await graphHelper.getUserAsync();
      const userEmail = user?.mail ?? user?.userPrincipalName;
  
      if (!userEmail) {
        console.log('Couldn\'t get your email address, canceling...');
        return;
      }
  
      await graphHelper.sendMailAsync('Testing Microsoft Graph',
        'Hello world!', userEmail);
      console.log('Mail sent.');
    } catch (err) {
      console.log(`Error sending mail: ${err}`);
    }
  }
  
  async function createEventAsync() {
    try {
      await graphHelper.createEventAsync();
    } catch (err) {
      console.log(`Error creating Event: ${err}`);
    }
  }

  async function listEventsAsync() {
    try {
      const eventsPage = await graphHelper.listEventsAsync();
      const events = eventsPage.value;
  
      // Output each message's details
      for (const event of events) {
        console.log(`Message: ${event.subject ?? 'NO SUBJECT'}`);
        console.log(`  Start: ${event.start.dateTime ?? 'UNKNOWN'}`);
        console.log(`  End: ${event.end.dateTime ?? 'UNKNOWN'}`);
        console.log(`  Location: ${event.location.displayName ?? 'UNKNOWN'}`);
      }
  
      // If @odata.nextLink is not undefined, there are more messages
      // available on the server
      const moreAvailable = eventsPage['@odata.nextLink'] != undefined;
      console.log(`\nMore events available? ${moreAvailable}`);
    } catch (err) {
      console.log(`Error getting events: ${err}`);
    }
  }

  async function makeGraphCallAsync() {
    try {
      await graphHelper.makeGraphCallAsync();
    } catch (err) {
      console.log(`Error making Graph call: ${err}`);
    }
  }

  