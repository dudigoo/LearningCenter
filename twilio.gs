function sendSms(customerPhoneNumber,msg) {
 msg='kkk';
 customerPhoneNumber='972543112109'
 var twilioAccountSID = 'AC45c3f9d90ee9d557d7a8dde5863d7494';
 var twilioAuthToken = 'fcc5880e9fd4a2b1d6387f2a7a504654';
 var twilioPhoneNumber = '+19705917335';
  var twilioUrl = 'https://api.twilio.com/2010-04-01/Accounts/' + twilioAccountSID + '/Messages.json';
 var authenticationString = twilioAccountSID + ':' + twilioAuthToken;
 try {
   UrlFetchApp.fetch(twilioUrl, {
     method: 'post',
     headers: {
       Authorization: 'Basic ' + Utilities.base64Encode(authenticationString)
     },
     payload: {
       To: "+972" + customerPhoneNumber.toString(),
       Body: msg ,
       From: twilioPhoneNumber,  // Your Twilio phone number
     },
   });
   return 'sent: ' + new Date();
 } catch (err) {
   return 'error: ' + err;
 }
};