import React, { useState } from 'react'; 
import {Client} from '@microsoft/microsoft-graph-client';

import {  
  StyleSheet,  
  View,  
  Text,  
  TouchableHighlight,  
} from 'react-native';  
  
import { authorize } from 'react-native-app-auth';  
  
const AuthConfig = {  
  appId: "88558549-cc10-4253-bc74-63430c2b4d8e",  
  tenantId: "2ca68321-0eda-4908-88b2-424a8cb4b0f9",  
  appScopes: [  
    'openid',  
    'offline_access',  
    'profile',  
    'User.Read',  
    'MailboxSettings.Read',
    'Calendars.ReadWrite',
  ],  
};  
  
const config = {  
  warmAndPrefetchChrome: true,  
  clientId: AuthConfig.appId,  
  redirectUrl: Platform.OS === 'ios' ? 'urn:ietf:wg:oauth:2.0:oob' : 'mlogin://react-native-auth',  
  scopes: AuthConfig.appScopes,  
  additionalParameters: { prompt: 'select_account' },  
  serviceConfiguration: {  
    authorizationEndpoint: 'https://login.microsoftonline.com/' + AuthConfig.tenantId + '/oauth2/v2.0/authorize',  
    tokenEndpoint: 'https://login.microsoftonline.com/' + AuthConfig.tenantId + '/oauth2/v2.0/token',  
  },  
};  

var eList;
var accessToken;
var bodyPreview;


export class GraphAuthProvider {
  getAccessToken = async () => {
    const token = accessToken;
    return token || '';
  };
}

  
const App: () => React$Node = () => {  
  const [result, setResult] = useState({});  
  const [eventsList, setEventsList] = useState({});  
  
  loginWithOffice365 = async () => {  
    let tempResult = await authorize(config);  
    setResult(tempResult); 
    accessToken = tempResult.accessToken;
    getEventsList();
  }; 


  getEventsList = async () => {
    // GET /me/events
    const clientOptions = {
      authProvider: new GraphAuthProvider(),
    };

    const graphClient = Client.initWithMiddleware(clientOptions);
    let tempEventsList = await graphClient
      .api('/me/events')
      .header('Prefer', 'outlook.timezone="Eastern Standard Time"')
      .select('subject')
      .top(5)
      .get();
      setEventsList(tempEventsList);
      eList = tempEventsList;
      console.log(eList);
      bodyPreview = JSON.stringify(eList, null, 2);
  };

  return (  
    <>  
      <View style={styles.container}>  
        <TouchableHighlight  
          style={[styles.buttonContainer, styles.loginButton]}  
          onPress={() => loginWithOffice365()}>  
          <Text style={styles.loginText}>Login with Office365</Text>  
        </TouchableHighlight>  
        <Text>{result.accessToken ? "Logged In" : ""}</Text>
      </View>  
    </>  
  );  
};  
  
const styles = StyleSheet.create({  
  container: {  
    flex: 1,  
    justifyContent: 'center',  
    alignItems: 'center',  
    backgroundColor: '#DCDCDC',  
  },  
  buttonContainer: {  
    height: 45,  
    flexDirection: 'row',  
    justifyContent: 'center',  
    alignItems: 'center',  
    marginBottom: 20,  
    width: 250,  
    borderRadius: 30,  
  },  
  loginButton: {  
    backgroundColor: '#3659b8',  
  },  
  loginText: {  
    color: 'white',  
  },  
});  
  
export default App; 