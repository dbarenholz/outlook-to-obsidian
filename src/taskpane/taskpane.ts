import axios from 'axios';
import * as msal from "@azure/msal-browser";

/**
 * Office is global, provided by OfficeJS API. 
 * When this is loaded, we do some initial work.
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Sets the sideload-msg to invisible.
    document.getElementById("sideload-msg").style.display = "none";
    // Sets the app to visible.
    document.getElementById("app-body").style.display = "flex";
    // Sets which function to run when clicked on the element
    document.getElementById("send").onclick = sendToObsidian;
  }
});

/**
 * Method that sends the current emailchain to Obsidian.
 */
export async function sendToObsidian() {
  // Set to false when publishing.
  const DEBUG = true;


  // ====== MSAL ATTEMPT ======

  /*

  // // Configuration object for microsoft authentication API
  const msalConfig = {
    auth: {
      clientId: '023a0425-378b-4662-8fe6-252e1de141ef'
    }
  };
  // Create an instance of the API
  const msalInstance = new msal.PublicClientApplication(msalConfig)

  // Try to login
  msalInstance.loginRedirect({
    redirectStartPage: "https://login.microsoftonline.com", // Tries to login to my school account, so can't test... I'm guessing this IS correct, but doesn't work for my school account
    scopes: ["mail.read"],
    redirectUri: "http://localhost:3000/blank.html"
  }).then((auth_result) => {
    document.getElementById("info").innerHTML += "MSAL: " + JSON.stringify(auth_result)
    console.log("MSAL:")
    console.log(JSON.stringify(auth_result))
  }).catch((error) => {
    console.log("MSAL:")
    console.log(error)
  })

  */
  // ====== AXIOS ATTEMPT ======

  const AUTHORIZE = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
  const token = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
  const API = "https://graph.microsoft.com/"

  // Login

  /*
  axios({
    method: 'get',
    url: AUTHORIZE,
    params: {
      client_id: "023a0425-378b-4662-8fe6-252e1de141ef",
      response_type: 'code',
      redirect_uri: "http://localhost",
      scope: "https://graph.microsoft.com/mail.read"
    }
  }).then((response) => {
    document.getElementById("info").innerHTML += "AXIOS: " + JSON.stringify(response)
    console.log("AXIOS:")
    console.log(`[outlook-to-obsidian]: response.status: ${response.status}`)
    console.log(`[outlook-to-obsidian]: response.statusText: ${response.statusText}`)
    console.log(`[outlook-to-obsidian]: response.headers: ${response.headers}`)
    console.log(`[outlook-to-obsidian]: response.data: ${response.data}`)
  }).catch((error) => {
    console.log("AXIOS:")
    console.log(`[outlook-to-obsidian]: err: ${error}`)
  })
  */

  // ====== OFFICEJS ATTEMPT ======

  // Get authorization token
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (asyncResult) => {
    // Success
    if (asyncResult.status.toString() === "succeeded") {
      // Store the token in a variable
      const TOKEN = asyncResult.value;

      if (DEBUG) {
        console.log(`[outlook-to-obsidian]: Got token: ${TOKEN}`);
      }

      // Get item ID in correct format
      let ITEM_ID = null
      if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
        ITEM_ID = Office.context.mailbox.item.itemId;
      } else {
        ITEM_ID = Office.context.mailbox.convertToRestId(
          Office.context.mailbox.item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );
      }

      // Call the API
      axios({
        method: 'get',
        url: `${API}/v1.0/me/messages/${ITEM_ID}`,
        headers: {
          'Authorization': `Bearer ${TOKEN}`
        }
      }).then((response) => {
        if (response.status == 200) {
          // Success
        } else {
          // Failure
          console.log(`[outlook-to-obsidian]: response: ${response}`)

        }
      }).catch((err) => {
        console.log(`[outlook-to-obsidian]: Error: ${err}`)
      })


    } else {
      // No success...
      console.log("[outlook-to-obsidian]: Could not get attachment token.")
      console.log(`[outlook-to-obsidian]: Error: ${asyncResult.error}`)
      console.log(`[outlook-to-obsidian]: Diagnostics: ${asyncResult.diagnostics}`)
    }
  })

}
