## Embed PowerBI into Quip

Very brief instructions on setting up this app:


1. Create PowerBI workspace and required Reports in PowerBI
2. Register a new app in Azure AD, and note the Application Id
3. Grant the AzureAD app permissions on the PowerBI service
4. Create a Auth Configuration in Quid developer console, call it `PBI`, populate the form using the AzureAD tenant OAUTH2 authorize and token endpoints, the clientid = AzureAD application Id, and set scope : https://analysis.windows.net/powerbi/api
5. Get the quip app Redirect ULR, and add it to the AzureAD apps reply URL
6. update the POWERBIWORKSPACE `<powerbi workspace id>` prop that is passed to App component in the root.jsx file (REACT_APP variables dont appear to work!)

