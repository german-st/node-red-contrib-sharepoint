[![NPM](https://nodei.co/npm/node-red-contrib-sharepoint.png?downloads=true&downloadRank=true&stars=true)](https://nodei.co/npm/node-red-contrib-sharepoint/)

[![npm](https://img.shields.io/npm/dt/node-red-contrib-sharepoint.svg)](https://www.npmjs.com/package/node-red-contrib-sharepoint)

# README #
A simple node for MS SharePoint REST API (NTLM authentication). Supports GET and POST methods.

### Install ###

Install latest release: `npm i -g node-red-contrib-sharepoint`

### Inputs

`msg.spURL` *string*
End point REST. Specified in the config(URL service) or passed in stream mode. URL must contain "/_api/" path.
*Example: http://server/site/_api/lists/getbytitle('listname')*


`msg.headers` *Object*
Special headers for SharePoint **POST** request (methods, etc). You must specify them yourself. `X-RequestDigest` not required to specify, it is written automatically.
*Example: 
`msg.headers= { 'X-HTTP-Method': 'MERGE','IF-MATCH': '*' }`*
[Additional info](https://docs.microsoft.com/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints#properties-used-in-rest-requests "Additional info")


`msg.payload` *Object*
Body(data) for our **POST** request.
*Example:
`msg.payload = { '__metadata': { 'type': 'SP.List' }, 'Title': 'TestList' }`*

[Additional info 1](https://github.com/s-KaiNet/sp-request#sprequestposturl-options "Additional info 1")
[Additional info 2](https://docs.microsoft.com/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service "Additional info 2")

### Outputs

`msg.sharepointresult` *Object*
First output. Result. The object that is returned by the service.

`msg.error` *Object*
Second output. Error. The error that service returned or Nodejs platform error.

### Details
[NTLM authentication](https://github.com/s-KaiNet/node-sp-auth/wiki/SharePoint%20on-premise%20user%20credentials%20authentication "Used NTLM authentication!")
The **GET** method does not need to be specified `msg.headers` and `msg.payload`, the request parameters are transmitted in the URL itself (path, filters, etc.)
[For more info](https://docs.microsoft.com/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service#sharepoint-rest-endpoint-examples "For more info")