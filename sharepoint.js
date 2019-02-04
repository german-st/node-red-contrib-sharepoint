var sprequest = require('sp-request');

  module.exports = function(RED) {
  
    function nodeRedSharepoint(config) {
      RED.nodes.createNode(this, config);
      var node = this;
      node.name = config.sharepoint;
      node.username = config.username;
      node.password = config.password;
	    node.method = config.method;
      node.domain = config.domain;
 
      this.on('input', function (msg) {
        var SharepointFlowNode = RED.nodes.getNode(node.name);
        
        // get service from config and method
        var serviceUri = config.serviceUri||msg.spURL;
		    var method = config.method||'GET';  //GET by default
		
        // exit if empty credentials
        if (SharepointFlowNode == null || SharepointFlowNode.credentials == null) {
          node.warn('Sharepoint credentials are missing.');
          return;
        }
        // error if no service url ot type
        if (serviceUri =='' || serviceUri == null || serviceUri.indexOf('_api')<0) {
          node.error('Service URL must be specified in Sharepoint node or passed by stream mode in msg.spURL. And URL must contain \"_api\" path');
          return;
        }
       
        //Set credentials for connect
        let credentialOptions = {
          username: SharepointFlowNode.credentials.username,
          password: SharepointFlowNode.credentials.password,
          domain: SharepointFlowNode.credentials.domain||''
        };
        
        //Base auth method for our requests
        let spr = sprequest.create(credentialOptions);
		
		    //node.warn("DATAS: "+JSON.stringify(credentialOptions)+'\n'+method+'\n'+serviceUri);
		
        //Select the method
        if (method == 'GET')  //GET
          {
            //Get data by selected or specified service
            spr.get(encodeURI(serviceUri))
              .then(response => {
              msg.sharepointresult = response.body.d;
              node.send([msg, null]);
              })
              .catch(err =>{
              msg.error = err;
              node.send([null, msg]);
              node.warn('Sharepoint node error GET: '+err);
              console.log('Sharepoint node error GET: '+err);
              }); 
          }
        else //POST
          {
              var digestURI = serviceUri.substring(0, serviceUri.indexOf('/_api'));
              spr.requestDigest(digestURI)
                .then(digest => {
                  //Set special headers by stream mode from msg.headers
                  var spHeaders = {
                    'X-RequestDigest': digest
                  };
                  Object.assign(spHeaders, msg.headers||''); //Headers joining
                  //Post query to SharePoint
                  return spr.post(encodeURI(serviceUri), {
                    body: JSON.stringify(msg.payload||''),
                    headers: {
                      spHeaders
                    }
                  });
                })
                .then(response => {
                    msg.sharepointresult = response.body.d;
                    msg.resCodeSharepoint = response.statusCode;
                    node.send([msg, null]);
                  }
                )
                .catch(err =>{
                  msg.error = err;
                  node.send([null, msg]);
                  node.warn('Sharepoint node error POST: '+err);
                  console.log('Sharepoint node error POST: '+err);
                  });
              }
        });
  }

    RED.nodes.registerType('nodeRedSharepoint', nodeRedSharepoint);
 
    function nodeRedSharepointSettings(n) {
      RED.nodes.createNode(this, n);
    }
 
    RED.nodes.registerType('nodeRedSharepoint-access', nodeRedSharepointSettings, {
      credentials: {
        username: {
          type: 'text'
        },
        password: {
          type: 'password'
        },
        domain: {
          type: 'text'
        },
      }
    });
};