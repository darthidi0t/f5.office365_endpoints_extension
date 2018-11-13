/**
*** Name       : office365_endpoints_extension
*** Authors    : Niels van Sluis, <niels@van-sluis.nl>
*** Modified by: Brett Smith @f5
*** Version    : 0.4
*** Date       : 2018-11-13
***
*** Changes
***            : v0.4 - added error handling on https.get request.
***            : v0.3 - rewrite to interact with Microsoft REST-based web service.
***            : v0.2 - added URL checking needed for Intelligent Proxy Steering.
***            :        See: https://devcentral.f5.com/articles/intelligent-proxy-steering-office365-31037
***            : v0.1 - initial version
**/
 
'use strict';
 
// Import the f5-nodejs module and others.
var f5 = require('f5-nodejs');
var https = require('https');
var repeat = require('repeat');
const uuidv4 = require('uuid/v4');
var loki = require('lokijs');
var concat = require('unique-concat');
var ipRangeCheck = require('ip-range-check');

// Create a new rpc server for listening to TCL iRule calls.
var ilx = new f5.ILXServer();

// Create (in-memory) LokiJS database and collections.
var db = new loki('db.json');
var products = db.addCollection('products');
var dbEndpointSet = db.addCollection('endpointSet', { unique: ["id"] });

// Set initial version
var latestVersion = '0000000000';

var updateNeeded = 0;

// helper to call the webservice
function getJson(methodName, instanceName, clientRequestId, callback) {
    var ws = "https://endpoints.office.com";
    var requestPath = ws + '/' + methodName + '/' + instanceName + '?clientRequestId=' + clientRequestId;
    
    var req = https.get(requestPath, function(res) {
        var data = '';
 
        res.on('data', function(chunk) {
            data += chunk;
        });
 
        res.on('error', function(e) {
            callback(e, null);
        }); 
 
        res.on('timeout', function(e) {
            callback(e, null);
        }); 
 
        res.on('end', function() {
            if(res.statusCode == 200) {
                callback(null, data);
            }
        });
    }).on('error', function(e) {
        console.log("Got error: " + e.message);
    });
}

function logStatistics() {
    
    // statistics about entries in loki database.
    var v, i;
    console.log('info: ' + products.count() + ' serviceAreas found.');
    products.find().forEach((v, i) => {
        console.log ('info: serviceArea ' + v.name + ' holds ' + v.ipAddresses.length + ' IP addresses and ' + v.urls.length + ' URLs');
    });
}

function checkVersion () {

   // check and update version
   getJson('version', 'worldwide', uuidv4(), function(err,data) {
        if(err) {
            console.log("Error: failed to fetch Office365 JSON.");
            return;
        }
        
        // if data happens to be empty due to an error, do not continue.
        if(!data) {
            console.log("Error: Office365 JSON contains no data.");
            return;
        }
        
        var o365json = JSON.parse(data);
        
        // check if update is needed
        if (latestVersion < o365json.latest ) {
            updateNeeded = 1;
            latestVersion = o365json.latest;
            console.log('info: update available');
        }
        
        // free memory
        o365json = null;
  
        console.log('info: latest version available is: ' + latestVersion);
    });

}

// Function that uses the Office365 endpoint data to create a database
// that can be used to perform IP address and URL lookups.
function getOffice365Endpoints() {
    console.log('info: getOffice365Endpoints function called');
    
    // only update when new version is available
    if ( updateNeeded == 1 || latestVersion == '0000000000' ) {
        console.log('info: fetching new endpoint database.');
        // Update Office365 endpoints
        getJson('endpoints', 'worldwide', uuidv4(), function(err,data) {
            if(err) {
                console.log("Error: failed to fetch Office365 JSON.");
                return;
            }
        
            // if data happens to be empty due to an error, do not continue.
            if(!data) {
                console.log("Error: Office365 JSON contains no data.");
                return;
            }
            
            // define arrays that will contain every Office 365 associated IP addresses and URLs.
            // used for queries on all product groups.
            var allIpAddresses = [];
            var allUrls = [];
        
            var o365json = JSON.parse(data);
            
            // clear current collections
            dbEndpointSet.clear();
            products.clear();
            
            // go through each endpoint set
            for(var endpointSet in o365json) {
                var ipAddresses = [];
                var urls = [];

                for(var url in o365json[endpointSet].urls) {
                    //console.log(o365json[endpointSet].urls[url]);
                    if(o365json[endpointSet].urls[url].indexOf('*') !== -1) {
                        // Start Version 0.2 - Added by Brett Smith @f5.com
                        // remove the wildcard (*) from the URL
                        o365json[endpointSet].urls[url] = o365json[endpointSet].urls[url].substr(o365json[endpointSet].urls[url].indexOf('*')+1);
                        // End Version 0.2
                    }
                    urls.push(o365json[endpointSet].urls[url]);
                }
            
                for(var ip in o365json[endpointSet].ips) {
                    //console.log(o365json[endpointSet].ips[ip]);
                    ipAddresses.push(o365json[endpointSet].ips[ip]);
                }
                
                // insert into lokijs database
                dbEndpointSet.insert({
                    id: o365json[endpointSet].id,
                    serviceArea: o365json[endpointSet].serviceArea,
                    serviceAreaDisplayName: o365json[endpointSet].serviceAreaDisplayName,
                    tcpPorts: o365json[endpointSet].tcpPorts,
                    udpPorts: o365json[endpointSet].udpPorts,
                    category: o365json[endpointSet].category,
                    required: o365json[endpointSet].required,
                    notes: o365json[endpointSet].notes,
                    expressRoute: o365json[endpointSet].expressRoute,
                    urls: urls,
                    ips: ipAddresses
                });
                
                // insert IP addresses and Urls by serviceArea to loki database
                var serviceArea = products.findObject({'name':o365json[endpointSet].serviceArea.toLowerCase()});
                if(!serviceArea || serviceArea.version < latestVersion) {
                    //console.log('DEBUG: products.serviceArea insert.');
                    products.insert({ name: o365json[endpointSet].serviceArea.toLowerCase(), ipAddresses: ipAddresses, urls: urls, version: latestVersion });
                } else {
                    //console.log('DEBUG: products.serviceArea update.');
                    serviceArea.ipAddresses = concat(serviceArea.ipAddresses, ipAddresses);
                    serviceArea.urls = concat(serviceArea.urls, urls);
                    products.update(serviceArea);
                }
                
                allIpAddresses = concat(allIpAddresses, ipAddresses);
                allUrls = concat(allUrls, urls);
            }
            
            // insert all IP addresses and URLs to loki database
            var allServiceAreas = products.findObject({'name':'any'});
            if(!allServiceAreas || allServiceAreas.version < latestVersion) {
                products.insert({ name: 'any', ipAddresses: allIpAddresses, urls: allUrls, version: latestVersion });
            }
            
            // log statistics
            logStatistics();
            
            // free memory
            o365json = null;
        });
        
        updateNeeded = 0;
        
    } else {
        console.log('info: no update needed');
        
        // log statistics
        logStatistics();
    }
}

// refresh Microsoft Office 365 endpoints every hour
repeat(checkVersion).every(15, 'minutes').start.now();
repeat(getOffice365Endpoints).every(1, 'hour').start.now();

// Start Version 0.2 - Added by Brett Smith @f5.com
// ILX::call to check if an URL is part of Office365
ilx.addMethod('checkProductURL', function(objArgs, objResponse) {
    
    // Valid values for serviceArea are: common, exchange, sharepoint, skype or any
    var serviceArea = objArgs.params()[0];
    var hostName = objArgs.params()[1];
    
    var verdict = false;
    
    var req = products.findObject( { 'name':serviceArea.toLowerCase()});
    if(Array.isArray(req.urls)) {
        for (let url of req.urls) {
            if(hostName.indexOf(url) !== -1) {
                console.log("match found:" + url);
                verdict = true;
                break;
            }
        }
    }
 
    // return verdict to Tcl iRule
    objResponse.reply(verdict);
});
// End Version 0.2
 
// ILX::call to check if an IP address is part of Office365
ilx.addMethod('checkProductIP', function(objArgs, objResponse) {
    
    // Valid values for serviceArea are: common, exchange, sharepoint, skype or any
    var serviceArea = objArgs.params()[0];
    var ipAddress = objArgs.params()[1];
    
    // fail-open = true, fail-close = false
    var verdict = true;
    
    var req = products.findObject( { 'name':serviceArea.toLowerCase()});
    if(req) {
        verdict = ipRangeCheck(ipAddress, req.ipAddresses);
    }
 
    // return AuthnRequest to Tcl iRule
    objResponse.reply(verdict);
});

ilx.listen();
