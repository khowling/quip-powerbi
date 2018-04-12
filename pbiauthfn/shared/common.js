const  secret = process.env.SECRET,
    SECRET_NAME = process.env.VAULT_SECRET_KEY || 'rftoken',
    VAULT_NAME = process.env.VAULT_NAME || 'pbivault' ,
    client_id = process.env.CLIENT_ID || "",
    client_secret = process.env.CLIENT_SECRET || "",
    callback_host = process.env.CALLBACK_HOST || "http://localhost:7071/api", // 'https://pbiauth.azurewebsites.net/api',
    client_directory = process.env.CLIENT_DIRECTORY || "",
    use_sp_auth = false && client_secret != null,
    embed_username = process.env.EMBED_USERNAME,
    embed_password = process.env.EMBED_PASSWORD,
    aad_hostname = 'login.microsoftonline.com',
    pbi_hostname = 'api.powerbi.com',
    powerbi_service_resource_v1 = 'https://analysis.windows.net/powerbi/api',
    powerbi_service_scope_v2 = 'https://graph.microsoft.com/mail.send', 
    aad_token_endpoint = `/${client_directory}/oauth2/token`,
    aad_auth_endpoint = `/${client_directory}/oauth2/authorize`,
    aad_token_endpoint_v2 = `/${client_directory}/oauth2/v2.0/token`,
    aad_auth_endpoint_v2 = `/${client_directory}/oauth2/v2.0/authorize`,
    oauth2_authorise_flow_v1 = `https://${aad_hostname}${aad_auth_endpoint}?client_id=${client_id}&redirect_uri=${encodeURIComponent(callback_host+'/callback')}&resource=${encodeURIComponent(powerbi_service_resource_v1)}&response_type=code&prompt=consent`,
    powerbi_group_name = process.env.POWERBI_GROUP_NAME

const 
    http = require('http'),
    https = require('https'),
    url = require('url')

function getAccessToken (hostname, creds, context) {
    context.log (`getAccessToken... ${hostname}  ${aad_token_endpoint}`)
    return new Promise((accept, reject) => {

        let flow_body
        if (creds.code) {
            flow_body = `client_id=${client_id}&scope=${encodeURIComponent(powerbi_service_resource_v1)}&code=${encodeURIComponent(creds.code)}&client_secret=${encodeURIComponent(client_secret)}&grant_type=authorization_code&redirect_uri=${encodeURIComponent(callback_host+'/callback')}`
        } else if (creds.refresh_token) {
            flow_body = `client_id=${client_id}&scope=${encodeURIComponent(powerbi_service_resource_v1)}&client_secret=${encodeURIComponent(client_secret)}&refresh_token=${encodeURIComponent(creds.refresh_token)}&grant_type=refresh_token`
        }
        //context.log (flow_body)

        let	authcode_req = https.request({
                hostname: hostname,
                path: aad_token_endpoint,
                method: 'POST',
                headers: {
                    'Content-Type': "application/x-www-form-urlencoded",
                    'Content-Length': Buffer.byteLength(flow_body)
                }
            }, (res) => {
                let rawData = '';
                res.on('data', (chunk) => {
                    rawData += chunk
                })

                res.on('end', () => {
                    //context.log (`response: ${rawData}`)
                    if (res.statusCode === 301 || res.statusCode === 302) {
                        context.log (`got ${res.statusCode} ${res.headers.location}`)
                        getAccessToken (url.parse(res.headers.location).hostname, creds, context).then((succ) => accept(succ), (err) => reject(err))
                    } else if(!(res.statusCode === 200 || res.statusCode === 201)) {
                        reject({code: res.statusCode, message: rawData})
                    } else {
                        context.log ('200 - success')
                        let authdata;
                        try {
                            authdata = JSON.parse(rawData)
                        } catch (e) {
                            reject({code: 500, message: `failed to parse result ${rawData}`})
                        }
                        context.log (`getAccessToken - Successfully got token_data`) //:  + ${rawData}`)
                        
                        // store refresh_token in vault
                        if (authdata.refresh_token && process.env.MSI_ENDPOINT && process.env.MSI_SECRET) {
                            context.log (`getAccessToken - got MSI_ENDPOINT:  ${process.env.MSI_ENDPOINT}`)
                            let keyvault_token_request = Object.assign(url.parse(`${process.env.MSI_ENDPOINT}?resource=${encodeURIComponent("https://vault.azure.net")}&api-version=2017-09-01`), {headers: {"secret": process.env.MSI_SECRET }})
                            context.log (`getAccessToken - keyvault_token_request  ${keyvault_token_request.href}`)
                            http.get(keyvault_token_request, (msi_res) => {
                                let msi_data = '';
                                msi_res.on('data', (d) => {
                                    msi_data+= d
                                });
                    
                                msi_res.on('end', () => {
                                    context.log (`msi end ${msi_res.statusCode}`)
                                    if(msi_res.statusCode === 200 || msi_res.statusCode === 201) {
                                        let keyvault_access = JSON.parse(msi_data),
                                            request_body = JSON.stringify({"value": authdata.refresh_token })
                                            
                                        context.log (`getAccessToken: writing to keyvault_access ${VAULT_NAME} ${SECRET_NAME} ${keyvault_access.token_type} ${keyvault_access.access_token}`)

                                        let putreq = https.request({
                                            method: "PUT",
                                            hostname : `${VAULT_NAME}.vault.azure.net`,
                                            path : `/secrets/${SECRET_NAME}?api-version=2016-10-01`,
                                            headers: {
                                                "Authorization": `${keyvault_access.token_type} ${keyvault_access.access_token}`,
                                                'Content-Type': "application/json",
                                                'Content-Length': Buffer.byteLength(request_body)
                                            }}, (res) => {
                                                let vault_data = '';
                                                res.on('data', (chunk) => {
                                                    vault_data += chunk
                                                })
                            
                                                res.on('end', () => {
                                                    context.log (`write secret ${res.statusCode} : ${vault_data}`)
                                                    if (res.statusCode === 200) {
                                                        context.log ("getAccessToken - successfully stored refresh token")
                                                        accept(authdata)
                                                    } else {
                                                        context.log(`Got vault setSecret error: ${res.statusCode}`)
                                                        reject({code: res.statusCode, message: `Got vault setSecret error ${vault_data}`})
                                                    }
                                                
                                                })
                                            }).on('error', (e) => {
                                                context.log(`Got vault setSecret error: ${e.message}`);
                                                reject({code: 400, message: `vault setSecret error ${e.message}`})
                                            });
                                        putreq.write(request_body);
                                        putreq.end();
                                    } else {
                                        context.error(`Got MSI error: ${msi_res.statusCode}`);
                                        reject({code: msi_res.statusCode, message: msi_data})
                                    }
                                })
                            }).on('error', (e) => {
                                context.error(`Got MSI error: ${e.message}`)
                                reject({code: 400, message: `Got MSI error ${e.message}`})
                            });
                            
                        } else {
                            context.log ("getAccessToken - no MSI, not storing refresh_token")
                            accept(authdata)
                        }
                    }
                })

            }).on('error', (e) => {
                reject({code: 400, message: e})
            })
        authcode_req.write(flow_body)
        authcode_req.end()
    })
}
    
module.exports = {
    oauth2_authorise_flow_v1,
    client_id,
    VAULT_NAME,
    SECRET_NAME,
    aad_hostname,
    aad_token_endpoint,
    powerbi_service_resource_v1,
    client_secret,
    callback_host,
    getAccessToken
}