const 
    http = require('http'),
    https = require('https'),
    url = require('url')
    consts = require ('../shared/common.js')


function setCurrentAccessToken(context) {
	return new Promise((accept, reject) => {
    
        let keyvault_token_request = Object.assign(url.parse(`${process.env.MSI_ENDPOINT}?resource=${encodeURIComponent("https://vault.azure.net")}&api-version=2017-09-01`), {headers: {"secret": process.env.MSI_SECRET }})
        context.log (`setCurrentAccessToken - got MSI_ENDPOINT, restoring refresh_token`) //  ${JSON.stringify(keyvault_token_request)}`)
        http.get(keyvault_token_request, (msi_res) => {
            let msi_data = '';
            msi_res.on('data', (d) => {
                msi_data+= d
            });

            msi_res.on('end', () => {
                context.log (`msi end ${msi_res.statusCode}`)
                if(msi_res.statusCode === 200 || msi_res.statusCode === 201) {
                    let keyvault_access = JSON.parse(msi_data)
                    let keyvaut_secret_request = Object.assign(url.parse(`https://${consts.VAULT_NAME}.vault.azure.net/secrets/${consts.SECRET_NAME}/?api-version=2016-10-01`), {headers: {
                        "Authorization": `${keyvault_access.token_type} ${keyvault_access.access_token}`}})
                    
                        context.log (`setCurrentAccessToken - getting from keyvault`)

                    https.get(keyvaut_secret_request, (secret_res) => {
                        let vault_data = '';
                        secret_res.on('data', (chunk) => {
                            vault_data += chunk
                        })
    
                        secret_res.on('end', () => {
                            context.log (`get secret ${secret_res.statusCode} : ${vault_data}`)
                            if (secret_res.statusCode === 200) {
                                context.log ("setCurrentAccessToken - successfully got  refresh token from vault")
                                let keyvault_secret = JSON.parse(vault_data)

                                consts.getAccessToken (consts.aad_hostname, {refresh_token: keyvault_secret.value}, context).then((auth) => {
                                    context.log ('setCurrentAccessToken -  successfully got access token!')
                                    accept(auth)
                                }, (err) => {
                                    context.log ('setCurrentAccessToken refresh failed, rejecting with:  ' + err.message)
                                    reject (err)
                                })
                            } else {
                                context.log(`Got vault setSecret error: ${secret_res.statusCode}`)
                                reject({code: secret_res.statusCode, message: `Got vault setSecret error ${vault_data}`})
                            }
                        
                        })
                    }).on('error', (e) => {
                        context.log(`Got vault setSecret error: ${e.message}`);
                        reject({code: 400, message: `vault setSecret error ${e.message}`})
                    });

                } else {
                    context.log(`Got MSI error: ${msi_res.statusCode}`);
                    reject({code: msi_res.statusCode, message: msi_data})
                }
            })
        }).on('error', (e) => {
            context.log(`Got MSI error: ${e.message}`)
            reject({code: 400, message: `Got MSI error ${e.message}`})
        })
    })
}

module.exports = function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request: ' + context.bindings.pbioken);
    if (process.env.MSI_ENDPOINT && process.env.MSI_SECRET) {
        setCurrentAccessToken(context).then ((auth)=> {
            context.done(null, {
                body: JSON.stringify({"access_token": auth.access_token})
            })
        }, (err) => {
            context.done(null, {
                status: 400,
                body: JSON.stringify(err)
            })
		})
    } else {
        context.done(null, {
            status: 400,
            body: "Please enable MSI and setup a keyvault"
        })
    }
}