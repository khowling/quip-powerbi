import Styles from "./App.less";
import  { service, factories, IEmbedConfiguration, Embed }  from 'powerbi-client';
import * as models from 'powerbi-models'

export default class App extends React.Component {

    static propTypes = {
        auth: React.PropTypes.instanceOf(quip.apps.Auth).isRequired,
        forceLogin: React.PropTypes.func.isRequired
    };

    state = {mode: "load"}

    componentDidMount() {
        console.log (`displaying power bi group/report ${this.props.group}`)
        this.getPBIServiceToken().then(auth => {
            
            this.listReports(auth, this.props.group).then(reports => {
                this.setState({auth: auth, reports: reports.value, mode: "list"})
            })
        }, err => {
            this.setState({error: err})
        })
    }

    getPBIServiceToken() {
        return new Promise((accept, reject) => { 
            fetch ('https://pbiauth.azurewebsites.net/api/gettoken?code=uLBgNuUCmyKd5Czy4iZzRz/CMnZIXOAeXQLj3wTcZewGgcQR1ipbhw==',
                    {headers : { 
                        "Authorization": `Bearer ${this.props.auth.getTokenResponseParam("id_token")}`
                    }
            }).then (response => {
                if (response.status !== 200) {
                    reject('Looks like there was a problem getting the token. Status Code: ' + response.status);
                }
                response.json().then((auth) => {
                    accept(auth)
                }, err => reject('Looks like there was a problem getting the powerbi token. Status Code: ' + err))
            }, err => reject('Looks like there was a problem getting the powerbi token. Status Code: ' + err))
        })
    }

    listReports(auth, group) {
        return new Promise((accept, reject) => {
            fetch (`https://api.powerbi.com/v1.0/myorg/groups/${group}/reports`,{
                headers: new Headers({
                "Authorization": `Bearer ${auth.access_token}`
            })}).then(res => {
                console.log (`listReports  response: ${res.statusCode} `)
            // if(!(res.statusCode === 200 || res.statusCode === 201)) {
            //     reject({code: res, message: "failted"})
            // } else {
                    res.json().then(body => {
                        console.log (`body : ${JSON.stringify(body)}`)
                        if (!body.error) {
                            accept(body)
                        } else {
                            reject({code: "Cannot generate embed token", message: JSON.stringify(body.error)})
                        }
                    }, err => reject({code: "error generate embed token response not json", message: err}))
            // }
            }, err => reject({code: "error, failed to generate embed token", message: err}))
        })
    }
 


    generateEmbedToken (auth, group, report)  {
        return new Promise((accept, reject) => {
          console.log (`generateEmbedToken: ${JSON.stringify(auth)}  : ${group} : ${report}`)
          let	embedtoken_body = JSON.stringify({"accessLevel": "View" })

          fetch(`https://api.powerbi.com/v1.0/myorg/groups/${group}/reports/${report}/GenerateToken`,
            {
                headers: new Headers({
                  "Authorization": `Bearer ${auth.access_token}`,
                  "Content-Type": "application/json",
                  'Content-Length': Buffer.byteLength(embedtoken_body)
                }),
               method: 'POST', // or 'PUT'
               body: embedtoken_body
            }).then(res => {
                console.log (`generateEmbedToken GenerateToken response: ${res.statusCode}`)
               // if(!(res.statusCode === 200 || res.statusCode === 201)) {
               //     reject({code: res, message: "failted"})
               // } else {
                    res.json().then(body => {
                        console.log (`body : ${JSON.stringify(body)}`)
                        if (!body.error) {
                            accept(body)
                        } else {
                            reject({code: "Cannot generate embed token", message: JSON.stringify(body.error)})
                        }
                    }, err => reject({code: "error generate embed token response not json", message: err}))
               // }
            }, err => reject({code: "error, failed to generate embed token", message: err}))
      })
    }

    displayReport (auth, group, report) {
        this.generateEmbedToken(auth, group, report.id).then (etoken => {
            console.log (`got embed token ${etoken.token}`)

            //quip.apps.registerEmbeddedIframe()
            let powerBiService = new service.Service(
                factories.hpmFactory,
                factories.wpmpFactory,
                factories.routerFactory);

            powerBiService.embed(this.pbicontent, {
                type: 'report',
                accessToken: etoken.token,
                tokenType: models.TokenType.Embed, //Aad
                embedUrl: `https://msit.powerbi.com/reportEmbed?reportId=${report.id}&amp;groupId=${group}`,
                permissions: models.Permissions.All,
                settings: {
                    filterPaneEnabled: false,
                    navContentPaneEnabled: false
                }
            })
            this.setState({mode: "display", report: report.name })
        }, err => {
            this.setState({error: 'Looks like there was a problem getting the embed token. Status Code: ' + JSON.stringify(err)})
        })
    }


    render() {
        console.log ('render')
        return (
            <div>
                { this.state.error &&
                    <div>
                    <div style={{"color": "red"}}>{this.state.error}</div>
                    <div><button onClick={this.props.forceLogin}>Logout</button></div>
                    </div>
                }
                { this.state.mode === "load" &&
                    <div>Loading....</div>
                }
                { this.state.mode === "list" &&
                <div>
                    <div></div>
                    <table class="table">
                        <thead>
                            <tr>
                            <th scope="col"></th>
                            <th scope="col"></th>
                            </tr>
                        </thead>
                        <tbody>
                            { this.state.reports.map(r =>
                                <tr>
                                    <td><h2>{r.name}</h2></td>
                                    <td><button style={{"marginLeft": "15px"}} onClick={this.displayReport.bind(this, this.state.auth, this.props.group, r)}>open</button></td>
                                </tr>
                            )}

                        </tbody>
                    </table>
                </div>
                }
                { this.state.mode === "display" &&
                    <div><h2>{this.state.report}</h2>
                    
                    </div>
                }
                <div style={{"width": "100%", "height": "500px"}} ref={(div) => { this.pbicontent = div; }}/>
            </div>
        )
    }
}
