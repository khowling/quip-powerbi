import Styles from "./App.less";
import  { service, factories, IEmbedConfiguration, Embed }  from 'powerbi-client';
import quip from "quip";
import * as models from 'powerbi-models'

export default class App extends React.Component {

    static propTypes = {
        auth: React.PropTypes.instanceOf(quip.apps.Auth).isRequired,
        forceLogin: React.PropTypes.func.isRequired
    };

    state = {mode: "load"}

    componentDidMount() {
        let group = this.props.POWERBIWORKSPACE,
            tokenurl = this.props.TOKENURL,
            bearer = this.props.auth.getTokenResponseParam("id_token") // this.props.auth.getTokenResponseParam("access_token")

        if (group && tokenurl) {
            console.log (`displaying power bi group/report ${group}`)
            this.getPBIServiceToken(tokenurl, bearer).then(auth => {
                let record = quip.apps.getRootRecord();
                if (record.get("gotreport")) {
                    this.displayReport(auth, group, record.getData())
                } else {
                    this.listReports(auth, group).then(reports => {
                        this.setState({mode: "list", auth: auth, group: group, reports: reports.value})
                    })
                }
            }, err => {
                this.setState({mode: "error", error: err})
            })
        } else {
            this.setState({mode: "error", error: `Application requires REACT_APP_TOKEN_URL ${tokenurl} & REACT_APP_POWERBI_WORKSPACE ${group} to be set, see youre developer ${JSON.stringify(Object.keys(process.env))}`})
        }
    }

    getPBIServiceToken(tokenurl, bearer, tryrefresh = true) {
        return new Promise((accept, reject) => { 
            fetch (tokenurl,{
                headers : { 
                    "Authorization": `Bearer ${bearer}`
                }
            }).then (response => {
                if (response.status === 401 && tryrefresh) {
                    console.log ('trying refesh')
                    this.props.auth.refreshToken().then(
                        refresh => {
                            //console.log (`got refesh ${JSON.stringify(refresh)}`)
                            this.getPBIServiceToken(tokenurl, refresh.access_token, false).then(auth => accept(auth), err => reject (err))
                        }, err => reject (`failed to refresh token ${err}`)
                    )
                } else if (response.status !== 200) {
                    reject('Looks like there was a problem getting the token. Status Code: ' + response.status);
                } else {
                    response.json().then((auth) => {
                        accept(auth)
                    }, err => reject('Looks like there was a problem getting the powerbi token. Status Code: ' + err))
                }
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
                        //console.log (`body : ${JSON.stringify(body)}`)
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
          //console.log (`generateEmbedToken: ${JSON.stringify(auth)}  : ${group} : ${report}`)
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
                //console.log (`generateEmbedToken GenerateToken response: ${res.statusCode}`)
               // if(!(res.statusCode === 200 || res.statusCode === 201)) {
               //     reject({code: res, message: "failted"})
               // } else {
                    res.json().then(body => {
                        //console.log (`body : ${JSON.stringify(body)}`)
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

    setReport (report, type) {
        let record = quip.apps.getRootRecord();
        record.set("gotreport", true)
        record.set("id", report.id)
        record.set("name", report.name)
        record.set("webUrl", report.webUrl)
        record.set("embedUrl", report.embedUrl)
        record.set("datasetId", report.datasetId)
        record.set("type", type)

        quip.apps.sendMessage(`Selected Report ${report.name}`)
        this.displayReport (this.state.auth, this.state.group, quip.apps.getRootRecord().getData())
    }

    displayReport (auth, group, report) {
        this.generateEmbedToken(auth, group, report.id).then (etoken => {
            //console.log (`got embed token ${etoken.token}`)
            //quip.apps.registerEmbeddedIframe()
            let powerBiService = new service.Service(
                factories.hpmFactory,
                factories.wpmpFactory,
                factories.routerFactory);

            if (report.type === "qna") {
                powerBiService.embed(this.pbicontent, {
                    type: 'qna',
                    accessToken: etoken.token,
                    tokenType: models.TokenType.Embed, //Aad
                    embedUrl: `https://msit.powerbi.com/qnaEmbed?groupId=${group}`,
                    datasetIds: [report.datasetId],
                    viewMode:  0 // interactive
                })
                
            } else if (report.type === "report") {

                console.log ('embed report')
                let settings = {
                    filterPaneEnabled: false,
                    navContentPaneEnabled: false,
                    layoutType: models.LayoutType.Master
                }

                if (quip.apps.isMobile()) {
                    console.log ('mobile layout')
                    settings.layoutType = models.LayoutType.MobilePortrait
                }

                powerBiService.embed(this.pbicontent, {
                    type: 'report',
                    accessToken: etoken.token,
                    tokenType: models.TokenType.Embed, //Aad
                    embedUrl: report.embedUrl,
                    permissions: models.Permissions.All,
                    settings: settings
                })
            }

            this.setState({mode: "display", report: `${report.type}: ${report.name}` })
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
                            <th scope="col"></th>
                            </tr>
                        </thead>
                        <tbody>
                            { this.state.reports.map(r =>
                                <tr>
                                    <td><h2>{r.name}</h2></td>
                                    <td>
                                        <quip.apps.ui.Button text="report"  onClick={this.setReport.bind(this, r, "report")}/>
                                    </td>
                                    <td>
                                        <quip.apps.ui.Button text="qna"  onClick={this.setReport.bind(this,  r, "qna")}/>
                                    </td>
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
