import quip from "quip";
import App from "./App.jsx";

export class Root extends React.Component {
    static propTypes = {
        auth: React.PropTypes.instanceOf(quip.apps.Auth).isRequired
    };
    constructor(props) {
        super(props);
        this.state = {
            isLoggedIn: props.auth.isLoggedIn(),//true, //props.auth.isLoggedIn(),
            error: null
        };
    }
    
    render() {
        const { auth } = this.props;
        const { isLoggedIn, error } = this.state;
        return (
            <div>
                {isLoggedIn && 
                    <div>
                        <App auth={auth} POWERBIWORKSPACE="<workspaceid>" TOKENURL="<token endpoint>"  forceLogin={this.forceLogin} />
                    </div>        
                }
                {!isLoggedIn &&
                    <button onClick={this.onLoginClick}>
                        Login
                    </button>}
                {error && <div>We encountered an error. Please try again</div>}
            </div>
        );
    }

    onLogoutClick = () => {
        console.log ('logout called')
        this.props.auth
            .logout()
            .then(
                () => {
                    console.log ('returned from login success')
                    this.setState({
                        isLoggedIn: this.props.auth.isLoggedIn(),
                        error: null
                    });
                },
                error => {
                    console.log ('returned from logout error')
                    this.setState({ error });
                }
            ).catch((e) => {
                console.log ('error '+ e)
            });
    }

    onLoginClick = () => {
        console.log ('login called')
        this.props.auth
            .login({
                //"resource": "https://analysis.windows.net/powerbi/api",
                //"prompt": "consent"
            })
            .then(
                () => {
                    console.log ('returned from login success')
                    this.setState({
                        isLoggedIn: this.props.auth.isLoggedIn(),
                        error: null
                    });
                },
                error => {
                    console.log ('returned from login error')
                    this.setState({ error });
                }
            ).catch((e) => {
                console.log ('error '+ e)
            });
    };
    forceLogin = () => {
        this.setState({ isLoggedIn: false });
    };
}

class BPIReport extends quip.apps.RootRecord {
    static getProperties() {
        return {
            gotreport: "boolean",
            id: "string",
            name: "string",
            webUrl: "string",
            embedUrl: "string",
            datasetId: "string",
            type: "string"
        };
    }
}
quip.apps.registerClass(BPIReport, "root");

quip.apps.initialize({
    initializationCallback: function(rootNode, { isCreation }) {

        let rootRecord = quip.apps.getRootRecord();
        if (isCreation) {
            rootRecord.set("gotreport", false);
        }
        ReactDOM.render(<Root auth={quip.apps.auth("EmbedDashboard")}/>, rootNode);
    },
});
