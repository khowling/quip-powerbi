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
                        <App auth={auth} group="6a29f8ce-2d64-444c-90a0-0af82c0e330c" report="cdb5d336-726f-4ce1-8925-6c979fb50ce5" forceLogin={this.forceLogin} />
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

quip.apps.initialize({
    initializationCallback: function(rootNode) {
        ReactDOM.render(<Root auth={quip.apps.auth("EmbedDashboard")}/>, rootNode);
    },
});
