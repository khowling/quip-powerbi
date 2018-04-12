const consts = require ('../shared/common.js')

module.exports = function (context, req) {
    context.done(null, {
        status: 302,
        body: "",
        headers: {
            "Location": consts.oauth2_authorise_flow_v1,
            "Content-Type": 'text/plain',
        }
    });
};