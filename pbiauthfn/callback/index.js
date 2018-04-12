const  consts = require ('../shared/common.js')


module.exports = function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');

    context.log (process.version)
    if (req.query.code) {
        context.log (`/callback - got code, calling getAccessToken ${consts.aad_hostname}`)
		consts.getAccessToken (consts.aad_hostname, {code: req.query.code}, context).then((auth) => {
			context.log ('success, got token ')
			
            return context.done(null, {
                // status: 200, /* Defaults to 200 */
                body: "Success"
            })
        } , (err) => {
			return context.done(null, {
                status: 400,
                body: "Error: " + err.message
            })
		})
    }
    else {
        context.log ('/callback - no code')
        context.done(null, {
            status: 400,
            body: "Please pass a code on the query"
        })
    }
}