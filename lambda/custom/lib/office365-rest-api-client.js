'use strict';
const MicrosoftGraph = require("@microsoft/microsoft-graph-client");

let client;

function setAccessToken(token) {
    client = MicrosoftGraph.Client.init({
        authProvider: (done) => {
            done(null, token);
        }
    });

}

function unReadMailCount() {
    return new Promise(
        (resolve, reject) => {
            client
                .api('/me/mailfolders/inbox/messages')
                .filter("isRead eq false")
                .count(true)
                .select("odata.count")
                .get()
                .then(
                    (res) => { resolve(res.value.length); }
                )
                .catch(
                    (err) => { reject(new Error(err)) }
                );
        });
};

module.exports.setAccessToken = setAccessToken;
module.exports.unReadMailCount = unReadMailCount;