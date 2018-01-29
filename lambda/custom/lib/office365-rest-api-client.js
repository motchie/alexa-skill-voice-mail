'use strict';
const MicrosoftGraph = require("@microsoft/microsoft-graph-client");
var moment = require('moment-timezone');

let client;

function setAccessToken(token) {
    client = MicrosoftGraph.Client.init({
        authProvider: (done) => {
            done(null, token);
        }
    });

}

function countUnreadMails() {
    return new Promise(
        (resolve, reject) => {
            client
                .api('/me/mailfolders/inbox/messages')
                .top(25)
                .filter("isRead eq false")
                .count(true)
                .select("odata.count")
                .get()
                .then(
                    (res) => { resolve(res.value.length); }
                )
                .catch(
                    (err) => { reject(console.log(err)) }
                );
        });
};

function UnReadMails() {
    return new Promise(
        (resolve, reject) => {
            client
                .api('/me/mailfolders/inbox/messages')
                .top(25)
                .filter("isRead eq false")
                .select("id", "from", "subject", "bodyPreview", "receivedDateTime")
                .get()
                .then(
                    (res) => { resolve(processingEMails(res)); }
                ).catch(
                    (err) => {
                        reject(console.log(err));
                    }
                );
        });
}

function countTodayMails() {
    return new Promise(
        (resolve, reject) => {
            let today = todayString();
            client
                .api('/me/mailfolders/inbox/messages')
                .top(25)
                .filter('receivedDateTime ge ' + today)
                .count(true)
                .select("odata.count")
                .get()
                .then(
                    (res) => {
                        resolve(res.value.length);
                    }
                )
                .catch(
                    (err) => { reject(console.log(err)) }
                );
        });
};

function todayMails() {
    return new Promise(
        (resolve, reject) => {
            let today = todayString();
            client
                .api('/me/mailfolders/inbox/messages')
                .top(25)
                .filter('receivedDateTime ge ' + today)
                .select("id", "from", "subject", "bodyPreview", "receivedDateTime")
                .get()
                .then(
                    (res) => {
                        resolve(processingEMails(res));
                    }
                )
                .catch(
                    (err) => { reject(console.log(err)) }
                );
        });
};

function todayString() {
    moment.tz.setDefault("Asia/Tokyo");
    let today = moment().startOf('day').utc().format();

    return today;
}

function processingEMails(rawEmails) {
    const emails = [];

    for (let rawEmail of rawEmails.value) {
        const email = {};

        email.id = rawEmail.id;
        email.from = rawEmail.from.emailAddress.name;
        email.subject = rawEmail.subject;
        email.body = rawEmail.bodyPreview;
        email.body = email.body.replace(/\r\n+/ig, "");
        email.body = email.body.replace(/--+/ig, "");
        email.body = email.body.replace(/  +/ig, "");
        email.received = rawEmail.receivedDateTime;

        emails.push(email);
    }
    return emails;
}

module.exports.setAccessToken = setAccessToken;
module.exports.countUnreadMails = countUnreadMails;
module.exports.UnReadMails = UnReadMails;
module.exports.countTodayMails = countTodayMails;
module.exports.todayMails = todayMails;