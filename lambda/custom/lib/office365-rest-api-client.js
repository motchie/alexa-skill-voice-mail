'use strict';
const MicrosoftGraph = require('@microsoft/microsoft-graph-client');
var moment = require('moment-timezone');

let client;

function setAccessToken(token) {
    client = MicrosoftGraph.Client.init({
        authProvider: (done) => {
            done(null, token);
        }
    });
}

function countUnreadMessages() {
    return new Promise(
        (resolve, reject) => {
            client
                .api('/me/mailfolders/inbox/messages')
                .top(25)
                .filter('isRead eq false')
                .count(true)
                .select('odata.count')
                .get()
                .then(
                    (res) => { resolve(res.value.length); }
                )
                .catch(
                    (err) => { reject(console.log(err)) }
                );
        });
};

function retrieveUnreadMessages() {
    return new Promise(
        (resolve, reject) => {
            client
                .api('/me/mailfolders/inbox/messages')
                .top(25)
                .filter('isRead eq false')
                .select('id', 'from', 'subject', 'bodyPreview', 'receivedDateTime')
                .get()
                .then(
                    (res) => { resolve(processMessages(res)); }
                ).catch(
                    (err) => {
                        reject(console.log(err));
                    }
                );
        });
}

function countMessagesPerDay(date) {
    return new Promise(
        (resolve, reject) => {
            let dateUTCISOString = toUTCISOString(date);
            console.log(dateUTCISOString);

            let nextUTCISOString = nextDayUTCISOString(date);

            console.log(nextUTCISOString);

            client
                .api('/me/mailfolders/inbox/messages')
                .headers({
                    'Prefer': 'outlook.timezone="Tokyo Standard Time"'
                })
                .top(25)
                .filter('receivedDateTime ge ' + dateUTCISOString + ' and receivedDateTime lt ' + nextUTCISOString)
                .count(true)
                .select('odata.count')
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

function retrieveMessagesPerDay(date) {
    return new Promise(
        (resolve, reject) => {
            let dateUTCISOString = toUTCISOString(date);
            let nextUTCISOString = nextDayUTCISOString(date);

            client
                .api('/me/mailfolders/inbox/messages')
                .headers({
                    'Prefer': 'outlook.timezone="Tokyo Standard Time"'
                })
                .top(25)
                .filter('receivedDateTime ge ' + dateUTCISOString + ' and receivedDateTime lt ' + nextUTCISOString)
                .select('id', 'from', 'subject', 'bodyPreview', 'receivedDateTime')
                .get()
                .then(
                    (res) => {
                        resolve(processMessages(res));
                    }
                )
                .catch(
                    (err) => { reject(console.log(err)) }
                );
        });
};

function toUTCISOString(date) {
    moment.tz.setDefault('Asia/Tokyo');
    let dateUTCISOString = moment(date).startOf('day').utc().format();

    return dateUTCISOString;
}

function nextDayUTCISOString(date) {
    moment.tz.setDefault('Asia/Tokyo');
    let nextDayUTCISOString = moment(date).add(1, 'days').startOf('day').utc().format();

    return nextDayUTCISOString;
}

function processMessages(rawMessages) {
    const messages = [];

    for (let rawMessage of rawMessages.value) {
        const message = {};

        message.id = rawMessage.id;
        message.from = rawMessage.from.emailAddress.name;
        message.subject = rawMessage.subject;
        message.body = rawMessage.bodyPreview;
        message.body = message.body.replace(/\r\n+/ig, '');
        message.body = message.body.replace(/--+/ig, '');
        message.body = message.body.replace(/  +/ig, '');
        message.received = rawMessage.receivedDateTime;

        messages.push(message);
    }
    return messages;
}

module.exports.setAccessToken = setAccessToken;
module.exports.countUnreadMessages = countUnreadMessages;
module.exports.retrieveUnreadMessages = retrieveUnreadMessages;
module.exports.countMessagesPerDay = countMessagesPerDay;
module.exports.retrieveMessagesPerDay = retrieveMessagesPerDay;