'use strict';
var Alexa = require("alexa-sdk");
var Speech = require('ssml-builder');
var moment = require('moment-timezone');
var languageStrings = require("./lang/languageStrings");
var accessToken;
var client = require("./lib/office365-rest-api-client");
var UnreadMailsCount;

exports.handler = function(event, context) {
    var alexa = Alexa.handler(event, context);
    alexa.appId = "amzn1.ask.skill.0e676bea-fc2e-4cb5-bbfa-e5ad4d220fea";
    alexa.resources = languageStrings;

    accessToken = event.session.user.accessToken;
    if (typeof accessToken !== 'undefined') {
        client.setAccessToken(accessToken);
    }

    alexa.registerHandlers(handlers);
    alexa.execute();
};

var handlers = {
    'LaunchRequest': function() {
        if (typeof accessToken === 'undefined') {
            this.emit(':tellWithLinkAccountCard', this.t("PLEASE_LINK_ACCOUNT"));
        }

        client.countUnreadMails()
            .then(
                (value) => {
                    this.UnreadMailsCount = value;

                    if (Number(this.UnreadMailsCount) > 0) {
                        this.emit(':ask', this.t("WELCOME_TO_VOICEMAIL") + this.t("THERE_ARE_UNREAD_MAILS", this.UnreadMailsCount));
                        // this.attributes["mode"] = "read_unread_mail";
                    } else {
                        this.emit(':ask', this.t("WELCOME_TO_VOICEMAIL") + this.t("NO_UNREAD_MAIL"));
                    }
                }
            )
            .catch(
                (error) => { console.log(error); }
            );
    },
    'UnReadMailIntent': function() {
        this.emit('UnReadMail');
    },
    'UnReadMail': function() {
        if (typeof accessToken === 'undefined') {
            this.emit(':tellWithLinkAccountCard', this.t("PLEASE_LINK_ACCOUNT"));
        }

        client.UnReadMails()
            .then(
                (value) => {
                    const mails = [];
                    let count = 0;
                    for (let mail of value) {
                        let mailresponse = buildMailResponse(++count, mail);
                        mails.push(mailresponse);
                    }
                    this.emit(':ask', mails.join(''));
                }
            )
            .catch(
                (error) => { console.log(error); }
            );
    },
    'AMAZON.YesIntent': function() {
        var mode = this.attributes['mode'];

        if (mode == "read_unread_mail") {
            client.UnReadMails()
                .then(
                    (value) => {
                        const mails = [];
                        let count = 0;
                        for (let mail of value) {
                            let mailresponse = buildMailResponse(++count, mail);
                            mails.push(mailresponse);
                        }
                        this.response.speak(mails.join(""));
                        this.emit(':responseReady');
                    }
                )
                .catch(
                    (error) => { console.log(error); }
                );
        } else {
            this.response.speak('other mode');
            this.emit(':responseReady');
        }
    },
    'SessionEndedRequest': function() {
        console.log('Session ended with reason: ' + this.event.request.reason);
    },
    'AMAZON.StopIntent': function() {
        this.response.speak(this.t('Bye'));
        this.emit(':responseReady');
    },
    'AMAZON.HelpIntent': function() {
        this.emit(':ask', this.t('HELP'))
    },
    'AMAZON.CancelIntent': function() {
        this.response.speak(this.t('Bye'));
        this.emit(':responseReady');
    },
    'Unhandled': function() {
        this.response.speak(this.t('unhandled'));
    }
};

function buildMailResponse(count, mail) {
    let speech = new Speech();
    speech.say(`${count}件目。`);
    let receivedDate = new Date(mail.received);
    speech.say(`${moment(receivedDate).format("M月D日 hh時mm分")} に受信。`);
    speech.pause('1s');
    speech.say(`件名は「${mail.subject}」で、`);
    speech.say(`本文の冒頭は次の通りです。${mail.body}`);
    speech.pause('1s');
    var response = speech.ssml(true);

    return response;
}