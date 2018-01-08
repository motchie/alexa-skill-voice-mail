'use strict';
var Alexa = require("alexa-sdk");
var languageStrings = require("./lang/languageStrings");
var accessToken;
var client = require("./lib/office365-rest-api-client");
var unReadMailCount;

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

        client.unReadMailCount()
            .then(
                (value) => {
                    this.unReadMailCount = value;
                    if (Number(this.unReadMailCount) > 0) {
                        this.emit(':tell', this.t("WELCOME_TO_VOICEMAIL") + this.t("THERE_ARE") + this.unReadMailCount + this.t("UNREAD_MAILS"));
                    } else {
                        this.emit(':tell', this.t("WELCOME_TO_VOICEMAIL") + this.t("NO_UNREAD_MAIL"));
                    }
                    this.emit(':responseReady');
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

        client.unReadMailCount()
            .then(
                (value) => {
                    this.unReadMailCount = value;
                    this.response.speak(this.t("HELLO") + '未読メールは' + this.unReadMailCount + '通です。')
                        .cardRenderer('未読メールは', this.unReadMailCount + '通です。');
                    this.emit(':responseReady');
                }
            )
            .catch(
                (error) => { console.log(error); }
            );
    },
    'SessionEndedRequest': function() {
        console.log('Session ended with reason: ' + this.event.request.reason);
    },
    'AMAZON.StopIntent': function() {
        this.response.speak('Bye');
        this.emit(':responseReady');
    },
    'AMAZON.HelpIntent': function() {
        this.response.speak("You can try: 'alexa, hello world' or 'alexa, ask hello world my" +
            " name is awesome Aaron'");
        this.emit(':responseReady');
    },
    'AMAZON.CancelIntent': function() {
        this.response.speak('Bye');
        this.emit(':responseReady');
    },
    'Unhandled': function() {
        this.response.speak("Sorry, I didn't get that. You can try: 'alexa, hello world'" +
            " or 'alexa, ask hello world my name is awesome Aaron'");
    }
};