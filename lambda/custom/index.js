'use strict';
var Alexa = require("alexa-sdk");

exports.handler = function(event, context) {
    var alexa = Alexa.handler(event, context);
    alexa.registerHandlers(handlers);
    alexa.execute();
};

var handlers = {
    'LaunchRequest': function() {
        this.emit('音声メールへようこそ。');
    },
    'UnReadMailIntent': function() {
        this.emit('UnReadMail');
    },
    'UnReadMail': function() {
        let accessToken = this.event.session.user.accessToken;
        console.log(accessToken);
        this.response.speak('未読メールをチェックします。')
            .cardRenderer('hello world', 'hello world');
        this.emit(':responseReady');
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