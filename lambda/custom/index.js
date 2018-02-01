'use strict';
var Alexa = require('alexa-sdk');
var Speech = require('ssml-builder');
var moment = require('moment-timezone');

var client = require('./lib/office365-rest-api-client');
var languageStrings = require('./lang/languageStrings');

var accessToken;

exports.handler = function(event, context) {
    var alexa = Alexa.handler(event, context);
    alexa.appId = 'amzn1.ask.skill.0e676bea-fc2e-4cb5-bbfa-e5ad4d220fea';
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
            this.emit(':tellWithLinkAccountCard', this.t('PLEASE_LINK_ACCOUNT'));
        }
        let unreadMessagesCount;

        client.countUnreadMessages()
            .then(
                (value) => {
                    unreadMessagesCount = value;

                    if (Number(unreadMessagesCount) > 0) {
                        this.emit(':ask', this.t('WELCOME_TO_VOICEMAIL') + this.t('THERE_ARE_UNREAD_MESSAGES', unreadMessagesCount), this.t('SAY_SOMETHING'));
                    } else {
                        this.emit(':ask', this.t('WELCOME_TO_VOICEMAIL') + this.t('NO_UNREAD_MESSAGES'), this.t('SAY_SOMETHING'));
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
            this.emit(':tellWithLinkAccountCard', this.t('PLEASE_LINK_ACCOUNT'));
        }
        let unreadMessagesCount;

        client.countUnreadMessages()
            .then(
                (value) => {
                    unreadMessagesCount = value;

                    if (Number(unreadMessagesCount) > 0) {
                        client.retrieveUnreadMessages()
                            .then(
                                (value) => {
                                    const messages = [];
                                    let count = 0;
                                    for (let message of value) {
                                        let messageResponse = buildMessageResponse(++count, message);
                                        messages.push(messageResponse);
                                    }
                                    this.emit(':ask', this.t('THERE_ARE_UNREAD_MESSAGES', unreadMessagesCount) + messages.join(''), this.t('SAY_SOMETHING'));
                                }
                            )
                            .catch(
                                (error) => { console.log(error); }
                            );
                    } else {
                        this.emit(':ask', this.t('NO_UNREAD_MESSAGES'), this.t('SAY_SOMETHING'));
                    }
                }
            )
            .catch(
                (error) => { console.log(error); }
            );
    },
    'ReadMails': function() {
        if (typeof accessToken === 'undefined') {
            this.emit(':tellWithLinkAccountCard', this.t('PLEASE_LINK_ACCOUNT'));
        }

        const intentObj = this.event.request.intent;
        let date = intentObj.slots.ReceivedDate.value;

        if (!moment(date).isValid()) {
            this.emit(':ask', this.t('INVALID_DATE'), this.t('SAY_SOMETHING'));
        }

        if (moment(date).isAfter()) {
            this.emit(':ask', this.t('FUTURE_DATE'), this.t('SAY_SOMETHING'));
        }

        let messagesCount;

        client.countMessagesPerDay(date)
            .then(
                (value) => {
                    messagesCount = value;
                    moment.locale('ja');

                    if (Number(messagesCount) > 0) {
                        client.retrieveMessagesPerDay(date)
                            .then(
                                (value) => {

                                    client.retrieveMessagesPerDay(date)
                                        .then(
                                            (value) => {
                                                const messages = [];
                                                let count = 0;
                                                for (let message of value) {
                                                    let messageResponse = buildMessageResponse(++count, message);
                                                    messages.push(messageResponse);
                                                }
                                                this.emit(':ask', this.t('THERE_ARE_MESSAGES', moment(date).format('ll'), messagesCount) + messages.join(''), this.t('SAY_SOMETHING'));
                                            }
                                        )
                                        .catch(
                                            (error) => { console.log(error); }
                                        );
                                }
                            )
                            .catch(
                                (error) => { console.log(error); }
                            );
                    } else {
                        this.emit(':ask', this.t('NO_MESSAGES', moment(date).format('ll')), this.t('SAY_SOMETHING'));
                    }
                }
            ).catch(
                (error) => { console.log(error); }
            );
    },
    'SessionEndedRequest': function() {
        console.log('Session ended with reason: ' + this.event.request.reason);
    },
    'AMAZON.StopIntent': function() {
        this.emit(':tell', this.t('BYE'));
    },
    'AMAZON.HelpIntent': function() {
        this.emit(':ask', this.t('HELP'))
    },
    'AMAZON.CancelIntent': function() {
        this.emit(':ask', this.t('CANCEL'));
    },
    'Unhandled': function() {
        this.emit(':ask', this.t('unhandled'))
    }
};

function buildMessageResponse(count, message) {
    moment.locale('ja');
    let speech = new Speech();

    speech.say(`${count}通目。`);
    let receivedDate = moment(message.received).format('M月D日 hh時mm分');

    speech.say(`${receivedDate} に受信。`);
    speech.pause('1s');
    speech.say(`件名は「${message.subject}」で、`);
    speech.say(`本文の冒頭は次の通りです。${message.body}`);
    speech.pause('1s');

    var response = speech.ssml(true);

    return response;
}