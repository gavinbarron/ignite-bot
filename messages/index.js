// For more information about this template visit http://aka.ms/azurebots-node-qnamaker

"use strict";
var builder = require("botbuilder");
var botbuilder_azure = require("botbuilder-azure");
var builder_cognitiveservices = require("botbuilder-cognitiveservices");
var path = require('path');
const appInsights = require("applicationinsights");
appInsights.setup().start(); // assuming ikey in env var
const client = appInsights.defaultClient;
client.config.endpointUrl = "https://dc.services.visualstudio.com/v2/track";
var useEmulator = (process.env.NODE_ENV == 'development');

var connector = useEmulator ? new builder.ChatConnector() : new botbuilder_azure.BotServiceConnector({
    appId: process.env['MicrosoftAppId'],
    appPassword: process.env['MicrosoftAppPassword'],
    stateEndpoint: process.env['BotStateEndpoint'],
    openIdMetadata: process.env['BotOpenIdMetadata']
});

var bot = new builder.UniversalBot(connector, function(session) {
    var msg = new builder.Message(session);
    session.beginDialog("qna");
});

bot.localePath(path.join(__dirname, './locale'));

var recognizer = new builder_cognitiveservices.QnAMakerRecognizer({
                knowledgeBaseId: process.env.QnAKnowledgebaseId,
                subscriptionKey: process.env.QnASubscriptionKey,
                top: 4});

var basicQnAMakerDialog = new builder_cognitiveservices.QnAMakerDialog({
    recognizers: [recognizer],
                defaultMessage: 'Sorry, I\'m a bot and I\'m still learning.',
                qnaThreshold: 0.3,
                feedbackLib: qnaMakerTools}
);

var qnaMakerTools = new builder_cognitiveservices.QnAMakerTools();
bot.library(qnaMakerTools.createLibrary());

// Override to also include the knowledgebase question with the answer on confident matches
basicQnAMakerDialog.respondFromQnAMakerResult = function(session, qnaMakerResult){
	var result = qnaMakerResult;
	var response = 'Here is the match from FAQ:  \r\n  Q: ' + result.answers[0].questions[0] + '  \r\n A: ' + result.answers[0].answer;
	session.send(response);
};

// Override to log user query and matched Q&A before ending the dialog
basicQnAMakerDialog.defaultWaitNextMessage = function(session, qnaMakerResult){
    if(session.privateConversationData.qnaFeedbackUserQuestion != null) {
        console.log('User Query: ' + session.privateConversationData.qnaFeedbackUserQuestion);
        client.trackEvent({name: 'bot-question-asked', properties: {qnaQuestion: session.privateConversationData.qnaFeedbackUserQuestion}});

        if(qnaMakerResult.answers != null && qnaMakerResult.answers.length > 0
		&& qnaMakerResult.answers[0].questions != null && qnaMakerResult.answers[0].questions.length > 0 && qnaMakerResult.answers[0].answer != null){
			console.log('KB Question: ' + qnaMakerResult.answers[0].questions[0]);
            console.log('KB Answer: ' + qnaMakerResult.answers[0].answer);
        }
    }
	session.endDialog();
};

bot.dialog('qna', basicQnAMakerDialog);

if (useEmulator) {
    var restify = require('restify');
    var server = restify.createServer();
    server.listen(3978, function() {
        console.log('test bot endpont at http://localhost:3978/api/messages');
    });
    server.post('/api/messages', connector.listen());
} else {
    module.exports = { default: connector.listen() }
}
