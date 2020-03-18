// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ChoicePrompt, DialogSet, DialogTurnStatus, OAuthPrompt, TextPrompt, DateTimePrompt, WaterfallDialog } = require('botbuilder-dialogs');
const { LogoutDialog } = require('./logoutDialog');
const { OAuthHelpers } = require('../oAuthHelpers');

const MAIN_WATERFALL_DIALOG = 'mainWaterfallDialog';
const OAUTH_PROMPT = 'oAuthPrompt';
const CHOICE_PROMPT = 'choicePrompt';
const MEET_WATERFALL_DIALOG = 'meetWaterfallDialog';　
const TEXT_PROMPT = 'textPrompt';

class MainDialog extends LogoutDialog {
    constructor() {
        // super('MainDialog');
        super('MainDialog', process.env.connectionName);
        this.addDialog(new ChoicePrompt(CHOICE_PROMPT))
            .addDialog(new OAuthPrompt(OAUTH_PROMPT, {
                connectionName: process.env.ConnectionName,
                text: 'Please login',
                title: 'Login',
                timeout: 300000
            }))
            .addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
                this.promptStep.bind(this),
                this.loginStep.bind(this),
                this.commandStep.bind(this),
                this.processStep.bind(this)
            ]))
            .addDialog(new WaterfallDialog(MEET_WATERFALL_DIALOG, [
                async (step) => {
                    // Ask the meeting suject
                    return await step.prompt('sujectPrompt', `Meeting's suject is?`);
                },
                async (step) => {
                    // Remember the meeting suject
                    //step.values['subject'] = step.result;
                    step.stack[0].state.values['subject'] = step.result;
                    // Ask the meeting's content
                    return await step.prompt('textPrompt', `Meeting's content is?`);
                },
                async (step) => {
                    // Remember the meeting content
                    //step.values['subject'] = step.result;
                    step.stack[0].state.values['content'] = step.result;
                    // Ask the meeting's start time
                    return await step.prompt('startTimePrompt', `Meeting's start time is?`);
                },
                async (step) => {
                    // Remember the meeting start time
                    //step.values['startTime'] = step.result;
                    step.stack[0].state.values['startTime'] = step.result;
                    // Ask the meeting's end time
                    return await step.prompt('endTimePrompt', `Meeting's end time is?`);
                },
                async (step) => {
                     // Remember the meeting end time
                     //step.values['endTime'] = step.result;
                     step.stack[0].state.values['endTime'] = step.result;
                     // Ask the meeting's room
                     return await step.prompt('textPrompt', `Meeting's room is?`);
                },
                async (step) => {
                     // Remember the meeting localtion
                     step.stack[0].state.values['room'] = step.result;
                     // Ask the meeting's participants
                     return await step.prompt('textPrompt', `Meeting's participants have?(use ',' to separate)`);
                },
                async (step) => {
                    // Remember the meeting participants
                    step.stack[0].state.values['participants'] = step.result;
                    return await step.beginDialog(OAUTH_PROMPT);
                }
            ]));

        // Add prompts
        this.addDialog(new TextPrompt('sujectPrompt'));
        this.addDialog(new DateTimePrompt('startTimePrompt'));
        this.addDialog(new DateTimePrompt('endTimePrompt'));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
    }

    /**
     * The run method handles the incoming activity (in the form of a TurnContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {*} turnContext
     * @param {*} accessor
     */
    async run(turnContext, accessor) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(turnContext);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    async promptStep(step) {
        return step.beginDialog(OAUTH_PROMPT);
    }

    async loginStep(step) {
        // Get the token from the previous step. Note that we could also have gotten the
        // token directly from the prompt itself. There is an example of this in the next method.
        const tokenResponse = step.result;
        if (tokenResponse) {
            await step.context.sendActivity('You are now logged in.');
            return await step.prompt(TEXT_PROMPT, { prompt: 'Would you like to do? (type \'me\', \'send <EMAIL>\', \'recent\',\'rooms\',\'meeting\'  or \'schedule\')' });
        }
        await step.context.sendActivity('Login was not successful please try again.');
        return await step.endDialog();
    }

    async commandStep(step) {
 
        // Call the prompt again because we need the token. The reasons for this are:
        // 1. If the user is already logged in we do not need to store the token locally in the bot and worry
        // about refreshing it. We can always just call the prompt again to get the token.
        // 2. We never know how long it will take a user to respond. By the time the
        // user responds the token may have expired. The user would then be prompted to login again.
        //
        // There is no reason to store the token locally in the bot because we can always just call
        // the OAuth prompt to get the token or get a new token if needed.
        if (step.result) {

            // If we have the token use the user is authenticated so we may use it to make API calls.
            const parts = step.result.toLowerCase().split(' ');
            if (Array.isArray(parts)) {
                for (let cnt = 0; cnt < parts.length; cnt++) {
                    const command = parts[cnt];

                    switch (command) {
                    case 'meeting':
                        step.values['command'] = step.result;
                        return await step.beginDialog(MEET_WATERFALL_DIALOG);
                    default:
                        step.values['command'] = step.result;
                        return await step.beginDialog(OAUTH_PROMPT);
                    }
                }
            }
        } else {
            await step.context.sendActivity('We couldn\'t log you in. Please try again later.');
        }
        //return await step.beginDialog(OAUTH_PROMPT);
    }

    async processStep(step) {
        if (step.result) {
            // We do not need to store the token in the bot. When we need the token we can
            // send another prompt. If the token is valid the user will not need to log back in.
            // The token will be available in the Result property of the task.
            const tokenResponse = step.result;

            // If we have the token use the user is authenticated so we may use it to make API calls.
            if (tokenResponse && tokenResponse.token) {
                const parts = (step.values['command'] || '').toLowerCase().split(' ');
                if (Array.isArray(parts)) {
                    for (let cnt = 0; cnt < parts.length; cnt++) {
                        const command = parts[cnt];

                        switch (command) {
                        case 'me':
                            await OAuthHelpers.listMe(step.context, tokenResponse);
                            break;
                        case 'send':
                            await OAuthHelpers.sendMail(step.context, tokenResponse, parts[1]);
                            break;
                        case 'recent':
                            await OAuthHelpers.listRecentMail(step.context, tokenResponse);
                            break;
                        case 'rooms':
                            await OAuthHelpers.getFindRooms(step.context, tokenResponse);
                            break;
                        case 'schedule':
                            await OAuthHelpers.getSchedule(step.context, tokenResponse);
                            break;
                        case 'event':
                            await OAuthHelpers.getEvents(step.context, tokenResponse);
                            break;
                        case 'meeting':
                            var options = {
                                subject : step.values['subject'],
                                content : step.values['content'],
                                startTime : step.values['startTime'],
                                endTime : step.values['endTime'],
                                room : step.values['room'],
                                participants : step.values['participants'],
                                organizer : ''
                            }
                            await OAuthHelpers.addEvents(step.context, tokenResponse, options);
                            break;
                        default:
                            //await step.context.sendActivity(`Your token is ${ tokenResponse.token }`);
                            break;
                        }
                    }
                }
            }
            return await step.endDialog();
        } else {
            await step.context.sendActivity('We couldn\'t log you in. Please try again later.');
        }  
    }
}

module.exports.MainDialog = MainDialog;
