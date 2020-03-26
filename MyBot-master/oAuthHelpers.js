// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { AttachmentLayoutTypes, CardFactory } = require('botbuilder');
const { SimpleGraphClient } = require('./simple-graph-client');

/**
 * These methods call the Microsoft Graph API. The following OAuth scopes are used:
 * 'OpenId' 'email' 'Mail.Send.Shared' 'Mail.Read' 'profile' 'User.Read' 'User.ReadBasic.All'
 * for more information about scopes see:
 * https://developer.microsoft.com/en-us/graph/docs/concepts/permissions_reference
 */
class OAuthHelpers {
    /**
     * Enable the user to send an email via the bot.
     * @param {TurnContext} context A TurnContext instance containing all the data needed for processing this conversation turn.
     * @param {TokenResponse} tokenResponse A response that includes a user token.
     * @param {string} emailAddress The email address of the recipient.
     */
    static async sendMail(context, tokenResponse, emailAddress) {
        if (!context) {
            throw new Error('OAuthHelpers.sendMail(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.sendMail(): `tokenResponse` cannot be undefined.');
        }

        const client = new SimpleGraphClient(tokenResponse.token);
        const me = await client.getMe();

        await client.sendMail(
            emailAddress,
            'Message from a bot!',
            `Hi there! I had this message sent from a bot. - Your friend, ${me.displayName}`
        );
        await context.sendActivity(`I sent a message to ${emailAddress} from your account.`);
    }

    static async getSchedule(context, tokenResponse) {
        if (!context) {
            throw new Error('OAuthHelpers.getSchedule(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.getSchedule(): `tokenResponse` cannot be undefined.');
        }

        const client = new SimpleGraphClient(tokenResponse.token);
        const me = await client.getMe();
        const schedule = await client.getSchedule(me.mail) || '';
        var scheduleInfo = '';
        var moment = require('moment');
        if (schedule != '' && schedule.value.length > 0 && schedule.value[0].scheduleItems.length > 0) {
            for (let cnt = 0; cnt < schedule.value[0].scheduleItems.length; cnt++) {
                const scheduleItem = schedule.value[0].scheduleItems[cnt];
                if (scheduleItem.status == 'busy') {
                    if (scheduleInfo == '') {
                        scheduleInfo = 'Subject: ' + scheduleItem.subject + '\r\n';
                    } else {
                        scheduleInfo += 'Subject: ' + scheduleItem.subject + '\r\n';
                    }
                    scheduleInfo += 'Location: ' + scheduleItem.location + '\r\n';
                    var localeString = moment.parseZone(scheduleItem.start.dateTime).local().format('YYYY-MM-DD HH:mm:ss');
                    scheduleInfo += 'StartTime: ' + localeString + '\r\n';
                    localeString = moment.parseZone(scheduleItem.end.dateTime).local().format('YYYY-MM-DD HH:mm:ss');
                    scheduleInfo += 'EndTime: ' + localeString + '\r\n';
                }
            }
        }
        if (scheduleInfo != '') {
            scheduleInfo = 'Schedule information:\r\n' + scheduleInfo;
        } else {
            scheduleInfo = 'There are not Schedule information.';
        }

        await context.sendActivity(scheduleInfo);
    }

    static async getFindRooms(context, tokenResponse) {
        if (!context) {
            throw new Error('OAuthHelpers.getFindRooms(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.getFindRooms(): `tokenResponse` cannot be undefined.');
        }
        // Pull in the data from Microsoft Graph.
        const client = new SimpleGraphClient(tokenResponse.token);
        const findRooms = await client.getFindRooms() || '';

        // await context.sendActivity(`find rooms: ${ JSON.stringify(findRooms) }`);
        var roomMessage = '';
        for (let cnt = 0; cnt < findRooms.value.length; cnt++) {
            const room = findRooms.value[cnt];

            if (room.address != null && ('SH' == room.address.substr(0, 2)
                || 'DL' == room.address.substr(0, 2))) {
                var local = '';
                if ('SH' == room.address.substr(0, 2)) {
                    local = 'ShangHai';
                } else if ('DL' == room.address.substr(0, 2)) {
                    local = 'DaLian';
                }
                if (roomMessage == '') {
                    roomMessage = 'rooms[' + local + ']: ' + '\r\nname: ' + room.name + '\r\naddress: ' + room.address;
                } else {
                    roomMessage += '\r\nrooms[' + local + ']: ' + '\r\nname: ' + room.name + '\r\naddress: ' + room.address;
                }
            }
        }
        await context.sendActivity(roomMessage);
    }

    static async getEvents(context, tokenResponse) {
        if (!context) {
            throw new Error('OAuthHelpers.getFindRooms(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.getFindRooms(): `tokenResponse` cannot be undefined.');
        }
        // Pull in the data from Microsoft Graph.
        const client = new SimpleGraphClient(tokenResponse.token);
        const events = await client.getEvents() || '';
        await context.sendActivity(`Events: ${JSON.stringify(events)}`);
    }

    /**
     * Displays informau'r'ntion about the user in the bot.
     * @param {TurnContext} context A TurnContext instance containing all the data needed for processing this conversation turn.
     * @param {TokenResponse} tokenResponse A response that includes a user token.
     * @param {var} options A response that includes a user token.
     */
    static async addEvents(context, tokenResponse, options) {
        if (!context) {
            throw new Error('OAuthHelpers.getFindRooms(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.getFindRooms(): `tokenResponse` cannot be undefined.');
        }
        // Pull in the data from Microsoft Graph.
        const client = new SimpleGraphClient(tokenResponse.token);
        const me = await client.getMe();
        options['organizer'] = me.mail;
        const contacts = await client.getContacts();
        const parts = options['participants'].split(',');
        var mailaddress = '';
        if (Array.isArray(parts)) {
            var contactsList = '';
            for (let cnt = 0; cnt < parts.length; cnt++) {
                const name = parts[cnt];
                var isfind = false;
                for (let index = 0; index < contacts.value.length; index++) {
                    const person = contacts.value[index];
                    if (cnt == 0) {
                        if (contactsList == '') {
                            contactsList = person.displayName;
                        } else {
                            contactsList += ',' + person.displayName;
                        }
                    }
                    if (person.displayName == name) {
                        if (mailaddress == '') {
                            mailaddress = person.emailAddresses[0].address;
                        } else {
                            mailaddress += ',' + person.emailAddresses[0].address;
                        }
                        isfind = true;
                    }
                }
                if (!isfind) {
                    return await context.sendActivity(`Participant(${JSON.stringify(name)}) is not exist. Contacts List(${JSON.stringify(contactsList)})`);
                }
            }
        }
        options['participants'] = mailaddress;
        const rooms = await client.getRooms("SH-TH") || '';
        if (options['room'] != '') {
            var isfind = false;
            for (let cnt = 0; cnt < rooms.value.length; cnt++) {
                const room = rooms.value[cnt];
                if (room.givenName == options['room']) {
                    isfind = true;
                    const schedule = await client.getRoomSchedule(room.id, room.mail, options['startTime'], options['endTime']) || '';
                    var moment = require('moment');
                    if (schedule != '' && schedule.value.length > 0 && schedule.value[0].scheduleItems.length > 0) {
                        for (let cnt = 0; cnt < schedule.value[0].scheduleItems.length; cnt++) {
                            const scheduleItem = schedule.value[0].scheduleItems[cnt];
                            if (scheduleItem.status == 'busy') {
                                var startTime = moment.parseZone(scheduleItem.start.dateTime).local().format('YYYY-MM-DD HH:mm:ss');
                                var endTime = moment.parseZone(scheduleItem.end.dateTime).local().format('YYYY-MM-DD HH:mm:ss');
                                return await context.sendActivity(`Room(${JSON.stringify(room.name)}) is busy at ` + startTime + ` to ` + endTime);
                            }
                        }
                    }
                }
            }
            if (!isfind) {
                return await context.sendActivity(`Room(` + options['room'] + `) is not exist.`);
            }
        }
        const events = await client.addEvents(options) || '';
        await context.sendActivity(`Events: ${JSON.stringify(events)}`);
    }

    /**
     * Displays informau'r'ntion about the user in the bot.
     * @param {TurnContext} context A TurnContext instance containing all the data needed for processing this conversation turn.
     * @param {TokenResponse} tokenResponse A response that includes a user token.
     */
    static async listMe(context, tokenResponse) {
        if (!context) {
            throw new Error('OAuthHelpers.listMe(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.listMe(): `tokenResponse` cannot be undefined.');
        }
        // Pull in the data from Microsoft Graph.
        const client = new SimpleGraphClient(tokenResponse.token);
        const me = await client.getMe();
        const manager = await client.getManager();

        await context.sendActivity(`You are ${me.displayName} and you report to ${manager.displayName}.`);
    }

    /**
     * Lists the user's collected email.
     * @param {TurnContext} context A TurnContext instance containing all the data needed for processing this conversation turn.
     * @param {TokenResponse} tokenResponse A response that includes a user token.
     */
    static async listRecentMail(context, tokenResponse) {
        if (!context) {
            throw new Error('OAuthHelpers.listRecentMail(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.listRecentMail(): `tokenResponse` cannot be undefined.');
        }

        var client = new SimpleGraphClient(tokenResponse.token);
        var messages = await client.getRecentMail();
        if (Array.isArray(messages)) {
            let numberOfMessages = 0;
            if (messages.length > 5) {
                numberOfMessages = 5;
            }
            const reply = { attachments: [], attachmentLayout: AttachmentLayoutTypes.Carousel };
            for (let cnt = 0; cnt < numberOfMessages; cnt++) {
                const mail = messages[cnt];
                const card = CardFactory.heroCard(
                    mail.subject,
                    mail.bodyPreview,
                    [{ alt: 'Outlook Logo', url: 'https://botframeworksamples.blob.core.windows.net/samples/OutlookLogo.jpg' }],
                    [],
                    { subtitle: `${mail.from.emailAddress.name} <${mail.from.emailAddress.address}>` }
                );
                reply.attachments.push(card);
            }
            await context.sendActivity(reply);
        } else {
            await context.sendActivity('Unable to find any recent unread mail.');
        }
    }

    /**
 * Displays information about the user in the bot.
 * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
 * @param {TokenResponse} tokenResponse A response that includes a user token.
 */
    static async loginData(turnContext, tokenResponse) {
        if (!turnContext) {
            throw new Error('OAuthHelpers.loginData(): `turnContext` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.loginData(): `tokenResponse` cannot be undefined.');
        }

        try {
            // Pull in the data from Microsoft Graph.
            const client = new SimpleGraphClient(tokenResponse.token);
            const me = await client.getMe();
            const photoResponse = await client.getPhoto();

            // Attaches user's profile photo to the reply activity.
            if (photoResponse != null) {
                let replyAttachment;
                //const base64 = Buffer.from(photoResponse, 'binary').toString('base64');
                //let base64 = Buffer.from(photoResponse).toString('base64');
                //photoResponse.toString('base64');
                //var photoBytes = ReadAsBytes(photoResponse);
                //var result = Convert.ToBase64String(photoBytes);
                //var base64 = new Buffer(photoResponse, 'base64');
                //var base64 = Buffer.from(photoResponse.getBytes(), 'binary').toString('base64');
                replyAttachment = {
                    contentType: 'image/jpeg',
                    contentUrl: `data:image/jpeg;base64,${ null }`
                };
                replyAttachment.displayName = me.displayName;
                return (replyAttachment);
            }
        } catch (error) {
            throw error;
        }
    }
}

exports.OAuthHelpers = OAuthHelpers;
