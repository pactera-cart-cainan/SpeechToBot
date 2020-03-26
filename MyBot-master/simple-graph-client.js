// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { Client } = require('@microsoft/microsoft-graph-client');

/**
 * This class is a wrapper for the Microsoft Graph API.
 * See: https://developer.microsoft.com/en-us/graph for more information.
 */
class SimpleGraphClient {
    constructor(token) {
        if (!token || !token.trim()) {
            throw new Error('SimpleGraphClient: Invalid token received.');
        }

        this._token = token;

        // Get an Authenticated Microsoft Graph client using the token issued to the user.
        this.graphClient = Client.init({
            authProvider: (done) => {
                done(null, this._token); // First parameter takes an error if you can't get an access token.
            }
        });
    }

    /**
     * Sends an email on the user's behalf.
     * @param {string} toAddress Email address of the email's recipient.
     * @param {string} subject Subject of the email to be sent to the recipient.
     * @param {string} content Email message to be sent to the recipient.
     */
    async sendMail(toAddress, subject, content) {
        if (!toAddress || !toAddress.trim()) {
            throw new Error('SimpleGraphClient.sendMail(): Invalid `toAddress` parameter received.');
        }
        if (!subject || !subject.trim()) {
            throw new Error('SimpleGraphClient.sendMail(): Invalid `subject`  parameter received.');
        }
        if (!content || !content.trim()) {
            throw new Error('SimpleGraphClient.sendMail(): Invalid `content` parameter received.');
        }

        // Create the email.
        const mail = {
            body: {
                content: content, // `Hi there! I had this message sent from a bot. - Your friend, ${ graphData.displayName }!`,
                contentType: 'Text'
            },
            subject: subject, // `Message from a bot!`,
            toRecipients: [{
                emailAddress: {
                    address: toAddress
                }
            }]
        };

        // Send the message.
        return await this.graphClient
            .api('/me/sendMail')
            .post({ message: mail }, (error, res) => {
                if (error) {
                    throw error;
                } else {
                    return res;
                }
            });
    }

    /**
     * Gets recent mail the user has received within the last hour and displays up to 5 of the emails in the bot.
     */
    async getRecentMail() {
        return await this.graphClient
            .api('/me/messages')
            .version('beta')
            .top(5)
            .get().then((res) => {
                return res;
            });
    }

    /**
     * Collects information about the user in the bot.
     */
    async getMe() {
        return await this.graphClient
            .api('/me')
            .get().then((res) => {
                return res;
            });
    }

    /**
     * Collects the user's manager in the bot.
     */
    async getManager() {
        return await this.graphClient
            .api('/me/manager')
            .version('beta')
            .select('displayName')
            .get().then((res) => {
                return res;
            });
    }

    /**
     * @param {string} mailAddress Email address of the email's recipient.
     */
    async getSchedule(mailAddress) {
        const startDate = new Date();
        const endDate = new Date();
        endDate.setDate(startDate.getDate() + 1);
        endDate.setHours(0);
        endDate.setMinutes(0);
        endDate.setSeconds(0);
        endDate.setMilliseconds(0);
        const scheduleInformation = {
            schedules: [mailAddress],
            startTime: {
                dateTime: startDate.toJSON(),
                timeZone: "Pacific Standard Time"
            },
            endTime: {
                dateTime: endDate.toJSON(),
                timeZone: "Pacific Standard Time"
            },

            availabilityViewInterval: 60
        };

        let res = await this.graphClient.api('/me/calendar/getSchedule')
            .version('beta')
            .post(scheduleInformation);

        return res;
    }

    /**
     * @param {string} id room id.
     * @param {string} mailAddress Email address
     * @param {string} startTime start time.
     * @param {string} endtime end time.
     */
    async getRoomSchedule(id, mailAddress, startTime, endTime) {
        //options['startTime'];
        var date1 = new Date().getFullYear() + "-" + (new Date().getMonth() + 1) + "-" + new Date().getDate();
        var temp = date1 + ' ' + startTime;
        const startDate = new Date(Date.parse(temp));
        //options['endTime'];
        temp = date1 + ' ' + endTime;
        const endDate = new Date(Date.parse(temp));
        const scheduleInformation = {
            schedules: [mailAddress],
            startTime: {
                dateTime: startDate.toJSON(),
                timeZone: "China Standard Time"
            },
            endTime: {
                dateTime: endDate.toJSON(),
                timeZone: "China Standard Time"
            },

            availabilityViewInterval: 30
        };
        var roomSchedule = "/users/" + id + "/calendar/getSchedule";
        let res = await this.graphClient.api(roomSchedule)
            .version('v1.0')
            .post(scheduleInformation);

        return res;
    }

    async getFindRooms() {
        return await this.graphClient
            .api('/me/findRooms')
            .version('beta')
            .get().then((res) => {
                return res;
            });
    }

    /**
     * @param {string} roomName Rooms name.
     */
    async getRooms(roomName) {
        var filers = "startswith(givenName," + "'" + roomName + "'" + ")";
        return await this.graphClient
            .api('/users')
            .version('v1.0')
            .filter(filers)
            .get().then((res) => {
                return res;
            });
    }

    async getEvents() {
        return await this.graphClient
            .api('/me/events')
            .header('Prefer', 'outlook.timezone="Pacific Standard Time"')
            .version('beta')
            .select('subject,organizer,attendees,start,end,location')
            .get().then((res) => {
                return res;
            });
    }

    async getContacts() {
        return await this.graphClient
            .api('/me/contacts')
            .version('beta')
            .select('displayName,emailAddresses')
            .get().then((res) => {
                return res;
            });
    }

    /**
     * @param {var} options
     */
    async addEvents(options) {

        //options['startTime'];
        var date1 = new Date().getFullYear() + "-" + (new Date().getMonth() + 1) + "-" + new Date().getDate();
        var temp = date1 + ' ' + options['startTime'][0].value;
        const startDate = new Date(Date.parse(temp));
        //options['endTime'];
        temp = date1 + ' ' + options['endTime'][0].value;
        const endDate = new Date(Date.parse(temp));
        const event = {
            subject: options['subject'],
            body: {
                contentType: "HTML",
                content: options['content']
            },
            start: {
                //dateTime: options['startTime'],
                dateTime: startDate.toJSON(),
                timeZone: "China Standard Time"
            },
            end: {
                //dateTime: options['endTime'],
                dateTime: endDate.toJSON(),
                timeZone: "China Standard Time"
            },
            location: {
                displayName: options['room']
            },
            attendees: [],
            organizer: {
                emailAddress: {
                    //name: "cainan",
                    address: options['organizer']
                }
            }
        };
        const parts = options['participants'].split(',');
        if (Array.isArray(parts)) {
            for (let cnt = 0; cnt < parts.length; cnt++) {
                const name = parts[cnt];
                var temp = {
                    emailAddress: {
                        address: name
                    },
                    type: "required"
                }
                event.attendees.push(temp);
            }
        }
        let res = await this.graphClient
            .api('/me/events')
            .header('Prefer', 'outlook.timezone="Pacific Standard Time"')
            .post(event);
        return res;
    }

    /**
    * Collects the user's photo.
    */
    async getPhoto() {
        return await this.graphClient
            .api('/me/photo/$value')
            .responseType('ArrayBuffer')
            .version('beta')
            .get()
            .then((res) => {
                return res;
            })
            .catch((err) => {
                console.log(err);
            }); 
    }
}

exports.SimpleGraphClient = SimpleGraphClient;
