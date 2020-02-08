import { BotDeclaration, IBot } from "express-msteams-host";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import WelcomeCard from "./dialogs/WelcomeDialog";
import { TeamsContext, TeamsActivityProcessor, TeamsAdapter, TeamsChannelAccount } from "botbuilder-teams";
const fetch = require('node-fetch');

/**
 * Implementation for Power Pillager
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD
)

export class PowerPillager implements IBot {
    private readonly conversationState: ConversationState;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;
    private readonly activityProc = new TeamsActivityProcessor();
    
    private commands: string[] = [ 'king', 'me', 'stats', 'help', 'man', '?', 'wagewar', 'war' ];
    private async messageHandler(text: string, context: TurnContext, sender: TeamsChannelAccount): Promise<void> { 
        let args: string[] = text.trim().split(' ');
        const command: string = args[0].toLocaleLowerCase();
        if (this.commands.indexOf(command) !== -1) {
            switch(command) {
                case 'me':
                case 'stats':
                case 'king': {
                    let request = { email: sender.email.toLowerCase() };
                    if (args.indexOf('--user') !== -1) {
                        const arg = args[args.indexOf('--user') + 1]
                        if (arg) request.email = arg;
                    }

                    let kings: any;
                    try {
                        const response = await fetch(
                            `https://pillagers-storage-functions.azurewebsites.net/api/GetKing?email=${request.email}`,
                            { method: 'GET',  headers: { 'Content-Type': 'application/json' } }
                        );
                        kings = await (response as any).json();
                    } catch(e) {
                        this.onError(e, context);
                        console.error('### error (fetch azure):', e);
                    }

                    if (kings.value[0]) {
                        try {
                            const king = kings.value[0];
                            const kingName: string = this.capitalizeWords(`${king.FirstName} ${king.LastName}`);
                            await context.sendActivity({
                                type: 'message',
                                attachments: [
                                    {
                                        contentType: 'application/vnd.microsoft.card.adaptive',
                                        content: {
                                            type: 'AdaptiveCard',
                                            version: '1.0',
                                            body: [
                                                { type: 'TextBlock', text: `King Stats for ${king.email}`, size: "Large", weight: 'Bolder' },
                                                { type: 'TextBlock', text: `Name: ${kingName}`, size: 'Small' },
                                                { type: 'TextBlock', text: `Pennings: ${king.Penning}`, size: 'Small' },
                                                { type: 'TextBlock', text: `Latitude: ${king.lat}`, size: 'Small' },
                                                { type: 'TextBlock', text: `Longitude: ${king.lon}`, size: 'Small' }
                                            ],
                                            actions: [{ type: 'Action.Submit', title: 'Cheat - Add Pennings', data: { addMoney: 1000000000, king: king.email } }]
                                        }
                                    }
                                ]
                            });
                        } catch(e) {
                            this.onError(e, context);
                            console.error('### error (adaptiveCard):', e);
                        }
                    } else {
                        await context.sendActivity(`Cannot find a user registred with: <i>${request.email}</i>, registrer at <a href='http://pillagers.no'>pillagers.no<a/>.`)
                    }
                    return;
                }
                case '?':
                case 'man':
                case 'help': {
                    await context.sendActivity('Actions: man, help, ?, status, me, king, war, wagewar');
                    return;
                }
                case 'war':
                case 'wagewar': {
                    const senderKingEmail: string = sender.email.toLowerCase();

                    // TODO fetch from get all kings
                    const kings: any[] = (await this.getKings(context)).map((king: any) => {
                        return { 
                            name: this.capitalizeWords(`${king.FirstName} ${king.LastName}`), 
                            email: king.email 
                        }
                    });

                    if (kings.length <= 0) {
                        await context.sendActivity('Cannot find any enemy kings!');
                        return;
                    }

                    const actions = kings.map((item) => {
                        return { 
                            type: 'Action.Submit', title: item.name, iconUrl: "https://cdn0.iconfinder.com/data/icons/material-style/48/crown-512.png", 
                            data: { targetKingEmail: item.email, targetKingName: item.name } 
                        };
                    });

                    const response = await fetch(
                        `https://pillagers-storage-functions.azurewebsites.net/api/GetKing?email=${senderKingEmail}`,
                        { method: 'GET',  headers: { 'Content-Type': 'application/json' } }
                    );
                    let json = await (response as any).json();
                    if (json.value) {
                        const senderKing = json.value[0];
                        try {
                            await context.sendActivity({
                                type: 'message',
                                attachments: [
                                    {
                                        contentType: 'application/vnd.microsoft.card.adaptive',
                                        content: {
                                            type: 'AdaptiveCard',
                                            version: '1.0',
                                            body: [
                                                { type: 'TextBlock', text: `Wage War`, size: "Large", color: 'Attention', weight: 'Bolder' },
                                                { type: 'TextBlock', text: `${this.capitalizeWords(`${senderKing.FirstName} ${senderKing.LastName}`)}, who would you like to wage war on?` },
                                                { type: 'TextBlock', text: `Enemy kings:` }
                                            ],
                                            actions
                                        }
                                    }
                                ]
                            });
                        } catch(e) {
                            this.onError(e, context);
                            console.error('### Error (sendActivity senderKing stuff)', e);
                        }
                    } else {
                        await context.sendActivity(`${senderKingEmail} is not a valid king.`);
                    }
                }
                return;
            }
        } else {
            await context.sendActivity(`${command} is not a valid action.`);
            return;
        }
    }

    /**
     * Send back feedback to user
     * 
     * @param error Error
     * @param context Teams context
     */
    private async onError(error: Error, context: TurnContext): Promise<void> {
        await context.sendActivity(`Something went wrong: ${error}`);
    }

    public constructor(conversationState: ConversationState) {
        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog("help"));

        this.activityProc.messageActivityHandler = {
            onMessage: async (context: TurnContext): Promise<void> => {
                const teamsContext: TeamsContext = TeamsContext.from(context);
                const sender: TeamsChannelAccount = await this.getSenderInformation((context.adapter as TeamsAdapter), context);

                switch (context.activity.type) {
                    case ActivityTypes.Message:
                        let text: string = teamsContext ? (
                            teamsContext.getActivityTextWithoutMentions() ? 
                                teamsContext.getActivityTextWithoutMentions().toLowerCase() : 
                                context.activity.text
                        ) : context.activity.text;
                        if (text) {
                            await this.messageHandler(text, context, sender);
                        }
                    case ActivityTypes.Invoke: {
                        if (context.activity.value) {
                            if (context.activity.value.addMoney) {
                                const king: string = context.activity.value.king ? context.activity.value.king : sender.email.toLowerCase();
                                const get: any = await fetch(
                                    `https://pillagers-storage-functions.azurewebsites.net/api/GetKing?email=${king}`,
                                    { method: 'GET',  headers: { 'Content-Type': 'application/json' } }
                                );
                                const json: any = await get.json();
                                const currentPennings: number = parseFloat(json.value[0].Penning);
                                const addedPennings: number = parseFloat(context.activity.value.addMoney);
                                if (currentPennings && addedPennings) {
                                    const Penning: number = addedPennings + currentPennings;
                                    await fetch(
                                        'https://pillagers-storage-functions.azurewebsites.net/api/SetPenning',
                                        { method: 'POST', body: JSON.stringify({ king, Penning }), headers: { 'Content-Type': 'application/json' }}
                                    );
                                } else {
                                    console.error('### Error - missing monies');
                                    this.onError(new Error('Cannot find both sources of pennings'), context);
                                }
                            }
                            if (context.activity.value.targetKingName) {
                                const targetKingEmail: string = context.activity.value.targetKingEmail;
                                const targetKingName: string = context.activity.value.targetKingName;
                                try {
                                    await context.sendActivity({
                                        type: 'message',
                                        attachments: [
                                            {
                                                contentType: 'application/vnd.microsoft.card.adaptive',
                                                content: {
                                                    type: 'AdaptiveCard',
                                                    version: '1.0',
                                                    body: [
                                                        { type: 'TextBlock', text: `Waging war against ${targetKingName}`, size: "Large", color: 'Attention', weight: 'Bolder' },
                                                    ]
                                                }
                                            }
                                        ]
                                    });
                                } catch(e) {
                                    this.onError(e, context);
                                    console.error(`### Error (waging war on action)`, e);
                                }

                                await this.wageWar(context, { attacker: sender.email.toLowerCase(), defender: targetKingEmail });
                            }
                        }
                    }
                    default:
                        break;
                }

                return this.conversationState.saveChanges(context);
            }
        };

        this.activityProc.conversationUpdateActivityHandler = {
            onConversationUpdateActivity: async (context: TurnContext): Promise<void> => {
                if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                    for (const idx in context.activity.membersAdded) {
                        if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                            const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                            await context.sendActivity({ attachments: [welcomeCard] });
                        }
                    }
                }
            }
        };
   }

   /**
    * Wage war
    * 
    * @param context Teams context
    * @param data attacker and defender emails
    */
   private async wageWar(context: TurnContext, data: { attacker: string, defender: string }) {
        const response = await fetch(
            `https://pillagers-storage-functions.azurewebsites.net/api/WageWar?attacking=${data.attacker}&defending=${data.defender}`,
            { method: 'GET',  headers: { 'Content-Type': 'application/json' } }
        );
        const json = await (response as any).json();

        console.log(`### json feedback wageWar`, json);
   }

   private async getSenderInformation(adapter: TeamsAdapter, context: TurnContext): Promise<TeamsChannelAccount> {
       const activityMembers: TeamsChannelAccount[] = await adapter.getActivityMembers(context);
       let conversationMembers: TeamsChannelAccount[] = [];
       if (!activityMembers) {
        conversationMembers = await adapter.getConversationMembers(context);
       }

       const members: TeamsChannelAccount[] = activityMembers ? activityMembers : conversationMembers;
       if (members[0]) {
        return members[0] as TeamsChannelAccount;
       }

       return null;
   }

    /**
     * Capitalize start character of every word in string.
     * Used for name capitalization
     * 
     * @param str String to capitalize
     */
    private capitalizeWords(str: string): string {
        let words: string[] = str.toLowerCase().split(' ');
        for (let i = 0; i < words.length; i++) {
            words[i] = `${words[i].charAt(0).toUpperCase()}${words[i].substring(1)}`
        }
        return words.join(' ');
    }

    /**
     * Get all kings
     * 
     * @returns Array of kings or an empty array
     */
    private async getKings(context: TurnContext): Promise<any[]> {
        try {
            const response = await fetch(
                'https://pillagers-storage-functions.azurewebsites.net/api/GetKings',
                { method: 'GET',  headers: { 'Content-Type': 'application/json' } }
            );
            console.log('¤¤¤¤¤¤¤ get kings repsonse', response);
            const json = await (response as any).json();
            if (json) {
                return json.value;
            }
        } catch(e) {
            console.error('### Error (getKings())', e);
            this.onError(e, context);
            return [];
        }

        return [];
    }

   public async onTurn(context: TurnContext): Promise<any> {
       console.log('### activity type', context.activity.type);
        await this.activityProc.processIncomingActivity(context);
    }

}
