import { BotDeclaration, IBot } from "express-msteams-host";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, ChannelAccount, BotAdapter } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import WelcomeCard from "./dialogs/WelcomeDialog";
import { TeamsContext, TeamsActivityProcessor, TeamsConnectorClient, TeamsAdapter, TeamsChannelAccount } from "botbuilder-teams";
import * as AdaptiveCard from 'adaptivecards';
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
    
    private commands: string[] = [ 'king', 'me', 'stats', 'help', 'man' ];
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
                            { method: 'GET',  headers: { 'Content-Type': 'application/json' }}
                        );
                        kings = await (response as any).json();
                    } catch(e) {
                        console.error('### error:', e);
                    }

                    if (kings.value[0]) {
                        try {
                            const king = kings.value[0];
                            let adaptiveCard = new AdaptiveCard.AdaptiveCard();
                            const card = {
                                type: 'message',
                                attachments: [
                                    {
                                        contentType: 'application/vnd.microsoft.card.adaptive',
                                        content: {
                                            type: 'AdaptiveCard',
                                            version: '1.0',
                                            body: [
                                                { type: 'xml', text: '<b>King</b>' },
                                                { type: 'TextBlock', text: `name: ${king.FirstName} ${king.LastName}` },
                                                { type: 'TextBlock', text: `monies: ${king.Penning} Pennings` }
                                            ],
                                            actions: [
                                                { type: 'Action.OpenUrl', title: 'pillagers.no', url: 'http://pillagers.no' }
                                            ]
                                        }
                                    }
                                ]
                            }
                            adaptiveCard.onExecuteAction = (action) => { alert(`action ${action.title} executed`); }
                            adaptiveCard.parse(card);

                            await context.sendActivity(adaptiveCard);
                        } catch(e) {
                            console.error('### error:', e);
                        }
                    } else {
                        await context.sendActivity(`Cannot find a user registred with: <i>${request.email}</i>, registrer at <a href='http://pillagers.no'>pillagers.no<a/>.`)
                    }
                    return;
                }
                case 'man':
                case 'help': {
                    const dc = await this.dialogs.createContext(context);
                    await dc.beginDialog("help");
                    return;
                }
            }
        } else {
            await context.sendActivity(`${command} is not a valid action.`);
            return;
        }
    }

    public constructor(conversationState: ConversationState) {
        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog("help"));

        this.activityProc.messageActivityHandler = {
            onMessage: async (context: TurnContext): Promise<void> => { // NOTE Incoming messages
                const teamsContext: TeamsContext = TeamsContext.from(context); // NOTE will be undefined outside of teams
                
                switch (context.activity.type) {
                    case ActivityTypes.Message:
                        let text: string = teamsContext ? teamsContext.getActivityTextWithoutMentions().toLowerCase() : context.activity.text;
                        const sender: TeamsChannelAccount = await this.getSenderInformation((context.adapter as TeamsAdapter), context);

                        await this.messageHandler(text, context, sender);
                        console.log('### - Command finsihed');
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

   public async onTurn(context: TurnContext): Promise<any> {
        await this.activityProc.processIncomingActivity(context);
    }

}
