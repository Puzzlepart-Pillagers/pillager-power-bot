import { BotDeclaration, IBot } from "express-msteams-host";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, ChannelAccount, BotAdapter } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import WelcomeCard from "./dialogs/WelcomeDialog";
import { TeamsContext, TeamsActivityProcessor, TeamsConnectorClient, TeamsAdapter, TeamsChannelAccount } from "botbuilder-teams";
import request = require("request");
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
    
    private commands: string[] = [ 'king', 'me', 'help' ];
    private async messageHandler(text: string, context: TurnContext, sender: TeamsChannelAccount): Promise<void> {  
        let args: string[] = text.trim().split(' ');
        const command: string = args[0].toLocaleLowerCase();
        if (this.commands.indexOf(command) !== -1) {
            switch(command) {
                case 'me':
                case 'king': {
                    let request = { email: sender.email };
                    if (args.indexOf('--user') !== -1) {
                        const value = args[args.indexOf('--user') + 1]
                        if (value) request.email = value;
                    }

                    const response = await fetch(`https://pillagers-storage-functions.azurewebsites.net/api/GetKing?email=${request.email}`, { method: 'GET', headers: { 'Content-Type': 'application/json' } });
                    console.log('### response', response);
                    try {
                        const json = await response.json();
                        if (json) {
                            console.log('### json:', json);
                        }
                    } catch(e) {
                        console.error('### error -', e);
                    }

                    await context.sendActivity({ 
                        textFormat: 'xml', 
                        text: `<b>King: ${sender.name}</b>` 
                    });
                    return;
                }
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

        // Message reactions in Microsoft Teams
        this.activityProc.messageReactionActivityHandler = {
            onMessageReaction: async (context: TurnContext): Promise<void> => {
                const added = context.activity.reactionsAdded;
                if (added && added[0]) {
                    await context.sendActivity({
                        textFormat: "xml",
                        text: `That was an interesting reaction (<b>${added[0].type}</b>)`
                    });
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
