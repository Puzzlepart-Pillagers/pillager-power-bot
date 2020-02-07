import { BotDeclaration, IBot } from "express-msteams-host";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, ChannelAccount, BotAdapter } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import WelcomeCard from "./dialogs/WelcomeDialog";
import { TeamsContext, TeamsActivityProcessor, TeamsConnectorClient, TeamsAdapter, TeamsChannelAccount } from "botbuilder-teams";
const got = require('got');

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
    private emailRegex: RegExp = new RegExp('/^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/');

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
                        const adapter = (context.adapter as TeamsAdapter);
                        const activityMembers = await adapter.getActivityMembers(context);
                        const conversationMembers = await adapter.getConversationMembers(context);
                        const members: TeamsChannelAccount[] = activityMembers ? activityMembers : conversationMembers;
                        const sender: TeamsChannelAccount = members[0];

                        if (text.startsWith("get king")) {
                            if (sender) {
                                try {
                                    const response = await got(`https://pillagers-storage-functions.azurewebsites.net/api/GetKing?email=${'kim@pzl.onmicrosoft.com'}`);
                                    const json = response.json();
                                    const value = json.body.value;
                                    console.log(value);
                                    if (value === []) {
                                        await context.sendActivity(`Cannot find user`);
                                        return;
                                    }
                                    await context.sendActivity(`stats for ${sender.email}: stats`);
                                    return;
                                } catch(e) {
                                    console.error(e);
                                    return;
                                }
                            }
                            // else - cannot find user
                            await context.sendActivity(`Cannot find user`);
                            return;
                        } else if (text.startsWith('members')) {
                            const members = await (context.adapter as TeamsAdapter).getActivityMembers(context);
                            const convMembers = await (context.adapter as TeamsAdapter).getConversationMembers(context);
                            await context.sendActivity({ textFormat: 'xml', text: `<b>Activity members:</b><pre>${JSON.stringify(members, null, 2)}</pre><b>Conversation members:</b><pre>${JSON.stringify(convMembers, null, 2)}</pre>` });
                            return;
                        } else if (text.startsWith("help")) {
                            const dc = await this.dialogs.createContext(context);
                            await dc.beginDialog("help");
                        } else {
                            await context.sendActivity(`I\'m terribly sorry, but my master hasn\'t trained me to do anything yet...`);
                        }
                        break;
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

   public async onTurn(context: TurnContext): Promise<any> {
        await this.activityProc.processIncomingActivity(context);
    }

}
