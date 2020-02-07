import { BotDeclaration, IBot } from "express-msteams-host";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, ChannelAccount } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import WelcomeCard from "./dialogs/WelcomeDialog";
import { TeamsContext, TeamsActivityProcessor } from "botbuilder-teams";
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
                        let sender: ChannelAccount = context.activity.from;
                        let test = context.activity;

                        if (text.startsWith("stats")) {
                            if (sender) {
                                try {
                                    const response = await got(`https://pillagers-storage-functions.azurewebsites.net/api/GetUnits?email=${sender.name}`);
                                    //console.log(response);
                                    console.log('test', test);
                                    await context.sendActivity(`stats for ${sender.name}:`);
                                    return;
                                } catch(e) {
                                    console.error(e);
                                    return;
                                }
                            }
                            await context.sendActivity(`Cannot find user`);
                            return;
                        }  else if (text.startsWith("help")) {
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
