import { Dialog, DialogContext, DialogTurnResult } from "botbuilder-dialogs";

export default class HelpDialog extends Dialog {
    constructor(dialogId: string) {
        super(dialogId);
    }

    public async beginDialog(context: DialogContext, options?: any): Promise<DialogTurnResult> {
        context.context.sendActivity(`TODO - help`);
        return await context.endDialog();
    }
}
