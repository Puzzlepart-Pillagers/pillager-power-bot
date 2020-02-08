import { Dialog, DialogContext, DialogTurnResult } from "botbuilder-dialogs";

export default class HelpDialog extends Dialog {
    constructor(dialogId: string) {
        super(dialogId);
    }

    public async beginDialog(context: DialogContext, options?: any): Promise<DialogTurnResult> {
        //context.context.sendActivity({ type: 'xml', text: `<pre>help, man, stats, me, king, war, wagewar</pre>` });
        return await context.endDialog();
    }
}
