import { ActivityTypes, BotFrameworkAdapter } from "botbuilder";
import { ComponentDialog, DialogContext } from "botbuilder-dialogs";

export class LogoutDialog extends ComponentDialog {
  // really need this?
  connectionName: string | undefined;

  constructor(id: string, connectionName: string | undefined) {
    super(id);
    this.connectionName = connectionName;
  }

  async onBeginDialog(innerDc: DialogContext, options: {} | undefined) {
    const result = await this.interrupt(innerDc);
    if (result) {
      return result;
    }

    return await super.onBeginDialog(innerDc, options);
  }

  async onContinueDialog(innerDc: DialogContext) {
    const result = await this.interrupt(innerDc);
    if (result) {
      return result;
    }

    return await super.onContinueDialog(innerDc);
  }

  async interrupt(innerDc: DialogContext) {
    if (innerDc.context.activity.type === ActivityTypes.Message) {
      const text = innerDc.context.activity.text.toLowerCase().trim();
      if (text === "logout") {
        // The bot adapter encapsulates the authentication processes.
        const botAdapter = innerDc.context.adapter as BotFrameworkAdapter;
        await botAdapter.signOutUser(innerDc.context, this.connectionName); // ? no need to sign out since there is no token store
        await innerDc.context.sendActivity("You have been signed out.");
        return await innerDc.cancelAllDialogs();
      }
    }
  }
}
