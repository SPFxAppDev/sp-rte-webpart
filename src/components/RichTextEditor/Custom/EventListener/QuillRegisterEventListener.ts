import { IEventListenerResult } from "@spfxappdev/framework/lib/spfxframework/events/EventHandler";
import { EventListenerBase } from "@spfxappdev/framework/lib/spfxframework/events/EventListenerBase";
import { ReusableContentBlot } from "../QuillBlots/ReusableContentBlot";

export class QuillRegisterEventListener extends EventListenerBase {
    public Execute(name: string, lastEventResult: IEventListenerResult|null, ...args: any[]): QuillRegisterEventListener {
        
        const quill: any = args[0][0];
        quill.register(ReusableContentBlot, true);
        return this;
    }
}