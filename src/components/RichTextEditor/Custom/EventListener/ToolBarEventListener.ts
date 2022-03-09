import * as React from 'react';
import { IEventListenerResult } from "@spfxappdev/framework/lib/spfxframework/events/EventHandler";
import { EventListenerBase } from "@spfxappdev/framework/lib/spfxframework/events/EventListenerBase";
import { ICustomRichTextProps } from '../RichText';
import { CustomToolbar, ICustomToolbarProps } from '../Toolbar/CustomToolbar';

export class ToolBarEventListener extends EventListenerBase {
    public richTextproperties: ICustomRichTextProps = null;
    
    public Execute(name: string, lastEventResult: IEventListenerResult|null, quill: any): ToolBarEventListener {
        const element: React.ReactElement<ICustomToolbarProps> = React.createElement(
            CustomToolbar, {
                Editor: quill,
                richTextproperties: this.richTextproperties
            }
        );
        this.Result = element;
        
        return this;
      }  
}