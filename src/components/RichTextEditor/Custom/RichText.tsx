
import * as React from 'react';
import { IRichTextProps } from '../PnP/RichText.types';
import { IPlaceholderResolver, PlaceholderHandler } from '@spfxappdev/framework/lib/spfxframework/placeholder/PlaceholderHandler';
import { RichText as PnPRichText } from '../PnP/RichText';
import { issetDeep, isFunction } from '@spfxappdev/utility';
import { EventHandler } from '@spfxappdev/framework/lib/spfxframework/events/EventHandler';
import { ToolBarEventListener } from './EventListener/ToolBarEventListener';
import { QuillRegisterEventListener } from './EventListener/QuillRegisterEventListener';
import {
    IReadonlyTheme,
  } from '@microsoft/sp-component-base';

export interface ISPFxAppDevRichtTextToolbarItemProps {
    show: boolean;
}

export interface IPlaceHolderMenuItemProps {
    menuText: string;
    rteContent: string;
}

export interface IPlaceHolderProps extends ISPFxAppDevRichtTextToolbarItemProps {
    placeholderResolver?: IPlaceholderResolver[];
    menuItems: IPlaceHolderMenuItemProps[];
}

export interface IReusableContentItem {
    title: string;
    content: string;
    isStatic: boolean;
    id: string|number;
}

export interface IReusableContentProps extends ISPFxAppDevRichtTextToolbarItemProps {
    menuItems: IReusableContentItem[];
}


export interface ICustomRichTextProps extends IRichTextProps {
    placeholderProps?: IPlaceHolderProps;
    reusableContentProps?: IReusableContentProps;
}

export interface IRichTextComponentProps {
    themeVariant: IReadonlyTheme | undefined;
    rteProps: ICustomRichTextProps;
}

export class RichText extends React.Component<IRichTextComponentProps, {}> {
    private static isRegistered: boolean = false;
    private placeholderHandler: PlaceholderHandler;
    private originalOnChangeFn?(text: string): string;

    constructor(props: IRichTextComponentProps) {
        super(props);
        this.placeholderHandler = new PlaceholderHandler();
        if (issetDeep(this.props, "rteProps.placeholderProps.placeholderResolver")) {
            this.props.rteProps.placeholderProps.placeholderResolver.forEach((resolver: IPlaceholderResolver) => {
                this.placeholderHandler.Register(resolver);
            });
        }

        if(isFunction(this.props.rteProps.onChange)) {
            this.originalOnChangeFn = this.props.rteProps.onChange;
            this.props.rteProps.onChange = undefined;
        }
        this.registerHandler();
    }

    public render(): React.ReactElement<IRichTextComponentProps> {
        if(!this.props.rteProps.isEditMode) {
            this.props.rteProps.value = this.placeholderHandler.Replace(this.props.rteProps.value);
        }

        const { semanticColors }: IReadonlyTheme = this.props.themeVariant;
    
        return(
          <>
          <PnPRichText {...this.props.rteProps} style={{ backgroundColor: semanticColors.bodyBackground, color: semanticColors.bodyText }} onChange={(text: string) => {
              if(isFunction(this.originalOnChangeFn)) {
                  text = this.originalOnChangeFn(text);
              }

              return text;
          }} />
          </>
        );
    }

    public registerHandler(): void {
        if(!this.props.rteProps.isEditMode) {
            return;
        }
        if(!RichText.isRegistered){
          const eventName: string = "OnRTEHeaderControlsRender";
          const listener = new ToolBarEventListener();
          listener.Sequence = 10;
          listener.richTextproperties = this.props.rteProps;
          EventHandler.Listen(eventName, listener);
          EventHandler.Listen("OnQuillRegister", new QuillRegisterEventListener());
          RichText.isRegistered = true;
        }
      }
}