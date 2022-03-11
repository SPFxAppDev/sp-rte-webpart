import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneButton } from '@microsoft/sp-property-pane';
import { ISPFxAppDevClientSideWebPartProps, SPFxAppDevClientSideWebPart } from '@spfxappdev/framework';
import { IReusableContentItem, RichText } from '@src/components/RichTextEditor/Custom/RichText';
import { IRichTextComponentProps } from '@src/components/RichTextEditor/Custom/RichText';
import { SPPlaceholderResolver  } from '@spfxappdev/framework/lib/spfxframework/placeholder/SPPlaceholderResolver';
import { UserProfilePlaceholderResolver } from './UserProfilePlaceholderResolver';
// import { SPPlaceholderResolver  } from '@spfxappdev/framework/lib/spfxframework/placeholder';
import { CurrentUserProfile, PortalUser } from '@spfxappdev/framework/lib/spfxframework/sp/userprofile/CurrentUserProfile';
import { SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import * as strings from 'RteStringsWebPartStrings';
import ManageGreeting, { IManageGreetingProps } from './components/ManageGreeting';
import { GreetingPlaceholderResolver } from './GreetingPlaceholderResolver';
import { PropertyPaneAboutWebPart } from '../PropertyPaneAboutWebPart';
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme,
} from '@microsoft/sp-component-base';

export interface IGreeting {
  fromHr: number;
  fromMinutes: number;
  toHr: number;
  toMinutes: number;
  text: string;
}

export interface ITexteditorWebPartProps extends ISPFxAppDevClientSideWebPartProps {
  content: string;
  greetingSettings: IGreeting[];
}


export default class TexteditorWebPart extends SPFxAppDevClientSideWebPart<ITexteditorWebPartProps> {

    private currentUser: PortalUser = null;

    private reusableItems: IReusableContentItem[] = [];

    private themeProvider: ThemeProvider;
    private themeVariant: IReadonlyTheme | undefined;

    private handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
      this.themeVariant = args.theme;
      this.render();
    }

    public onInit(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
          super.onInit().then(async () => {
            this.currentUser = await CurrentUserProfile.Get(this.spfxContext);
            this.properties.greetingSettings = this.properties.greetingSettings||[];
            this.themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
            // If it exists, get the theme variant
            this.themeVariant = this.themeProvider.tryGetTheme();
            // Register a handler to be notified if the theme variant changes
            this.themeProvider.themeChangedEvent.add(this, this.handleThemeChangedEvent);
            return resolve();
          });
        });
    }

    public render(): void {

      if(!this.helper.functions.isset(this.properties)) {
        (this.properties as any) = {
          content: "",
          greetingSettings: []
        };
      }

      if(!this.IsPageInEditMode) {
        this.renderDisplayMode();
        return;
      }

      this.loadReUsableContentFromList().then((reusableItems: IReusableContentItem[]) => {
        this.reusableItems = reusableItems;
        this.renderEditMode();
      });     
      
    }

    public getLogCategory(): string {
        return 'TexteditorWebPart';
    }

    protected renderDisplayMode(): void {
      const element: React.ReactElement<IRichTextComponentProps> = React.createElement(
        RichText,
        {
          themeVariant: this.themeVariant,
          rteProps: {
            isEditMode: this.IsPageInEditMode,
            value: this.properties.content,
            placeholderProps: {
              show: false,
              menuItems: [],
              placeholderResolver: [
                new SPPlaceholderResolver(this.spfxContext.pageContext),
                new UserProfilePlaceholderResolver(this.currentUser),
                new GreetingPlaceholderResolver(this.properties.greetingSettings)
              ]
            }
          }
        }
      );
  
      ReactDom.render(element, this.domElement);
    }
  
    protected async renderEditMode(): Promise<void> {

      
      const element: React.ReactElement<IRichTextComponentProps> = React.createElement(
        RichText,
        {
          themeVariant: this.themeVariant,
          rteProps: {
            isEditMode: this.IsPageInEditMode,
            value: this.properties.content,
            onChange: (text: string): string =>  {
              this.properties.content = text;
              return text;
            },
            placeholderProps: {
              show: true,
              menuItems: [
              {
                menuText: "UserProfile Properties",
                rteContent: "{UserProfile.ProfilePropertyName}"
              },
              {
                menuText: "Greeting text",
                rteContent: "{Greeting}"
              },
              {
                menuText: "Absolute Web-URL",
                rteContent: "{Web.Url}"
              },
              {
                menuText: "Relative Web-URL",
                rteContent: "{Web.RelativeUrl}"
              },
              {
                menuText: "Web Title",
                rteContent: "{Web.Title}"
              },
              {
                menuText: "Absolute Site-Url",
                rteContent: "{Site.Url}"
              },
              {
                menuText: "Relative Site-Url",
                rteContent: "{Site.RelativeUrl}"
              }],
              placeholderResolver: [
                new SPPlaceholderResolver(this.spfxContext.pageContext),
                new UserProfilePlaceholderResolver(this.currentUser),
                new GreetingPlaceholderResolver(this.properties.greetingSettings)
              ]
            },
            reusableContentProps: {
              show: true,
              menuItems: this.reusableItems
            }
          }
        }
      );
  
      ReactDom.render(element, this.domElement);
    }
  

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
      return {
        pages: [
          {
            groups: [
              {
                groupName: strings.WebPartPropertyGroupGreetings,
                groupFields: [
                  PropertyPaneButton(null, {
                    text: strings.WebPartManageGreetingTextButtonText,
                    onClick: (val: any) => {
                      const dummyElement: HTMLDivElement = document.createElement("div");
                      document.body.appendChild(dummyElement);
    
                      const element: React.ReactElement<IManageGreetingProps> = React.createElement(ManageGreeting, {
                        greetingSettings: this.properties.greetingSettings||[],
                        onDismiss: () => {
                          dummyElement.remove();
                        },
                        onGreetingChanged: (greetingSettings: IGreeting[]) => {
                          dummyElement.remove();
                          this.properties.greetingSettings = greetingSettings;
                        }
                      });
    
                      ReactDom.render(element, dummyElement);
                      
                      return null;
                    }
                  })
                ]
              },
              {  
                groupName: strings.WebPartPropertyGroupAbout,
                groupFields: [
                  PropertyPaneAboutWebPart({
                    key: `bbc94d63-5630-4077-b6a8-6b8d37551766_${this.context.instanceId}`,
                    html: `
                    <div>
                      <h3>Author</h3> 
                      <a href="https://spfx-app.dev/" data-interception="off" target="_blank">SPFx-App.dev</a>
                      <h3>Version</h3>
                      ${this.context.manifest.version}
                      <h3>Web Part Instance id</h3>
                      ${this.context.instanceId}
                    </div>
                    <div>
                      <a href="https://github.com/SPFxAppDev/sp-rte-webpart" target="_blank">More info</a>
                    </div>
                    `
                  })    
                ]
              }
            ],
            // displayGroupsAsAccordion: true,
          }
        ]
      };
    }

    private loadReUsableContentFromList(): Promise<IReusableContentItem[]> {
      return new Promise<IReusableContentItem[]>((resolve, reject) => {
        if(!this.helper.functions.isNullOrEmpty(this.reusableItems)) {
          return resolve(this.reusableItems);
        }
  
        const relativeListUrl = this.helper.url.MakeRelativeSiteUrl("Lists/ReusableContent");
        const endPoint = this.helper.url.MakeAbsoluteWebUrl("/_api/web/getlist('" + relativeListUrl + "')/items?$select=Id,Title,spfxAppDevReusableContent,spfxAppDevReusableContentIsStatic");
        this.context.spHttpClient.get(endPoint, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
          response.json().then((responseJSON: any) => {
            this.log(responseJSON);
  
            const listItems: any[] = this.helper.functions.getDeepOrDefault(responseJSON, "value", []);
            const result: IReusableContentItem[] = listItems.map((listItem: any, index: number): IReusableContentItem => {
                let reusableContent: IReusableContentItem = {
                  content: listItem.spfxAppDevReusableContent,
                  id: listItem.Id,
                  isStatic: this.helper.functions.toBoolean(listItem.spfxAppDevReusableContentIsStatic),
                  title: listItem.Title
                };
                return reusableContent;
            });
            return resolve(result);
          });
        }).catch((err) => {
          resolve([]);
        });
      });
    }

    
}