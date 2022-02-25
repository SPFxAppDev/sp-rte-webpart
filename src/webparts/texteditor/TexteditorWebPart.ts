import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { ISPFxAppDevClientSideWebPartProps, SPFxAppDevClientSideWebPart } from '@spfxappdev/framework';
import { IReusableContentItem, RichText } from '@src/components/RichTextEditor/Custom/RichText';
import { IRichTextComponentProps } from '@src/components/RichTextEditor/Custom/RichText';
import { SPPlaceholderResolver  } from '@spfxappdev/framework/lib/spfxframework/placeholder/SPPlaceholderResolver';
import { UserProfilePlaceholderResolver } from './UserProfilePlaceholderResolver';
// import { SPPlaceholderResolver  } from '@spfxappdev/framework/lib/spfxframework/placeholder';
import { CurrentUserProfile, PortalUser } from '@spfxappdev/framework/lib/spfxframework/sp/userprofile/CurrentUserProfile';
import { SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';


export interface ITexteditorWebPartProps extends ISPFxAppDevClientSideWebPartProps {
  content: string;
}


export default class TexteditorWebPart extends SPFxAppDevClientSideWebPart<ITexteditorWebPartProps> {

    private currentUser: PortalUser = null;

    private reusableItems: IReusableContentItem[] = [];

    public onInit(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
          super.onInit().then(async () => {
            this.currentUser = await CurrentUserProfile.Get(this.spfxContext);
            return resolve();
          });
        });
    }

    public render(): void {

      if(!this.helper.functions.isset(this.properties)) {
        this.properties.content = "";
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
          rteProps: {
            isEditMode: this.IsPageInEditMode,
            value: this.properties.content,
            placeholderProps: {
              show: false,
              menuItems: [],
              placeholderResolver: [
                new SPPlaceholderResolver(this.spfxContext.pageContext),
                new UserProfilePlaceholderResolver(this.currentUser)
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
          rteProps: {
            isEditMode: this.IsPageInEditMode,
            value: this.properties.content,
            onChange: (text: string): string =>  {
              this.properties.content = text;
              return text;
            },
            placeholderProps: {
              show: true,
              menuItems: [{
                menuText: "UserProfile Properties",
                rteContent: "{UserProfile.ProfilePropertyName}"
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
                new UserProfilePlaceholderResolver(this.currentUser)
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