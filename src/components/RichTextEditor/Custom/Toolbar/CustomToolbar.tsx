import * as React from 'react';
import { RichText as PnPRichText } from '../../PnP/RichText';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { TooltipHost } from 'office-ui-fabric-react/lib/Tooltip';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import styles from '../../PnP/RichText.module.scss';
import customStyles from './Toolbar.module.scss';
import { ICustomRichTextProps, IPlaceHolderMenuItemProps, IReusableContentItem } from '../RichText';
import { isNullOrEmpty } from '@spfxappdev/utility';

export interface ICustomToolbarProps {
    Editor: PnPRichText;
    richTextproperties: ICustomRichTextProps;
}

export class CustomToolbar extends React.Component<ICustomToolbarProps> {

    private onRenderTitle(title: string, iconName: string): JSX.Element {      
        return (
          <div>
            <div className={`${styles.toolbarSubmenuDisplayButton} ${customStyles.dropdownsubmenu}`}>
                <Icon className={styles.toolbarDropDownTitleIcon}
                        iconName={iconName}
                        aria-hidden="true" />
                <span className={customStyles.dropdowntitle}>{title}</span>
            </div>
          </div>
        );
    }

    public render() : React.ReactElement<ICustomToolbarProps> {
        const reusableContentItems = this.getReUsableContentItems();
        const placeholderItems = this.getPlaceholderItems();
        return (
            <div className={customStyles.customtoolbarwrapper}>
                {this.props.richTextproperties.placeholderProps.show &&
                    <TooltipHost content={"Insert Placeholder"}
                            id="placeholder-richtextbutton"
                            calloutProps={{ gapSpace: 0 }}>
                        <Dropdown className={styles.toolbarDropDown}
                        id="DropDownPlaceholders"
                                    onRenderCaretDown={() => <Icon className={styles.toolbarSubmenuCaret} iconName="CaretDownSolid8" />}
                                    options={placeholderItems}
                                    /*onRenderOption={(option: IDropdownOption): JSX.Element => {
                                        return this.onRenderListOption(option);
                                    }}*/
                                    placeholder="Placeholder"
                                    onRenderTitle={(p: any) => {return this.onRenderTitle("Placeholder", "WebAppBuilderSlot");}}
                                    onRenderPlaceholder={(p: any) => {return this.onRenderTitle("Placeholder", "WebAppBuilderSlot");}}
                                    selectedKey={placeholderItems[0].key}
                                    onChanged={(option: IDropdownOption, index: number) => {this.onPlaceholderChanged(option, index);}}
                        />
                </TooltipHost>
                }

            {this.props.richTextproperties.reusableContentProps.show && !isNullOrEmpty(reusableContentItems[0]) &&
              <TooltipHost content={"Insert Reusable content"}
                           id="reusable-richtextbutton"
                           calloutProps={{ gapSpace: 0 }}>
              <Dropdown className={styles.toolbarDropDown}
              id="DropDownReusableContents"
                        onRenderCaretDown={() => <Icon className={styles.toolbarSubmenuCaret} iconName="CaretDownSolid8" />}
                        options={reusableContentItems}
                        /*onRenderOption={(option: IDropdownOption): JSX.Element => {
                            return this.onRenderListOption(option);
                        }}*/
                        placeholder="Reusable Content"
                        onRenderTitle={(p: any) => {return this.onRenderTitle("Reusable Content", "Sprint");}}
                        onRenderPlaceholder={(p: any) => {return this.onRenderTitle("Reusable Content", "Sprint");}}
                        selectedKey={reusableContentItems[0].key}
                        onChanged={(option: IDropdownOption, index: number) => {this.onReUsabeleContentChanched(option, index);}}
              />
              </TooltipHost>
                }
            </div>
        );
    }

    private getReUsableContentItems(): IDropdownOption[] {
        return this.props.richTextproperties.reusableContentProps.menuItems.map((reusablecontent: IReusableContentItem, i: number): IDropdownOption => {
            return {
                key: `item_${i}`,
                text: reusablecontent.title,
                data: { content: reusablecontent.content, isStatic: reusablecontent.isStatic, itemId: reusablecontent.id }
            };
        });
    }

    private getPlaceholderItems(): IDropdownOption[] {
        return this.props.richTextproperties.placeholderProps.menuItems.map((placeholder: IPlaceHolderMenuItemProps, i: number): IDropdownOption => {
            return {
                key: `placeholder_${i}`,
                text: placeholder.menuText,
                data: { content: placeholder.rteContent }
            };
        });
    }

    private handleClick(): void {
        const quill = this.props.Editor.getEditor();
        const range = quill.getSelection(true);
        const value = `{UserProfile.ProfilePropertyName}`;
        
        const cursorPosition: number = range!.index;
        
        quill.insertText(range.index, value);
        quill.setSelection(range.index + value.length + 5, 0);
    }

    private onPlaceholderChanged(option: IDropdownOption, index: number): void {
        if(option.key == 'default') {
            return;
        }

        const quill = this.props.Editor.getEditor();
        const range = quill.getSelection(true);
        const value = option.data.content;
        
        const cursorPosition: number = range!.index;
        
        quill.insertText(range.index, value);
        quill.setSelection(range.index + value.length + 5, 0);
    }

    private onReUsabeleContentChanched(option: IDropdownOption, index: number): void {
        if(option.key == 'default') {
            return;
        }
        const quill = this.props.Editor.getEditor();
        const range = quill.getSelection(true);
        const value = option.data.content;
        
        const cursorPosition: number = range!.index;
        const delta = quill.insertEmbed(range.index, "reusable", { isStatic: option.data.isStatic, value: value });
        quill.insertText(range.index + value.length + 5, ' ');
        quill.setSelection(range.index + value.length + 5, 0);
    }

}