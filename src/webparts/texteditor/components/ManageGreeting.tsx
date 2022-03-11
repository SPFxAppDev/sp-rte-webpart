import { cloneDeep } from '@microsoft/sp-lodash-subset';
import { cssClasses, isFunction, isNullOrEmpty, isset } from '@spfxappdev/utility';
import { DefaultButton, Dialog, DialogContent, DialogFooter, DialogType, Dropdown, IDropdownOption, PrimaryButton, TextField } from 'office-ui-fabric-react';
import * as React from 'react';
import * as strings from 'RteStringsWebPartStrings';
import { IGreeting } from '../TexteditorWebPart';
import styles from './ManageGreeting.module.scss';
import "@spfxappdev/utility/lib/extensions/ArrayExtensions";

export interface IManageGreetingProps {
    greetingSettings: IGreeting[];
    onDismiss(): void;
    onGreetingChanged(greetingSettings: IGreeting[]);
}

interface IManageGreetingState {
    greetingSettings: IGreeting[];
    isDialogVisible: boolean;
    isSaveButtonDisabled: boolean;
    isAddNewButtonVisible: boolean;
}

type DropDownTimeValues = { hrOptions: IDropdownOption[], minuteOptions: IDropdownOption[] };

export default class ManageGreeting extends React.Component<IManageGreetingProps, IManageGreetingState> {

    public state: IManageGreetingState = {
        greetingSettings: this.getSettingsOrInitial(this.props.greetingSettings),
        isDialogVisible: true,
        isSaveButtonDisabled: false,
        isAddNewButtonVisible: this.isAddNewButtonVisible(this.props.greetingSettings)
    };
    
    public render(): React.ReactElement<IManageGreetingProps> {
        return (
        <Dialog 
            hidden={!this.state.isDialogVisible}
            onDismiss={() => { this.onDialogDismiss(); }}
            dialogContentProps={{
                title: strings.WebPartManageGreetingTextButtonText,
                type: DialogType.close
            }}
            minWidth={800}
            modalProps={{
                isBlocking: true,
                className: "greeting-dialog"
                
            }}>
                <DialogContent>
                <div className={cssClasses(styles['greeting-container'], "spfxappdev-grid")}>
                    <div className='spfxappdev-grid-row grid-header'>
                        <div className='spfxappdev-grid-col spfxappdev-sm4'>{strings.GreetingTextLabel}</div>
                        <div className='spfxappdev-grid-col spfxappdev-sm2'>{strings.FromHourLabel}</div>
                        <div className='spfxappdev-grid-col spfxappdev-sm2'>{strings.FromMinuteLabel}</div>
                        <div className='spfxappdev-grid-col spfxappdev-sm2'>{strings.ToHourLabel}</div>
                        <div className='spfxappdev-grid-col spfxappdev-sm2'>{strings.ToMinuteLabel}</div>
                    </div>
                    {this.state.greetingSettings.map((greetings: IGreeting, index: number): JSX.Element => {
                        
                        const startDropDownOptions: DropDownTimeValues = this.getStartTimeDropDownValues(index);
                        const endDropDownOptions: DropDownTimeValues = this.getEndTimeDropDownValues(index);


                        return(<div key={`greeting_${index}`} className='spfxappdev-grid-row greetings-grid'>
                            <div className='spfxappdev-grid-col spfxappdev-sm4'>
                                <TextField 
                                    defaultValue={greetings.text}
                                    onChange={(ev: any, newText: string) => {
                                        this.state.greetingSettings[index].text = newText;
                                        this.setState({
                                            greetingSettings: this.state.greetingSettings
                                        });
                                    }}
                                />
                            </div>
                            <div className='spfxappdev-grid-col spfxappdev-sm2'>
                                <Dropdown
                                    disabled={true}
                                    defaultSelectedKey={greetings.fromHr}
                                    onChange={(ev: any, option: IDropdownOption) => {
                                        this.state.greetingSettings[index].fromHr = option.key as number;
                                        this.onTimeChanged(this.state.greetingSettings);
                                    }}
                                    options={startDropDownOptions.hrOptions}
                                />
                            </div>
                            <div className='spfxappdev-grid-col spfxappdev-sm2'>
                                <Dropdown
                                    disabled={true}
                                    defaultSelectedKey={greetings.fromMinutes}
                                    onChange={(ev: any, option: IDropdownOption) => {
                                        this.state.greetingSettings[index].fromMinutes = option.key as number;
                                        this.onTimeChanged(this.state.greetingSettings);
                                    }}
                                    options={startDropDownOptions.minuteOptions}
                                />
                            </div>
                            <div className='spfxappdev-grid-col spfxappdev-sm2'>
                                <Dropdown
                                    defaultSelectedKey={greetings.toHr}
                                    onChange={(ev: any, option: IDropdownOption) => {
                                        this.state.greetingSettings[index].toHr = option.key as number;
                                        this.onTimeChanged(this.state.greetingSettings);
                                    }}
                                    options={endDropDownOptions.hrOptions}
                                />
                            </div>
                            <div className='spfxappdev-grid-col spfxappdev-sm2'>
                                <Dropdown
                                    defaultSelectedKey={greetings.toMinutes}
                                    onChange={(ev: any, option: IDropdownOption) => {
                                        this.state.greetingSettings[index].toMinutes = option.key as number;
                                        this.onTimeChanged(this.state.greetingSettings);
                                    }}
                                    options={endDropDownOptions.minuteOptions}
                                />
                            </div>
                        </div>);
                    })}

                    <div className='spfxappdev-grid-row grid-footer'>

                        <div className='spfxappdev-grid-col spfxappdev-sm12'>
                            <PrimaryButton 
                                text={strings.AddLabel} 
                                disabled={!this.isAddNewButtonVisible(this.state.greetingSettings)} 
                                onClick={() => {
                                    this.setState({
                                        greetingSettings: this.addNewGreetingRow(this.state.greetingSettings)
                                    });
                                }}
                            />
                        </div>
                    </div>
                </div>
                </DialogContent>
                <DialogFooter>
                    <PrimaryButton 
                        onClick={() => {

                            if(isFunction(this.props.onGreetingChanged)) {
                                this.props.onGreetingChanged(this.state.greetingSettings);
                            }

                            this.setState({
                                isDialogVisible: false
                            }); 
                        }} 
                        text={strings.SaveLabel} 
                        disabled={this.state.isSaveButtonDisabled} 
                    />
                    <DefaultButton onClick={() => {
                        this.onDialogDismiss();
                    }} text={strings.CancelLabel} />
                </DialogFooter>
        </Dialog>);
    }

    private onDialogDismiss(): void {
        this.setState({
            isDialogVisible: false
        }); 
        this.props.onDismiss();
    }

    private getStartTimeDropDownValues(currentIndex: number): DropDownTimeValues {
        const result: DropDownTimeValues = {
            hrOptions: [],
            minuteOptions: []
        };

        for (let index = 0; index < 24; index++) {
            
            const option: IDropdownOption = {
                key: index,
                text: ('0' + index).slice(-2)
            };

            result.hrOptions.push(option);
            
        }

        for (let index = 0; index < 60; index++) {
            
            const option: IDropdownOption = {
                key: index,
                text: ('0' + index).slice(-2)
            };

            result.minuteOptions.push(option);
            
        }

        return result;
    }

    private getEndTimeDropDownValues(currentIndex: number): DropDownTimeValues {
        const settings: IGreeting[] = this.state.greetingSettings;        

        // const prevElement: IGreeting = currentIndex == 0 ? null : settings[currentIndex-1];
        const currentElement: IGreeting = settings[currentIndex];
        // const nextElement: IGreeting = currentIndex + 1 < settings.length ? settings[currentIndex+1] : null;


        const result: DropDownTimeValues = {
            hrOptions: [],
            minuteOptions: []
         };

        for (let index = 0; index < 24; index++) {
            
            const option: IDropdownOption = {
                key: index,
                text: ('0' + index).slice(-2)
            };

            // if(index == currentElement.fromHr && currentElement.toMinutes) {
            //     option.disabled = currentElement.fromMinutes <= 58;
            // }

            if(index < currentElement.fromHr) {
                option.disabled = true;
            }

            if(index == currentElement.fromHr && currentElement.fromMinutes >= currentElement.toMinutes) {
                option.disabled = true;
            }

            result.hrOptions.push(option);
            
        }

        for (let index = 0; index < 60; index++) {
            
            const option: IDropdownOption = {
                key: index,
                text: ('0' + index).slice(-2)
            };

            if(index <= currentElement.fromMinutes && currentElement.fromHr == currentElement.toHr) {
                option.disabled = true;
            }

            result.minuteOptions.push(option);
            
        }

        return result;
    }

    private isAddNewButtonVisible(greetingSettings: IGreeting[]): boolean {
        const lastItem: IGreeting|null = greetingSettings.LastOrDefault();

        const isEmpty: boolean = !isset(lastItem);
        
        if(isEmpty) {
            return true;
        }

        return !(lastItem.toHr == 23 && lastItem.toMinutes == 59);
    }

    private getSettingsOrInitial(greetingSettings: IGreeting[]): IGreeting[] {
        let settings: IGreeting[] = cloneDeep(greetingSettings);
        const isEmpty: boolean = isNullOrEmpty(settings);
        
        if(isEmpty) {
            settings = [];
        }

        const newGreetingSettings: IGreeting = {
            fromHr: 0,
            fromMinutes: 0,
            toHr: 23,
            toMinutes: 59,
            text: "Hello"
        };

        if(isEmpty) {
            settings.push(newGreetingSettings);
        }

        return settings;
    }

    private addNewGreetingRow(greetingSettings: IGreeting[]): IGreeting[] {

        // let settings: IGreeting[] = cloneDeep(greetingSettings);
        // const isEmpty: boolean = isNullOrEmpty(settings);
        
        // if(isEmpty) {
        //     settings = [];
        // }

        

        // if(isEmpty) {
        //     settings.push(newGreetingSettings);
        //     return settings;
        // }

        const settings: IGreeting[] = this.getSettingsOrInitial(greetingSettings);
        
        const lastItem: IGreeting|null = this.state.greetingSettings.LastOrDefault();

        const newGreetingSettings: IGreeting = {
            fromHr: lastItem.toMinutes == 59 && lastItem.toHr != 23 ? lastItem.toHr + 1 : lastItem.toHr,
            fromMinutes: lastItem.toMinutes == 59 ? 0 : lastItem.toMinutes + 1,
            toHr: 23,
            toMinutes: 59,
            text: "Hello"
        };

        settings.push(newGreetingSettings);
        return settings;

    }

    private onTimeChanged(greetingSettings: IGreeting[]): IGreeting[] {

        const settings: IGreeting[] = this.getSettingsOrInitial(greetingSettings);

        for (let index = 0; index < settings.length; index++) {
            
            if(index == 0) {
                continue;
            }
            
            const prevElement = settings[index-1];
            const currentElement = settings[index];
            const nextElement = index + 1 < settings.length ? settings[index+1] : null;

            currentElement.fromHr = prevElement.toMinutes == 59 && prevElement.toHr != 23 ? prevElement.toHr + 1 : prevElement.toHr;
            currentElement.fromMinutes = prevElement.toMinutes == 59 ? 0 : prevElement.toMinutes + 1;

            //If current hours are in past ==> Set to same hours as prev
            if(prevElement.toHr >= currentElement.fromHr) {
                currentElement.fromHr = prevElement.toHr;
            }

            //If current "to hours" is not in future
            if(currentElement.fromHr > currentElement.toHr) {
                currentElement.toHr = currentElement.fromHr;
            }

            //If prev and current hours equal AND Minutes of current are in past
            if(prevElement.toHr == currentElement.fromHr && currentElement.fromMinutes <= prevElement.toMinutes) {
                
                if(prevElement.toMinutes == 59) {
                    currentElement.fromMinutes = 0;
                    currentElement.fromHr = prevElement.toHr + 1;
                }
                else {
                    currentElement.fromMinutes += prevElement.toMinutes + 1;
                }
            }

            
            
            if(currentElement.fromHr == currentElement.toHr && currentElement.fromMinutes >= currentElement.toMinutes) {
                
                if(currentElement.fromMinutes == 59) {
                    currentElement.toMinutes = 0;
                    currentElement.toHr = currentElement.toHr + 1;
                }
                else {
                    currentElement.toMinutes += 1;
                }
            }

            

            // if (isNullOrEmpty(nextElement)) {

            //     currentElement.toHr = 23;
            //     currentElement.toMinutes = 59;
            //     continue;
            // }




        }

        const lastIndexOfEoD: number = settings.IndexOf(s => s.toHr >= 23 && s.toMinutes >=58);

        if (lastIndexOfEoD >= 0 && settings.length-1 > lastIndexOfEoD) {
            settings.RemoveAt((lastIndexOfEoD + 1), settings.length-lastIndexOfEoD);
        }

        this.setState({
            greetingSettings: settings
        });

        return settings;
    }
}