import { IPlaceholderResolver } from '@spfxappdev/framework/lib/spfxframework/placeholder/PlaceholderHandler';
import { IGreeting } from './TexteditorWebPart';
import { isNullOrEmpty } from '@spfxappdev/utility';
import "@spfxappdev/utility/lib/extensions/ArrayExtensions";

export class GreetingPlaceholderResolver implements IPlaceholderResolver {

    

    public placeHolderData: any = {
        Greeting: ""
    };

    public constructor(settings: IGreeting[]) {

        if(isNullOrEmpty(settings)) {
            return;
        }

        const fromDate: Date = new Date();
        const currentDate: Date = new Date();
        const toDate: Date = new Date();

        const setting: IGreeting = settings.FirstOrDefault(s => fromDate.setHours(s.fromHr, s.fromMinutes) && currentDate >= fromDate && toDate.setHours(s.toHr, s.toMinutes) && currentDate <= toDate);
        

        if(isNullOrEmpty(setting)) {
            return;
        }

        this.placeHolderData.Greeting = setting.text;        
    }
}