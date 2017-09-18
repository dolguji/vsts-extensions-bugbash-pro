import { BaseStore } from "MB/Flux/Stores/BaseStore";

import { LongTextActionsHub } from "../Actions/ActionsHub";
import { ILongText } from "../Interfaces";
import { LongText } from "../ViewModels/LongText";

export class LongTextStore extends BaseStore<IDictionaryStringTo<LongText>, LongText, string> {
    constructor() {
        super();
        this.items = {};
    }

    public getItem(id: string): LongText {
        if (id) {
            return this.items[id.toLowerCase()];
        }
        return null;
    }

    protected initializeActionListeners() {
        LongTextActionsHub.FireStoreChange.addListener(() => {
            this.emitChanged();
        });

        LongTextActionsHub.AddOrUpdateLongText.addListener((longTextModel: ILongText) => {
            if (longTextModel) {
                this.items[longTextModel.id.toLowerCase()] = new LongText(longTextModel);
            }

            this.emitChanged();
        });
    }

    public getKey(): string {
        return "LongTextStore";
    }

    protected convertItemKeyToString(key: string): string {
        return key;
    }
}