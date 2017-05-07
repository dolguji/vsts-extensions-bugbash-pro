import { Constants, IBugBash } from "./Models";
import Utils_String = require("VSS/Utils/String");

import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

export class BugBashManager {
    public static async readBugBashes(): Promise<IBugBash[]> {        
        let items = await ExtensionDataManager.readDocuments<IBugBash>("bugbashes", false);
        items = items.filter((item: IBugBash) => Utils_String.equals(VSS.getWebContext().project.id, item.projectId, true));
        for(let item of items) {
            BugBashManager._translateDates(item);
        }

        return items;        
    }

    public static async readBugBash(bugBashId: string): Promise<IBugBash> {
        let item = await ExtensionDataManager.readDocument<IBugBash>("bugbashes", bugBashId, null, false);
        if (item) {
            if (!Utils_String.equals(VSS.getWebContext().project.id, item.projectId, true)) {
                return null;
            }

            BugBashManager._translateDates(item);
        }

        return item;
    }

    public static async writeBugBash(bugBashModel: IBugBash): Promise<IBugBash> {
        if (!Utils_String.equals(VSS.getWebContext().project.id, bugBashModel.projectId, true)) {
            return null;
        }
        
        let item = await ExtensionDataManager.writeDocument<IBugBash>("bugbashes", bugBashModel, false);
        BugBashManager._translateDates(item);

        return item;
    }

    public static async deleteBugBash(bugBashModel: IBugBash): Promise<boolean> {
        if (!Utils_String.equals(VSS.getWebContext().project.id, bugBashModel.projectId, true)) {
            return null;
        }

        try {        
            await ExtensionDataManager.deleteDocument<IBugBash>("bugbashes", bugBashModel.id, false);
            return true;
        }
        catch (e) {
            return false;
        }
    }

    private static _translateDates(item: any) {
        if (typeof item.startTime === "string") {
            if (item.startTime.trim() === "") {
                item.startTime = undefined;
            }
            else {
                item.startTime = new Date(item.startTime);
            }
        }
        if (typeof item.endTime === "string") {
            if (item.endTime.trim() === "") {
                item.endTime = undefined;
            }
            else {
                item.endTime = new Date(item.endTime);
            }
        }
    }
}