import Service = require("VSS/Service");
import { Constants, IBugBash } from "./Models";
import Utils_String = require("VSS/Utils/String");

export class BugBashManager {
    public static async readBugBashes(): Promise<IBugBash[]> {
        let dataService: IExtensionDataService = await VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData);

        try {            
            let items = await dataService.getDocuments(Constants.BUGBASH_DOCUMENT_COLLECTION_NAME);            
            items = items.filter((item: IBugBash) => Utils_String.equals(VSS.getWebContext().project.id, item.projectId, true));

            for(let item of items) {
                BugBashManager._translateDates(item);
            }

            return items;
        }
        catch (e) {
            return [];
        }
    }

    public static async readBugBash(bugBashId: string): Promise<IBugBash> {
        try {
            let dataService: IExtensionDataService = await VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData);
            let item = await dataService.getDocument(Constants.BUGBASH_DOCUMENT_COLLECTION_NAME, bugBashId);
            if (!Utils_String.equals(VSS.getWebContext().project.id, item.projectId, true) || !Utils_String.equals(VSS.getWebContext().team.id, item.teamId, true)) {
                return null;
            }

            BugBashManager._translateDates(item);

            return item;
        }
        catch (e) {
            return null;
        }
    }

    public static async writeBugBash(bugBashModel: IBugBash): Promise<IBugBash> {
        try {
            if (!Utils_String.equals(VSS.getWebContext().project.id, bugBashModel.projectId, true) || !Utils_String.equals(VSS.getWebContext().team.id, bugBashModel.teamId, true)) {
                return null;
            }

            let dataService: IExtensionDataService = await VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData);
            let item = await dataService.setDocument(Constants.BUGBASH_DOCUMENT_COLLECTION_NAME, bugBashModel);
            BugBashManager._translateDates(item);

            return item;
        }
        catch (e) {
            return null;
        }
    }

    public static async deleteBugBash(bugBashModel: IBugBash): Promise<boolean> {
        try {
            if (!Utils_String.equals(VSS.getWebContext().project.id, bugBashModel.projectId, true) || !Utils_String.equals(VSS.getWebContext().team.id, bugBashModel.teamId, true)) {
                return false;
            }

            let dataService: IExtensionDataService = await VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData);
            await dataService.deleteDocument(Constants.BUGBASH_DOCUMENT_COLLECTION_NAME, bugBashModel.id);
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