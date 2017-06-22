import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import * as CoreClient from "TFS/Core/RestClient";
import { WebApiTeam } from "TFS/Core/Contracts";

export class TeamStore extends BaseStore<WebApiTeam[], WebApiTeam, string> {
    constructor() {
        super();
        this.items = [];    
    }
    
    protected getItemByKey(id: string): WebApiTeam {
         return Utils_Array.first(this.items, (item: WebApiTeam) => Utils_String.equals(item.id, id, true));
    }

    protected async initializeItems(): Promise<void> {
        this.items = await CoreClient.getClient().getTeams(VSS.getWebContext().project.id);
    }

    public getKey(): string {
        return "TeamStore";
    }
}