import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import * as CoreClient from "TFS/Core/RestClient";
import { WebApiTeam } from "TFS/Core/Contracts";

export class TeamStore extends BaseStore<WebApiTeam[], WebApiTeam, string> {
    protected getItemByKey(idOrName: string): WebApiTeam {
         return Utils_Array.first(this.items, (team: WebApiTeam) => Utils_String.equals(team.id, idOrName, true) || Utils_String.equals(team.name, idOrName, true));
    }

    protected async initializeItems(): Promise<void> {
        const teams = await CoreClient.getClient().getTeams(VSS.getWebContext().project.id, 300);
        this.items = teams.sort((t1: WebApiTeam, t2: WebApiTeam) => {
            return Utils_String.ignoreCaseComparer(t1.name, t2.name);
        });
    }

    public getKey(): string {
        return "TeamStore";
    }
}