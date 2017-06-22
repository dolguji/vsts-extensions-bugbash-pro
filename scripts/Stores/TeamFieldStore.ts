import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { TeamFieldValues } from "TFS/Work/Contracts";
import { TeamContext } from "TFS/Core/Contracts";
import * as WorkClient from "TFS/Core/WorkClient";

export class TeamStore extends BaseStore<IDictionaryStringTo<TeamFieldValues>, TeamFieldValues, string> {
    constructor() {
        super();
        this.items = {};    
    }

    protected getItemByKey(teamId: string): TeamFieldValues {
         return this.items[teamId];
    }

    protected async initializeItems(): Promise<void> {
        return;   
    }

    public getKey(): string {
        return "TeamFieldStore";
    }

    public async ensureItem(teamId: string): Promise<boolean> {
        if (!this.itemExists(teamId)) {
            const projectId = VSS.getWebContext().project.id;

            const teamContext: TeamContext = {
                project: "",
                projectId: projectId,
                team: "",
                teamId: teamId
            };

            let key = `${projectId}_${teamId}`.toLowerCase();

            const teamFieldValues = await WorkClient.getClient().getTeamFieldValues(teamContext);
            this._onAdd(teamId, teamFieldValues);
        }
        else {
            this.emitChanged();
            return true;
        }
    }

    private _onAdd(teamId: string, item: TeamFieldValues): void {
        if (!item) {
            return;
        }

        if (!this.items) {
            this.items = {};
        }

        this.items[teamId] = item;

        this.emitChanged();
    }
}