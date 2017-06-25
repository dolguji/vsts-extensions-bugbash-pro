import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

import * as GitClient from "TFS/VersionControl/GitRestClient";
import { GitRepository } from "TFS/VersionControl/Contracts";

export class GitRepoStore extends BaseStore<GitRepository[], GitRepository, string> {
    protected getItemByKey(id: string): GitRepository {
         return Utils_Array.first(this.items, (repo: GitRepository) => Utils_String.equals(repo.id, id, true));
    }

    protected async initializeItems(): Promise<void> {
        this.items = await GitClient.getClient().getRepositories(VSS.getWebContext().project.id);
    }

    public getKey(): string {
        return "GitRepoStore";
    }
}