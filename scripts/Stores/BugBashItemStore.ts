import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { Store } from "VSS/Flux/Store";

import { IBugBash } from "../Models";
import { ActionsHub } from "../Actions/ActionsCreator";

export interface IBugBashItemStore {
    isLoaded(): boolean;
    itemExists(id: string): boolean;
    getItem(id: string): IBugBash;
    getAll(): IBugBash[];
}

export class BugBashItemStore extends Store implements IBugBashItemStore {
    private _items: IBugBash[];

    constructor(actions: ActionsHub) {
        super();

        this._items = null;

        actions.InitializeBugBashItems.addListener((items: IBugBash[]) => {
            if (!items) {
                this.emitChanged();
            }
            this._onAdd(items);
        });

        actions.ClearBugBashItems.addListener(() => {
            this._items = null;
            this.emitChanged();
        });
        
        actions.BugBashItemDeleted.addListener((item: IBugBash) => {
            this._onRemove(item);
        });

        actions.BugBashItemAdded.addListener((item: IBugBash) => {
            this._onAdd(item);
        });

        actions.BugBashItemUpdated.addListener((item: IBugBash) => {
            this._onAdd(item);
        });
    }

    public isLoaded(): boolean {
        return this._items ? true : false;
    }

    public itemExists(id: string): boolean {
        return this._getById(id) ? true : false;
    }

    public getItem(id: string): IBugBash {
        return this._getById(id);
    }

    public getAll(): IBugBash[] {
        return this._items || [];
    }

    private _getById(id: string): IBugBash {
        if (!this.isLoaded()) {
            return null;
        }

        return Utils_Array.first(this._items, (item: IBugBash) => Utils_String.equals(item.id, id, true));
    }

    private _onAdd(items: IBugBash | IBugBash[]): void {
        if (!items) {
            return;
        }

        if (!this._items) {
            this._items = [];
        }

        if (Array.isArray(items)) {
            for (let item of items) {
                this._addItem(item);
            }
        }
        else {
            this._addItem(items);
        }

        this.emitChanged();
    }

    private _onRemove(items: IBugBash | IBugBash[]): void {
        if (!items || !this.isLoaded()) {
            return;
        }

        if (Array.isArray(items)) {
            for (let item of items) {
                this._removeItem(item);
            }
        }
        else {
            this._removeItem(items);
        }

        this.emitChanged();
    }

    private _addItem(item: IBugBash): void {
        let existingItemIndex = Utils_Array.findIndex(this._items, (existingItem: IBugBash) => Utils_String.equals(item.id, existingItem.id, true));
        if (existingItemIndex != -1) {
            // Overwrite the item data
            this._items[existingItemIndex] = item;
        }
        else {
            this._items.push(item);
        }
    }

    private _removeItem(item: IBugBash): void {
        let existingItemIndex = Utils_Array.findIndex(this._items, (existingItem: IBugBash) => Utils_String.equals(item.id, existingItem.id, true));

        if (existingItemIndex != -1) {
            this._items.splice(existingItemIndex, 1);
        }
    }
}