import * as saveAs from "jszip/vendor/FileSaver";
import * as XLSX from "better-xlsx/dist/xlsx.min.js";
import { INameValuePair } from "./Interfaces";

export class ExcelExporter {
    private _data: INameValuePair[];

    constructor(data: INameValuePair[]) {
        this._data = data;
    }

    public export(fileName: string) {
        let total = 0;
        const file = new XLSX.File();
        const sheet = file.addSheet("Sheet1");
    
        const header = sheet.addRow();
        header.setHeightCM(0.8);
        const headerLabels = ["Row Labels", "Number of Bugs"];
        for (const headerLabel of headerLabels) {
            const hc = header.addCell();
            hc.value = headerLabel;
            
            this._applyHeaderStyle(hc);
        }
    
        for (const data of this._data) {
            total += data.value;

            const row = sheet.addRow();
            row.setHeightCM(0.7);
    
            const cell1 = row.addCell();
            cell1.value = data.name;
            this._applyTeamNameStyle(cell1);

            const cell2 = row.addCell();
            cell2.value = data.value;
            this._applyTeamNameStyle(cell2);
    
            for (const dataMember of (data.members || [])) {
                const memberRow = sheet.addRow();
                memberRow.setHeightCM(0.6);
    
                const memberCell1 = memberRow.addCell();
                memberCell1.value = dataMember.name;
                this._applyCommonStyle(memberCell1);
                memberCell1.style.align.indent = 2;

                const memberCell2 = memberRow.addCell();
                memberCell2.value = dataMember.value;
                this._applyCommonStyle(memberCell2);
            }
        }
    
        const footer = sheet.addRow();
        footer.setHeightCM(0.8);

        const fc1 = footer.addCell();
        fc1.value = "Grand Total";    
        this._applyHeaderStyle(fc1);

        const fc2 = footer.addCell();
        fc2.value = total;    
        this._applyHeaderStyle(fc2);

        sheet.col(0).width = 50;
        sheet.col(1).width = 15;
    
        file.saveAs("blob").then((content) => {
          saveAs(content, fileName);
        });
    }

    private _applyCommonStyle(cell) {
        cell.style.align.v = "center";
        cell.style.font.color = "ff000000";
        cell.style.font.size = 11;
        cell.style.font.name= "Calibri";
        cell.style.fill.patternType = "solid";
        cell.style.fill.bgColor = "ffffffff";
        cell.style.fill.fgColor = "ffffffff";
    }

    private _applyHeaderStyle(cell) {
        this._applyCommonStyle(cell);

        cell.style.align.h = "center";
        cell.style.font.size = 12;
        cell.style.fill.fgColor = "ffd9e2f3";
        cell.style.font.bold = true;        
    }

    private _applyTeamNameStyle(cell) {
        this._applyCommonStyle(cell);

        cell.style.font.bold = true;
        cell.style.border.bottom = "thin";
        cell.style.border.bottomColor = "ff6ba9e7";
        cell.style.border.top = "thin";
        cell.style.border.top = "ff6ba9e7";
    }
}