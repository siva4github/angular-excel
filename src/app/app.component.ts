import { Component } from '@angular/core';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'angular-excel';

  constructor() {
    this.excelRead();
  }

  async excelRead() {

    const srcWB = this.getSourceExcel();
    const destWB = this.getDestinationExcel();

    if ((await srcWB).worksheets.length == (await destWB).worksheets.length) {
      console.log('Source Work Book and Destination Work Book sheets count matched');

      for (let k = 0; k < (await srcWB).worksheets.length; k++) {
        
        console.log('Processing Sheet Index - ' + k);
        
        let srcSheet = (await srcWB).worksheets[k];
        let destSheet = (await destWB).worksheets[k];

        if (srcSheet.actualRowCount == destSheet.actualRowCount && srcSheet.actualColumnCount == destSheet.actualColumnCount) {
          for (let i = 1; i <= srcSheet.rowCount; i++) {
            for (let j = 1; j <= destSheet.rowCount; j++) {

              if (srcSheet.getRow(i).getCell(j).toString() != destSheet.getRow(i).getCell(j).toString()) {
                console.log('Data not matched at row: ' + i + ', column: ' + j);
                console.log(srcSheet.getRow(i).getCell(j));
                console.log(destSheet.getRow(i).getCell(j));
                destSheet.getRow(i).getCell(j).fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFFF0000' },
                  bgColor: { argb: 'FFFF7D7D' }
                };
                destSheet.getRow(i).commit();
              }
              else {
                console.log('Data matched at row: ' + i + ', column: ' + j, srcSheet.getRow(i).getCell(j).toString(), destSheet.getRow(i).getCell(j).toString());
              }
            }
          }
        }
        else {
          console.log('Source Excel - Sheet: ' + srcSheet.name + ' Rows & Columns count not matched with Destination Excel - Sheet: ' + destSheet.name);
          console.log(srcSheet.actualRowCount, srcSheet.actualColumnCount, destSheet.actualRowCount, destSheet.actualColumnCount);
        }

      }

    }
    else {
      console.log('Source Work Book and Destination Work Book sheets count not matched');
    }

    // if (srcSheet.actualRowCount == destSheet.actualRowCount && srcSheet.actualColumnCount == destSheet.actualColumnCount) {
    //   for (var i = 1; i <= srcSheet.actualRowCount; i++) {
    //     for (var j = 1; j <= srcSheet.actualColumnCount; j++) {
    //       if (srcSheet.getRow(i).getCell(j).toString() != destSheet.getRow(i).getCell(j).toString()) {
    //         console.log('Data not matched at row: ' + i + ', column: ' + j);
    //         console.log(destSheet.getRow(i).getCell(j));
    //         destSheet.getRow(i).getCell(j).fill = {
    //           type: 'pattern',
    //           pattern: 'solid',
    //           fgColor: { argb: 'FFFF0000' },
    //           bgColor: { argb: 'FFFF7D7D' }
    //         };
    //         destSheet.getRow(i).commit();
    //       }
    //       else {
    //         console.log('Data matched at row: ' + i + ', column: ' + j);
    //       }
    //     }
    //     console.log();
    //   }

    //   //generate output file after completion of loop
    //   (await destWB).xlsx.writeBuffer().then((data) => {
    //     let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    //     //fs.saveAs(blob, 'Destination_up.xlsx');
    //   });
    // }
    // else {
    //   console.log('Excel sheet not matched');
    // }
  }

  async getSourceExcel() {
    const resp = await fetch('../assets/Source.xlsx');
    const buf = await resp.arrayBuffer();
    const wb = new Workbook();

    return wb.xlsx.load(buf);
  }

  async getDestinationExcel() {
    const resp = await fetch('../assets/Destination.xlsx');
    const buf = await resp.arrayBuffer();
    const wb = new Workbook();

    return wb.xlsx.load(buf);
  }

}
