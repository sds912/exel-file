import { EtatFinancierModel1Service } from './services/etat-financier-model1.service';
import { Component } from '@angular/core';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
import { ExportExcelService } from './services/export-excel.service';
import * as XLSX from 'xlsx';
import { Router } from '@angular/router';
import { InfoApi, Configuration } from "groupdocs-viewer-cloud";


const excelData = {
  title: '',
  headers: [''],
  data: []
}

const title = excelData.title;
const header = excelData.headers
const data = excelData.data;

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  data: any;
  fileName: 'exel.xlsx'
  



 
  

  constructor(
    public ete: ExportExcelService, 
    private router: Router,
    private _etatFinancierModel1 : EtatFinancierModel1Service, ) { }



export(): void {
    this._etatFinancierModel1.exportmodel();

}

 
    

    
}




