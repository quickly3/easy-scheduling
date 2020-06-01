import { Component, ChangeDetectorRef } from "@angular/core";
import * as Excel from "exceljs";
import { Workbook } from "exceljs";
import * as fs from "file-saver";
import { NgbModal } from "@ng-bootstrap/ng-bootstrap";

declare var electron: any;
const ipc = electron.ipcRenderer;

@Component({
  selector: "app-root",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.less"]
})
export class AppComponent {
  title = "scheduling";
  sch_len = 0;
  titles = [];
  data = [];
  members = [];
  schedules = [
    {name: "全班A", alias: "", time: "8:00-12:00；13:00-17:00；18:00-19:00", color: "#009900"},
    {name: "全班B", alias: "", time: "9:00-13:00; 14:00-18:00; 19:00-20:00", color: "#00CC33"},
    {name: "大夜班", alias: "", time: "20:00-24:00; 00:00-8:00", color: "#0099CC"},
    {name: "派单班", alias: "", time: "09:00-13:00; 14:00-17:30", color: "#CC9900"},
    {name: "正白A班", alias: "", time: "09:00-13:00 14:00-17:30", color: "#00FF66"},
    {name: "正白B班", alias: "", time: "09:00-12:00 13:00-17:30", color: "#FF9933"},
    {name: "休息", alias: "", time: "休息", color: "#FF0000"},
  ];
  schArray = [];
  exchageLog = [];
  month!: number;
  days = [];
  wb!: Workbook;
  main_worksheet: Excel.Worksheet;
  log_worksheet: Excel.Worksheet;
  logging = false;
  wb_changed = false;
  saving = false;
  stop_action = true;
  message = "";
  fix0Active = false;
  fix1Active = false;
  constructor(
    private cd: ChangeDetectorRef,
    private modalService: NgbModal
  ) {

  }

  // tslint:disable-next-line: use-life-cycle-interface
  ngOnInit(): void {
    ipc.on("save-buffer-done", (event, arg) => {
      const message = `存档成功`;
      this.saving = false;
      this.stop_action = false;
      this.wb_changed = false;
      this.cd.detectChanges();
    });

    ipc.on("send-cache-buffer", (event, buffer) => {
      console.log("接收buffer");
      this.stateInit();
      this.wb = new Excel.Workbook();
      this.wb.xlsx.load(buffer).then(() => {
        this.wbFormater();
        this.cd.detectChanges();
      });
    });

    this.onLoadCache(undefined);
  }

  stateInit() {
    this.wb_changed = false;
    this.saving = false;
    this.stop_action = true;
  }

  onLoadCache(saveConfirm) {

    if (saveConfirm) {
      this.message = "确定重新载入保存的调班表？";
      this.modalService.open(saveConfirm, {ariaLabelledBy: "modal-basic-title"}).result.then((resp) => {
        this.message = "";

        if ("sure" === resp) {
          ipc.send("load-cache", "ping");
        }
      }, (resean) => {});
    } else {
      ipc.send("load-cache", "ping");
    }
  }

  onSave(saveConfirm): void {

    this.message = "确定保存？";
    this.modalService.open(saveConfirm, {ariaLabelledBy: "modal-basic-title"}).result.then((resp) => {
      if ("sure" === resp) {
        this.message = "";
        this.saving = true;
        this.stop_action = true;
        this.wb_changed = false;
        this.wb.xlsx.writeBuffer().then((buffer) => {
          ipc.send("save-buffer", {buffer: buffer});
        });
      }
    }, (resean) => {});
  }

  download(notice): void {

    if (this.wb_changed) {
      this.message = "下载前请保存或重置未保存的修改。";
      this.modalService.open(notice, {ariaLabelledBy: "modal-basic-title"}).result.then(() => {
        this.message = "";
      }, (resean) => {});
    } else {
      this.wb.xlsx.writeBuffer().then((data) => {
        const blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const date = new Date();
        const file = "班次" + date.toLocaleString("zh", {hour12: false}) + ".xlsx";
        fs.saveAs(blob, file);
      });
    }
  }



  wbFormater() {
    this.data = [];
    this.exchageLog = [];

    // play with workbook and worksheet now
    this.main_worksheet = this.wb.getWorksheet(1);
    this.log_worksheet = this.wb.getWorksheet(2);

    this.main_worksheet.eachRow( (row, rowNumber) => {
      const table_row = [];

      row.eachCell(cell => {
        const cell_data: any = {
          address: cell.address,
          value: cell.value
        };

        if (cell.style && cell.style.fill) {
          if ((<Excel.FillPattern>cell.style.fill).fgColor) {
            const colorObj = (<Excel.FillPattern>cell.style.fill).fgColor;
            if (colorObj.argb) {
              cell_data.fgColor = colorObj.argb;
            }
            if (colorObj.theme && colorObj.theme === 7) {
              cell_data.fgColor = "FFffda65";
            }
          }
        }

        if (cell_data.fgColor) {
          cell_data.fgColor = "#" + cell_data.fgColor.slice(2);
        }

        table_row.push(cell_data);
      });
      this.data.push(table_row);
    });

    this.log_worksheet.eachRow( (row, rowNumber) => {
      const logArr = [];
      if (row.getCell(1).value !== "操作日期") {
        row.eachCell(cell => {
          logArr.push(cell.value);
        });
        this.exchageLog.push(logArr);
      }

    });

    this.dataAnalysis();
    this.stop_action = false;
  }

  readExcel(event) {
    this.wb = new Excel.Workbook();
    const target: DataTransfer = <DataTransfer>(event.target);
    if (target.files.length !== 1) {
      return false;
    }

    this.wb_changed = true;

    const arryBuffer = new Response(target.files[0]).arrayBuffer();
    arryBuffer.then( (data) => {
      this.wb.xlsx.load(data)
        .then(() => {
          this.wbFormater();
        });
    });
  }

  dataAnalysis() {
    this.findScheduleTable();
  }

  openModel() {

    const tpl = `
      <div class="modal-body">
        <span class="text-danger">读取中...</span>
      </div>
      `;

    this.modalService.open(tpl, {ariaLabelledBy: "modal-basic-title", backdrop: "static"}).result.then((result) => {

    }, (reason) => {
    });
  }

  cellChecked() {

  }

  changeSch(notice, exchConfirm) {
    let checkedCount = 0;
    this.schArray = [];
    this.data.map(( row, rowI) => {
      row.map((cell, cellI) => {
        if (cell.checked) {
          this.schArray.push(JSON.parse(JSON.stringify(cell)));
          checkedCount++;
        }
      });
    });

    if (checkedCount < 2) {
      this.message = "请选择要调换的两个班次。";
      this.modalService.open(notice, {ariaLabelledBy: "modal-basic-title"}).result.then(() => {
        this.message = "";
      }, (resean) => {});
      return false;
    }

    this.schArray[0].fixValue = this.schArray[1].value;
    this.schArray[1].fixValue = this.schArray[0].value;
    this.fix0Active = false;
    this.fix1Active = false;

    this.modalService.open(exchConfirm, {ariaLabelledBy: "modal-basic-title", size: "lg"}).result.then((result) => {
      if ("sure" === result) {
        this.logging = true;
        const logArr = [];

        const date = new Date();
        const logDate = "【" + date.toLocaleString("zh", {hour12: false}) + "】";
        logArr.push(logDate);
        logArr.push(this.schArray[0].member);
        logArr.push(this.schArray[0].date);
        logArr.push(this.schArray[0].value);
        logArr.push(this.schArray[0].fixValue);
        logArr.push(" --- ");
        logArr.push(this.schArray[1].member);
        logArr.push(this.schArray[1].date);
        logArr.push(this.schArray[1].value);
        logArr.push(this.schArray[1].fixValue);

        // const log = logArr.join(" ");
        this.changeCells();
        this.exchageLog.push(logArr);
        this.addLogWs(logArr);
        this.wb_changed = true;
        this.logging = false;
      }
    }, (resean) => {});
  }

  showExchLogs(logtpl) {
    this.modalService.open(logtpl, {ariaLabelledBy: "modal-basic-title", size: "lg"});
  }

  changeCells() {
    this.changeWb();
    this.wbFormater();
    // this.cd.detectChanges();
  }

  changeWb() {

    const add1 = this.schArray[0].address;
    const add2 = this.schArray[1].address;

    const cell1 = this.main_worksheet.getCell(add1);
    const cell2 = this.main_worksheet.getCell(add2);

    const cell1_value = cell1.value;
    const cell1_style = cell1.style ;

    const fixStyle: any = {
      fill: {type: "pattern", pattern: "solid", fgColor: {argb: "FFCC0099"}},
      border: {
        "left": {"style": "thin", "color": {"indexed": 64}},
        "right": {"style": "thin", "color": {"indexed": 64}},
        "top": {"style": "thin", "color": {"indexed": 64}},
        "bottom": {"style": "thin", "color": {"indexed": 64}}
      }
    };

    if (this.schArray[0].fixValue !== this.schArray[1].value) {
      cell1.value = this.schArray[0].fixValue;
      cell1.style = fixStyle;
    } else {
      cell1.value = cell2.value;
      cell1.style = cell2.style;
    }

    if (this.schArray[1].fixValue !== this.schArray[0].value) {
      cell2.value = this.schArray[1].fixValue;
      cell2.style = fixStyle;
    } else {
      cell2.value = cell1_value;
      cell2.style = cell1_style;
    }
    // (<Excel.FillPattern>cell2.style.fill).fgColor = cell1_fgColor;

  }

  addLogWs(logArr) {
    if (this.log_worksheet.rowCount === 0) {
      const header = ["操作日期", "姓名", "日期", "调班前", "调班后", " --- ", "姓名", "日期", "调班前", "调班后"];
      this.log_worksheet.addRow(header);
    }
    this.log_worksheet.addRow(logArr);
  }

  getCheckedCellCnt() {
    let checkedCount = 0;
    this.data.map(( row, rowI) => {
      row.map((cell, cellI) => {
        if (cell.checked) {
          checkedCount++;
        }
      });
    });
    return checkedCount;
  }

  cellClick(e, _cell) {
    const checkedCount = this.getCheckedCellCnt();
    if (checkedCount >= 2 && _cell.checked === false) {
      e.stopPropagation();
      e.preventDefault();
    }
  }

  findScheduleTable() {
    const member_start = 1;
    const member_end = 0;

    this.data.map(( row, rowI) => {
      if (row[0] && row[0].value && !isNaN(parseInt(row[0].value))) {
        row.map((cell, cellI) => {
          if (cellI > 2) {
            cell.showCK = true;
            cell.checked = false;
            cell.member = row[1].value;
            cell.date = this.data[0][cellI].value;
          }
        });
      }
    });

    // this.sch_len = this.data[0].length - 3;
    // this.days = this.titles.slice(3);

    // for (let i = member_start; i <= member_end; i++) {
    //   const member = {
    //     id: this.data[i][0],
    //     name: this.data[i][1],
    //     number: this.data[i][2],
    //     schedules: []
    //   };

    //   const schedules = this.data[i].slice(3);
    //   const schs = this.days.map((day, i) => {
    //     return {day: day, sch_alias: schedules[i]};
    //   });

    //   member.schedules = schs;
    //   this.members.push(member);
    // }
  }
}
