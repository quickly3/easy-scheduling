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

  constructor(
    private cd: ChangeDetectorRef,
    private modalService: NgbModal
  ) {

  }

  // tslint:disable-next-line: use-life-cycle-interface
  ngOnInit(): void {
    ipc.on("asynchronous-reply", (event, arg) => {
      const message = `异步消息回复: ${arg}`;
      console.log(message);
    });

    ipc.on("send-cache-buffer", (event, buffer) => {
      this.wb = new Excel.Workbook();
      this.wb.xlsx.load(buffer).then(() => {
        this.wbFormater();
        this.cd.detectChanges();
      });
    });
    this.onLoadCache();
  }


  onLoadCache() {
    ipc.send("load-cache", "ping");
  }

  download(): void {
    this.wb.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const date = new Date();
      const file = "班次" + date.toLocaleString("zh", {hour12: false}) + ".xlsx";
      fs.saveAs(blob, file);
    });
  }

  onIpc(): void {
    this.wb.xlsx.writeBuffer().then((buffer) => {
      ipc.send("asynchronous-message", {buffer: buffer});
    });
  }

  wbFormater() {
    this.data = [];
    // play with workbook and worksheet now
    const worksheet = this.wb.getWorksheet(1);
    // console.log("rowCount: ", worksheet.rowCount);

    worksheet.eachRow( (row, rowNumber) => {
      const table_row = [];

      row.eachCell(cell => {

        const cell_data: any = {
          address: cell.address,
          value: cell.value
        };

        if (cell.style && cell.style.fill) {
          if ((<Excel.FillPattern>cell.style.fill).fgColor) {
            cell_data.fgColor = (<Excel.FillPattern>cell.style.fill).fgColor.argb;
          }
        }

        if (cell_data.fgColor) {
          cell_data.fgColor = "#" + cell_data.fgColor.slice(2);
        }

        table_row.push(cell_data);
      });
      this.data.push(table_row);
      this.dataAnalysis();
    });

  }

  readExcel(event) {
    this.wb = new Excel.Workbook();
    const target: DataTransfer = <DataTransfer>(event.target);
    if (target.files.length !== 1) {
      // console.log("Cannot use multiple files");
      return false;
    }

    const arryBuffer = new Response(target.files[0]).arrayBuffer();
    arryBuffer.then( (data) => {
      this.wb.xlsx.load(data)
        .then(() => {
          this.wbFormater();
          this.cd.detectChanges();
        });
    });
  }

  dataAnalysis() {
    this.titles = this.data[0];
    this.findScheduleTable();
  }

  openModel() {

    const tpl = `
      <div class="modal-body">
        <span class="text-danger">读取中...</span>
      </div>
      `;

    this.modalService.open(tpl, {ariaLabelledBy: "modal-basic-title", backdrop: "static"}).result.then((result) => {
      console.log("done");
    }, (reason) => {
      console.log("close");
    });
  }

  cellChecked(){


  }


  changeSch(notice,exchConfirm){
    let checkedCount = 0
    this.schArray = [];
    this.data.map(( row, rowI) => {
      row.map((cell,cellI)=>{
        if(cell.checked){
          this.schArray.push(cell)
          checkedCount++;
        }
      })
    })

    if(checkedCount < 2){
      this.modalService.open(notice, {ariaLabelledBy: "modal-basic-title"});
      return false;
    }

    this.modalService.open(exchConfirm, {ariaLabelledBy: "modal-basic-title", backdrop: "static"}).result.then((result) => {
      if("sure" === result){
        const logArr = [];


        logArr.push(this.schArray[0].member);
        logArr.push(this.schArray[0].date);
        logArr.push(this.schArray[0].value);
        logArr.push(" --- ");
        logArr.push(this.schArray[1].member);
        logArr.push(this.schArray[1].date);
        logArr.push(this.schArray[1].value);

        const log = logArr.join(" ");
        this.changeCells(this.schArray[0].address, this.schArray[1].address);
        this.exchageLog.push(log);
        console.log(this.exchageLog);
      }
    });
  }

  showExchLogs(logtpl){
    this.modalService.open(logtpl, {ariaLabelledBy: "modal-basic-title", backdrop: "static", size: "lg"});
  }

  changeCells(add1,add2){

  }

  getCheckedCellCnt(){
    let checkedCount = 0
    this.data.map(( row, rowI) => {
      row.map((cell,cellI)=>{
        if(cell.checked){
          checkedCount++;
        }
      })
    })
    return checkedCount;
  }

  cellClick(e,_cell){
    const checkedCount = this.getCheckedCellCnt();
    if(checkedCount >=2 && _cell.checked === false){
      e.stopPropagation();
      e.preventDefault();
    }
  }

  findScheduleTable() {
    const member_start = 1;
    const member_end = 0;

    this.data.map(( row, rowI) => {
      if (row[0] && row[0].value && !isNaN(parseInt(row[0].value))) {
        row.map((cell,cellI)=>{
          if(cellI > 2){
            cell.showCK = true;
            cell.checked = false;
            cell.member = row[1].value;
            cell.date = this.data[0][cellI].value;
          }
        })
      }
    })

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
