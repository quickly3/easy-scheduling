import { Component } from "@angular/core";
import * as Excel from "exceljs";
import { Workbook } from "exceljs";
import * as fs from "file-saver";


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
  month!: number;
  days = [];
  wb!: Workbook;

  ngOnInit(): void {
    ipc.on("asynchronous-reply", function (event, arg) {
      const message = `异步消息回复: ${arg}`;
      console.log(message);

    });

  }

  onSave(): void {
    this.wb.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      fs.saveAs(blob, "test.xlsx");
    });
  }

  onIpc(): void {
    this.wb.xlsx.writeBuffer().then((buffer)=>{
      ipc.send("asynchronous-message", {buffer:buffer});
    })


  }

  readExcel(event) {
    this.wb = new Excel.Workbook();
    const target: DataTransfer = <DataTransfer>(event.target);
    if (target.files.length !== 1) {
      // console.log("Cannot use multiple files");
      return false;
    }

    /**
     * Final Solution For Importing the Excel FILE
     */

    const arryBuffer = new Response(target.files[0]).arrayBuffer();
    arryBuffer.then( (data) => {
      this.wb.xlsx.load(data)
        .then(() => {

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
          });

        });
    });
  }

  dataAnalysis() {
    this.titles = this.data[0];
    this.findScheduleTable();
  }

  findScheduleTable() {
    const member_start = 1;
    let member_end = 0;
    for (const i in this.data) {
      if (this.data[i][0] === undefined) {
        // tslint:disable-next-line: radix
        member_end = parseInt(i) - 1;
        break;
      }
    }

    this.sch_len = this.data[0].length - 3;
    this.days = this.titles.slice(3);

    for (let i = member_start; i <= member_end; i++) {
      const member = {
        id: this.data[i][0],
        name: this.data[i][1],
        number: this.data[i][2],
        schedules: []
      };

      const schedules = this.data[i].slice(3);
      const schs = this.days.map((day, i) => {
        return {day: day, sch_alias: schedules[i]};
      });

      member.schedules = schs;
      this.members.push(member);
    }
  }
}
