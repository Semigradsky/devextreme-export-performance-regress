import React from "react";
import Button from "devextreme-react/button";
import DataGrid, {
  Scrolling,
  Column,
  Export
} from "devextreme-react/data-grid";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { exportDataGrid } from "devextreme/excel_exporter";

import service from "./data.js";

let s = 123456789;
function random() {
  s = (1103515245 * s + 12345) % 2147483647;
  return s % (10 - 1);
}

function generateData(count) {
  var i;
  var surnames = [
    "Smith",
    "Johnson",
    "Brown",
    "Taylor",
    "Anderson",
    "Harris",
    "Clark",
    "Allen",
    "Scott",
    "Carter"
  ];
  var names = [
    "James",
    "John",
    "Robert",
    "Christopher",
    "George",
    "Mary",
    "Nancy",
    "Sandra",
    "Michelle",
    "Betty"
  ];
  var gender = ["Male", "Female"];
  var items = [],
    startBirthDate = Date.parse("1/1/1975"),
    endBirthDate = Date.parse("1/1/1992");

  for (i = 0; i < count; i++) {
    var birthDate = new Date(
      startBirthDate +
        Math.floor((random() * (endBirthDate - startBirthDate)) / 10)
    );
    birthDate.setHours(12);

    var nameIndex = random();
    var item = {
      id: i + 1,
      firstName: names[nameIndex],
      lastName: surnames[random()],
      gender: gender[Math.floor(nameIndex / 5)],
      birthDate: birthDate
    };
    items.push(item);
  }
  return items;
}

const dataSource = generateData(9999);

class App extends React.Component {
  constructor(props) {
    super(props);
    this.employees = service.getEmployees();
    this.export = this.export.bind(this);
    this.gridRef = React.createRef();

    DataGrid.defaultProps = {
      loadPanel: {
        enabled: false
      }
    };
  }

  render() {
    return (
      <React.Fragment>
        <Button onClick={this.export}>Export</Button>
        <DataGrid
          ref={this.gridRef}
          id="gridContainer"
          dataSource={dataSource}
          showBorders={true}
          onExporting={this.onExporting}
        >
          <Column
            name="column0"
            dataField="firstName"
            cellRender={(x) => {
              return <div>{x.value}</div>;
            }}
          />
          <Column
            name="column1"
            dataField="birthDate"
            cellRender={(x) => {
              return <div>{x.text}</div>;
            }}
            dataType="date"
            width={100}
          />
          <Column
            name="column2"
            dataField="firstName"
            cellRender={(x) => {
              return <div>{x.value}</div>;
            }}
          />
          <Column
            name="column3"
            dataField="birthDate"
            cellRender={(x) => {
              return <div>{x.text}</div>;
            }}
            dataType="date"
            width={100}
          />
          <Column
            name="column4"
            dataField="firstName"
            cellRender={(x) => {
              return <div>{x.value}</div>;
            }}
          />
          <Column
            name="column5"
            dataField="birthDate"
            cellRender={(x) => {
              return <div>{x.text}</div>;
            }}
            dataType="date"
            width={100}
          />
          <Column
            name="column6"
            dataField="firstName"
            cellRender={(x) => {
              return <div>{x.value}</div>;
            }}
          />
          <Column
            name="column6"
            dataField="birthDate"
            cellRender={(x) => {
              return <div>{x.text}</div>;
            }}
            dataType="date"
            width={100}
          />

          <Scrolling mode="virtual" />
          <Export enabled={false} />
        </DataGrid>
      </React.Fragment>
    );
  }

  async export() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Main sheet");

    await exportDataGrid({
      component: this.gridRef.current.instance,
      worksheet: worksheet
    });

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(
      new Blob([buffer], { type: "application/octet-stream" }),
      "DataGrid.xlsx"
    );
  }
}

export default App;
