let ExcelData = null;
document.addEventListener("DOMContentLoaded", (event) => {
  //   const files = [];
  window.versions.getFiles("get excel files");
  window.versions.getFileList((event, files) => {
    return generateFilesRows(files);
  });
});

function generateFilesRows(files) {
  const tbody = document.getElementById("filelist-body");

  let html = "";
  files.forEach((file, index) => {
    html += `<tr>
        <td class="padding">${index + 1}</td>
        <td class="padding hover" onclick="readFile('${file.name}')">${
      file.name
    }</td>
      </tr>`;
  });
  tbody.insertAdjacentHTML("beforeend", html);
}

function readFile(fileName) {
  window.versions.sendFile(fileName);
  window.versions.getFile((event, jsonData, sheets) => {
    ExcelData = jsonData;

    const container = document.getElementById("file-view-container");
    const ListContainer = document.getElementById("file-list-container");
    const sheetContainer = document.getElementById("sheet-list");

    let html = "";

    sheets.forEach((sheet, sheetIndex) => {
      html += `<button type="button" onclick="readSheet(event, '${sheetIndex}', '${fileName}')">${sheet}</button>`;
    });

    sheetContainer.innerHTML = "";
    sheetContainer.insertAdjacentHTML("beforeend", html);

    createTable(0, fileName);
    ListContainer.classList.toggle("d-none");
    container.classList.toggle("d-none");
  });
}

function readSheet(e, sheetIndex, fileName) {
  e.preventDefault();
  createTable(Number(sheetIndex), fileName);
}

function createTable(sheetIndex = 0, fileName) {
  let html = `<table class="table">
               <thead>
                <tr>`;

  console.log(ExcelData[sheetIndex][0]);
  ExcelData[sheetIndex][0].forEach((header) => {
    html += `<th class="padding  backgroundColor">${header}</th>`;
  });
  html += `</tr>
          </thead>
        <tbody>`;

  ExcelData[sheetIndex].forEach((row, rowIndex) => {
    if (rowIndex !== 0) {
      html += `<tr>`;
      row.forEach((column, columnIndex) => {
        html += `<td class="border">
                   <input class="inputValue"
                    onchange="readColumn(event,'${rowIndex}','${columnIndex}','${fileName}','${sheetIndex}')" 
                    type="text" value="${column ? column : ""}" />
                  </td>`;
      });
      html += "</tr>";
    }
  });
  html += `</tbody></table>`;

  const tableContainer = document.getElementById("table-container");
  tableContainer.innerHTML = "";
  tableContainer.insertAdjacentHTML("beforeend", html);
}

function back() {
  const container = document.getElementById("file-view-container");
  const ListContainer = document.getElementById("file-list-container");
  ListContainer.classList.toggle("d-none");
  container.classList.toggle("d-none");
}

function readColumn(event, rowIndex, columnIndex, fileName, sheetIndex) {
  console.log(event);
  console.log(rowIndex);
  console.log(columnIndex);
  console.log(sheetIndex);
  console.log(fileName);
  let columnValue = event.target.value;
  if (columnValue !== "") {
    // console.log(ExcelData);
    // console.log(ExcelData[Number(rowIndex)]);
    // console.log(ExcelData[Number(rowIndex) + 1]);
    ExcelData[Number(sheetIndex)][Number(rowIndex)][Number(columnIndex)] =
      columnValue;
    // console.log(ExcelData);
    window.versions.sendUpdatedData(ExcelData, fileName);
    window.versions.getUpdatedData((event, msg) => {
      document.getElementById("message").innerText = msg;
      // window.alert(msg);
      console.log(msg);
    });
  }
}
