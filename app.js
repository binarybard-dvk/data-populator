async function readTemplate(arrayBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  return workbook;
}

function replacePlaceholders(worksheet, data) {
  worksheet.eachRow((row, _rowNumber) => {
    row.eachCell((cell, _colNumber) => {
      if (
        typeof cell.value === "string" &&
        cell.value.startsWith("{{") &&
        cell.value.endsWith("}}")
      ) {
        const key = cell.value.slice(2, -2).trim();
        if (data.hasOwnProperty(key) && !Array.isArray(data[key])) {
          cell.value = data[key] ?? "";
        }
      }
    });
  });
}

function insertTableData(worksheet, tableStartPlaceholder, tableData) {
  let tableStartRow;
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, _colNumber) => {
      if (cell.value === tableStartPlaceholder) {
        tableStartRow = rowNumber;
        console.log(
          `Found placeholder '${tableStartPlaceholder}' at row ${rowNumber}`,
        );
        cell.value = ""; // Clear the placeholder
      }
    });
  });

  if (tableStartRow) {
    tableData.forEach((rowData, rowIndex) => {
      const newRow = worksheet.insertRow(
        tableStartRow + rowIndex,
        Object.values(rowData),
      );
      newRow.eachCell((cell, colNumber) => {
        const templateCell = worksheet.getRow(tableStartRow).getCell(colNumber);
        cell.style = { ...templateCell.style };
      });
    });
  } else {
    console.error(`Placeholder '${tableStartPlaceholder}' not found.`);
  }
}

async function autoPopulateData(templateFile, data) {
  const arrayBuffer = await templateFile.arrayBuffer();
  const workbook = await readTemplate(arrayBuffer);
  const worksheet = workbook.getWorksheet(1);

  // Replace placeholders with simple data
  replacePlaceholders(worksheet, data);

  // Insert table data for applicants
  if (data.applicants) {
    insertTableData(worksheet, "{{applicant_row}}", data.applicants);
  }

  // Insert table data for coapplicants
  if (data.coapplicants) {
    insertTableData(worksheet, "{{coapplicant_row}}", data.coapplicants);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "merged.xlsx";
  link.click();
}

document.getElementById("mergeButton").addEventListener("click", async () => {
  const fileInput = document.getElementById("templateInput");
  const jsonDataInput = document.getElementById("jsonData").value;

  if (!fileInput.files.length) {
    alert("Please upload an Excel template.");
    return;
  }

  let jsonData;
  try {
    jsonData = JSON.parse(jsonDataInput);
  } catch (error) {
    alert("Invalid JSON data.");
    return;
  }

  const templateFile = fileInput.files[0];
  try {
    await autoPopulateData(templateFile, jsonData);
    alert("Data populated successfully");
  } catch (error) {
    console.error("Error populating data:", error);
    alert("Error populating data. Check the console for details.");
  }
});
