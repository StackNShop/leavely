let uploadedFile = null;
let fullData = [];
let lastFilteredResult = [];

document.getElementById("uploadExcel").addEventListener("change", function (e) {
  uploadedFile = e.target.files[0];
});

document.getElementById("showDataBtn").addEventListener("click", function () {
  if (!uploadedFile) return alert("Please choose a file first.");

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "", header: 1 });

    if (json.length < 2) return;

    fullData = json;

    const headers = json[0];
    const rows = json.slice(1);

    let html =
      "<table class='table table-bordered table-hover table-sm'><thead><tr>";
    headers.forEach((header) => (html += `<th>${header}</th>`));
    html += "</tr></thead><tbody>";

    rows.forEach((row) => {
      html += "<tr>";
      headers.forEach((_, i) => (html += `<td>${row[i] || ""}</td>`));
      html += "</tr>";
    });

    html += "</tbody></table>";

    document.getElementById("tableContent").innerHTML = html;
    document.getElementById("uploadSection").classList.add("d-none");
    document.getElementById("tableBox").classList.remove("d-none");

    const actionBtn = document.getElementById("actionBtn");
    actionBtn.textContent = "Perform Operation";
    actionBtn.className = "btn btn-primary btn-custom";
    actionBtn.onclick = performOperation;
  };
  reader.readAsArrayBuffer(uploadedFile);
});

function goBack() {
  document.getElementById("uploadSection").classList.remove("d-none");
  document.getElementById("tableBox").classList.add("d-none");
  document.getElementById("uploadExcel").value = "";
  uploadedFile = null;
  fullData = [];
  lastFilteredResult = [];
  document.getElementById("resultContent").classList.add("d-none");
  document.getElementById("tableContent").classList.remove("d-none");

  const actionBtn = document.getElementById("actionBtn");
  actionBtn.textContent = "Perform Operation";
  actionBtn.className = "btn btn-primary btn-custom";
  actionBtn.onclick = performOperation;
}

function performOperation() {
  if (!fullData || fullData.length < 2) return;

  const month = parseInt(document.getElementById("month").value);
  const year = parseInt(document.getElementById("year").value);
  if (!month || !year) return alert("Enter valid month and year");

  const headers = fullData[0];
  const rows = fullData.slice(1);
  const empCodeIndex = headers.findIndex((h) =>
    h.toString().toLowerCase().includes("emp")
  );
  if (empCodeIndex === -1) return alert("Emp Code column not found");

  const result = [["Emp Code", "Leave Type", "Start Date", "End Date"]];
  let saturdayCounter = 0;

  headers.forEach((header, i) => {
    const day = parseInt(header.toString().trim());
    if (!isNaN(day) && day >= 1 && day <= 31) {
      const dateObj = new Date(year, month - 1, day);
      const dayOfWeek = dateObj.getDay();

      let skip = false;
      if (dayOfWeek === 0) skip = true;
      else if (dayOfWeek === 6) {
        saturdayCounter++;
        if (saturdayCounter === 2 || saturdayCounter === 3) skip = true;
      }

      if (skip) return;

      rows.forEach((row) => {
        const empCode = row[empCodeIndex];
        const leaveVal = (row[i] || "").toString().trim().toUpperCase();
        if (leaveVal && leaveVal !== "DP" && leaveVal !== "WO") {
          result.push([
            empCode,
            leaveVal,
            `${day}-${month}-${year}`,
            `${day}-${month}-${year}`,
          ]);
        }
      });
    }
  });

  lastFilteredResult = result;

  let html = "";
  if (result.length === 1) {
    html =
      "<div class='alert alert-warning'>No valid leave entries found.</div>";
  } else {
    html +=
      "<h5>Filtered Leave Data</h5><table class='table table-bordered table-striped table-hover table-sm'><thead><tr>";
    result[0].forEach((h) => (html += `<th>${h}</th>`));
    html += "</tr></thead><tbody>";
    result.slice(1).forEach((r) => {
      html += "<tr>";
      r.forEach((c) => (html += `<td>${c}</td>`));
      html += "</tr>";
    });
    html += "</tbody></table>";
  }

  document.getElementById("resultContent").innerHTML = html;
  document.getElementById("resultContent").classList.remove("d-none");
  document.getElementById("tableContent").classList.add("d-none");

  const actionBtn = document.getElementById("actionBtn");
  actionBtn.innerHTML = `<i class="fas fa-download"></i> Download File`;
  actionBtn.className = "btn btn-success btn-custom";
  actionBtn.onclick = downloadFiltered;
}

function downloadFiltered() {
  if (!lastFilteredResult || lastFilteredResult.length <= 1) return;
  const ws = XLSX.utils.aoa_to_sheet(lastFilteredResult);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Filtered Leaves");
  XLSX.writeFile(wb, "Filtered_Leave_Report.xlsx");
}
