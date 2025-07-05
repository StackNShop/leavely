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

                if (json.length < 2) return alert("Empty or invalid file.");

                fullData = json;
                const headers = json[0];
                const rows = json.slice(1);

                let html = "<table class='table table-bordered table-hover table-sm'><thead><tr>";
                headers.forEach((header) => (html += `<th>${header}</th>`));
                html += "</tr></thead><tbody>";

                rows.forEach((row) => {
                    html += "<tr>";
                    headers.forEach((_, i) => {
                        html += `<td>${row[i] || ""}</td>`;
                    });
                    html += "</tr>";
                });

                html += "</tbody></table>";

                document.getElementById("tableContent").innerHTML = html;
                document.getElementById("uploadSection").classList.add("d-none");
                document.getElementById("tableBox").classList.remove("d-none");

                const actionBtn = document.getElementById("actionBtn");
                actionBtn.innerHTML = `<i class="fas fa-cogs fa-spin text-light"></i> Perform Operation`;
                actionBtn.className = "btn btn-info btn-custom text-white";
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
            const empCodeIndex = headers.findIndex((h) => h.toString().toLowerCase().includes("emp"));
            if (empCodeIndex === -1) return alert("Emp Code column not found");

            const result = [["Emp Code", "Leave Type", "Start Date", "End Date"]];

            headers.forEach((header, i) => {
                const day = parseInt(header.toString().trim());
                if (!isNaN(day) && day >= 1 && day <= 31) {
                    const dateStr = `${day}-${month}-${year}`;
                    rows.forEach((row) => {
                        const empCode = row[empCodeIndex];
                        const leaveVal = (row[i] || "").toString().trim();
                        if (leaveVal) {
                            result.push([empCode, leaveVal, dateStr, dateStr]);
                        }
                    });
                }
            });

            window.originalLeaveResult = result;

            let html = `
        <div class="d-flex justify-content-between align-items-center mb-2">
          <h5>All Leave Entries (Unfiltered)</h5>
          <div class="dropdown">
            <button class="btn btn-outline-dark btn-sm dropdown-toggle" type="button" data-bs-toggle="dropdown" aria-expanded="false">
              <i class="fas fa-filter"></i> Filter
            </button>
            <div class="dropdown-menu p-3" style="width: 260px;">
              <div id="leaveFilterList" class="mb-2">            
                ${["DP", "WO", "WOP", "CL", "ABS", "ABS/DP", "CL/DP", "DP/ABS"]
                    .map(
                        (type) => `
                  <div class="form-check">
                    <input class="form-check-input leave-filter-checkbox" type="checkbox" value="${type}" id="filter_${type}">
                    <label class="form-check-label" for="filter_${type}">${type}</label>
                  </div>`
                    )
                    .join("")}
              </div>
              <button class="btn btn-sm btn-primary w-100" onclick="applyLeaveFilter()">Apply Filters</button>
            </div>
          </div>
        </div>
        <div id="leaveTableBox"></div>
      `;

            document.getElementById("resultContent").innerHTML = html;
            document.getElementById("resultContent").classList.remove("d-none");
            document.getElementById("tableContent").classList.add("d-none");

            applyLeaveFilter();

            const actionBtn = document.getElementById("actionBtn");
            actionBtn.innerHTML = `<i class="fas fa-download"></i> Download File`;
            actionBtn.className = "btn btn-success btn-custom";
            actionBtn.onclick = downloadFiltered;
        }

        function applyLeaveFilter() {
            const checkboxes = document.querySelectorAll(".leave-filter-checkbox");
            const included = new Set();
            checkboxes.forEach((cb) => {
                if (cb.checked) included.add(cb.value.trim().toUpperCase());
            });

            const allResults = window.originalLeaveResult || [];
            const headers = allResults[0];
            const dataRows = allResults.slice(1);
            const filtered = [headers];

            if (included.size === 0) {
                filtered.push(...dataRows);
            } else {
                dataRows.forEach((row) => {
                    const leaveType = row[1].toString().trim().toUpperCase();
                    if (included.has(leaveType)) {
                        filtered.push(row);
                    }
                });
            }

            lastFilteredResult = filtered;

            let html = "";
            if (filtered.length === 1) {
                html = "<div class='alert alert-warning'>No matching leave types selected.</div>";
            } else {
                html += "<table class='table table-bordered table-striped table-hover table-sm'><thead><tr>";
                headers.forEach((h) => (html += `<th>${h}</th>`));
                html += "</tr></thead><tbody>";
                filtered.slice(1).forEach((r) => {
                    html += "<tr>";
                    r.forEach((c) => (html += `<td>${c}</td>`));
                    html += "</tr>";
                });
                html += "</tbody></table>";
            }

            document.getElementById("leaveTableBox").innerHTML = html;
        }

        function downloadFiltered() {
            if (!lastFilteredResult || lastFilteredResult.length <= 1) return;
            const ws = XLSX.utils.aoa_to_sheet(lastFilteredResult);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Filtered Leaves");
            XLSX.writeFile(wb, "Filtered_Leave_Report.xlsx");
        }
        // Set current month and year on page load
        window.addEventListener("DOMContentLoaded", () => {
            const now = new Date();
            document.getElementById("month").value = now.getMonth() + 1;
            document.getElementById("year").value = now.getFullYear();
        });
