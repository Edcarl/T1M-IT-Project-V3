document.getElementById('fileInput').addEventListener('change', function(event) {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        const employeeIdFilter = document.getElementById('employeeIdFilter');
        const statusFilter = document.getElementById('statusFilter');
        employeeIdFilter.innerHTML = '<option value="">All ID</option>';
        statusFilter.innerHTML = '<option value="">All Status</option>';
        const tableBody = document.querySelector('#outputTable tbody');
        tableBody.innerHTML = '';
        const employeeData = {};

        rows.forEach((row, index) => {
            if (index > 0) {
                const employeeId = row[0];
                const excelDate = row[1];
                const jsDate = getJsDateFromExcel(excelDate);
                const dateString = jsDate.toLocaleDateString('en-GB'); // Format as dd/mm/yyyy

                if (!employeeData[employeeId]) {
                    employeeData[employeeId] = {};
                }
                if (!employeeData[employeeId][dateString]) {
                    employeeData[employeeId][dateString] = [];
                }
                employeeData[employeeId][dateString].push(jsDate);
                
                if (!employeeIdFilter.querySelector(`option[value="${employeeId}"]`)) {
                    const option = document.createElement('option');
                    option.value = employeeId;
                    option.textContent = employeeId;
                    employeeIdFilter.appendChild(option);
                }
            }
        });

        for (const employeeId in employeeData) {
            for (const date in employeeData[employeeId]) {
                const timestamps = employeeData[employeeId][date];
                timestamps.sort((a, b) => a - b); // Sort timestamps

                const timeIn = timestamps[0];
                    let timeOut = timestamps[timestamps.length - 1];

                    const tr = document.createElement('tr');
                    const employeeIdCell = document.createElement('td');
                    const dateCell = document.createElement('td');
                    const timeInCell = document.createElement('td');
                    const timeOutCell = document.createElement('td');
                    const lateCell = document.createElement('td');
                    const totalHoursCell = document.createElement('td');
                    const statusCell = document.createElement('td');

                    employeeIdCell.textContent = employeeId;
                    dateCell.textContent = date;
                    timeInCell.textContent = timeIn.toLocaleTimeString('en-US', { hour12: true });

                    const timeOutInput = document.createElement('input');
                    timeOutInput.type = 'time';
                    timeOutInput.value = timestamps.length > 1 ? timeOut.toLocaleTimeString('en-US', { hour12: false }) : '';
                    timeOutCell.appendChild(timeOutInput);

                    const allowedStatuses = ["Regular time", "Under time", "Over time", "Didn't clock out", "Late"];
                    const scheduledTimeIn = new Date(timeIn);
                    scheduledTimeIn.setHours(8, 30, 0, 0);

                    let lateness = 0;
                    let status;
                    if (timeIn > scheduledTimeIn) {
                        lateness = (timeIn - scheduledTimeIn) / (1000 * 60); // in minutes
                    }

                    const hoursLate = Math.floor(lateness / 60);
                    const minutesLate = Math.floor(lateness % 60);

                    if (hoursLate === 0 && minutesLate === 0) {
                        lateCell.textContent = 'NA';
                    } else if (hoursLate === 0) {
                        lateCell.textContent = `${minutesLate} min/s late`;
                        status = "Late";
                    }
                    else {
                        lateCell.textContent = `${hoursLate}hr/s and ${minutesLate}min/s`;
                        status = "Late";
                    }

                    const updateTotalHours = () => {
                        const timeOutValue = timeOutInput.value;
                        if (timeOutValue) {
                            const [hours, minutes] = timeOutValue.split(':');
                            timeOut.setHours(hours, minutes);
                    
                            const totalHours = ((timeOut - timeIn) / (1000 * 60 * 60)).toFixed(2); // Calculate total hours
                            const totalMins = ((timeOut - timeIn) / (1000 * 60)).toFixed(2);
                            totalHoursCell.textContent = totalHours;
                    
                            if (totalHours-1 < 7.5 && totalHours-1 > 0) {
                                let deficit = Math.round(480 - totalMins + 60);
                                status = `Under time: ${deficit} mins`;
                            } else if (totalHours-1 > 8.5) {
                                let OT = Math.round(totalHours-1 - 8);
                                status = `Over time: ${OT} hour/s`;
                            } else if (totalHours <= 0) {
                                status = "Didn't clock out";
                            } else {
                                status = "Regular time";
                            }
                            statusCell.textContent = status;
                    
                            const filteredStatus = allowedStatuses.find(allowedStatus => status.includes(allowedStatus));
                            if (filteredStatus && !statusFilter.querySelector(`option[value="${filteredStatus}"]`)) {
                                const option = document.createElement('option');
                                option.value = filteredStatus;
                                option.textContent = filteredStatus;
                                statusFilter.appendChild(option);
                            }
                        } else {
                            totalHoursCell.textContent = "N/A";
                            statusCell.textContent = "Didn't clock out";
                        }
                };

                timeOutInput.addEventListener('change', updateTotalHours);
                updateTotalHours();

                tr.appendChild(employeeIdCell);
                tr.appendChild(dateCell);
                tr.appendChild(timeInCell);
                tr.appendChild(timeOutCell);
                tr.appendChild(lateCell);
                tr.appendChild(totalHoursCell);
                tr.appendChild(statusCell);

                tableBody.appendChild(tr);
            }
        }
    };
    reader.readAsArrayBuffer(file);
});

function filterTable() {
    const selectedEmployeeId = document.getElementById('employeeIdFilter').value;
    const selectedStatus = document.getElementById('statusFilter').value;
    const selectedDate = document.getElementById('dateFilter').value;
    const table = document.getElementById('outputTable');
    const rows = table.getElementsByTagName('tr');

    for (let i = 1; i < rows.length; i++) {
        const cells = rows[i].getElementsByTagName('td');
        const employeeId = cells[0].textContent;
        const status = cells[5].textContent;
        const filterDate = cells[1].textContent;

        if ((selectedEmployeeId === "" || employeeId === selectedEmployeeId) &&
            (selectedStatus === "" || status.includes(selectedStatus)) && 
            (selectedDate === "" || filterDate === new Date(selectedDate).toLocaleDateString('en-GB'))) {
            rows[i].style.display = "";
        } else {
            rows[i].style.display = "none";
        }
    }
}

document.getElementById('searchButton').addEventListener('click', function() {
    const employeeId = document.getElementById('searchEmployeeId').value.toLowerCase();
    const startDate = new Date(document.getElementById('searchStartDate').value.split('/').reverse().join('-'));
    const endDate = new Date(document.getElementById('searchEndDate').value.split('/').reverse().join('-'));
    const searchTimeIn = document.getElementById('timeInSearch').value.trim();
    const searchTimeInRange = document.getElementById('timeInRange').value.trim();
    const employeeTableBody = document.querySelector('#employeeTable tbody');
    employeeTableBody.innerHTML = '';
    const tableBody = document.querySelector('#outputTable tbody');
    const rows = tableBody.getElementsByTagName('tr');

    for (let i = 0; i < rows.length; i++) {
        const cells = rows[i].getElementsByTagName('td');
        const rowEmployeeId = cells[0].textContent.toLowerCase();
        const rowDate = new Date(cells[1].textContent.split('/').reverse().join('-'));

        const timeOutInput = cells[3].querySelector('input');
        const timeOutValue = timeOutInput ? timeOutInput.value : '';
        const timeInText = cells[2].textContent;

        // Parse timeInText to a Date object
        const timeInParts = timeInText.split(/[: ]/);
        const timeInHours = parseInt(timeInParts[0], 10) + (timeInParts[2] === 'PM' && timeInParts[0] !== '12' ? 12 : 0);
        const timeInMinutes = parseInt(timeInParts[1], 10);
        const rowTime = new Date(1970, 0, 1, timeInHours, timeInMinutes);

        // rowDate format (dd/mm/yyyy)
        const formattedRowDate = rowDate.toLocaleDateString('en-US');

        // Parse search times
        const [searchStartHours, searchStartMinutes = 0] = searchTimeIn.split(':').map(Number); 
        const [searchEndHours, searchEndMinutes = 0] = searchTimeInRange.split(':').map(Number); 

        const searchStartTime = new Date(1970, 0, 1, searchStartHours, searchStartMinutes);
        const searchEndTime = new Date(1970, 0, 1, searchEndHours, searchEndMinutes);

        let matchesTime = false;
        if (searchTimeIn && searchTimeInRange) {
            matchesTime = rowTime >= searchStartTime && rowTime <= searchEndTime;
        } else {
            matchesTime = true;
        }

        if ((employeeId === '' || rowEmployeeId.includes(employeeId)) &&
            (isNaN(startDate) || isNaN(endDate) || (rowDate >= startDate && rowDate <= endDate)) &&
            (searchTimeIn === '' || matchesTime)) {
            const tr = document.createElement('tr');
            for (let j = 0; j < cells.length; j++) {
                const td = document.createElement('td');
                
                if (j === 1) {
                    td.textContent = formattedRowDate;
                } else if (j === 2 && timeInText) {
                    td.textContent = timeInText;
                } else if (j === 3 && timeOutInput) {
                    // Convert timeOutInput value to text and format without seconds
                    const timeOutParts = timeOutValue.split(':');
                    const formattedTimeOut = `${timeOutParts[0]}:${timeOutParts[1]}`;
                    td.textContent = formattedTimeOut;
                } else if (cells[j].querySelector('input')) {
                    const input = cells[j].querySelector('input');
                    td.appendChild(input.cloneNode(true));
                } else {
                    td.textContent = cells[j].textContent;
                }
                tr.appendChild(td);
            }
            employeeTableBody.appendChild(tr);
        }
    }
});


document.getElementById('clearButton').addEventListener('click', function() {
    const employeeTableBody = document.getElementById('employeeTable').getElementsByTagName('tbody')[0];
    const employeeId = document.getElementById('searchEmployeeId').value = '';
    const startDate = document.getElementById('searchStartDate').value = '';
    const endDate = document.getElementById('searchEndDate').value = '';
    const searchTimeIn = document.getElementById('timeInSearch').value = '';
    const searchTimeInRange = document.getElementById('timeInRange').value = '';
    employeeTableBody.innerHTML = '';
});

// to convert Excel date to JavaScript date
function getJsDateFromExcel(excelDate) {
    const excelEpoch = new Date(1899, 11, 30); 
    const msPerDay = 86400000; 
    const jsDate = new Date(excelEpoch.getTime() + excelDate * msPerDay);
    return jsDate;
}   

document.getElementById('btn-export').addEventListener('click', function() {
    const table = document.getElementById('employeeTable');
    const workbook = XLSX.utils.table_to_book(table, { sheet: 'Sheet1' });
    const fileInput = document.getElementById('fileInput');

    // Get today's date
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, '0');
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const yyyy = today.getFullYear();

    const todayDate = yyyy + '-' + mm + '-' + dd;

    // Generate filename with today's date
    const filename = `EmployeeAttendance_${todayDate}.xlsx`;

    if (!fileInput.files.length) {
        alert('Please select a file before exporting.');
        return;
    }

    XLSX.writeFile(workbook, filename);
});

document.getElementById('btn-reload').addEventListener('click', function() { 
    location.reload();
});