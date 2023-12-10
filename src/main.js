// Imports
const ExcelJS = require('exceljs');
const FileSaver = require('file-saver');

// Globals.
const validCustomerId = 'd1c5eb68-dc63-42a7-bc71-52e026e994d2';
const validApiKey = '688ec158-7550-42d6-abf0-3eceb2af664a';

const customerIdInput = document.getElementById('customerId');
const apiKeyInput = document.getElementById('apiKey');
const rememberMeInput = document.getElementById('rememberMe');

const downloadFileButton = document.getElementById('downloadFile');

const statusFeedbackBlock = document.getElementById('statusFeedback');
const statusMessagesSpan = document.getElementById('statusMessages');
const statusProgressBar = document.getElementById('statusProgress');

const errorMessageBox = document.getElementById('errorMessage');

const completionStates = {
    validatingCredentials: { message: 'Validating credentials.', value: 5, waitDuration: 1000 },
    fetchingJpiSettings: { message: 'Fetching JPI settings.', value: 10, waitDuration: 1000 },
    fetchingJpiMachines: { message: 'Fetching the machine names from JPI.', value: 25, waitDuration: 1000 },
    fetchingJpiJobs: { message: 'Fetching JPI job and task details.', value: 40, waitDuration: 1000 },
    fetchingJpiResources: { message: 'Fetching resource details and availability.', value: 65, waitDuration: 1000 },
    chartingTheSchedule: { message: 'Transforming the schedule to an Excel file.', value: 90, waitDuration: 1000 },
    downloadingTheExcelFile: { message: 'Downloading the Excel file.', value: 95, waitDuration: 1500 },
    allDone: { message: 'The schedule has been successfully exported.', value: 100, waitDuration: 1000 },
}

const jpiEndpointUrls = {
    settings: 'https://api.just-plan-it.com/v1/settings/',
    machines: 'https://api.just-plan-it.com/v1/resourcegroups/5345152d-48f0-47aa-b86a-300e0442da3a',
    jobs: 'https://api.just-plan-it.com/v1/jobs/',
    resources: 'https://api.just-plan-it.com/v1/resources/',
    // settings: './data/settings.json',
    // machines: './data/machines.json',
    // jobs: './data/jobs.json',
    // resources: './data/resources.json',
}

class JpiExporterError extends Error {
    constructor(message) {
        super(message);
        this.name = 'MyCustomError';
    }
}

function updateCredentials() {
    if (rememberMeInput.checked) {
        localStorage.setItem(customerIdInput.id, customerIdInput.value);
        localStorage.setItem(apiKeyInput.id, apiKeyInput.value);
        localStorage.setItem(rememberMeInput.id, rememberMeInput.checked);
    } else {
        localStorage.removeItem(customerIdInput.id, customerIdInput.value);
        localStorage.removeItem(apiKeyInput.id, apiKeyInput.value);
        localStorage.removeItem(rememberMeInput.id, rememberMeInput.checked);
    }
}

function prepareDocument() {
    // Restore credentials.
    customerIdInput.value = localStorage.getItem(customerIdInput.id) || '';
    apiKeyInput.value = localStorage.getItem(apiKeyInput.id) || '';
    rememberMeInput.checked = localStorage.getItem(rememberMeInput.id) || false;

    // Add event listeners.
    customerIdInput.addEventListener('change', updateCredentials);
    apiKeyInput.addEventListener('change', updateCredentials);
    rememberMeInput.addEventListener('change', updateCredentials);
    downloadFileButton.addEventListener('click', exportJpiSchedule);
}

function throwError(message) {
    throw new JpiExporterError(message);
}

function handleError(error) {
    errorMessageBox.hidden = false;
    if (error instanceof JpiExporterError) {
        errorMessageBox.innerText = error.message;
    } else {
        // TODO: Remove.
        // errorMessageBox.innerText = error.message
        errorMessageBox.innerText = 'An unexpected error occurred. Please try again.';
    }
}

function sleep(duration) {
    return new Promise(resolve => setTimeout(resolve, duration));
}

function roundToNearestHour(date) {
    date.setMilliseconds(0);
    date.setSeconds(0);
    date.setMinutes(Math.round(date.getMinutes() / 60) * 60);
    return date;
}

async function validateUserInputsAndGetApiKey() {
    // GUID Regex
    const guidRegex = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;

    // Provide Feedback
    await provideStatusFeedback(completionStates.validatingCredentials);

    // Customer ID
    const customerId = customerIdInput.value;
    if (!customerId || customerId === '') {
        throwError('Please provide a Customer ID.');
    }
    if (!guidRegex.test(customerId)) {
        throwError('Invalid format. Please provide a valid Customer ID.');
    }
    if (customerId !== validCustomerId) {
        throwError('Customer ID is not authorised for this application.');
    }

    // API Key
    const apiKey = apiKeyInput.value;
    if (!apiKey || apiKey === '') {
        throwError('Please provide an API Key.');
    }
    if (!guidRegex.test(apiKey)) {
        throwError('Invalid format. Please provide a valid API Key.');
    }
    if (apiKey !== validApiKey) {
        throwError('API Key is not authorised for this application.');
    }

    return apiKey;
}

async function provideStatusFeedback(statusFeedback) {
    statusFeedbackBlock.hidden = false;
    statusMessagesSpan.innerText = statusFeedback.message;
    statusProgressBar.value = statusFeedback.value;
    if (statusFeedback.value === 100) {
        // statusProgressBar.classList.remove('is-info');
        // statusProgressBar.classList.add('is-success');
        statusProgressBar.classList = ('progress is-success');
    } else {
        statusProgressBar.classList = 'progress is-info';
    }
    await sleep(statusFeedback.waitDuration); // Delay for UI updates.
}

function determineAndThrowApiError(response) {
    if (response.status === 401) {
        throwError('API key is not authorised to fetch information. Please verify information on JPI website.');
    } else {
        throwError('Unable to fetch information from JPI. Please try again.');
    }
}

async function fetchData(url, apiKey) {
    const headers = {
        'X-Api-Key': apiKey,
        'Content-Type': 'application/json'
    }
    const options = {
        method: 'GET',
        headers: headers
    }
    const response = await fetch(url, options);

    if (!response.ok) {
        determineAndThrowApiError(response)
    }

    const data = await response.json();

    return data;
}

async function fetchJpiData(apiKey) {
    const jpiData = {
        settings: {},
        machines: {},
        jobs: [],
        resources: [],
    }

    await provideStatusFeedback(completionStates.fetchingJpiSettings);
    jpiData.settings = await fetchData(jpiEndpointUrls.settings, apiKey);
    await provideStatusFeedback(completionStates.fetchingJpiMachines);
    jpiData.machines = await fetchData(jpiEndpointUrls.machines, apiKey);
    await provideStatusFeedback(completionStates.fetchingJpiJobs);
    jpiData.jobs = await fetchData(jpiEndpointUrls.jobs, apiKey);
    await provideStatusFeedback(completionStates.fetchingJpiResources);
    jpiData.resources = await fetchData(jpiEndpointUrls.resources, apiKey);

    return jpiData;
}

async function transformJpiData(jpiData) {
    const scheduleData = {}

    const jpiSettings = jpiData.settings;
    const jpiMachines = jpiData.machines;
    const jpiJobs = jpiData.jobs;
    const jpiResources = jpiData.resources;

    // Timeline
    let scheduleStart = new Date(jpiSettings.PlanningStart);
    let scheduleEnd = new Date(jpiSettings.PlanningStart);

    scheduleStart.setHours(0, 0, 0, 0);
    scheduleEnd.setHours(0, 0, 0, 0);

    scheduleStart.setDate(scheduleStart.getDate() - jpiSettings.DaysBeforePlanningStart);
    scheduleEnd.setDate(scheduleStart.getDate() + 7 * jpiSettings.PlanningHorizon + 1);

    const timeline = [];
    const timeOffsets = [];
    let currentDateTime = new Date(scheduleStart);
    while (currentDateTime < scheduleEnd) {
        timeline.push(new Date(currentDateTime));
        timeOffsets.push(currentDateTime - scheduleStart);
        currentDateTime.setHours(currentDateTime.getHours() + 1);
    }

    // Machines
    const machines = jpiMachines.Resources;
    const machineGuids = machines.map(element => element.Guid);
    const machineNames = machines.map(element => element.Name);

    // Tasks
    const tasks = jpiJobs
        .filter(job => ['Planned', 'Started'].includes(job.ExecuteStatus))
        .reduce((accumulator, job) => {
            const jobTasks = job.Tasks
                .filter(task => ['Planned', 'Started'].includes(task.TaskStatus))
                .map(task => {
                    task.job = job;
                    return task;
                });
            return accumulator.concat(jobTasks);
        }, []);

    // Cells
    const matrix = []
    for (let i = 0; i < timeline.length; i++) {
        const rowsArray = [];
        for (let j = 0; j < machines.length; j++) {
            rowsArray.push({});
        }
        matrix.push(rowsArray);
    }

    // Chart
    for (const task of tasks) {
        const taskName = task.Name || 'N/A';
        const taskCustomer = task.job.Customer || 'N/A';
        const taskLines = task.job.CustomFieldValue1 || 'N/A';

        const taskSetup = task.SetupTime > 0 || /SET\s?UP/i.test(task.Name);

        const start = roundToNearestHour(new Date(task.Start));
        const end = roundToNearestHour(new Date(task.End));

        if (start > scheduleEnd || end < scheduleStart) {
            continue;
        }

        const taskStart = Math.max(start, scheduleStart);
        const taskEnd = Math.min(end, scheduleEnd);

        const taskStartOffset = taskStart - scheduleStart;
        const taskEndOffset = taskEnd - scheduleStart;

        const rowStart = timeOffsets.indexOf(taskStartOffset);
        const rowEnd = timeOffsets.indexOf(taskEndOffset);

        if (rowStart < 0 || rowEnd < 0) {
            continue
        }

        for (const resource of task.AssignedResources) {
            const col = machineGuids.indexOf(resource.Guid);
            if (col < 0) {
                continue
            }
            if (rowStart === rowEnd) {
                const cellNote = `${taskName}\n${taskCustomer}\n${taskLines}`;
                matrix[rowStart][col].note = cellNote;
                continue;
            }
            for (let index = rowStart; index < rowEnd; index++) {
                let cellText = '';
                const fieldCount = index - rowStart;
                if (fieldCount === 0) {
                    cellText = taskName;
                } else if (fieldCount === 1) {
                    cellText = taskCustomer;
                } else if (fieldCount === 2 && taskSetup) {
                    cellText = taskLines;
                } else {
                    cellText = '---';
                }
                matrix[index][col].text = cellText;
                matrix[index][col].bold = taskSetup;
            }
        }
    }

    // Exceptions
    const resources = jpiResources
        .filter(resource => machineGuids.includes(resource.Guid))
        .map(resource => {
            const offDays = resource.CalendarExceptions
                .map(element => new Date(element.Date))
                .filter(offDay => offDay >= scheduleStart && offDay < scheduleEnd);
            resource.exceptions = offDays;
            return resource;
        });

    for (const resource of resources) {
        const col = machineGuids.indexOf(resource.Guid);
        for (const offDay of resource.exceptions) {
            const nextDay = new Date(offDay);
            nextDay.setDate(nextDay.getDate() + 1);

            const dayStartOffset = offDay - scheduleStart;
            const dayEndOffset = nextDay - scheduleStart;

            const rowStart = timeOffsets.indexOf(dayStartOffset);
            const rowEnd = timeOffsets.indexOf(dayEndOffset);

            if (rowStart < 0 || rowEnd < 0) {
                continue
            }

            for (let index = rowStart; index < rowEnd; index++) {
                matrix[index][col].color = true;
            }
        }
    }

    scheduleData.matrix = matrix;
    scheduleData.timeline = timeline;
    scheduleData.machineNames = machineNames;

    return scheduleData;
}

async function createScheduleWorkbook(scheduleData) {
    await provideStatusFeedback(completionStates.chartingTheSchedule);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet 1');

    const boldFont = { bold: true };
    const mediumBorder = { bottom: { style: 'medium' } };
    const fillColor = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'BFB1D1' } }

    worksheet.getRow(1).font = boldFont;
    worksheet.getRow(1).border = mediumBorder;
    worksheet.getColumn(1).font = boldFont;
    worksheet.getColumn(2).font = boldFont;
    worksheet.getColumn(3).font = boldFont;

    const columnHeaders = ['Date', 'Day', 'Time'].concat(scheduleData.machineNames);
    columnHeaders.forEach((machineName, index) => {
        const col = index + 1;
        worksheet.getCell(1, col).value = machineName;
    });

    const rowHeaders = scheduleData.timeline;
    rowHeaders.forEach((currentTime, index) => {
        const row = index + 2;
        const dailyIndex = index % 24;
        if (dailyIndex === 6) {
            worksheet.getRow(row).border = mediumBorder;
        }
        if (dailyIndex === 7) {
            const currentDate = currentTime.toLocaleDateString('en-US', { month: 'short', day: '2-digit' });
            worksheet.getCell(row, 1).value = currentDate;
            const currentDay = currentTime.toLocaleDateString('en-US', { weekday: 'short' });
            worksheet.getCell(row, 2).value = currentDay;
        }
        const currentHour = currentTime.getHours();
        const formattedHour = currentHour < 10 ? `0${currentHour}` : `${currentHour}`;
        worksheet.getCell(row, 3).value = formattedHour;
    });

    const cells = scheduleData.matrix;
    cells.forEach((rowData, rowIndex) => {
        const row = rowIndex + 2;
        rowData.forEach((cell, index) => {
            const col = index + 4;
            if (cell.text) {
                worksheet.getCell(row, col).value = cell.text;
            }
            if (cell.bold) {
                worksheet.getCell(row, col).font = boldFont;
            }
            if (cell.note) {
                worksheet.getCell(row, col).note = cell.note;
                worksheet.getCell(row, col).fill = fillColor;
            }
            if (cell.color) {
                worksheet.getCell(row, col).fill = fillColor;
            }
        })
    });

    // Freeze one row and three columns.
    worksheet.views = [{ state: 'frozen', xSplit: 3, ySplit: 1 }];

    // Auto-size all columns
    const minimalWidth = 5;
    worksheet.columns.forEach((column) => {
        let maxColumnLength = 0;
        if (column && typeof column.eachCell === 'function') {
            column.eachCell({ includeEmpty: true }, (cell) => {
                maxColumnLength = Math.max(
                    maxColumnLength,
                    minimalWidth,
                    cell.value ? cell.value.toString().length : 0
                );
            });
            column.width = maxColumnLength + 2;
        }
    });

    return workbook;
}

async function sendWorkbookAsExcel(scheduleWorkbook) {
    await provideStatusFeedback(completionStates.downloadingTheExcelFile);

    scheduleWorkbook.xlsx.writeBuffer().then(function (buffer) {
        // Create a Blob from the buffer
        var blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        FileSaver.saveAs(blob, "JPI-Schedule.xlsx");
    });
}

async function exportJpiSchedule(event) {
    try {
        event.preventDefault();
        const apiKey = await validateUserInputsAndGetApiKey();
        const jpiData = await fetchJpiData(apiKey);
        const scheduleData = await transformJpiData(jpiData);
        const scheduleWorkbook = await createScheduleWorkbook(scheduleData);
        await sendWorkbookAsExcel(scheduleWorkbook)
        await provideStatusFeedback(completionStates.allDone);
    } catch (error) {
        handleError(error);
    }
}

prepareDocument();
